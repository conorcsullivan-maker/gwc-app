"""GWC weekly automation — runs every Tuesday 8:00 AM via launchd.

What it does, idempotently (safe to run any time):
1. Finds the current/upcoming NFL week from ESPN.
2. If that week isn't in the database yet: pulls the slate, randomizes
   assignments with the fairness rules, and marks it announced. The group-chat
   announcement text is written to gwc-app/announcements/.
3. Grades any finished games in current-season weeks.
4. Syncs graded picks into the official workbook and refreshes the
   dashboard's live-week feed (skipped gracefully if the workbook is open).
5. Posts a macOS notification with what happened.
"""
from __future__ import annotations
import shutil
import sys
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from subprocess import run
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parent))

from sqlalchemy import and_, select

from gwc import assign, export, rules, schedule
from gwc.db import (CURRENT_SEASON, SEASON_YEAR, assignments, engine, games,
                    players, rows, weeks)

ET = ZoneInfo("America/New_York")
APP_DIR = Path(__file__).resolve().parent
WORKBOOK = (APP_DIR.parent / "2025:2026" /
            "Gentlemens Wagering Cooperative.xlsx")


def log(msg):
    stamp = datetime.now(ET).strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{stamp}] {msg}", flush=True)


def notify(text):
    try:
        run(["osascript", "-e",
             f'display notification "{text}" with title "🎩 GWC Commissioner"'],
            timeout=10)
    except Exception:
        pass


def find_target_week(now):
    """Smallest regular-season week whose slate isn't fully in the past."""
    import time
    for w in range(1, 19):
        glist = None
        for attempt in range(3):     # network may not be up right after wake
            try:
                glist = schedule.fetch_week(SEASON_YEAR, w)
                break
            except Exception as e:
                log(f"ESPN fetch week {w} failed (try {attempt + 1}/3): {e}")
                time.sleep(30)
        if glist is None:
            log("giving up on ESPN this run — will try again next time")
            return None, None
        if not glist:
            continue
        last_ko = max(datetime.fromisoformat(g["kickoff_et"]) for g in glist)
        if last_ko > now - timedelta(hours=36):
            return w, glist
    return None, None


def restore_workbook_if_missing():
    """The workbook has gone missing twice (moved while being shared?).
    If it's gone, quietly restore the newest backup so automation never
    stalls — the backups are complete copies."""
    if WORKBOOK.exists():
        return
    backups = sorted(WORKBOOK.parent.glob("*backup*.xlsx"))
    if backups:
        shutil.copy2(backups[-1], WORKBOOK)
        log(f"workbook was missing — restored from {backups[-1].name}")
        notify("Workbook was missing; restored it from the latest backup.")


def ensure_week(week_num, fetched):
    """Create + randomize the week if it doesn't exist. Returns week row."""
    with engine().begin() as conn:
        existing = rows(conn, select(weeks).where(and_(
            weeks.c.season == CURRENT_SEASON, weeks.c.week_num == week_num)))
        if existing:
            return existing[0], False
        wid = conn.execute(weeks.insert().values(
            season=CURRENT_SEASON, week_num=week_num, status="announced",
            deadline_et=schedule.default_deadline(fetched),
        )).inserted_primary_key[0]
        for g in fetched:
            g = dict(g)
            g.pop("final", None)
            conn.execute(games.insert().values(week_id=wid, **g))

    with engine().connect() as conn:
        glist = rows(conn, select(games).where(and_(
            games.c.week_id == wid, games.c.excluded == False)))  # noqa: E712
        hist = rows(conn, select(weeks.c.week_num,
                                 players.c.name.label("player"),
                                 games.c.day, games.c.away, games.c.home)
                    .select_from(assignments
                                 .join(players, assignments.c.player_id == players.c.id)
                                 .join(games, assignments.c.game_id == games.c.id)
                                 .join(weeks, assignments.c.week_id == weeks.c.id))
                    .where(weeks.c.week_num < week_num))
        active = rows(conn, select(players).where(players.c.active == True))  # noqa: E712

    result = assign.randomize(glist, [p["name"] for p in active],
                              assign.build_history(hist))
    ids = {p["name"]: p["id"] for p in active}
    with engine().begin() as conn:
        for gid, name in result.items():
            conn.execute(assignments.insert().values(
                week_id=wid, game_id=gid, player_id=ids[name]))
        week = rows(conn, select(weeks).where(weeks.c.id == wid))[0]
    return week, True


def write_announcement(week):
    with engine().connect() as conn:
        alist = rows(conn, select(
            players.c.name.label("player"), games.c.away, games.c.home,
            games.c.kickoff_et, games.c.favorite, games.c.spread,
            games.c.ou_total)
            .select_from(assignments
                         .join(players, assignments.c.player_id == players.c.id)
                         .join(games, assignments.c.game_id == games.c.id))
            .where(assignments.c.week_id == week["id"])
            .order_by(players.c.name, games.c.kickoff_et))
    lines = [f"🎩🏈 GWC WEEK {week['week_num']} ASSIGNMENTS 🏈🎩", ""]
    current = None
    for a in alist:
        if a["player"] != current:
            current = a["player"]
            lines.append(f"{current}:")
        ko = datetime.fromisoformat(a["kickoff_et"])
        bits = []
        if a["favorite"] and a["spread"]:
            bits.append(f"{a['favorite']} -{a['spread']:g}")
        if a["ou_total"]:
            bits.append(f"O/U {a['ou_total']:g}")
        line = " · ".join(bits) or "no line yet"
        lines.append(f"  • {a['away']} @ {a['home']} "
                     f"({ko.strftime('%a %-m/%-d %-I:%M %p')} ET) — {line}")
    dl = week["deadline_et"]
    if dl:
        dl = datetime.fromisoformat(dl).strftime("%A %-m/%-d %-I:%M %p ET")
    lines += ["", f"Picks due {dl}. Spreads & totals; ML only if the spread "
                  "is under 3. Lines are ESPN's — send your book's line with "
                  "your pick."]
    out = APP_DIR / "announcements"
    out.mkdir(exist_ok=True)
    f = out / f"week-{week['week_num']:02d}.txt"
    f.write_text("\n".join(lines))
    return f


def grade_pending():
    graded = 0
    with engine().connect() as conn:
        wks = rows(conn, select(weeks).where(weeks.c.season == CURRENT_SEASON))
    for week in wks:
        with engine().connect() as conn:
            pending = rows(conn, select(assignments).where(and_(
                assignments.c.week_id == week["id"],
                assignments.c.pick_selection.isnot(None),
                assignments.c.result.is_(None))))
        if not pending:
            continue
        try:
            live = schedule.fetch_week_live(SEASON_YEAR, week["week_num"])
        except Exception:
            continue
        with engine().begin() as conn:
            glist = rows(conn, select(games).where(
                games.c.week_id == week["id"]))
            for g in glist:
                f = live.get(g["espn_id"])
                if f and f["final"]:
                    conn.execute(games.update().where(games.c.id == g["id"])
                                 .values(final=True,
                                         home_score=f["home_score"],
                                         away_score=f["away_score"]))
                    g.update(final=True, home_score=f["home_score"],
                             away_score=f["away_score"])
                for a in [p for p in pending if p["game_id"] == g["id"]]:
                    res = rules.grade(g, a)
                    if res:
                        conn.execute(assignments.update()
                                     .where(assignments.c.id == a["id"])
                                     .values(result=res))
                        graded += 1
    return graded


def main():
    now = datetime.now(ET)
    log("weekly update starting")
    restore_workbook_if_missing()

    week_num, fetched = find_target_week(now)
    created = False
    if week_num:
        week, created = ensure_week(week_num, fetched)
        if created:
            f = write_announcement(week)
            log(f"week {week_num}: slate fetched, assignments randomized, "
                f"announcement -> {f.name}")
        else:
            log(f"week {week_num}: already set up, leaving assignments alone")
    else:
        log("no upcoming NFL week found (offseason?)")

    graded = grade_pending()
    if graded:
        log(f"graded {graded} finished pick(s)")

    if WORKBOOK.exists():
        try:
            added, backup = export.sync_workbook(str(WORKBOOK))
            n = export.write_live_feed(str(WORKBOOK))
            log(f"workbook: {added} graded pick(s) synced, live feed "
                f"refreshed ({n} rows)")
        except PermissionError:
            log("workbook busy (open in Excel?) — will sync next run")
        except Exception as e:
            log(f"workbook sync failed: {e}")
    else:
        log(f"workbook not found at {WORKBOOK}")

    if created:
        notify(f"Week {week_num} randomized — announcement ready to copy "
               "from gwc-app/announcements ✅")
    elif graded:
        notify(f"Graded {graded} picks and refreshed the dashboard.")
    log("weekly update done")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log("FAILED:\n" + traceback.format_exc())
        notify("Weekly update hit an error — check weekly_update.log")
        sys.exit(1)
