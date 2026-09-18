"""Closing line value (CLV) tracking — did the market move for or against
each pick between lock-in and kickoff, and how did those picks do?

CLV is stored per pick (assignments.clv) once it can be computed: the
closing line is the last tracked snapshot at or before kickoff, and the
edge is measured from our side of the bet, so positive always means we
hold a better number than the market closed at.
"""
from __future__ import annotations
from collections import defaultdict
from datetime import datetime

from sqlalchemy import and_, select

from .db import assignments, games, odds_snapshots, players, rows, weeks

GOOD, BAD = 0.5, -0.5     # thresholds in points


def market_from_our_side(a: dict, snap: dict) -> float | None:
    """The snapshot's number expressed for our side of the bet."""
    period_half = a.get("period") == "1H"
    if a["pick_type"] == "Over/Under":
        t = snap.get("ou_total")
        if t is None:
            return None
        return t / 2 if period_half else t
    sp = snap.get("spread")
    if sp is None:
        return None
    if period_half:
        sp = sp / 2
    if snap.get("favorite") is None:
        return 0.0
    return -sp if a.get("pick_team") == snap["favorite"] else sp


def edge(a: dict, closing: dict, opening: dict | None = None) -> float | None:
    """Points of edge we hold vs. the closing line (positive = good)."""
    if not a.get("pick_selection"):
        return None
    mkt = market_from_our_side(a, closing)
    if mkt is None:
        return None
    if a["pick_type"] == "Over/Under":
        over = str(a["pick_selection"]).startswith("Over")
        return (mkt - a["pick_line"]) if over else (a["pick_line"] - mkt)
    if a["pick_type"] == "Moneyline":
        if opening is None:
            return None
        o = market_from_our_side(a, opening)
        return None if o is None else o - mkt     # spread drifted toward us?
    return a["pick_line"] - mkt


def backfill(engine) -> int:
    """Compute and store CLV for every graded pick that lacks one."""
    with engine.connect() as conn:
        todo = rows(conn, select(
            assignments.c.id, assignments.c.pick_type, assignments.c.period,
            assignments.c.pick_selection, assignments.c.pick_team,
            assignments.c.pick_line, games.c.espn_id, games.c.kickoff_et,
            weeks.c.week_num)
            .join(games, assignments.c.game_id == games.c.id)
            .join(weeks, assignments.c.week_id == weeks.c.id)
            .where(and_(assignments.c.result.isnot(None),
                        assignments.c.clv.is_(None),
                        assignments.c.pick_selection.isnot(None))))
        if not todo:
            return 0
        ids = {t["espn_id"] for t in todo}
        snaps = rows(conn, select(odds_snapshots)
                     .where(odds_snapshots.c.espn_id.in_(ids))
                     .order_by(odds_snapshots.c.ts))
    by_game = defaultdict(list)
    for s in snaps:
        by_game[s["espn_id"]].append(s)
    updates = []
    for a in todo:
        hist = by_game.get(a["espn_id"])
        if not hist:
            continue
        ko = a["kickoff_et"]
        pre = [s for s in hist if s["ts"] <= ko] if ko else hist
        closing = (pre or hist)[-1]
        e = edge(a, closing, hist[0])
        if e is not None:
            updates.append((a["id"], round(e, 2)))
    if updates:
        with engine.begin() as conn:
            for aid, e in updates:
                conn.execute(assignments.update()
                             .where(assignments.c.id == aid).values(clv=e))
    return len(updates)


def _rec(picks):
    w = sum(1 for p in picks if p["result"] == "Win")
    l = sum(1 for p in picks if p["result"] == "Loss")
    pu = sum(1 for p in picks if p["result"] == "Push")
    return {"W": w, "L": l, "P": pu, "Picks": len(picks),
            "Win %": (w / (w + l)) if w + l else 0.0}


def bucket(clv):
    if clv >= GOOD:
        return "✅ Market moved FOR us"
    if clv <= BAD:
        return "❌ Market moved AGAINST us"
    return "➖ Flat (within ½ pt)"


def report(engine, season: str):
    """Outcome breakdown by CLV bucket, by member, and by week."""
    with engine.connect() as conn:
        picks = rows(conn, select(
            assignments.c.result, assignments.c.clv, assignments.c.pick_type,
            assignments.c.pick_selection, players.c.name.label("player"),
            weeks.c.week_num, games.c.away_abbr, games.c.home_abbr)
            .join(players, assignments.c.player_id == players.c.id)
            .join(games, assignments.c.game_id == games.c.id)
            .join(weeks, assignments.c.week_id == weeks.c.id)
            .where(and_(weeks.c.season == season,
                        assignments.c.result.isnot(None),
                        assignments.c.clv.isnot(None)))
            .order_by(weeks.c.week_num))
    order = ["✅ Market moved FOR us", "➖ Flat (within ½ pt)",
             "❌ Market moved AGAINST us"]
    by_bucket = defaultdict(list)
    for p in picks:
        by_bucket[bucket(p["clv"])].append(p)
    buckets = [{"When…": b, **_rec(by_bucket[b]),
                "Avg CLV": (sum(p["clv"] for p in by_bucket[b]) / len(by_bucket[b]))
                if by_bucket[b] else 0.0}
               for b in order]

    by_member = defaultdict(list)
    for p in picks:
        by_member[p["player"]].append(p)
    members = []
    for m, pp in by_member.items():
        r = _rec(pp)
        members.append({"Member": m, "Avg CLV (pts)": sum(p["clv"] for p in pp) / len(pp),
                        "Beat the close": sum(1 for p in pp if p["clv"] >= GOOD),
                        "Lost to the close": sum(1 for p in pp if p["clv"] <= BAD),
                        "Record": f"{r['W']}-{r['L']}-{r['P']}", "Win %": r["Win %"]})
    members.sort(key=lambda r: -r["Avg CLV (pts)"])

    by_week = defaultdict(list)
    for p in picks:
        by_week[p["week_num"]].append(p)
    weekly = []
    for wk in sorted(by_week):
        pp = by_week[wk]
        r = _rec(pp)
        weekly.append({"Week": wk, "Net CLV (pts)": sum(p["clv"] for p in pp),
                       "Legs for us": sum(1 for p in pp if p["clv"] >= GOOD),
                       "Legs against us": sum(1 for p in pp if p["clv"] <= BAD),
                       "Record": f"{r['W']}-{r['L']}-{r['P']}", "Win %": r["Win %"]})

    detail = [{"Wk": p["week_num"], "Member": p["player"],
               "Game": f"{p['away_abbr']} @ {p['home_abbr']}",
               "Pick": p["pick_selection"], "CLV": p["clv"],
               "Market": bucket(p["clv"]).split(" ", 1)[1].replace("Market moved ", ""),
               "Result": p["result"]} for p in picks]
    overall = _rec(picks)
    overall["Avg CLV"] = (sum(p["clv"] for p in picks) / len(picks)) if picks else 0.0
    return {"overall": overall, "buckets": buckets, "members": members,
            "weekly": weekly, "detail": detail}
