"""The Gentlemen's Wagering Cooperative — coop web app.

Members: log in with name + PIN, make picks, watch the live board.
Commissioner: everything above plus slate management, grading, and Excel.
"""
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from sqlalchemy import and_, select

from gwc import assign, rules, schedule, stats
from gwc.db import (CURRENT_SEASON, MEMBERS, SEASON_YEAR, assignments,
                    engine, games, odds_snapshots, players, rows, weeks)

ET = ZoneInfo("America/New_York")

st.set_page_config(page_title="Gentlemen's Wagering Cooperative",
                   page_icon="🎩", layout="wide")

DEFAULT_WORKBOOK = str((Path(__file__).resolve().parent.parent
                        / "2025:2026" / "Gentlemens Wagering Cooperative.xlsx"))
IS_LOCAL = Path(DEFAULT_WORKBOOK).exists()


# ---------------------------------------------------------------- data access
def q(stmt):
    with engine().connect() as conn:
        return rows(conn, stmt)


def get_players():
    return q(select(players).order_by(players.c.id))


def current_week():
    r = q(select(weeks).where(weeks.c.season == CURRENT_SEASON)
          .order_by(weeks.c.week_num.desc()))
    return r[0] if r else None


def week_games(week_id, include_excluded=False):
    stmt = select(games).where(games.c.week_id == week_id)
    if not include_excluded:
        stmt = stmt.where(games.c.excluded == False)  # noqa: E712
    return q(stmt.order_by(games.c.kickoff_et))


def week_assignments(week_id):
    stmt = (select(assignments, players.c.name.label("player"),
                   games.c.away, games.c.home, games.c.day, games.c.kickoff_et,
                   games.c.away_abbr, games.c.home_abbr, games.c.espn_id,
                   games.c.favorite, games.c.spread, games.c.ou_total,
                   games.c.away_score, games.c.home_score, games.c.final)
            .join(players, assignments.c.player_id == players.c.id)
            .join(games, assignments.c.game_id == games.c.id)
            .where(assignments.c.week_id == week_id)
            .order_by(games.c.kickoff_et, players.c.name))
    return q(stmt)


def graded_picks(season=None):
    stmt = (select(assignments.c.result, assignments.c.pick_type,
                   assignments.c.pick_selection, assignments.c.fav_dog,
                   assignments.c.home_away, assignments.c.pick_team,
                   assignments.c.pick_line, players.c.name.label("player"),
                   weeks.c.week_num, weeks.c.season, games.c.day,
                   games.c.away, games.c.home)
            .join(players, assignments.c.player_id == players.c.id)
            .join(weeks, assignments.c.week_id == weeks.c.id)
            .join(games, assignments.c.game_id == games.c.id)
            .where(assignments.c.result.isnot(None)))
    if season:
        stmt = stmt.where(weeks.c.season == season)
    return stats.enrich(q(stmt))


@st.cache_data(ttl=45)
def get_live(week_num):
    try:
        return schedule.fetch_week_live(SEASON_YEAR, week_num)
    except Exception:
        return {}


# ------------------------------------------------------------------ formatting
def fmt_pct(x):
    return f"{x:.1%}"


def matchup(g):
    return f"{g['away']} @ {g['home']}"


def line_summary(g):
    bits = []
    if g["favorite"] and g["spread"]:
        bits.append(f"{g['favorite']} -{g['spread']:g}")
    elif g["spread"] == 0:
        bits.append("Pick'em")
    if g["ou_total"]:
        bits.append(f"O/U {g['ou_total']:g}")
    return " · ".join(bits) if bits else "no line yet"


def kickoff_label(g):
    ko = datetime.fromisoformat(g["kickoff_et"])
    return ko.strftime("%a %-m/%-d %-I:%M %p ET")


def pct_df(data, pct_keys=("Win %",)):
    df = pd.DataFrame(data)
    for k in pct_keys:
        if k in df.columns:
            df[k] = df[k].map(lambda v: fmt_pct(v) if isinstance(v, float) else v)
    return df


def deadline_state(week):
    if not week or not week["deadline_et"]:
        return None, False
    dl = datetime.fromisoformat(week["deadline_et"])
    return dl, datetime.now(ET) > dl


def deadline_label(week):
    dl, _ = deadline_state(week)
    return dl.strftime("%A %-m/%-d %-I:%M %p ET") if dl else "not set"


# ---------------------------------------------------------------------- login
def login_gate():
    st.title("🎩 The Gentlemen's Wagering Cooperative")
    st.caption("Outcomes are uncertain, losses are shared, and results speak "
               "for themselves. — Charter, Section I")
    plist = get_players()
    with st.form("login"):
        name = st.selectbox("Who goes there?", [p["name"] for p in plist])
        pin = st.text_input("PIN", type="password", max_chars=10)
        if st.form_submit_button("Enter the Cooperative", type="primary"):
            match = next((p for p in plist if p["name"] == name), None)
            if match and pin == str(match["pin"]):
                st.session_state.user = match
                st.rerun()
            else:
                st.error("Incorrect PIN. Ignorance of the rules shall not "
                         "constitute exemption (Charter §XVI).")


# ------------------------------------------------------------------- game day
def cover_status(a, live):
    """(score_text, verdict_text, emoji) for one pick right now."""
    g = live.get(a["espn_id"], {})
    state = g.get("state", "pre")
    if a["result"]:
        v = {"Win": ("WON", "✅"), "Loss": ("LOST", "❌"),
             "Push": ("PUSH", "➖")}[a["result"]]
        score = (f"{a['away_abbr']} {a['away_score']} – "
                 f"{a['home_abbr']} {a['home_score']} · Final"
                 if a["away_score"] is not None else "Final")
        return score, v[0], v[1]
    if state == "pre" or not g:
        return kickoff_label(a), "not started", "🕐"
    hs, as_ = g["home_score"], g["away_score"]
    score = f"{a['away_abbr']} {as_} – {a['home_abbr']} {hs} · {g['detail']}"
    if not a["pick_selection"]:
        return score, "no pick submitted", "⚠️"
    final_ish = state == "post"
    if a["pick_type"] == "Over/Under":
        diff = (hs + as_) - (a["pick_line"] or 0)
        over = a["pick_selection"].startswith("Over")
        margin = diff if over else -diff
    elif a["pick_type"] == "Moneyline":
        my = hs if a["pick_team"] == a["home"] else as_
        margin = my - (as_ if a["pick_team"] == a["home"] else hs)
    else:  # spread
        my = hs if a["pick_team"] == a["home"] else as_
        opp = as_ if a["pick_team"] == a["home"] else hs
        margin = my + (a["pick_line"] or 0) - opp
    if margin > 0:
        verb = "covering" if a["pick_type"] != "Moneyline" else "leading"
        return score, f"{verb} by {margin:g}", "✅"
    if margin < 0:
        verb = "down" if a["pick_type"] != "Moneyline" else "trailing by"
        return score, f"{verb} {-margin:g}", "❌"
    return score, "dead on the number", "➖"


def page_game_day():
    week = current_week()
    if not week:
        st.info("No week is set up yet.")
        return
    alist = week_assignments(week["id"])
    if not alist:
        st.info("No assignments yet this week.")
        return
    st.subheader(f"🔴 Week {week['week_num']} — Live Board")
    st.caption("Refreshes itself every 60 seconds on game day. "
               f"Lines locked {deadline_label(week)}.")

    @st.fragment(run_every=60)
    def live_board():
        live = get_live(week["week_num"])
        n_live = sum(1 for g in live.values() if g.get("state") == "in")
        done = [a for a in alist if a["pick_selection"]]
        wins = losses = 0
        rows_out = []
        for a in alist:
            score, verdict, emoji = cover_status(a, live)
            if emoji == "✅":
                wins += 1
            elif emoji == "❌":
                losses += 1
            rows_out.append({"": emoji, "Day": a["day"],
                             "Member": a["player"],
                             "Matchup": f"{a['away_abbr']} @ {a['home_abbr']}",
                             "Pick": a["pick_selection"] or "—",
                             "Score": score, "Status": verdict})
        c1, c2, c3 = st.columns(3)
        c1.metric("Legs alive / dead", f"{wins} ✅ · {losses} ❌",
                  f"{len(done)} picks in")
        c2.metric("Games live right now", n_live)
        if not done:
            parlay = "⏳ waiting on picks"
        elif losses > 0:
            parlay = "💀 dead"
        else:
            parlay = "😤 STILL ALIVE"
        c3.metric("The parlay", parlay,
                  f"as of {datetime.now(ET):%-I:%M:%S %p ET}")
        st.dataframe(pd.DataFrame(rows_out), use_container_width=True,
                     hide_index=True)

    live_board()


# ------------------------------------------------------------------- my picks
def pick_controls(a, key_prefix=""):
    """Inline pick entry for one assignment. Returns True if changed."""
    g = a
    types = rules.allowed_pick_types(g)
    k = f"{key_prefix}{a['id']}"
    c1, c2, c3, c4, c5 = st.columns([2.2, 1.8, 1.2, 0.7, 0.7],
                                    vertical_alignment="bottom")
    ptype = c1.radio("Type", types, key=f"t{k}", horizontal=True,
                     index=types.index(a["pick_type"]) if a["pick_type"] in types else 0,
                     label_visibility="collapsed")
    team = ou = line = None
    if ptype == "Over/Under":
        default = 0
        if a["pick_selection"] and a["pick_selection"].startswith("Under"):
            default = 1
        ou = c2.radio("Side", ["Over", "Under"], key=f"s{k}",
                      horizontal=True, index=default, label_visibility="collapsed")
        saved = (a["pick_line"] if a["pick_type"] == "Over/Under"
                 and a["pick_line"] is not None else None)
        line = c3.number_input(
            "Total", value=float(saved if saved is not None
                                 else g["ou_total"] or 44.5), step=0.5,
            format="%.1f", key=f"l{k}-ou",
            help="Edit if your book's total differs from ESPN's")
    elif ptype == "Spread":
        opts = [g["away"], g["home"]]
        default = opts.index(a["pick_team"]) if a["pick_team"] in opts else 0
        team = c2.radio("Team", opts, key=f"m{k}", horizontal=True,
                        index=default, label_visibility="collapsed")
        saved = (a["pick_line"] if a["pick_type"] == "Spread"
                 and a["pick_team"] == team and a["pick_line"] is not None
                 else None)
        if g["spread"] is None or g["favorite"] is None:
            derived = 0.0
        else:
            derived = -g["spread"] if team == g["favorite"] else g["spread"]
        line = c3.number_input(
            "Line", value=float(saved if saved is not None else derived),
            step=0.5, format="%.1f", key=f"l{k}-{team}",
            help="Spread for your team — negative means laying points. "
                 "Enter the line your book gave you")
    else:
        opts = [g["away"], g["home"]]
        default = opts.index(a["pick_team"]) if a["pick_team"] in opts else 0
        team = c2.radio("Team", opts, key=f"m{k}", horizontal=True,
                        index=default, label_visibility="collapsed")
        c3.markdown("&nbsp;")
    if c4.button("Lock it in 🔒", key=f"b{k}", type="primary"):
        try:
            fields = rules.build_pick(g, ptype, team, ou, line)
            with engine().begin() as conn:
                conn.execute(assignments.update()
                             .where(assignments.c.id == a["id"])
                             .values(submitted_at=datetime.utcnow(), **fields))
            return True
        except ValueError as e:
            st.error(str(e))
    if a["pick_selection"] and c5.button("Clear", key=f"x{k}"):
        with engine().begin() as conn:
            conn.execute(assignments.update()
                         .where(assignments.c.id == a["id"])
                         .values(pick_type=None, pick_selection=None,
                                 pick_team=None, pick_line=None, fav_dog=None,
                                 home_away=None, submitted_at=None, result=None))
        return True
    return False


def page_my_picks(user):
    week = current_week()
    if not week:
        st.info("No week is set up yet. Pester the Commissioner.")
        return
    st.subheader(f"Week {week['week_num']} — Your Assignments")
    dl, past = deadline_state(week)
    locked = (past or week["status"] == "final") and not user["is_commissioner"]
    if dl and not past:
        left = dl - datetime.now(ET)
        hrs, rem = divmod(int(left.total_seconds()), 3600)
        st.info(f"⏳ Picks lock **{deadline_label(week)}** — "
                f"{hrs}h {rem // 60}m to go. You can change your pick "
                "until then.")
    elif past and locked:
        st.error(f"⏰ Deadline passed ({deadline_label(week)}). "
                 "Late changes require the Commissioner (Charter §XIII).")

    mine = [a for a in week_assignments(week["id"])
            if a["player_id"] == user["id"]]
    if not mine:
        st.warning("You weren't assigned any games this week.")
        return
    for a in mine:
        icon = {"Win": "✅", "Loss": "❌", "Push": "➖"}.get(a["result"], "🏈")
        with st.container(border=True):
            st.markdown(f"**{icon} {matchup(a)}**  \n"
                        f"{kickoff_label(a)} · {line_summary(a)}")
            if a["result"]:
                st.markdown(f"**{a['pick_selection']}** — {a['result']} "
                            f"({a['away']} {a['away_score']} - "
                            f"{a['home']} {a['home_score']})")
                continue
            if locked:
                st.markdown(f"Your pick: **{a['pick_selection'] or '— none submitted —'}**")
                continue
            if pick_controls(a, "me"):
                st.rerun()
            if a["pick_selection"]:
                st.success(f"Submitted: **{a['pick_selection']}**")


# ------------------------------------------------------------------- bet slip
def page_bet_slip():
    week = current_week()
    if not week:
        st.info("No week is set up yet.")
        return
    st.subheader(f"Week {week['week_num']} — Consolidated Bet Slip")
    alist = week_assignments(week["id"])
    missing = sorted({a["player"] for a in alist if not a["pick_selection"]})
    done = [a for a in alist if a["pick_selection"]]
    if missing:
        st.warning("Still waiting on: " + ", ".join(missing))
    if not done:
        st.info("No picks submitted yet.")
        return
    df = pd.DataFrame([{
        "Day": a["day"], "Kickoff": kickoff_label(a), "Member": a["player"],
        "Matchup": matchup(a), "Pick": a["pick_selection"] or "—",
        "Result": a["result"] or "—",
    } for a in alist])
    st.dataframe(df, use_container_width=True, hide_index=True)
    if not missing:
        legs = len(done)
        payout = 5 * (1.909 ** legs)
        slip = [f"🎩🏈 GWC WEEK {week['week_num']} BET SLIP 🏈🎩", ""]
        slip += [f"{a['day']} — {a['player']}: {a['pick_selection']} "
                 f"({matchup(a)})" for a in done]
        slip += ["", f"{legs}-leg parlay · $5 to win ≈ ${payout:,.0f} "
                     "(at -110 per leg)", "Good luck, gentlemen. 🍀"]
        st.markdown("**Copy for the group chat:**")
        st.code("\n".join(slip), language=None)


# ------------------------------------------------------------------ standings
def page_standings():
    seasons = sorted({w["season"] for w in q(select(weeks))}, reverse=True)
    if not seasons:
        st.info("No data yet.")
        return
    season = st.selectbox("Season", seasons)
    picks = graded_picks(season)
    if not picks:
        st.info("No graded picks yet this season.")
        return
    table, coop = stats.standings(picks, MEMBERS)
    c1, c2, c3 = st.columns(3)
    c1.metric("Coop Record", f"{coop['Wins']}-{coop['Losses']}-{coop['Pushes']}")
    c2.metric("Coop Win %", fmt_pct(coop["Win %"]))
    leader = table[0]
    c3.metric("Season Champion (leader)", leader["Player"],
              fmt_pct(leader["Win %"]))
    st.dataframe(pct_df(table), use_container_width=True, hide_index=True)
    st.markdown("#### Weekly Champions")
    rep = stats.weekly_report(picks)
    if rep:
        st.dataframe(pct_df(rep, pct_keys=("Coop Win %",)),
                     use_container_width=True, hide_index=True)


def page_analytics():
    seasons = sorted({w["season"] for w in q(select(weeks))}, reverse=True)
    if not seasons:
        st.info("No data yet.")
        return
    season = st.selectbox("Season", ["All time"] + seasons)
    picks = graded_picks(None if season == "All time" else season)
    if not picks:
        st.info("No graded picks yet.")
        return
    tab_over, tab_trend, tab_sit, tab_brag = st.tabs(
        ["📊 Overview", "📉 Trends", "🎯 Situations", "🏆 Bragging Rights"])
    with tab_over:
        overview, per_player = stats.analytics(picks, MEMBERS)
        c1, c2 = st.columns(2)
        c1.markdown("#### Group Splits")
        c1.dataframe(pct_df(overview), use_container_width=True, hide_index=True)
        c2.markdown("#### Per Player")
        c2.dataframe(pct_df(per_player), use_container_width=True, hide_index=True)
        st.markdown("#### Situational Matrix")
        st.dataframe(pd.DataFrame(stats.player_matrix(picks, MEMBERS)),
                     use_container_width=True, hide_index=True)
    with tab_trend:
        matrix = stats.weekly_matrix(picks, MEMBERS)
        if len(matrix) >= 2:
            st.markdown("#### Win % by week")
            chart = pd.DataFrame.from_dict(matrix, orient="index")[MEMBERS]
            chart.index.name = "Week"
            st.line_chart(chart.map(
                lambda v: round(v * 100, 1) if v is not None else None))
        st.markdown("#### 🔥 Hot / Cold — last 4 weeks vs season")
        st.dataframe(
            pct_df(stats.form(picks, MEMBERS),
                   pct_keys=("Last 4W Win %", "Season Win %", "Diff")),
            use_container_width=True, hide_index=True)
        heat = pd.DataFrame.from_dict(matrix, orient="index")[MEMBERS]
        heat.index = [f"Week {w}" for w in heat.index]
        st.markdown("#### Week-by-week board")
        st.dataframe(heat.map(
            lambda v: fmt_pct(v) if v is not None else "—"),
            use_container_width=True)
    with tab_sit:
        c1, c2 = st.columns(2)
        c1.markdown("#### By day of week")
        c1.dataframe(pct_df(stats.day_splits(picks)),
                     use_container_width=True, hide_index=True)
        c2.markdown("#### By spread size")
        c2.dataframe(pct_df(stats.spread_buckets(picks)),
                     use_container_width=True, hide_index=True)
        pt_dogs = [p for p in picks if p.get("day") in ("Thu", "Mon", "Fri", "Sat")
                   and p["fav_dog"] == "Underdog"]
        w = sum(1 for p in pt_dogs if p["result"] == "Win")
        l = sum(1 for p in pt_dogs if p["result"] == "Loss")
        st.metric("🌙 JR's Primetime Dogs Tracker (Thu/Sat/Mon underdogs)",
                  f"{w}-{l}" + (f"  ({w / (w + l):.0%})" if w + l else ""))
        teams = stats.team_records(picks)
        if teams:
            st.markdown("#### Riding the teams")
            c1, c2 = st.columns(2)
            eligible = [t for t in teams if t["Wins"] + t["Losses"] >= 3]
            c1.markdown("**Best to back** (3+ picks)")
            c1.dataframe(pct_df(sorted(eligible, key=lambda t: -t["Win %"])[:8]),
                         use_container_width=True, hide_index=True)
            c2.markdown("**Cursed teams** (3+ picks)")
            c2.dataframe(pct_df(sorted(eligible, key=lambda t: t["Win %"])[:8]),
                         use_container_width=True, hide_index=True)
    with tab_brag:
        awards, perfect, goose = stats.bragging_rights(picks, MEMBERS)
        if awards:
            st.dataframe(pd.DataFrame(awards), use_container_width=True,
                         hide_index=True)
        st.markdown("#### Pick streaks")
        st.dataframe(pd.DataFrame(stats.pick_streaks(picks, MEMBERS)),
                     use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        c1.markdown("**Perfect Weeks (100%)**")
        c1.dataframe(pd.DataFrame(perfect), use_container_width=True,
                     hide_index=True)
        c2.markdown("**Goose Eggs (0%)**")
        c2.dataframe(pd.DataFrame(goose), use_container_width=True,
                     hide_index=True)


# --------------------------------------------------------------- market watch
@st.cache_data(ttl=600)
def get_halftime(espn_id):
    try:
        return schedule.fetch_halftime(espn_id)
    except Exception:
        return None


def signed_home_spread(row, game):
    """Home-team spread (negative = home favored) for one snapshot row."""
    if row["spread"] is None:
        return None
    if row["favorite"] is None:
        return 0.0
    return -row["spread"] if row["favorite"] == game["home"] else row["spread"]


def page_market():
    week = current_week()
    if not week:
        st.info("No week is set up yet.")
        return
    tab_lines, tab_dogs = st.tabs(["📈 Line moves", "🌙 1H Primetime Dogs"])

    with tab_lines:
        st.subheader(f"Week {week['week_num']} — how the lines have moved")
        glist = week_games(week["id"])
        snaps = q(select(odds_snapshots)
                  .where(and_(odds_snapshots.c.season_year == SEASON_YEAR,
                              odds_snapshots.c.week_num == week["week_num"]))
                  .order_by(odds_snapshots.c.ts))
        by_game = {}
        for s in snaps:
            by_game.setdefault(s["espn_id"], []).append(s)
        if not by_game:
            st.info("No line history yet — snapshots collect every 6 hours.")
            return
        first_ts = min(s["ts"] for s in snaps)
        st.caption(f"Tracked every 6 hours since "
                   f"{datetime.fromisoformat(first_ts):%b %-d} — the further "
                   "into the week, the richer this gets. A line that moves "
                   "*against* the favorite usually means sharp money on "
                   "the dog.")
        move_rows = []
        for g in glist:
            hist = by_game.get(g["espn_id"], [])
            if len(hist) < 1:
                continue
            o, n = hist[0], hist[-1]
            def label(r):
                if r["spread"] is None:
                    return "—"
                return (f"{r['favorite']} -{r['spread']:g}"
                        if r["favorite"] else "PK")
            so, sn = signed_home_spread(o, g), signed_home_spread(n, g)
            if so is not None and sn is not None and sn != so:
                toward = g["home"] if sn < so else g["away"]
                move = f"{abs(sn - so):g} pt toward {toward}"
                dog = (g["away"] if g["favorite"] == g["home"] else g["home"])
                flag = "🚨 toward the dog" if toward == dog else ""
            else:
                move, flag = "—", ""
            move_rows.append({
                "Matchup": f"{g['away_abbr']} @ {g['home_abbr']}",
                "Opened (our tracking)": label(o), "Now": label(n),
                "O/U now": g["ou_total"], "Move": move, "": flag})
        st.dataframe(pd.DataFrame(move_rows), use_container_width=True,
                     hide_index=True)
        choice = st.selectbox(
            "Chart a game", glist,
            format_func=lambda g: f"{matchup(g)} ({kickoff_label(g)})")
        hist = by_game.get(choice["espn_id"], [])
        if len(hist) >= 2:
            df = pd.DataFrame({
                "when": [datetime.fromisoformat(s["ts"]) for s in hist],
                f"{choice['home_abbr']} spread (neg = favored)":
                    [signed_home_spread(s, choice) for s in hist],
                "total": [s["ou_total"] for s in hist],
            }).set_index("when")
            c1, c2 = st.columns(2)
            c1.line_chart(df.iloc[:, [0]])
            c2.line_chart(df.iloc[:, [1]])
        else:
            st.caption("Need at least two snapshots to chart this one.")

    with tab_dogs:
        st.subheader("🌙 The 1H Primetime Dogs Tracker")
        st.caption("JR's baby. Every primetime game (7 PM ET or later), take "
                   "the underdog on the first-half line — house convention: "
                   "half the full-game spread. Not part of the parlay; purely "
                   "for the culture.")
        wks = q(select(weeks).where(weeks.c.season == CURRENT_SEASON)
                .order_by(weeks.c.week_num))
        track, wins, losses, pushes = [], 0, 0, 0
        for wk in wks:
            for g in week_games(wk["id"], include_excluded=True):
                ko = datetime.fromisoformat(g["kickoff_et"])
                if ko.hour < 19 or not g["favorite"] or not g["spread"]:
                    continue
                dog = g["away"] if g["favorite"] == g["home"] else g["home"]
                line1h = round(g["spread"] / 2 * 2) / 2  # half line, ½-pt steps
                ht = get_halftime(g["espn_id"]) if g["espn_id"] else None
                if ht:
                    dog_pts = (ht["away_1h"] if dog == g["away"]
                               else ht["home_1h"])
                    fav_pts = (ht["home_1h"] if dog == g["away"]
                               else ht["away_1h"])
                    margin = dog_pts + line1h - fav_pts
                    if margin > 0:
                        res, wins = "✅ WIN", wins + 1
                    elif margin < 0:
                        res, losses = "❌ LOSS", losses + 1
                    else:
                        res, pushes = "➖ PUSH", pushes + 1
                    half = f"{ht['away_1h']}–{ht['home_1h']} at the half"
                else:
                    res, half = "🕐 pending", kickoff_label(g)
                track.append({"Week": wk["week_num"],
                              "Game": f"{g['away_abbr']} @ {g['home_abbr']}",
                              "The bet": f"{dog} +{line1h:g} (1H)",
                              "Half score": half, "Result": res})
        rec = f"{wins}-{losses}" + (f"-{pushes}" if pushes else "")
        st.metric("Season record — primetime dogs, first half", rec or "0-0")
        if track:
            st.dataframe(pd.DataFrame(track), use_container_width=True,
                         hide_index=True)
        st.caption("Last season's coop record on full-game primetime dogs: "
                   "7-3. The man may be onto something.")


# --------------------------------------------------------------- commissioner
def create_week(wk_num):
    fetched = schedule.fetch_week(SEASON_YEAR, wk_num)
    if not fetched:
        st.error("ESPN returned no games for that week.")
        return
    with engine().begin() as conn:
        wid = conn.execute(weeks.insert().values(
            season=CURRENT_SEASON, week_num=wk_num, status="draft",
            deadline_et=schedule.default_deadline(fetched),
        )).inserted_primary_key[0]
        for g in fetched:
            g = dict(g)
            g.pop("final", None)
            conn.execute(games.insert().values(week_id=wid, **g))


def refresh_lines(week):
    fetched = {g["espn_id"]: g for g in
               schedule.fetch_week(SEASON_YEAR, week["week_num"])}
    with engine().begin() as conn:
        for g in week_games(week["id"], include_excluded=True):
            f = fetched.get(g["espn_id"])
            if f:
                conn.execute(games.update().where(games.c.id == g["id"]).values(
                    favorite=f["favorite"], spread=f["spread"],
                    ou_total=f["ou_total"]))


def run_randomize(week):
    glist = week_games(week["id"])
    hist_stmt = (select(weeks.c.week_num, players.c.name.label("player"),
                        games.c.day, games.c.away, games.c.home)
                 .select_from(assignments
                              .join(players, assignments.c.player_id == players.c.id)
                              .join(games, assignments.c.game_id == games.c.id)
                              .join(weeks, assignments.c.week_id == weeks.c.id))
                 .where(weeks.c.week_num < week["week_num"]))
    history = assign.build_history(q(hist_stmt))
    active = [p for p in get_players() if p["active"]]
    result = assign.randomize(glist, [p["name"] for p in active], history)
    ids = {p["name"]: p["id"] for p in active}
    with engine().begin() as conn:
        conn.execute(assignments.delete()
                     .where(assignments.c.week_id == week["id"]))
        for gid, name in result.items():
            conn.execute(assignments.insert().values(
                week_id=week["id"], game_id=gid, player_id=ids[name]))


def grade_week(week):
    live = schedule.fetch_week_live(SEASON_YEAR, week["week_num"])
    graded = 0
    with engine().begin() as conn:
        for g in week_games(week["id"], include_excluded=True):
            f = live.get(g["espn_id"])
            if f and f["final"]:
                conn.execute(games.update().where(games.c.id == g["id"]).values(
                    final=True, home_score=f["home_score"],
                    away_score=f["away_score"]))
                g.update(final=True, home_score=f["home_score"],
                         away_score=f["away_score"])
            for a in rows(conn, select(assignments).where(and_(
                    assignments.c.game_id == g["id"],
                    assignments.c.pick_selection.isnot(None)))):
                res = rules.grade(g, a)
                if res:
                    conn.execute(assignments.update()
                                 .where(assignments.c.id == a["id"])
                                 .values(result=res))
                    graded += 1
    return graded


def assignments_message(week, alist):
    by_member = {}
    for a in alist:
        by_member.setdefault(a["player"], []).append(a)
    lines = [f"🎩🏈 GWC WEEK {week['week_num']} ASSIGNMENTS 🏈🎩", ""]
    for m in MEMBERS:
        if m not in by_member:
            continue
        lines.append(f"{m}:")
        for a in by_member[m]:
            lines.append(f"  • {matchup(a)} ({kickoff_label(a)}) — {line_summary(a)}")
        lines.append("")
    lines.append(f"Picks due {deadline_label(week)} — make them on the site! "
                 "Spreads & totals; ML only if the spread is under 3.")
    return "\n".join(lines)


def page_this_week():
    st.subheader("🧢 Commissioner's Office")
    wlist = q(select(weeks).where(weeks.c.season == CURRENT_SEASON)
              .order_by(weeks.c.week_num))
    labels = {w["id"]: f"Week {w['week_num']}  ·  {w['status']}" for w in wlist}
    options = list(labels) + ["➕ Start a new week"]
    c1, c2 = st.columns([2.5, 1])
    sel = c1.selectbox("Week", options,
                       index=len(wlist) - 1 if wlist else 0,
                       format_func=lambda x: labels.get(x, x))
    if sel == "➕ Start a new week":
        wk_num = c2.number_input("NFL week #", 1, 18,
                                 value=(wlist[-1]["week_num"] + 1) if wlist else 1)
        if st.button("📡 Fetch slate from ESPN", type="primary"):
            create_week(int(wk_num))
            st.rerun()
        return
    week = next(w for w in wlist if w["id"] == sel)
    alist = week_assignments(week["id"])
    missing = [m for m in MEMBERS
               if any(a["player"] == m and not a["pick_selection"] for a in alist)]

    with st.expander("1️⃣ Slate — games & lines", expanded=not alist):
        all_games = week_games(week["id"], include_excluded=True)
        df = pd.DataFrame([{
            "id": g["id"], "Kickoff": kickoff_label(g), "Matchup": matchup(g),
            "Favorite": g["favorite"] or "", "Spread": g["spread"],
            "Total": g["ou_total"], "Include": not g["excluded"],
        } for g in all_games])
        edited = st.data_editor(
            df, hide_index=True, use_container_width=True,
            disabled=["id", "Kickoff", "Matchup"],
            column_config={"id": None}, key="slate_ed")
        c1, c2, _ = st.columns([1, 1, 2])
        if c1.button("💾 Save slate changes"):
            teams_by_id = {g["id"]: (g["away"], g["home"]) for g in all_games}
            problems = []
            with engine().begin() as conn:
                for _, r in edited.iterrows():
                    gid = int(r["id"])
                    fav = (r["Favorite"] or "").strip() or None
                    if fav and fav not in teams_by_id[gid]:
                        problems.append(f"{r['Matchup']}: {fav!r} isn't a team")
                        continue
                    conn.execute(games.update().where(games.c.id == gid).values(
                        excluded=not bool(r["Include"]), favorite=fav,
                        spread=None if pd.isna(r["Spread"]) else float(r["Spread"]),
                        ou_total=None if pd.isna(r["Total"]) else float(r["Total"])))
            if problems:
                st.error("\n".join(problems))
            else:
                st.rerun()
        if c2.button("🔄 Refresh lines from ESPN"):
            refresh_lines(week)
            st.rerun()

    with st.expander("2️⃣ Assignments — randomize & announce",
                     expanded=bool(week_games(week["id"])) and not alist):
        label = "🎲 Randomize assignments" if not alist else "🎲 Re-roll assignments"
        if st.button(label, type="primary"):
            run_randomize(week)
            st.rerun()
        if alist:
            st.dataframe(pd.DataFrame([{
                "Day": a["day"], "Kickoff": kickoff_label(a),
                "Matchup": matchup(a), "Line": line_summary(a),
                "Member": a["player"],
            } for a in alist]), use_container_width=True, hide_index=True)
            c1, c2, c3 = st.columns([2, 1, 1])
            target = c1.selectbox(
                "Move a game (pick switching, Charter IV)", alist,
                format_func=lambda a: f"{matchup(a)} — {a['player']}")
            new_owner = c2.selectbox("To", MEMBERS)
            if c3.button("Switch"):
                pid = next(p["id"] for p in get_players()
                           if p["name"] == new_owner)
                with engine().begin() as conn:
                    conn.execute(assignments.update()
                                 .where(assignments.c.id == target["id"])
                                 .values(player_id=pid))
                st.rerun()
            st.markdown("**📣 Copy for the group chat:**")
            st.code(assignments_message(week, alist), language=None)
            if week["status"] == "draft":
                if st.button("Mark as announced ✅"):
                    with engine().begin() as conn:
                        conn.execute(weeks.update().where(weeks.c.id == week["id"])
                                     .values(status="announced"))
                    st.rerun()

    if alist:
        with st.expander("3️⃣ Picks — log any that come in by text",
                         expanded=False):
            st.caption("The guys should use My Picks themselves — this is "
                       "your override for texted picks and fixes.")
            for m in MEMBERS:
                theirs = [a for a in alist if a["player"] == m]
                if not theirs:
                    continue
                n_in = sum(1 for a in theirs if a["pick_selection"])
                icon = "✅" if n_in == len(theirs) else f"{n_in}/{len(theirs)}"
                st.markdown(f"#### {m} — {icon}")
                for a in theirs:
                    status = a["pick_selection"] or "*no pick yet*"
                    res = {"Win": " ✅", "Loss": " ❌", "Push": " ➖"}.get(a["result"], "")
                    st.markdown(f"**{matchup(a)}** · {kickoff_label(a)} · "
                                f"{line_summary(a)} → **{status}**{res}")
                    if pick_controls(a, "cm"):
                        st.rerun()
                    st.markdown("")

        with st.expander("4️⃣ Results & grading"):
            if st.button("🏁 Fetch scores from ESPN & grade picks",
                         type="primary"):
                n = grade_week(week)
                st.success(f"Graded {n} picks.")
                st.rerun()
            graded = [a for a in week_assignments(week["id"]) if a["result"]]
            pending = [a for a in week_assignments(week["id"])
                       if a["pick_selection"] and not a["result"]]
            if graded and not pending and week["status"] != "final":
                if st.button("Mark week FINAL 🏁"):
                    with engine().begin() as conn:
                        conn.execute(weeks.update().where(weeks.c.id == week["id"])
                                     .values(status="final"))
                    st.rerun()

    with st.expander("⚙️ Week admin & member PINs"):
        dl = st.text_input("Deadline (ET, ISO — e.g. 2026-09-10T16:00:00-04:00)",
                           value=week["deadline_et"] or "")
        if st.button("Update deadline"):
            with engine().begin() as conn:
                conn.execute(weeks.update().where(weeks.c.id == week["id"])
                             .values(deadline_et=dl or None))
            st.rerun()
        pdf = pd.DataFrame([{"Member": p["name"], "PIN": str(p["pin"])}
                            for p in get_players()])
        pedit = st.data_editor(pdf, hide_index=True, disabled=["Member"],
                               key="pin_ed")
        if st.button("Save PINs"):
            with engine().begin() as conn:
                for _, r in pedit.iterrows():
                    conn.execute(players.update()
                                 .where(players.c.name == r["Member"])
                                 .values(pin=str(r["PIN"])))
            st.success("PINs updated.")
        st.divider()
        if st.button("🗑️ Delete this week (games and picks included)"):
            with engine().begin() as conn:
                conn.execute(assignments.delete()
                             .where(assignments.c.week_id == week["id"]))
                conn.execute(games.delete().where(games.c.week_id == week["id"]))
                conn.execute(weeks.delete().where(weeks.c.id == week["id"]))
            st.rerun()


def page_excel():
    from gwc import export
    st.subheader("📤 Excel")
    path = st.text_input("Workbook path", value=DEFAULT_WORKBOOK)
    c1, c2 = st.columns(2)
    if c1.button("🔄 Sync graded picks to workbook", type="primary"):
        try:
            added, backup = export.sync_workbook(path)
            st.success(f"Added {added} pick(s). Backup: `{backup}`")
        except Exception as e:
            st.error(str(e))
    if c2.button("📺 Refresh live-week panel"):
        try:
            n = export.write_live_feed(path)
            st.success(f"Live panel refreshed — {n} game(s).")
        except Exception as e:
            st.error(str(e))
    st.divider()
    st.download_button(
        "⬇️ Download GWC report (.xlsx)", data=export.snapshot_bytes(),
        file_name=f"GWC Report {datetime.now(ET):%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------------------------------------------------- main
def main():
    if "user" not in st.session_state:
        login_gate()
        return
    user = st.session_state.user
    st.sidebar.title("🎩 G.W.C.")
    st.sidebar.caption(f"Signed in as **{user['name']}**"
                       + (" · Commissioner" if user["is_commissioner"] else ""))
    pages = ["🔴 Game Day", "🏈 My Picks", "🧾 Bet Slip", "🏆 Standings",
             "📈 Analytics", "📉 Market"]
    if user["is_commissioner"]:
        pages.append("🧢 This Week")
        if IS_LOCAL:
            pages.append("📤 Excel")
    page = st.sidebar.radio("Go to", pages, label_visibility="collapsed")
    if st.sidebar.button("Log out"):
        del st.session_state.user
        st.rerun()
    st.title("The Gentlemen's Wagering Cooperative")
    if page == "🔴 Game Day":
        page_game_day()
    elif page == "🏈 My Picks":
        page_my_picks(user)
    elif page == "🧾 Bet Slip":
        page_bet_slip()
    elif page == "🏆 Standings":
        page_standings()
    elif page == "📈 Analytics":
        page_analytics()
    elif page == "📉 Market":
        page_market()
    elif page == "🧢 This Week":
        page_this_week()
    else:
        page_excel()


main()
