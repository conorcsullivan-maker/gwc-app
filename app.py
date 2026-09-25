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

from gwc import assign, cashout, clv, rules, schedule, stats
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
                   games.c.away_score, games.c.home_score, games.c.final,
                   games.c.away_1h, games.c.home_1h)
            .join(players, assignments.c.player_id == players.c.id)
            .join(games, assignments.c.game_id == games.c.id)
            .where(assignments.c.week_id == week_id)
            .order_by(games.c.kickoff_et, players.c.name))
    return q(stmt)


def graded_picks(season=None):
    stmt = (select(assignments.c.result, assignments.c.pick_type,
                   assignments.c.pick_selection, assignments.c.fav_dog,
                   assignments.c.home_away, assignments.c.pick_team,
                   assignments.c.pick_line, assignments.c.period,
                   players.c.name.label("player"),
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


def half_line(x):
    """House convention: a first-half number is half the full-game number,
    rounded to the nearest half point (round(x)/2 does exactly that)."""
    return round(float(x)) / 2


def deadline_state(week):
    if not week or not week["deadline_et"]:
        return None, False
    dl = datetime.fromisoformat(week["deadline_et"])
    return dl, datetime.now(ET) > dl


def deadline_label(week):
    dl, _ = deadline_state(week)
    return dl.strftime("%A %-m/%-d %-I:%M %p ET") if dl else "not set"


def lock_state(week):
    """Picks harden at the week's FIRST kickoff. The posted deadline is a
    soft target — late picks are still allowed right up to kickoff, but
    nobody gets to pick a game that has already started."""
    if not week:
        return None, False
    kos = [datetime.fromisoformat(g["kickoff_et"])
           for g in week_games(week["id"]) if g["kickoff_et"]]
    if not kos:
        return None, False
    lock = min(kos)
    return lock, datetime.now(ET) > lock


def lock_label(week):
    lock, _ = lock_state(week)
    return lock.strftime("%A %-m/%-d %-I:%M %p ET") if lock else "kickoff"


def countdown(target):
    left = target - datetime.now(ET)
    total = int(left.total_seconds())
    if total <= 0:
        return "now"
    d, rem = divmod(total, 86400)
    h, rem = divmod(rem, 3600)
    m = rem // 60
    return (f"{d}d {h}h" if d else f"{h}h {m}m")


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
    """(score_text, verdict_text, emoji) for one pick right now.

    House rule: no wins, losses, or verdicts of any kind until a game is
    FINAL and graded — in-progress games just show the score.
    """
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
        if not a["pick_selection"]:
            return kickoff_label(a), "no pick submitted", "⚠️"
        return kickoff_label(a), "not started", "🕐"
    hs, as_ = g["home_score"], g["away_score"]
    score = f"{a['away_abbr']} {as_} – {a['home_abbr']} {hs} · {g['detail']}"
    if not a["pick_selection"]:
        return score, "no pick submitted", "⚠️"
    if state == "post":
        return score, "final — grading…", "⏳"
    if a.get("period") == "1H":
        return score, "in play (1st-half bet)", "🏈"
    return score, "in play", "🏈"


def auto_grade(week, alist, live):
    """Grade any pick whose game just went final. Runs on every board
    refresh, so standings and analytics update the moment ESPN calls it —
    for every viewer, no commissioner action needed. Idempotent."""
    finals = {eid for eid, g in live.items() if g.get("final")}
    # a 1H bet can only settle once the first half is over: halftime, Q3+,
    # or final. Never mid-2nd-quarter.
    half_done = {eid for eid, g in live.items()
                 if g.get("final") or (g.get("period") or 0) >= 3
                 or g.get("status_name") == "STATUS_HALFTIME"}
    todo = [a for a in alist
            if a["pick_selection"] and not a["result"]
            and (a["espn_id"] in finals
                 or (a.get("period") == "1H" and a["espn_id"] in half_done))]
    if not todo:
        return False
    with engine().begin() as conn:
        for a in todo:
            g = live[a["espn_id"]]
            game = dict(a)
            if a["espn_id"] in finals:
                game.update(final=True, home_score=g["home_score"],
                            away_score=g["away_score"])
                conn.execute(games.update().where(games.c.id == a["game_id"])
                             .values(final=True, home_score=g["home_score"],
                                     away_score=g["away_score"]))
            if a.get("period") == "1H" and game.get("home_1h") is None:
                ht = get_halftime(a["espn_id"])
                if ht:
                    game.update(home_1h=ht["home_1h"], away_1h=ht["away_1h"])
                    conn.execute(games.update()
                                 .where(games.c.id == a["game_id"])
                                 .values(home_1h=ht["home_1h"],
                                         away_1h=ht["away_1h"]))
            res = rules.grade(game, a)
            if res:
                conn.execute(assignments.update()
                             .where(assignments.c.id == a["id"])
                             .values(result=res))
    return True


def page_game_day():
    week = current_week()
    if not week:
        st.info("No week is set up yet.")
        return
    st.subheader(f"🔴 Week {week['week_num']} — Live Board")
    st.caption("Refreshes every 60 seconds. Picks are only marked won or "
               "lost when a game goes FINAL — no premature obituaries. "
               f"Picks lock at kickoff — {lock_label(week)}.")

    @st.fragment(run_every=60)
    def live_board():
        alist = week_assignments(week["id"])
        if not alist:
            st.info("No assignments yet this week.")
            return
        live = get_live(week["week_num"])
        if auto_grade(week, alist, live):
            alist = week_assignments(week["id"])   # pick up fresh grades
        done = [a for a in alist if a["pick_selection"]]
        w = sum(1 for a in alist if a["result"] == "Win")
        l = sum(1 for a in alist if a["result"] == "Loss")
        p = sum(1 for a in alist if a["result"] == "Push")
        in_play = sum(1 for a in alist
                      if not a["result"]
                      and live.get(a["espn_id"], {}).get("state") == "in")
        rows_out = []
        for a in alist:
            score, verdict, emoji = cover_status(a, live)
            rows_out.append({"": emoji, "Day": a["day"],
                             "Member": a["player"],
                             "Matchup": f"{a['away_abbr']} @ {a['home_abbr']}",
                             "Pick": a["pick_selection"] or "—",
                             "Score": score, "Status": verdict})
        c1, c2, c3 = st.columns(3)
        c1.metric("Final results", f"{w}-{l}-{p}",
                  f"{w + l + p} of {len(done)} picks settled")
        c2.metric("In play right now", in_play)
        if not done:
            parlay = "⏳ waiting on picks"
        elif l > 0:
            parlay = "💀 dead"
        elif w + l + p == len(done) == len(alist):
            parlay = "🏆 IT HIT?!"
        else:
            parlay = "😤 still alive"
        c3.metric("The parlay", parlay,
                  f"as of {datetime.now(ET):%-I:%M:%S %p ET}")
        st.dataframe(pd.DataFrame(rows_out), use_container_width=True,
                     hide_index=True)

        wk = current_week()               # fresh row: bet details may change
        if wk.get("payout"):
            st.markdown("#### 🎟️ The ticket")
            b1, b2, b3 = st.columns(3)
            b1.metric("Stake", f"${wk['stake']:,.2f}" if wk.get("stake") else "—")
            b2.metric("Odds", f"{wk['odds']:+d}" if wk.get("odds") else "—")
            b3.metric("Pays if it hits", f"${wk['payout']:,.2f}",
                      (f"+${wk['payout'] - wk['stake']:,.2f} profit"
                       if wk.get("stake") else None))
            with st.expander("💸 Cash-out estimate — what a book would offer "
                             "right now", expanded=in_play > 0):
                margin = st.slider("Book's margin (haircut on fair value)",
                                   0, 15, 8, format="%d%%",
                                   help="Sportsbooks price cash-out at fair "
                                        "value minus a cut — usually 5–10%.")
                v = cashout.valuation(wk, alist, live, margin / 100)
                e1, e2, e3 = st.columns(3)
                def money(x):
                    return f"${x:,.2f}" if x < 100 else f"${x:,.0f}"
                e1.metric("Chance the ticket still hits",
                          f"{v['prob']:.1%}" if v["prob"] >= 0.001
                          else f"{v['prob']:.4%}")
                e2.metric("Fair value right now", money(v["fair"]))
                e3.metric("Estimated cash-out offer", money(v["offer"]),
                          "ticket is dead" if v["dead"] else
                          (f"pushes trimmed payout to ${v['payout']:,.0f}"
                           if v["pushes"] else None))
                st.dataframe(pd.DataFrame([{
                    "Member": l["member"], "Pick": l["pick"],
                    "Status": {"pre": "not started", "in": "in play",
                               "post": "final"}.get(l["status"], l["status"]),
                    "Win prob": f"{l['p']:.0%}"} for l in v["legs"]]),
                    use_container_width=True, hide_index=True)
                st.caption("Model: fair value = payout × P(every open leg "
                           "wins). In-play legs use the current margin vs. "
                           "the line and time left (NFL margins ≈ normal, "
                           "σ 13.5 pts full game). Unstarted spread/total "
                           "legs are 50/50. Charter §VI: cashing out takes a "
                           "4-of-6 vote.")

    live_board()


# ------------------------------------------------------------------- my picks
def pick_controls(a, key_prefix=""):
    """Inline pick entry for one assignment. Returns True if changed."""
    g = a
    types = rules.allowed_pick_types(g)
    k = f"{key_prefix}{a['id']}"
    # Charter amendment (4-of-6 vote, Sept 2026): anyone may take a bet on
    # the first half instead of the full game.
    period = st.radio("Period", ["FG", "1H"], key=f"p{k}", horizontal=True,
                      index=1 if a.get("period") == "1H" else 0,
                      format_func=lambda v: {"FG": "Full game",
                                             "1H": "1st half"}[v],
                      label_visibility="collapsed",
                      help="First-half bets settle at halftime. The line "
                           "defaults to half the full-game number — edit it "
                           "to whatever your book actually posts.")
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
                 and (a.get("period") or "FG") == period
                 and a["pick_line"] is not None else None)
        derived = g["ou_total"] or 44.5
        if period == "1H":
            derived = half_line(derived)
        line = c3.number_input(
            "Total" if period == "FG" else "1H total",
            value=float(saved if saved is not None else derived), step=0.5,
            format="%.1f", key=f"l{k}-ou-{period}",
            help="Edit if your book's total differs from this one")
    elif ptype == "Spread":
        opts = [g["away"], g["home"]]
        default = opts.index(a["pick_team"]) if a["pick_team"] in opts else 0
        team = c2.radio("Team", opts, key=f"m{k}", horizontal=True,
                        index=default, label_visibility="collapsed")
        saved = (a["pick_line"] if a["pick_type"] == "Spread"
                 and a["pick_team"] == team
                 and (a.get("period") or "FG") == period
                 and a["pick_line"] is not None else None)
        if g["spread"] is None or g["favorite"] is None:
            derived = 0.0
        else:
            derived = -g["spread"] if team == g["favorite"] else g["spread"]
            if period == "1H":
                derived = half_line(derived)
        line = c3.number_input(
            "Line" if period == "FG" else "1H line",
            value=float(saved if saved is not None else derived),
            step=0.5, format="%.1f", key=f"l{k}-{team}-{period}",
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
            fields = rules.build_pick(g, ptype, team, ou, line, period)
            with engine().begin() as conn:
                conn.execute(assignments.update()
                             .where(assignments.c.id == a["id"])
                             .values(submitted_at=datetime.utcnow(),
                                     result=None, **fields))
            return True
        except ValueError as e:
            st.error(str(e))
    if a["pick_selection"] and c5.button("Clear", key=f"x{k}"):
        with engine().begin() as conn:
            conn.execute(assignments.update()
                         .where(assignments.c.id == a["id"])
                         .values(pick_type=None, pick_selection=None,
                                 pick_team=None, pick_line=None, fav_dog=None,
                                 home_away=None, submitted_at=None, result=None,
                                 period=None))
        return True
    return False


def page_my_picks(user):
    week = current_week()
    if not week:
        st.info("No week is set up yet. Pester the Commissioner.")
        return
    st.subheader(f"Week {week['week_num']} — Your Assignments")
    dl, past_due = deadline_state(week)
    lock, past_lock = lock_state(week)
    locked = ((past_lock or week["status"] == "final")
              and not user["is_commissioner"])
    if past_lock:
        st.error(f"🔒 Picks locked at kickoff ({lock_label(week)}). "
                 "Changes now require the Commissioner (Charter §XIII).")
    elif past_due and dl:
        st.warning(f"⚠️ The {dl:%-I:%M %p} deadline has passed — get your "
                   f"picks in. They stay open until kickoff "
                   f"**{lock_label(week)}** ({countdown(lock)} left), but "
                   "you're on the Commissioner's list (Charter §XIII).")
    elif dl:
        st.info(f"⏳ Picks due **{deadline_label(week)}** "
                f"({countdown(dl)} to go). Hard lock is kickoff — "
                f"{lock_label(week)} — and you can change your pick any "
                "time before that.")

    st.caption("Spreads and totals, full game **or first half** (coop vote, "
               "Sept 2026 — first-half bets settle at halftime). Moneyline "
               "only when the game spread is under 3 (Charter §III). Enter "
               "the line your own book gave you.")
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
        if week.get("payout"):
            slip += ["", f"{legs}-leg parlay · ${week['stake']:,.0f} "
                         f"at {week['odds']:+d} → pays ${week['payout']:,.0f}"
                         if week.get("odds") else
                         f"{legs}-leg parlay · ${week['stake']:,.0f} → pays "
                         f"${week['payout']:,.0f}", "Good luck, gentlemen. 🍀"]
        else:
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
@st.cache_data(ttl=90)
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
                   f"{datetime.fromisoformat(first_ts):%b %-d}. Everything "
                   "below is measured against **the line each of us actually "
                   "locked in** — positive edge means the market has since "
                   "moved our way and we're holding a better number than "
                   "you could get now (closing line value).")

        alist = week_assignments(week["id"])
        live = get_live(week["week_num"])
        rows_out, tot_edge, good, bad, flat = [], 0.0, 0, 0, 0
        for a in alist:
            hist = by_game.get(a["espn_id"], [])
            g_live = live.get(a["espn_id"], {})
            started = g_live.get("state", "pre") != "pre"
            n = hist[-1] if hist else None
            mkt_lbl = "—"
            edge = None
            if n is not None and a.get("period") == "1H":
                # 1st-half markets ≈ half the full-game numbers (house convention)
                n = dict(n, spread=(n["spread"] / 2 if n["spread"] else n["spread"]),
                         ou_total=(n["ou_total"] / 2 if n["ou_total"] else None))
            if a["pick_selection"] and n is not None:
                if a["pick_type"] == "Over/Under" and n["ou_total"] is not None:
                    mkt = n["ou_total"]
                    over = a["pick_selection"].startswith("Over")
                    edge = (mkt - a["pick_line"]) if over else (a["pick_line"] - mkt)
                    mkt_lbl = f"{'Over' if over else 'Under'} {mkt:g}"
                elif a["pick_type"] in ("Spread", "Moneyline") and n["spread"] is not None:
                    if n["favorite"] is None:
                        mkt = 0.0
                    elif a["pick_team"] == n["favorite"]:
                        mkt = -n["spread"]
                    else:
                        mkt = n["spread"]
                    ours = a["pick_line"] if a["pick_type"] == "Spread" else mkt
                    # for ML the number doesn't apply; use movement of the
                    # spread toward/away from our team since we locked in
                    if a["pick_type"] == "Moneyline" and hist:
                        o = hist[0]
                        o_line = (0.0 if o["favorite"] is None else
                                  -o["spread"] if a["pick_team"] == o["favorite"]
                                  else o["spread"])
                        edge = o_line - mkt
                    else:
                        edge = ours - mkt
                    mkt_lbl = f"{a['pick_team']} {mkt:+g}"
                if a.get("period") == "1H":
                    mkt_lbl += " (1H ≈ ½ game line)"
            if edge is None:
                verdict = "—"
            elif edge >= 0.5:
                verdict, good = "✅ Good — market moved our way", good + 1
            elif edge <= -0.5:
                verdict, bad = "❌ Bad — market moved against us", bad + 1
            else:
                verdict, flat = "➖ Flat", flat + 1
            if edge is not None:
                tot_edge += edge
            rows_out.append({
                "Member": a["player"],
                "Matchup": f"{a['away_abbr']} @ {a['home_abbr']}",
                "Our locked line": a["pick_selection"] or "no pick yet",
                "Closing line" if started else "Market now": mkt_lbl,
                "Edge (pts)": f"{edge:+g}" if edge is not None else "—",
                "Verdict": verdict,
            })
        picks_in = sum(1 for a in alist if a["pick_selection"])
        m1, m2, m3 = st.columns(3)
        m1.metric("Market agrees with us on", f"{good} of {picks_in} legs")
        m2.metric("Market disagrees on", f"{bad} of {picks_in} legs")
        m3.metric("Net line edge", f"{tot_edge:+g} pts",
                  "we're holding better numbers than the market"
                  if tot_edge > 0 else
                  "the market got better numbers than us" if tot_edge < 0
                  else None)
        st.dataframe(pd.DataFrame(rows_out), use_container_width=True,
                     hide_index=True)
        st.caption("Edge = our number minus what the market offers now, from "
                   "our side of the bet. +0.5 or better is a good sign; "
                   "−0.5 or worse means the money went the other way. Once a "
                   "game kicks off, the last tracked line becomes its closing "
                   "line.")

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
                line1h = half_line(g["spread"])
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
                 f"Stragglers have until kickoff ({lock_label(week)}); after "
                 "that you're locked out. Spreads & totals, full game or "
                 "1st half; ML only if the spread is under 3.")
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

    with st.expander("🎟️ The ticket — stake, odds, payout",
                     expanded=not week.get("payout")):
        st.caption("Enter the parlay once it's placed. Odds are American "
                   "(e.g. +219500). Leave payout blank to compute it from "
                   "stake × odds, or type the book's exact number.")
        t1, t2, t3 = st.columns(3)
        stake = t1.number_input("Stake ($)", min_value=0.0, step=1.0,
                                value=float(week.get("stake") or 5.0))
        odds = t2.number_input("Odds (American)", step=100,
                               value=int(week.get("odds") or 0))
        payout_in = t3.number_input("Payout if it hits ($)", min_value=0.0,
                                    step=100.0,
                                    value=float(week.get("payout") or 0.0))
        if st.button("💾 Save ticket"):
            payout = payout_in
            if not payout and odds and stake:
                mult = (odds / 100 + 1) if odds > 0 else (100 / abs(odds) + 1)
                payout = round(stake * mult, 2)
            with engine().begin() as conn:
                conn.execute(weeks.update().where(weeks.c.id == week["id"])
                             .values(stake=stake or None, odds=odds or None,
                                     payout=payout or None))
            st.rerun()
        if week.get("payout"):
            st.success(f"On file: ${week['stake']:,.2f} at {week['odds']:+d} "
                       f"→ pays ${week['payout']:,.2f}"
                       if week.get("odds") else
                       f"On file: ${week['stake']:,.2f} → pays "
                       f"${week['payout']:,.2f}")

    with st.expander("📉 Market vs. outcomes — does beating the line matter? "
                     "(commissioner only)"):
        clv.backfill(engine())
        rep = clv.report(engine(), CURRENT_SEASON)
        ov = rep["overall"]
        if not ov["Picks"]:
            st.info("Fills in as picks get graded — needs the odds tracker "
                    "running through the week (it is).")
        else:
            st.caption("Closing line value = our locked number minus where "
                       "the market closed at kickoff, from our side. The "
                       "question: when the market moves against us, do we "
                       "still hit? When it moves for us, do we cash more?")
            k1, k2, k3 = st.columns(3)
            k1.metric("Graded picks with CLV", ov["Picks"])
            k2.metric("Average CLV", f"{ov['Avg CLV']:+.2f} pts",
                      "we beat the close on average" if ov["Avg CLV"] > 0
                      else "the market beat us on average" if ov["Avg CLV"] < 0
                      else None)
            k3.metric("Record", f"{ov['W']}-{ov['L']}-{ov['P']}",
                      fmt_pct(ov["Win %"]))
            st.markdown("**Outcomes by how the market moved**")
            st.dataframe(pct_df(rep["buckets"]).assign(
                **{"Avg CLV": [f"{b['Avg CLV']:+.2f}" for b in rep["buckets"]]}),
                use_container_width=True, hide_index=True)
            c1, c2 = st.columns(2)
            c1.markdown("**Who gets the best numbers**")
            mdf = pct_df(rep["members"])
            mdf["Avg CLV (pts)"] = mdf["Avg CLV (pts)"].map(lambda v: f"{v:+.2f}")
            c1.dataframe(mdf, use_container_width=True, hide_index=True)
            c2.markdown("**Week by week**")
            wdf = pct_df(rep["weekly"])
            wdf["Net CLV (pts)"] = wdf["Net CLV (pts)"].map(lambda v: f"{v:+g}")
            c2.dataframe(wdf, use_container_width=True, hide_index=True)
            with st.expander("Every pick"):
                ddf = pd.DataFrame(rep["detail"])
                ddf["CLV"] = ddf["CLV"].map(lambda v: f"{v:+g}")
                st.dataframe(ddf, use_container_width=True, hide_index=True)

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
    week = current_week()          # grade fresh finals on any page view
    if week:
        try:
            auto_grade(week, week_assignments(week["id"]),
                       get_live(week["week_num"]))
        except Exception:
            pass
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
