"""Standings, weekly reports, analytics, and bragging rights.

All functions take graded pick rows:
{player, week_num, result, pick_type, fav_dog, home_away, pick_selection}
Richer analytics also use day, away, home, pick_team, and pick_line when the
rows carry them (see enrich()).
"""
import re
from collections import defaultdict

DAY_ORDER = {"Wed": 0, "Thu": 1, "Fri": 2, "Sat": 3, "Sun": 4, "Mon": 5}


def enrich(picks):
    """Derive pick_team / pick_line for imported rows that only carry the
    selection string (e.g. "Bears +3.0", "Under 43.5", "Falcons ML")."""
    for p in picks:
        sel = str(p.get("pick_selection") or "")
        if sel.startswith(("Over ", "Under ")):  # fix mislabeled source rows
            p["pick_type"] = "Over/Under"
        if p.get("pick_team") is None and p["pick_type"] in ("Spread", "Moneyline"):
            for t in (p.get("away"), p.get("home")):
                if t and sel.startswith(t):
                    p["pick_team"] = t
                    break
        if p.get("pick_line") is None and p["pick_type"] in ("Spread", "Over/Under"):
            m = re.search(r"([+-]?\d+(?:\.\d+)?)\s*$", sel)
            if m:
                p["pick_line"] = float(m.group(1))
    return picks


def _record(picks):
    w = sum(1 for p in picks if p["result"] == "Win")
    l = sum(1 for p in picks if p["result"] == "Loss")
    push = sum(1 for p in picks if p["result"] == "Push")
    decided = w + l
    return {"Wins": w, "Losses": l, "Pushes": push, "Total": len(picks),
            "Win %": (w / decided) if decided else 0.0}


def standings(picks, members):
    per = defaultdict(list)
    for p in picks:
        per[p["player"]].append(p)
    table = []
    for m in members:
        rec = _record(per.get(m, []))
        weekly = defaultdict(list)
        for p in per.get(m, []):
            weekly[p["week_num"]].append(p)
        best, worst = None, None
        for wk, wp in weekly.items():
            r = _record(wp)
            if r["Wins"] + r["Losses"] == 0:
                continue
            label = f"Wk {wk}: {r['Wins']}-{r['Losses']}"
            if best is None or r["Win %"] > best[0]:
                best = (r["Win %"], label)
            if worst is None or r["Win %"] < worst[0]:
                worst = (r["Win %"], label)
        table.append({"Player": m, **rec,
                      "Best Week": best[1] if best else "-",
                      "Worst Week": worst[1] if worst else "-"})
    table.sort(key=lambda r: (-r["Win %"], -r["Wins"]))
    coop = _record(picks)
    return table, coop


def weekly_report(picks):
    """Champion(s) and runner-up per week, by win % then wins."""
    by_week = defaultdict(lambda: defaultdict(list))
    for p in picks:
        by_week[p["week_num"]][p["player"]].append(p)
    report = []
    for wk in sorted(by_week):
        scores = []
        for player, pp in by_week[wk].items():
            r = _record(pp)
            if r["Wins"] + r["Losses"] == 0:
                continue
            scores.append((r["Win %"], r["Wins"], player,
                           f"{r['Wins']}-{r['Losses']}"))
        if not scores:
            continue
        scores.sort(reverse=True)
        top = scores[0][:2]
        champs = [s for s in scores if s[:2] == top]
        rest = [s for s in scores if s[:2] != top]
        week_all = [p for pp in by_week[wk].values() for p in pp]
        report.append({
            "Week": f"Week {wk}",
            "Champion(s)": ", ".join(s[2] for s in champs),
            "Score": champs[0][3],
            "Runner-Up": rest[0][2] if rest else "-",
            "Runner-Up Score": rest[0][3] if rest else "-",
            "Coop Win %": _record(week_all)["Win %"],
        })
    return report


def analytics(picks, members):
    def slice_rec(label, subset):
        return {"Category": label, **_record(subset)}

    overview = [
        slice_rec("All Picks", picks),
        slice_rec("Spreads", [p for p in picks if p["pick_type"] == "Spread"]),
        slice_rec("Over/Under", [p for p in picks if p["pick_type"] == "Over/Under"]),
        slice_rec("Overs", [p for p in picks if p["pick_type"] == "Over/Under"
                            and str(p["pick_selection"]).startswith("Over")]),
        slice_rec("Unders", [p for p in picks if p["pick_type"] == "Over/Under"
                             and str(p["pick_selection"]).startswith("Under")]),
        slice_rec("Moneylines", [p for p in picks if p["pick_type"] == "Moneyline"]),
        slice_rec("Favorites", [p for p in picks if p["fav_dog"] == "Favorite"]),
        slice_rec("Underdogs", [p for p in picks if p["fav_dog"] == "Underdog"]),
        slice_rec("Home Teams", [p for p in picks if p["home_away"] == "Home"]),
        slice_rec("Away Teams", [p for p in picks if p["home_away"] == "Away"]),
    ]
    per_player = [
        {"Category": m, **_record([p for p in picks if p["player"] == m])}
        for m in members
    ]
    return overview, per_player


def weekly_matrix(picks, members):
    """{week_num: {member: win% or None}} for trend charts / heatmaps."""
    by_week = defaultdict(lambda: defaultdict(list))
    for p in picks:
        by_week[p["week_num"]][p["player"]].append(p)
    out = {}
    for wk in sorted(by_week):
        row = {}
        for m in members:
            r = _record(by_week[wk].get(m, []))
            row[m] = r["Win %"] if r["Wins"] + r["Losses"] else None
        out[wk] = row
    return out


def form(picks, members, window=4):
    """Rolling last-N-weeks form vs season average (the hot/cold board)."""
    weeks_present = sorted({p["week_num"] for p in picks})
    recent = set(weeks_present[-window:])
    out = []
    for m in members:
        mine = [p for p in picks if p["player"] == m]
        season = _record(mine)
        last = _record([p for p in mine if p["week_num"] in recent])
        if last["Wins"] + last["Losses"] == 0:
            continue
        diff = last["Win %"] - season["Win %"]
        status = ("🔥 HOT" if diff >= 0.10 else
                  "🧊 COLD" if diff <= -0.10 else "➡️ NEUTRAL")
        out.append({"Player": m,
                    f"Last {window}W": f"{last['Wins']}-{last['Losses']}",
                    f"Last {window}W Win %": last["Win %"],
                    "Season Win %": season["Win %"],
                    "Diff": diff, "Status": status})
    out.sort(key=lambda r: -r["Diff"])
    return out


def day_splits(picks):
    """Record by day of week — who shows up in primetime."""
    by_day = defaultdict(list)
    for p in picks:
        if p.get("day"):
            by_day[p["day"]].append(p)
    return [{"Category": d, **_record(by_day[d])}
            for d in sorted(by_day, key=lambda d: DAY_ORDER.get(d, 9))]


def spread_buckets(picks):
    """Spread picks bucketed by how many points were laid or taken."""
    buckets = [
        ("Laying 7+", lambda l: l <= -7),
        ("Laying 3.5–6.5", lambda l: -7 < l <= -3.5),
        ("Laying up to 3", lambda l: -3.5 < l < 0),
        ("Pick'em", lambda l: l == 0),
        ("Getting up to 3", lambda l: 0 < l < 3.5),
        ("Getting 3.5–6.5", lambda l: 3.5 <= l < 7),
        ("Getting 7+", lambda l: l >= 7),
    ]
    spreads = [p for p in picks
               if p["pick_type"] == "Spread" and p.get("pick_line") is not None]
    out = []
    for label, test in buckets:
        subset = [p for p in spreads if test(p["pick_line"])]
        if subset:
            out.append({"Category": label, **_record(subset)})
    return out


def team_records(picks, min_picks=1):
    """How the coop does when riding each team."""
    by_team = defaultdict(list)
    for p in picks:
        if p.get("pick_team"):
            by_team[p["pick_team"]].append(p)
    out = [{"Team": t, **_record(pp)} for t, pp in by_team.items()
           if len(pp) >= min_picks]
    out.sort(key=lambda r: (-r["Total"], -r["Win %"]))
    return out


def pick_streaks(picks, members):
    """Longest single-pick win and loss streaks per player."""
    out = []
    for m in members:
        mine = sorted((p for p in picks if p["player"] == m),
                      key=lambda p: (p["week_num"],
                                     DAY_ORDER.get(p.get("day"), 9)))
        best_w = best_l = cur_w = cur_l = 0
        for p in mine:
            if p["result"] == "Win":
                cur_w, cur_l = cur_w + 1, 0
            elif p["result"] == "Loss":
                cur_l, cur_w = cur_l + 1, 0
            else:
                continue  # pushes don't break streaks
            best_w, best_l = max(best_w, cur_w), max(best_l, cur_l)
        out.append({"Player": m, "Longest Heater": best_w,
                    "Longest Skid": best_l,
                    "Riding Now": (f"{cur_w} W" if cur_w else
                                   f"{cur_l} L" if cur_l else "—")})
    out.sort(key=lambda r: -r["Longest Heater"])
    return out


def player_matrix(picks, members):
    """One row per player: record in every situational split."""
    def rec_str(subset):
        r = _record(subset)
        if r["Wins"] + r["Losses"] == 0:
            return "—"
        return f"{r['Wins']}-{r['Losses']} ({r['Win %']:.0%})"

    out = []
    for m in members:
        mine = [p for p in picks if p["player"] == m]
        if not mine:
            continue
        out.append({
            "Player": m,
            "Overall": rec_str(mine),
            "Spreads": rec_str([p for p in mine if p["pick_type"] == "Spread"]),
            "Totals": rec_str([p for p in mine if p["pick_type"] == "Over/Under"]),
            "Moneylines": rec_str([p for p in mine if p["pick_type"] == "Moneyline"]),
            "Favorites": rec_str([p for p in mine if p["fav_dog"] == "Favorite"]),
            "Underdogs": rec_str([p for p in mine if p["fav_dog"] == "Underdog"]),
            "Home": rec_str([p for p in mine if p["home_away"] == "Home"]),
            "Away": rec_str([p for p in mine if p["home_away"] == "Away"]),
            "Primetime (Thu/Mon)": rec_str(
                [p for p in mine if p.get("day") in ("Thu", "Mon")]),
        })
    return out


def bragging_rights(picks, members):
    by_pw = defaultdict(list)
    for p in picks:
        by_pw[(p["player"], p["week_num"])].append(p)

    perfect, goose = [], []
    champ_count = defaultdict(int)
    for row in weekly_report(picks):
        for name in row["Champion(s)"].split(", "):
            champ_count[name] += 1
    for (player, wk), pp in sorted(by_pw.items(), key=lambda kv: kv[0][1]):
        r = _record(pp)
        if r["Total"] < 2 or r["Wins"] + r["Losses"] == 0:
            continue
        if r["Losses"] == 0 and r["Wins"] >= 2:
            perfect.append({"Week": f"Week {wk}", "Player": player,
                            "Record": f"{r['Wins']}-{r['Losses']}"})
        if r["Wins"] == 0 and r["Losses"] >= 2:
            goose.append({"Week": f"Week {wk}", "Player": player,
                          "Record": f"{r['Wins']}-{r['Losses']}"})

    # longest streak of winning weeks (win % > .5) per player
    streaks = {}
    for m in members:
        wks = sorted(wk for (pl, wk) in by_pw if pl == m)
        best = cur = 0
        for wk in wks:
            r = _record(by_pw[(m, wk)])
            if r["Wins"] + r["Losses"] and r["Win %"] > 0.5:
                cur += 1
                best = max(best, cur)
            else:
                cur = 0
        streaks[m] = best

    def top(d, fmt):
        if not d:
            return None
        name = max(d, key=d.get)
        return {"Winner": name, "Details": fmt.format(d[name])} if d[name] else None

    awards = []
    if champ_count:
        awards.append({"Award": "Most Championship Weeks",
                       **top(champ_count, "{} weeks")})
    pf = defaultdict(int)
    for row in perfect:
        pf[row["Player"]] += 1
    if pf:
        awards.append({"Award": "Most Perfect Weeks", **top(pf, "{} perfect week(s)")})
    ge = defaultdict(int)
    for row in goose:
        ge[row["Player"]] += 1
    if ge:
        awards.append({"Award": "Most Goose Eggs", **top(ge, "{} goose egg week(s)")})
    st = top(streaks, "{} consecutive winning weeks")
    if st:
        awards.append({"Award": "Best Win Streak", **st})
    return awards, perfect, goose
