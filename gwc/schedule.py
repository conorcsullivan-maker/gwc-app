"""Fetch NFL schedules, odds, and scores from ESPN's public scoreboard API."""
from __future__ import annotations
from datetime import datetime
from zoneinfo import ZoneInfo

import requests

ET = ZoneInfo("America/New_York")
SCOREBOARD = "https://site.api.espn.com/apis/site/v2/sports/football/nfl/scoreboard"


def fetch_week(year: int, week: int) -> list[dict]:
    """Return one dict per game for a regular-season week.

    `year` is the season year (2026 for the 2026/2027 season).
    """
    resp = requests.get(
        SCOREBOARD,
        params={"seasontype": 2, "week": week, "dates": year},
        timeout=20,
    )
    resp.raise_for_status()
    out = []
    for event in resp.json().get("events", []):
        comp = event["competitions"][0]
        teams = {c["homeAway"]: c for c in comp["competitors"]}
        home_t, away_t = teams["home"]["team"], teams["away"]["team"]
        kickoff = datetime.strptime(event["date"], "%Y-%m-%dT%H:%MZ")
        kickoff = kickoff.replace(tzinfo=ZoneInfo("UTC")).astimezone(ET)

        favorite, spread, ou_total = None, None, None
        odds_list = comp.get("odds") or []
        if odds_list:
            odds = odds_list[0]
            ou_total = odds.get("overUnder")
            details = odds.get("details") or ""  # e.g. "SEA -3.5" or "EVEN"
            parts = details.split()
            if len(parts) == 2:
                abbr, line = parts
                try:
                    spread = abs(float(line))
                    favorite = (home_t["name"] if abbr == home_t["abbreviation"]
                                else away_t["name"])
                except ValueError:
                    pass
            if spread == 0:
                favorite = None

        status = comp.get("status", {}).get("type", {})
        scores = {}
        if status.get("completed"):
            scores = {
                "home_score": int(teams["home"].get("score", 0)),
                "away_score": int(teams["away"].get("score", 0)),
            }

        out.append({
            "espn_id": event["id"],
            "kickoff_et": kickoff.isoformat(),
            "day": kickoff.strftime("%a"),
            "away": away_t["name"], "home": home_t["name"],
            "away_abbr": away_t["abbreviation"], "home_abbr": home_t["abbreviation"],
            "favorite": favorite, "spread": spread, "ou_total": ou_total,
            "final": bool(status.get("completed")),
            **scores,
        })
    out.sort(key=lambda g: g["kickoff_et"])
    return out


def fetch_week_live(year: int, week: int) -> dict[str, dict]:
    """Live status per ESPN game id: state (pre/in/post), detail, scores."""
    resp = requests.get(
        SCOREBOARD,
        params={"seasontype": 2, "week": week, "dates": year},
        timeout=20,
    )
    resp.raise_for_status()
    out = {}
    for event in resp.json().get("events", []):
        comp = event["competitions"][0]
        teams = {c["homeAway"]: c for c in comp["competitors"]}
        status = comp.get("status", {})
        stype = status.get("type", {})
        out[event["id"]] = {
            "state": stype.get("state", "pre"),          # pre / in / post
            "detail": stype.get("shortDetail", ""),
            "home_score": int(teams["home"].get("score") or 0),
            "away_score": int(teams["away"].get("score") or 0),
            "final": bool(stype.get("completed")),
        }
    return out


SUMMARY = "https://site.api.espn.com/apis/site/v2/sports/football/nfl/summary"


def fetch_halftime(espn_id: str) -> dict | None:
    """First-half score {away_1h, home_1h} once a game reaches halftime."""
    resp = requests.get(SUMMARY, params={"event": espn_id}, timeout=20)
    resp.raise_for_status()
    comp = resp.json().get("header", {}).get("competitions", [{}])[0]
    out = {}
    for c in comp.get("competitors", []):
        ls = c.get("linescores") or []
        if len(ls) < 2:
            return None                      # halftime not reached
        try:
            half = sum(int(float(q.get("displayValue", 0))) for q in ls[:2])
        except (TypeError, ValueError):
            return None
        out[f"{c['homeAway']}_1h"] = half
    return out if len(out) == 2 else None


def default_deadline(games: list[dict]) -> str | None:
    """Thursday 4:00 PM ET (charter section V), but never after first kickoff."""
    if not games:
        return None
    first_ko = min(datetime.fromisoformat(g["kickoff_et"]) for g in games)
    deadline = None
    for g in games:
        ko = datetime.fromisoformat(g["kickoff_et"])
        if ko.strftime("%a") == "Thu":
            deadline = ko.replace(hour=16, minute=0, second=0, microsecond=0)
            break
    if deadline is None:
        deadline = first_ko.replace(hour=16, minute=0, second=0, microsecond=0)
    return min(deadline, first_ko).isoformat()
