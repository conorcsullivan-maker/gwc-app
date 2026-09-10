"""Snapshot current NFL betting lines into the database every few hours.

Runs via launchd (com.gwc.odds). Each run stores one timestamped row per
game for the current and next week, building the line-movement history the
coop wants ("how has this spread moved over the last 5 days?").
History only exists from the moment collection starts — keep this running.
"""
from __future__ import annotations
import sys
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parent))

from gwc import schedule
from gwc.db import SEASON_YEAR, engine, odds_snapshots

ET = ZoneInfo("America/New_York")


def current_week(now):
    for w in range(1, 19):
        try:
            glist = schedule.fetch_week(SEASON_YEAR, w)
        except Exception:
            return None
        if not glist:
            continue
        last_ko = max(datetime.fromisoformat(g["kickoff_et"]) for g in glist)
        if last_ko > now - timedelta(hours=36):
            return w
    return None


def main():
    now = datetime.now(ET)
    w = current_week(now)
    if not w:
        print(f"[{now:%Y-%m-%d %H:%M}] no active week (offseason?)")
        return
    ts = now.isoformat(timespec="minutes")
    stored = 0
    with engine().begin() as conn:
        for week in (w, w + 1):
            if week > 18:
                continue
            try:
                glist = schedule.fetch_week(SEASON_YEAR, week)
            except Exception as e:
                print(f"[{ts}] week {week} fetch failed: {e}")
                continue
            for g in glist:
                if g.get("final"):
                    continue          # line is dead once the game ends
                conn.execute(odds_snapshots.insert().values(
                    ts=ts, season_year=SEASON_YEAR, week_num=week,
                    espn_id=g["espn_id"],
                    away_abbr=g["away_abbr"], home_abbr=g["home_abbr"],
                    favorite=g["favorite"], spread=g["spread"],
                    ou_total=g["ou_total"],
                ))
                stored += 1
    print(f"[{ts}] stored {stored} line snapshots (weeks {w}-{min(w + 1, 18)})")


if __name__ == "__main__":
    main()
