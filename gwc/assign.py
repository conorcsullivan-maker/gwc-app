"""Randomized slate assignment with the Commissioner's fairness rules:

- every included game is assigned to exactly one member
- each member gets an even share (2-3 games in a typical week)
- premium slots (Thu / Mon / Sat) rotate away from whoever had them recently
- avoid handing someone a team they were assigned in recent weeks
"""
from __future__ import annotations
import random
from collections import defaultdict

PREMIUM_DAYS = {"Thu", "Mon", "Sat", "Fri"}
LOOKBACK_WEEKS_DAYS = 4   # rotation window for premium day slots
LOOKBACK_WEEKS_TEAMS = 2  # window for repeat-team avoidance


def build_history(past: list[dict]) -> dict:
    """`past` rows: {week_num, player, day, away, home} from recent weeks."""
    if not past:
        return {"day_counts": {}, "recent_teams": {}}
    latest = max(r["week_num"] for r in past)
    day_counts, recent_teams = defaultdict(int), defaultdict(set)
    for r in past:
        age = latest - r["week_num"]
        if r["day"] in PREMIUM_DAYS and age < LOOKBACK_WEEKS_DAYS:
            day_counts[(r["player"], r["day"])] += 1
        if age < LOOKBACK_WEEKS_TEAMS:
            recent_teams[r["player"]].update({r["away"], r["home"]})
    return {"day_counts": dict(day_counts), "recent_teams": dict(recent_teams)}


def randomize(games: list[dict], members: list[str], history: dict | None = None,
              seed: int | None = None) -> dict[int, str]:
    """Return {game_id: member}. Games need id, day, away, home keys."""
    rng = random.Random(seed)
    history = history or {"day_counts": {}, "recent_teams": {}}
    n, m = len(games), len(members)
    base, extra = divmod(n, m)
    # who gets the extra game rotates randomly each week
    lucky = rng.sample(members, extra)
    quota = {p: base + (1 if p in lucky else 0) for p in members}

    def weight(player, game):
        w = 1.0
        if game["day"] in PREMIUM_DAYS:
            w /= 1.0 + 2.0 * history["day_counts"].get((player, game["day"]), 0)
        teams = history["recent_teams"].get(player, set())
        if game["away"] in teams or game["home"] in teams:
            w *= 0.25
        return w

    # premium games first so the rotation weighting gets first choice of people
    ordered = sorted(games, key=lambda g: (g["day"] not in PREMIUM_DAYS, rng.random()))
    result = {}
    for game in ordered:
        pool = [p for p in members if quota[p] > 0]
        weights = [weight(p, game) for p in pool]
        choice = rng.choices(pool, weights=weights, k=1)[0]
        result[game["id"]] = choice
        quota[choice] -= 1
    return result
