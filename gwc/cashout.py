"""Live parlay valuation — a fair-value cash-out estimate.

How books do it (in short): the ticket is worth
    payout_if_it_hits × P(every remaining leg wins)
and the cash-out offer is that fair value minus a margin (typically 5–10%).

Per-leg win probability for a game in progress uses the standard model:
the final margin (relative to the line) is ~Normal(current margin,
sigma·sqrt(fraction of game remaining)), with sigma ≈ 13.5 pts for NFL
spreads/moneylines and ≈ 10 pts for totals. Legs not yet started are a
vig-free coin flip (0.50) for spreads and totals.
"""
from __future__ import annotations
import math
import re

SIGMA_MARGIN = 13.5      # full-game NFL point-margin std dev
SIGMA_TOTAL = 10.0       # full-game NFL total-points std dev
PUSH_LEG_FACTOR = 1.909  # a pushed leg at -110 drops out of the multiplier


def _phi(z):
    return 0.5 * (1.0 + math.erf(z / math.sqrt(2.0)))


def fraction_remaining(detail: str, period: str = "FG") -> float:
    """Fraction of the (game or first half) still to be played, from
    ESPN's shortDetail like '8:32 - 3rd', 'Halftime', 'End of 1st', 'OT'."""
    d = (detail or "").strip()
    q_left = {"1st": 3, "2nd": 2, "3rd": 1, "4th": 0}
    m = re.search(r"(\d+):(\d+)\s*-\s*(1st|2nd|3rd|4th|OT)", d)
    if m:
        mins = int(m.group(1)) + int(m.group(2)) / 60.0
        q = m.group(3)
        if q == "OT":
            game_left = 0.03
            half_left = 0.0
        else:
            game_left = (q_left[q] * 15 + mins) / 60.0
            half_left = ((q_left[q] - 2) * 15 + mins) / 30.0 if q in ("1st", "2nd") else 0.0
    elif "Halftime" in d:
        game_left, half_left = 0.5, 0.0
    elif "End of 1st" in d:
        game_left, half_left = 0.75, 0.5
    elif "End of 3rd" in d:
        game_left, half_left = 0.25, 0.0
    elif "OT" in d:
        game_left, half_left = 0.03, 0.0
    else:
        game_left, half_left = 0.5, 0.25   # unknown: assume midway
    return max(0.0, min(1.0, half_left if period == "1H" else game_left))


def leg_win_prob(a: dict, g: dict | None) -> float:
    """P(this pick wins) right now. `a` is an assignment row joined with
    its game; `g` is the live ESPN status dict (or None before kickoff)."""
    if a.get("result") == "Win":
        return 1.0
    if a.get("result") == "Loss":
        return 0.0
    if a.get("result") == "Push":
        return 1.0                      # handled via payout adjustment
    if not a.get("pick_selection"):
        return 0.0
    state = (g or {}).get("state", "pre")
    if state == "pre" or not g:
        return 0.5 if a["pick_type"] != "Moneyline" else _ml_prior(a)
    period = a.get("period") or "FG"
    hs, as_ = g["home_score"], g["away_score"]
    if a["pick_type"] == "Over/Under":
        margin = (hs + as_) - (a["pick_line"] or 0)
        if a["pick_selection"].startswith("Under"):
            margin = -margin
        sigma = SIGMA_TOTAL
    else:
        my = hs if a["pick_team"] == a["home"] else as_
        opp = as_ if a["pick_team"] == a["home"] else hs
        line = a["pick_line"] if a["pick_type"] == "Spread" else 0.0
        margin = my + (line or 0) - opp
        sigma = SIGMA_MARGIN
    f = fraction_remaining(g.get("detail", ""), period)
    if period == "1H":
        sigma = sigma / math.sqrt(2)       # a half has half the variance
    if f <= 0.0:
        return 1.0 if margin > 0 else 0.0 if margin < 0 else 0.5
    if state == "post":
        return 1.0 if margin > 0 else 0.0 if margin < 0 else 0.5
    return _phi(margin / (sigma * math.sqrt(f)))


def _ml_prior(a):
    """Pre-game moneyline win probability from the game spread: the
    favorite by s points wins about Φ(s / 13.5) of the time."""
    spread = a.get("spread")
    if not spread or not a.get("favorite"):
        return 0.5
    p_fav = _phi(spread / SIGMA_MARGIN)
    return p_fav if a.get("pick_team") == a.get("favorite") else 1 - p_fav


def valuation(week: dict, alist: list[dict], live: dict,
              margin: float = 0.08) -> dict:
    """Fair value and estimated book offer for the ticket right now."""
    payout = week.get("payout") or 0.0
    stake = week.get("stake") or 0.0
    legs = []
    prob = 1.0
    pushes = 0
    dead = False
    for a in alist:
        g = live.get(a.get("espn_id"))
        p = leg_win_prob(a, g)
        if a.get("result") == "Push":
            pushes += 1
        if a.get("result") == "Loss":
            dead = True
        legs.append({"member": a["player"], "pick": a.get("pick_selection") or "—",
                     "p": p, "status": a.get("result") or
                     ((g or {}).get("state", "pre"))})
        prob *= p
    adj_payout = payout / (PUSH_LEG_FACTOR ** pushes) if pushes else payout
    fair = 0.0 if dead else adj_payout * prob
    offer = fair * (1 - margin)
    return {"stake": stake, "payout": adj_payout, "prob": 0.0 if dead else prob,
            "fair": fair, "offer": offer, "dead": dead, "pushes": pushes,
            "legs": legs}
