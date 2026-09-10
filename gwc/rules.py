"""Charter rules: allowed pick types, pick construction, and grading."""
from __future__ import annotations

MONEYLINE_MAX_SPREAD = 3.0  # charter III: moneyline only if spread < 3


def allowed_pick_types(game: dict) -> list[str]:
    types = ["Spread", "Over/Under"]
    spread = game.get("spread")
    if spread is not None and spread < MONEYLINE_MAX_SPREAD:
        types.append("Moneyline")
    return types


def build_pick(game: dict, pick_type: str, team: str | None,
               over_under: str | None, line: float | None = None) -> dict:
    """Return the assignment fields for a validated pick.

    `line` overrides the game's ESPN line — different sportsbooks post
    different numbers, so picks are logged at the line actually bet.
    For spreads it's signed relative to the picked team (e.g. -2.5 means
    the picked team is laying 2.5); for totals it's the total itself.
    """
    if pick_type not in allowed_pick_types(game):
        raise ValueError(f"{pick_type} is not allowed for this game "
                         f"(spread is {game.get('spread')}).")
    fields = {"pick_type": pick_type, "pick_team": None, "pick_line": None,
              "fav_dog": None, "home_away": None}

    if pick_type == "Over/Under":
        total = line if line is not None else game.get("ou_total")
        if total is None:
            raise ValueError("No total is set for this game.")
        if over_under not in ("Over", "Under"):
            raise ValueError("Choose Over or Under.")
        fields["pick_line"] = float(total)
        fields["pick_selection"] = f"{over_under} {total:g}"
        return fields

    if team not in (game["away"], game["home"]):
        raise ValueError("Pick one of the two teams in this game.")
    fields["pick_team"] = team
    fields["home_away"] = "Home" if team == game["home"] else "Away"

    if pick_type == "Moneyline":
        fields["fav_dog"] = ("Favorite" if team == game.get("favorite")
                             else "Underdog")
        fields["pick_selection"] = f"{team} ML"
        return fields

    if line is not None:
        signed = float(line)
    else:
        spread = game.get("spread")
        if spread is None:
            raise ValueError("No spread is set for this game.")
        if game.get("favorite") is None:  # pick'em
            signed = 0.0
        elif team == game["favorite"]:
            signed = -spread
        else:
            signed = spread
    fields["fav_dog"] = ("Favorite" if signed < 0
                         else "Underdog" if signed > 0 else "Pick'em")
    fields["pick_line"] = signed
    fields["pick_selection"] = f"{team} {signed:+g}"
    return fields


def grade(game: dict, pick: dict) -> str | None:
    """Return Win / Loss / Push, or None if the game isn't final."""
    if not game.get("final") or game.get("home_score") is None:
        return None
    home_s, away_s = game["home_score"], game["away_score"]

    if pick["pick_type"] == "Over/Under":
        total = home_s + away_s
        line = pick["pick_line"]
        if total == line:
            return "Push"
        went_over = total > line
        picked_over = pick["pick_selection"].startswith("Over")
        return "Win" if went_over == picked_over else "Loss"

    picked_home = pick["pick_team"] == game["home"]
    my_score = home_s if picked_home else away_s
    their_score = away_s if picked_home else home_s

    if pick["pick_type"] == "Moneyline":
        if my_score == their_score:
            return "Push"
        return "Win" if my_score > their_score else "Loss"

    adjusted = my_score + (pick["pick_line"] or 0.0)
    if adjusted == their_score:
        return "Push"
    return "Win" if adjusted > their_score else "Loss"
