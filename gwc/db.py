"""Database layer for the Gentlemen's Wagering Cooperative.

Uses SQLite locally by default. Set a DATABASE_URL secret (e.g. a Supabase
Postgres connection string) in Streamlit Cloud for durable hosted storage.
"""
import os
from pathlib import Path

from sqlalchemy import (
    Boolean, Column, DateTime, Float, ForeignKey, Integer, MetaData,
    String, Table, create_engine, select, text,
)

APP_DIR = Path(__file__).resolve().parent.parent

metadata = MetaData()

players = Table(
    "players", metadata,
    Column("id", Integer, primary_key=True),
    Column("name", String(40), nullable=False, unique=True),
    Column("pin", String(10), nullable=False, default="1111"),
    Column("is_commissioner", Boolean, nullable=False, default=False),
    Column("active", Boolean, nullable=False, default=True),
)

weeks = Table(
    "weeks", metadata,
    Column("id", Integer, primary_key=True),
    Column("season", String(20), nullable=False),
    Column("week_num", Integer, nullable=False),
    # draft -> published -> final
    Column("status", String(12), nullable=False, default="draft"),
    Column("deadline_et", String(30)),  # ISO string, America/New_York
)

games = Table(
    "games", metadata,
    Column("id", Integer, primary_key=True),
    Column("week_id", Integer, ForeignKey("weeks.id"), nullable=False),
    Column("espn_id", String(20)),
    Column("day", String(5)),            # Thu / Fri / Sat / Sun / Mon
    Column("kickoff_et", String(30)),    # ISO string, America/New_York
    Column("away", String(30), nullable=False),   # nickname, e.g. "Bears"
    Column("home", String(30), nullable=False),
    Column("away_abbr", String(5)),
    Column("home_abbr", String(5)),
    Column("favorite", String(30)),      # nickname of favorite, None if pick'em
    Column("spread", Float),             # positive number, e.g. 3.5
    Column("ou_total", Float),
    Column("away_score", Integer),
    Column("home_score", Integer),
    Column("final", Boolean, nullable=False, default=False),
    Column("excluded", Boolean, nullable=False, default=False),
)

assignments = Table(
    "assignments", metadata,
    Column("id", Integer, primary_key=True),
    Column("week_id", Integer, ForeignKey("weeks.id"), nullable=False),
    Column("game_id", Integer, ForeignKey("games.id"), nullable=False),
    Column("player_id", Integer, ForeignKey("players.id"), nullable=False),
    Column("pick_type", String(12)),     # Spread / Over/Under / Moneyline
    Column("pick_selection", String(60)),  # e.g. "Bears +3.0", "Under 43.5"
    Column("pick_team", String(30)),     # nickname; None for totals
    Column("pick_line", Float),
    Column("fav_dog", String(10)),       # Favorite / Underdog
    Column("home_away", String(5)),      # Home / Away
    Column("submitted_at", DateTime),
    Column("result", String(6)),         # Win / Loss / Push
)

odds_snapshots = Table(
    "odds_snapshots", metadata,
    Column("id", Integer, primary_key=True),
    Column("ts", String(30), nullable=False),      # ISO, America/New_York
    Column("season_year", Integer, nullable=False),
    Column("week_num", Integer, nullable=False),
    Column("espn_id", String(20), nullable=False),
    Column("away_abbr", String(5)),
    Column("home_abbr", String(5)),
    Column("favorite", String(30)),
    Column("spread", Float),
    Column("ou_total", Float),
)

MEMBERS = ["Conor", "Parker", "Grundy", "JR", "Jimmy", "Spencer"]
COMMISSIONER = "Conor"
CURRENT_SEASON = "2026 / 2027"
SEASON_YEAR = 2026            # ESPN season year for CURRENT_SEASON


def get_engine():
    """Resolve the database, in priority order:
    1. DATABASE_URL env var
    2. Streamlit secrets (cloud deployment)
    3. database_url.txt next to the app (flips ALL local scripts — the
       Streamlit app, weekly_update, snapshot_odds — to the cloud DB at once)
    4. local SQLite file (the pre-cloud default)
    """
    url = os.environ.get("DATABASE_URL")
    if not url:
        try:  # Streamlit secrets, when running in the app
            import streamlit as st
            url = st.secrets.get("DATABASE_URL", None)
        except Exception:
            url = None
    if not url:
        f = APP_DIR / "database_url.txt"
        if f.exists() and f.read_text().strip():
            url = f.read_text().strip()
    if not url:
        url = f"sqlite:///{APP_DIR / 'gwc.db'}"
    if url.startswith("postgres://"):     # Supabase/Heroku style -> SQLAlchemy
        url = url.replace("postgres://", "postgresql://", 1)
    return create_engine(url, pool_pre_ping=True)


_engine = None


def engine():
    global _engine
    if _engine is None:
        _engine = get_engine()
        init_db(_engine)
    return _engine


def init_db(eng):
    metadata.create_all(eng)
    with eng.begin() as conn:
        existing = {r.name for r in conn.execute(select(players.c.name))}
        for name in MEMBERS:
            if name not in existing:
                conn.execute(players.insert().values(
                    name=name, pin="1111",
                    is_commissioner=(name == COMMISSIONER), active=True,
                ))


def rows(conn, stmt):
    return [dict(r._mapping) for r in conn.execute(stmt)]
