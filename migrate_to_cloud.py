"""One-time copy of the local SQLite database into the cloud database.

Usage:
  python3 migrate_to_cloud.py "postgresql://...supabase connection string..."

Copies players (with PINs), all weeks, games, picks, and odds snapshots,
preserving ids so every relationship survives. Refuses to run against a
cloud DB that already has picks (so it can't clobber live data). On
success it writes the URL to database_url.txt, which flips the local app,
weekly_update, and snapshot_odds over to the cloud automatically.
"""
from __future__ import annotations
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from sqlalchemy import create_engine, func, select

from gwc import db


def main(url):
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    src = create_engine(f"sqlite:///{db.APP_DIR / 'gwc.db'}")
    dst = create_engine(url, pool_pre_ping=True)
    db.metadata.create_all(dst)

    with dst.connect() as conn:
        n = conn.execute(select(func.count()).select_from(db.assignments)).scalar()
        if n:
            sys.exit(f"Cloud DB already has {n} picks — refusing to overwrite. "
                     "Wipe it in Supabase first if you really mean to re-migrate.")

    tables = [db.players, db.weeks, db.games, db.assignments,
              db.odds_snapshots]
    with src.connect() as s, dst.begin() as d:
        for t in tables:
            data = [dict(r._mapping) for r in s.execute(select(t))]
            if data:
                d.execute(t.insert(), data)
            print(f"{t.name}: {len(data)} rows copied")
        # postgres sequences must catch up to the copied explicit ids
        if dst.dialect.name == "postgresql":
            from sqlalchemy import text
            for t in tables:
                d.execute(text(
                    f"SELECT setval(pg_get_serial_sequence('{t.name}', 'id'), "
                    f"COALESCE((SELECT MAX(id) FROM {t.name}), 1))"))

    (db.APP_DIR / "database_url.txt").write_text(url + "\n")
    print("\ndatabase_url.txt written — the local app, weekly_update, and "
          "snapshot_odds now all use the cloud DB.")
    print("Keep gwc.db as a frozen pre-migration backup; don't delete it.")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
