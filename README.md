# 🎩 The Gentlemen's Wagering Cooperative — Commissioner's Dashboard

Your private weekly command center. It replaces the spreadsheet wrangling:
pull the slate, randomize assignments, copy the announcement to the group
chat, log picks as the texts roll in, copy the finished bet slip, grade
results with one click, and keep the season stats forever.

## Start it

Double-click **`GWC Dashboard.command`** (or run `streamlit run app.py` in
this folder). It opens in your browser at `http://localhost:8501`.

All data lives in `gwc.db` right here in this folder — already seeded with
the full 2025/2026 season from the Excel workbook. Nothing leaves your Mac.

## The weekly routine

Everything is on the **📋 This Week** page, top to bottom:

1. **New week** → *Fetch slate from ESPN* — grabs the games, spreads, totals,
   and sets the pick deadline (Thu 4 PM ET, or first kickoff if earlier).
2. **Slate** — untick games you want to skip; *Refresh lines* re-pulls odds.
3. **Assignments** — 🎲 *Randomize* (rotates Thu/Mon duty away from whoever
   had it recently, avoids repeat teams; re-roll freely). Copy the
   ready-made announcement into the group chat, then *Mark as announced*.
4. **Picks** — as each guy texts his picks, log them under his name. The
   charter is enforced: spreads & totals only, moneyline unlocked only when
   the spread is under 3. The header shows who you're still waiting on.
   *Move a game* handles pick switching (Charter IV).
5. **Bet slip** — appears as picks come in; complete slip is ready to
   copy-paste once everyone's in.
6. **Results** — after the games, *Fetch scores & grade* pulls final scores
   from ESPN and grades every pick (pushes included). *Mark week FINAL* and
   the standings, analytics, and bragging rights update themselves.

**🏆 Standings** and **📈 Analytics** cover season records, weekly champions,
splits (fav/dog, home/away, overs/unders), perfect weeks, and goose eggs —
filterable by season or all time.

## Housekeeping

- Re-import history (wipes nothing, skips seasons already present):
  `python3 seed_history.py "../2025:2026/Gentlemens Wagering Cooperative.xlsx"`
- Next season: bump `CURRENT_SEASON` in `gwc/db.py` and `SEASON_YEAR` in `app.py`.
- Back up `gwc.db` occasionally — it's the whole record book.
- If you ever want it online after all (e.g. to run it from your phone),
  it deploys to Streamlit Cloud like your other apps; ask Claude to walk
  you through it.
