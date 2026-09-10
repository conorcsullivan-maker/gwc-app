"""Rebuild the stats tabs as live, season-switchable pages.

Usage: python3 build_stats_tabs.py "path/to/workbook.xlsx"

Replaces Player Dashboard, Pick Analytics, Weekly Report, Bragging Rights,
and Historical Standings with formula-driven versions computed off the raw
Games table. Each tab gets its own season dropdown (cell H2) — including
"2026 / 2027" — so you stay on a tab and toggle seasons. Hidden helper
matrices live in columns AA:AS of each sheet. A timestamped backup is saved.
"""
from __future__ import annotations
import shutil
import sys
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from build_dashboard import (BASE, BLUE, CRIT, G, GOLD, GOOD, GRID, INK,
                             INK2, MEMBERS, MUTED, N, NEUTRAL, RED, SEASONS,
                             TILE, WEEKS)

STATS_DEFAULT = "2025 / 2026"     # what the coop should see today
SEL = "$H$2"                      # per-sheet season selector
SCRIT = "$AA$1"                   # per-sheet criterion cell

thin = Side(style="thin", color=GRID)
rule = Side(style="thin", color=BASE)


def mk(ws):
    def sty(cell, *, size=10, bold=False, color=INK, fill=None, align="left",
            fmt=None):
        c = ws[cell]
        c.font = Font(name="Arial", size=size, bold=bold, color=color)
        c.alignment = Alignment(horizontal=align, vertical="center")
        if fill:
            c.fill = PatternFill("solid", fgColor=fill)
        if fmt:
            c.number_format = fmt

    def underline(rng):
        for row in ws[rng]:
            for c in row:
                c.border = Border(bottom=rule)

    def line(rng):
        for row in ws[rng]:
            for c in row:
                c.border = Border(bottom=thin)

    return sty, underline, line


def mcol(w):
    return get_column_letter(27 + w)      # week 1 -> AB ... week 18 -> AS


def header(ws, sty, title, caption, default=STATS_DEFAULT, width=10):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 14
    for col in "BCDEFGHIJK":
        ws.column_dimensions[col].width = width
    ws.row_dimensions[1].height = 26
    ws["A1"] = title
    sty("A1", size=14, bold=True)
    ws["A2"] = caption
    sty("A2", size=8.5, color=INK2)
    ws["G2"] = "SEASON →"
    sty("G2", size=8, bold=True, color=MUTED, align="right")
    ws["H2"] = default
    sty("H2", size=10, bold=True, color="0000FF", fill="FFFFFF",
        align="center")
    for row in ws["H2:I2"]:
        for c in row:
            c.border = Border(top=Side(style="thin", color=GOLD),
                              bottom=Side(style="thin", color=GOLD),
                              left=Side(style="thin", color=GOLD),
                              right=Side(style="thin", color=GOLD))
    ws.merge_cells("H2:I2")
    dv = DataValidation(type="list", formula1='"' + ",".join(SEASONS) + '"',
                        allow_blank=False, showDropDown=False)
    ws.add_data_validation(dv)
    dv.add("H2")
    ws["AA1"] = f'=IF({SEL}="All seasons","*",{SEL})'
    for i in range(27, 46):                     # hide AA..AS
        ws.column_dimensions[get_column_letter(i)].hidden = True


def matrix_block(ws):
    """Hidden per-week engine in AA..AS keyed to this sheet's selector.

    rows: 3 week#s · 4-9 wins · 11-16 losses · 18-23 decided · 25-30 pct ·
    32-35 coop w/l/pct/max · 37 last week · 39-44 l4 wins · 46-51 l4 losses
    """
    for w in range(1, WEEKS + 1):
        ws[f"{mcol(w)}3"] = w
    for i, m in enumerate(MEMBERS):
        for base in (4, 11, 18, 25, 39, 46):
            ws[f"AA{base + i}"] = m
        for w in range(1, WEEKS + 1):
            c = mcol(w)
            rw, rl, rd, rp = 4 + i, 11 + i, 18 + i, 25 + i
            ws[f"{c}{rw}"] = (f'=COUNTIFS({G["person"]},$AA{rw},{G["week"]},'
                              f'{c}$3,{G["result"]},"Win",{G["season"]},{SCRIT})')
            ws[f"{c}{rl}"] = (f'=COUNTIFS({G["person"]},$AA{rl},{G["week"]},'
                              f'{c}$3,{G["result"]},"Loss",{G["season"]},{SCRIT})')
            ws[f"{c}{rd}"] = f"={c}{rw}+{c}{rl}"
            ws[f"{c}{rp}"] = f'=IF({c}{rd}=0,"",{c}{rw}/{c}{rd})'
            ws[f"{c}{39 + i}"] = f"=IF({c}$3>$AB$37-4,{c}{rw},0)"
            ws[f"{c}{46 + i}"] = f"=IF({c}$3>$AB$37-4,{c}{rl},0)"
    for w in range(1, WEEKS + 1):
        c = mcol(w)
        ws[f"{c}32"] = f"=SUM({c}4:{c}9)"
        ws[f"{c}33"] = f"=SUM({c}11:{c}16)"
        ws[f"{c}34"] = f'=IF({c}32+{c}33=0,"",{c}32/({c}32+{c}33))'
        ws[f"{c}35"] = f"=MAX({c}25:{c}30)"
    ws["AA37"] = "last week with picks"
    ws["AB37"] = "=SUMPRODUCT(MAX(((AB32:AS32+AB33:AS33)>0)*AB3:AS3))"


def wl_pct(w_formula, l_formula):
    """'11-8 (58%)' composite from two COUNTIFS bodies."""
    return (f'=IF(({w_formula})+({l_formula})=0,"—",({w_formula})&"-"&'
            f'({l_formula})&" ("&TEXT(({w_formula})/(({w_formula})+'
            f'({l_formula})),"0%")&")")')


def cnt(cond, result):
    return f'COUNTIFS({cond},{G["result"]},"{result}",{G["season"]},{SCRIT})'


def person_cnt(name_cell, result, extra=""):
    return (f'COUNTIFS({G["person"]},{name_cell},{G["result"]},"{result}"'
            f'{extra},{G["season"]},{SCRIT})')


# ------------------------------------------------------------ player dashboard
def build_player_dashboard(ws):
    sty, underline, line = mk(ws)
    header(ws, sty, "🏈 PLAYER DASHBOARD",
           "Standings and week-by-week form — pick a season top right, "
           "everything recomputes.")
    matrix_block(ws)

    ws["A4"] = "SEASON STANDINGS"
    sty("A4", size=11, bold=True)
    underline("A4:K4")
    heads = ["Player", "W", "P", "L", "Total", "Win %", "", "Best Week",
             "", "Worst Week"]
    for j, h in enumerate(heads):
        c = get_column_letter(1 + j)
        ws[f"{c}5"] = h
        sty(f"{c}5", size=8, bold=True, color=MUTED,
            align="left" if j == 0 else "center")
    ws.merge_cells("H5:I5")
    ws.merge_cells("J5:K5")
    for i, m in enumerate(MEMBERS):
        r = 6 + i
        rng = f"AB25:AS25".replace("25", str(25 + i))
        wrow, lrow = f"AB{4 + i}:AS{4 + i}", f"AB{11 + i}:AS{11 + i}"
        ws[f"A{r}"] = m
        sty(f"A{r}", bold=True)
        ws[f"B{r}"] = f"={person_cnt(f'$A{r}', 'Win')}"
        ws[f"C{r}"] = f"={person_cnt(f'$A{r}', 'Push')}"
        ws[f"D{r}"] = f"={person_cnt(f'$A{r}', 'Loss')}"
        ws[f"E{r}"] = f"=B{r}+C{r}+D{r}"
        ws[f"F{r}"] = f"=IF(B{r}+D{r}=0,0,B{r}/(B{r}+D{r}))"
        m_ = f"MATCH(MAX({rng}),{rng},0)"
        ws.merge_cells(f"H{r}:I{r}")
        ws[f"H{r}"] = (f'=IF(COUNT({rng})=0,"—","Wk "&{m_}&": "'
                       f'&INDEX({wrow},{m_})&"-"&INDEX({lrow},{m_}))')
        m2 = f"MATCH(MIN({rng}),{rng},0)"
        ws.merge_cells(f"J{r}:K{r}")
        ws[f"J{r}"] = (f'=IF(COUNT({rng})=0,"—","Wk "&{m2}&": "'
                       f'&INDEX({wrow},{m2})&"-"&INDEX({lrow},{m2}))')
        for col in "BCDE":
            sty(f"{col}{r}", align="center")
        sty(f"F{r}", bold=True, align="center", fmt="0.0%")
        sty(f"H{r}", size=9, align="center", color=INK2)
        sty(f"J{r}", size=9, align="center", color=INK2)
        line(f"A{r}:K{r}")
    r = 12
    ws[f"A{r}"] = "COOP TOTAL"
    sty(f"A{r}", bold=True, color=INK2)
    ws[f"B{r}"] = "=SUM(B6:B11)"
    ws[f"C{r}"] = "=SUM(C6:C11)"
    ws[f"D{r}"] = "=SUM(D6:D11)"
    ws[f"E{r}"] = "=SUM(E6:E11)"
    ws[f"F{r}"] = f"=IF(B{r}+D{r}=0,0,B{r}/(B{r}+D{r}))"
    for col in "BCDE":
        sty(f"{col}{r}", bold=True, align="center", color=INK2)
    sty(f"F{r}", bold=True, align="center", color=INK2, fmt="0.0%")
    underline(f"A{r}:K{r}")

    ws["A15"] = "WEEK-BY-WEEK WIN % TRACKER"
    sty("A15", size=11, bold=True)
    underline("A15:K15")
    ws.column_dimensions["A"].width = 14
    for w in range(1, WEEKS + 1):
        c = get_column_letter(1 + w)
        ws.column_dimensions[c].width = 5.8
        ws[f"{c}16"] = w
        sty(f"{c}16", size=8, color=MUTED, align="center")
    for i, m in enumerate(MEMBERS):
        r = 17 + i
        ws[f"A{r}"] = m
        sty(f"A{r}", size=9, color=INK2)
        for w in range(1, WEEKS + 1):
            c, mc = get_column_letter(1 + w), mcol(w)
            ws[f"{c}{r}"] = f'=IF({mc}{25 + i}="","",{mc}{25 + i})'
            sty(f"{c}{r}", size=8, align="center", fmt="0%")
    r = 23
    ws[f"A{r}"] = "COOP"
    sty(f"A{r}", size=9, bold=True)
    for w in range(1, WEEKS + 1):
        c, mc = get_column_letter(1 + w), mcol(w)
        ws[f"{c}{r}"] = f'=IF({mc}34="","",{mc}34)'
        sty(f"{c}{r}", size=8, bold=True, align="center", fmt="0%")
    ws.conditional_formatting.add(
        "B17:S23",
        ColorScaleRule(start_type="num", start_value=0, start_color=RED,
                       mid_type="num", mid_value=0.5, mid_color=NEUTRAL,
                       end_type="num", end_value=1, end_color=BLUE))


# ------------------------------------------------------------- pick analytics
def build_pick_analytics(ws):
    sty, underline, line = mk(ws)
    header(ws, sty, "📈 PICK ANALYTICS",
           "How the coop wins and loses, by bet type and situation.")

    conds = [
        ("All Picks", f'{G["season"]},{SCRIT}'),
        ("Spreads", f'{G["ptype"]},"Spread"'),
        ("Over/Under", f'{G["ptype"]},"Over/Under"'),
        ("Overs", f'{G["pick"]},"Over *"'),
        ("Unders", f'{G["pick"]},"Under *"'),
        ("Moneylines", f'{G["ptype"]},"Moneyline"'),
        ("Favorites", f'{G["favdog"]},"Favorite"'),
        ("Underdogs", f'{G["favdog"]},"Underdog"'),
        ("Home Team Picks", f'{G["homeaway"]},"Home"'),
        ("Away Team Picks", f'{G["homeaway"]},"Away"'),
    ]
    ws["A4"] = "GROUP OVERVIEW"
    sty("A4", size=11, bold=True)
    underline("A4:K4")
    for j, h in enumerate(["Category", "Wins", "Losses", "Pushes", "Total",
                           "Win %"]):
        c = get_column_letter(1 + j)
        ws[f"{c}5"] = h
        sty(f"{c}5", size=8, bold=True, color=MUTED,
            align="left" if j == 0 else "center")
    for i, (label, cond) in enumerate(conds):
        r = 6 + i
        base = cond if label == "All Picks" else f"{cond},{G['season']},{SCRIT}"
        ws[f"A{r}"] = label
        sty(f"A{r}", size=9, color=INK2)
        for col, res in [("B", "Win"), ("C", "Loss"), ("D", "Push")]:
            body = base.replace(f'{G["season"]},{SCRIT}',
                                f'{G["result"]},"{res}",{G["season"]},{SCRIT}')
            ws[f"{col}{r}"] = f"=COUNTIFS({body})"
            sty(f"{col}{r}", size=9, align="center")
        ws[f"E{r}"] = f"=B{r}+C{r}+D{r}"
        sty(f"E{r}", size=9, align="center")
        ws[f"F{r}"] = f"=IF(B{r}+C{r}=0,0,B{r}/(B{r}+C{r}))"
        sty(f"F{r}", size=9, bold=True, align="center", fmt="0.0%")
        line(f"A{r}:F{r}")

    ws["A18"] = "PER-PLAYER SITUATIONAL MATRIX — record (win %)"
    sty("A18", size=11, bold=True)
    underline("A18:K18")
    cols = [
        ("B", "Overall", ""),
        ("C", "Spreads", f',{G["ptype"]},"Spread"'),
        ("D", "Totals", f',{G["ptype"]},"Over/Under"'),
        ("E", "Moneylines", f',{G["ptype"]},"Moneyline"'),
        ("F", "Favorites", f',{G["favdog"]},"Favorite"'),
        ("G", "Underdogs", f',{G["favdog"]},"Underdog"'),
        ("H", "Home", f',{G["homeaway"]},"Home"'),
        ("I", "Away", f',{G["homeaway"]},"Away"'),
        ("J", "Overs", f',{G["pick"]},"Over *"'),
        ("K", "Unders", f',{G["pick"]},"Under *"'),
    ]
    ws["A19"] = "Player"
    sty("A19", size=8, bold=True, color=MUTED)
    for col, label, _ in cols:
        ws[f"{col}19"] = label
        sty(f"{col}19", size=8, bold=True, color=MUTED, align="center")
    for i, m in enumerate(MEMBERS):
        r = 20 + i
        ws[f"A{r}"] = m
        sty(f"A{r}", size=9, bold=True)
        for col, _, extra in cols:
            w = person_cnt(f"$A{r}", "Win", extra)
            l = person_cnt(f"$A{r}", "Loss", extra)
            ws[f"{col}{r}"] = wl_pct(w, l)
            sty(f"{col}{r}", size=8.5, align="center", color=INK2)
        line(f"A{r}:K{r}")
    ws["A27"] = "Primetime (Thu + Mon):"
    sty("A27", size=8, bold=True, color=MUTED)
    thu, mon = f',{G["day"]},"Thu"', f',{G["day"]},"Mon"'
    for i, m in enumerate(MEMBERS):
        c = get_column_letter(2 + i)
        ws[f"{c}27"] = m
        sty(f"{c}27", size=8, bold=True, color=MUTED, align="center")
        w = (f'{person_cnt(f"{c}$27", "Win", thu)}+'
             f'{person_cnt(f"{c}$27", "Win", mon)}')
        l = (f'{person_cnt(f"{c}$27", "Loss", thu)}+'
             f'{person_cnt(f"{c}$27", "Loss", mon)}')
        ws[f"{c}28"] = wl_pct(w, l)
        sty(f"{c}28", size=8.5, align="center", color=INK2)


# -------------------------------------------------------------- weekly report
def build_weekly_report(ws):
    sty, underline, line = mk(ws)
    header(ws, sty, "🏆 WEEKLY REPORT",
           "Champions and form, week by week. Ties go to the first listed "
           "member.")
    matrix_block(ws)

    ws["A4"] = "WEEKLY CHAMPIONS"
    sty("A4", size=11, bold=True)
    underline("A4:E4")
    for j, h in enumerate(["Week", "Champion", "Record", "Coop W-L",
                           "Coop Win %"]):
        c = get_column_letter(1 + j)
        ws[f"{c}5"] = h
        sty(f"{c}5", size=8, bold=True, color=MUTED,
            align="left" if j == 0 else "center")
    for w in range(1, WEEKS + 1):
        r, c = 5 + w, mcol(w)
        ws[f"A{r}"] = f"Week {w}"
        sty(f"A{r}", size=9, color=INK2)
        m_ = f"MATCH({c}35,{c}25:{c}30,0)"
        ws[f"B{r}"] = (f'=IF({c}32+{c}33=0,"—",'
                       f'INDEX($AA$4:$AA$9,{m_}))')
        sty(f"B{r}", size=9, bold=True, align="center")
        ws[f"C{r}"] = (f'=IF({c}32+{c}33=0,"",'
                       f'INDEX({c}4:{c}9,{m_})&"-"&INDEX({c}11:{c}16,{m_}))')
        sty(f"C{r}", size=9, align="center", color=INK2)
        ws[f"D{r}"] = f'=IF({c}32+{c}33=0,"",{c}32&"-"&{c}33)'
        sty(f"D{r}", size=9, align="center", color=INK2)
        ws[f"E{r}"] = f'=IF({c}34="","",{c}34)'
        sty(f"E{r}", size=9, bold=True, align="center", fmt="0.0%")
        line(f"A{r}:E{r}")

    ws["G4"] = "HOT / COLD — LAST 4 WEEKS VS SEASON"
    sty("G4", size=11, bold=True)
    underline("G4:K4")
    for j, h in enumerate(["Player", "Last 4", "L4 %", "Season %", "Form"]):
        c = get_column_letter(7 + j)
        ws[f"{c}5"] = h
        sty(f"{c}5", size=8, bold=True, color=MUTED,
            align="left" if j == 0 else "center")
    for i, m in enumerate(MEMBERS):
        r = 6 + i
        w4 = f"SUM(AB{39 + i}:AS{39 + i})"
        l4 = f"SUM(AB{46 + i}:AS{46 + i})"
        sw = f"SUM(AB{4 + i}:AS{4 + i})"
        sl = f"SUM(AB{11 + i}:AS{11 + i})"
        ws[f"G{r}"] = m
        sty(f"G{r}", size=9, bold=True)
        ws[f"H{r}"] = f'=IF({w4}+{l4}=0,"—",{w4}&"-"&{l4})'
        sty(f"H{r}", size=9, align="center", color=INK2)
        ws[f"I{r}"] = f'=IF({w4}+{l4}=0,"",{w4}/({w4}+{l4}))'
        sty(f"I{r}", size=9, align="center", fmt="0%")
        ws[f"J{r}"] = f'=IF({sw}+{sl}=0,"",{sw}/({sw}+{sl}))'
        sty(f"J{r}", size=9, align="center", fmt="0%")
        ws[f"K{r}"] = (f'=IF(OR(I{r}="",J{r}=""),"—",'
                       f'IF(I{r}-J{r}>=0.1,"▲ HOT",'
                       f'IF(I{r}-J{r}<=-0.1,"▼ COLD","► STEADY")))')
        sty(f"K{r}", size=9, bold=True, align="center")
        line(f"G{r}:K{r}")
    ws.conditional_formatting.add(
        "K6:K11", FormulaRule(formula=['ISNUMBER(SEARCH("HOT",K6))'],
                              font=Font(name="Arial", size=9, bold=True,
                                        color=GOOD)))
    ws.conditional_formatting.add(
        "K6:K11", FormulaRule(formula=['ISNUMBER(SEARCH("COLD",K6))'],
                              font=Font(name="Arial", size=9, bold=True,
                                        color=CRIT)))


# ------------------------------------------------------------ bragging rights
def build_bragging_rights(ws):
    sty, underline, line = mk(ws)
    header(ws, sty, "🏅 BRAGGING RIGHTS",
           "Counting what matters. Ties share the honor — first listed "
           "member shown.")
    matrix_block(ws)

    ws["A4"] = "THE LEDGER"
    sty("A4", size=11, bold=True)
    underline("A4:K4")
    heads = ["Player", "Crowns 🏆", "Perfect 💯", "Goose Eggs 🥚",
             "Total Picks", "Best Week", "Worst Week"]
    widths = ["A", "B", "C", "D", "E", "F", "H"]
    for h, c in zip(heads, widths):
        ws[f"{c}5"] = h
        sty(f"{c}5", size=8, bold=True, color=MUTED,
            align="left" if c == "A" else "center")
    ws.merge_cells("F5:G5")
    ws.merge_cells("H5:I5")
    for i, m in enumerate(MEMBERS):
        r = 6 + i
        rng = f"AB{25 + i}:AS{25 + i}"
        wrow, lrow = f"AB{4 + i}:AS{4 + i}", f"AB{11 + i}:AS{11 + i}"
        drow = f"AB{18 + i}:AS{18 + i}"
        ws[f"A{r}"] = m
        sty(f"A{r}", bold=True)
        ws[f"B{r}"] = f"=SUMPRODUCT(({rng}=AB$35:AS$35)*({drow}>0))"
        ws[f"C{r}"] = f"=SUMPRODUCT(({wrow}>=2)*({lrow}=0))"
        ws[f"D{r}"] = f"=SUMPRODUCT(({wrow}=0)*({lrow}>=2))"
        ws[f"E{r}"] = (f"={person_cnt(f'$A{r}', 'Win')}+"
                       f"{person_cnt(f'$A{r}', 'Loss')}+"
                       f"{person_cnt(f'$A{r}', 'Push')}")
        m_ = f"MATCH(MAX({rng}),{rng},0)"
        m2 = f"MATCH(MIN({rng}),{rng},0)"
        ws.merge_cells(f"F{r}:G{r}")
        ws[f"F{r}"] = (f'=IF(COUNT({rng})=0,"—","Wk "&{m_}&": "'
                       f'&INDEX({wrow},{m_})&"-"&INDEX({lrow},{m_}))')
        ws.merge_cells(f"H{r}:I{r}")
        ws[f"H{r}"] = (f'=IF(COUNT({rng})=0,"—","Wk "&{m2}&": "'
                       f'&INDEX({wrow},{m2})&"-"&INDEX({lrow},{m2}))')
        for col in "BCDE":
            sty(f"{col}{r}", align="center")
        sty(f"F{r}", size=9, align="center", color=INK2)
        sty(f"H{r}", size=9, align="center", color=INK2)
        line(f"A{r}:K{r}")

    ws["A14"] = "AWARD PODIUM"
    sty("A14", size=11, bold=True)
    underline("A14:K14")
    awards = [
        ("🏆 Most Championship Weeks", "B"),
        ("💯 Most Perfect Weeks", "C"),
        ("🥚 Most Goose Eggs", "D"),
        ("📚 Most Total Picks", "E"),
    ]
    for i, (label, col) in enumerate(awards):
        r = 15 + i
        rng = f"{col}6:{col}11"
        ws[f"A{r}"] = label
        sty(f"A{r}", size=9, color=INK2)
        ws.merge_cells(f"D{r}:F{r}")
        ws[f"D{r}"] = (f"=INDEX($A$6:$A$11,MATCH(MAX({rng}),{rng},0))"
                       f'&IF(COUNTIF({rng},MAX({rng}))>1,'
                       f'" +"&(COUNTIF({rng},MAX({rng}))-1)&" tied","")'
                       f'&"  —  "&MAX({rng})')
        sty(f"D{r}", size=10, bold=True)
        line(f"A{r}:K{r}")


# ------------------------------------------------------- historical standings
def build_historical_standings(ws):
    sty, underline, line = mk(ws)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 14
    for col in "BCDEF":
        ws.column_dimensions[col].width = 11
    ws["A1"] = "📜 HISTORICAL STANDINGS — ALL SEASONS COMBINED"
    sty("A1", size=14, bold=True)
    ws.row_dimensions[1].height = 26
    ws["A2"] = "Every pick ever logged. Updates itself as seasons are played."
    sty("A2", size=8.5, color=INK2)
    for j, h in enumerate(["Player", "Wins", "Pushes", "Losses", "Total",
                           "Win % (excl. pushes)"]):
        c = get_column_letter(1 + j)
        ws[f"{c}4"] = h
        sty(f"{c}4", size=8, bold=True, color=MUTED,
            align="left" if j == 0 else "center")
    underline("A4:F4")
    all_crit = '"*"'
    for i, m in enumerate(MEMBERS):
        r = 5 + i
        ws[f"A{r}"] = m
        sty(f"A{r}", bold=True)
        for col, res in [("B", "Win"), ("C", "Push"), ("D", "Loss")]:
            ws[f"{col}{r}"] = (f'=COUNTIFS({G["person"]},$A{r},{G["result"]},'
                               f'"{res}",{G["season"]},{all_crit})')
            sty(f"{col}{r}", align="center")
        ws[f"E{r}"] = f"=B{r}+C{r}+D{r}"
        sty(f"E{r}", align="center")
        ws[f"F{r}"] = f"=IF(B{r}+D{r}=0,0,B{r}/(B{r}+D{r}))"
        sty(f"F{r}", bold=True, align="center", fmt="0.0%")
        line(f"A{r}:F{r}")
    r = 11
    ws[f"A{r}"] = "COOP TOTAL"
    sty(f"A{r}", bold=True, color=INK2)
    for col in "BCDE":
        ws[f"{col}{r}"] = f"=SUM({col}5:{col}10)"
        sty(f"{col}{r}", bold=True, align="center", color=INK2)
    ws[f"F{r}"] = f"=IF(B{r}+D{r}=0,0,B{r}/(B{r}+D{r}))"
    sty(f"F{r}", bold=True, align="center", color=INK2, fmt="0.0%")
    underline(f"A{r}:F{r}")


BUILDERS = {
    "Player Dashboard": build_player_dashboard,
    "Pick Analytics": build_pick_analytics,
    "Weekly Report": build_weekly_report,
    "Bragging Rights": build_bragging_rights,
    "Historical Standings": build_historical_standings,
}


def main(path):
    src = Path(path).expanduser()
    if not src.exists():
        sys.exit(f"Workbook not found: {src}")
    stamp = datetime.now().strftime("%Y-%m-%d %H%M%S")
    backup = src.with_name(f"{src.stem} (backup {stamp}){src.suffix}")
    shutil.copy2(src, backup)
    print(f"backup: {backup.name}")

    wb = openpyxl.load_workbook(src)
    for name, builder in BUILDERS.items():
        idx = wb.sheetnames.index(name) if name in wb.sheetnames else None
        if idx is not None:
            del wb[name]
        ws = wb.create_sheet(name, idx)
        builder(ws)
        print(f"rebuilt: {name}")
    for scratch in ("ESPN Games", "Sheet1"):
        if scratch in wb.sheetnames:
            wb[scratch].sheet_state = "hidden"
    wb.calculation.fullCalcOnLoad = True
    wb.save(src)
    print(f"saved: {src}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
