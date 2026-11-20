# -*- coding: utf-8 -*-
"""RUN ONCE. Builds the dashboard template (GitHub dark theme). Daily automation never runs this.

Dashboard uses cells, formulas, one logo image and charts only. No text boxes/shapes,
because openpyxl drops those when update_excel.py saves the file.
"""

import os
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from openpyxl.chart.data_source import NumFmt
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.marker import Marker
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.image import Image
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties, Font as DFont, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from update_excel import SHEETS, TEMPLATE_PATH

# GitHub dark palette
BG, HEADER, CARD, ALT, BORDER = "0D1117", "010409", "161B22", "11161D", "30363D"
TEXT, MUTED = "E6EDF3", "8B949E"
GREEN, BLUE, PURPLE, ORANGE, RED, TEAL = "3FB950", "58A6FF", "A371F7", "D29922", "F85149", "39C5CF"
FONT = "Arial"
CREATOR = "Dishant Bhere"
LOGO = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "github_logo.png")

fill = lambda c: PatternFill("solid", start_color=c, end_color=c)
line = Side(style="thin", color=BORDER)


def font(size=10, bold=False, color=TEXT):
    return Font(name=FONT, size=size, bold=bold, color=color)


def box(ws, rng, bg=CARD):
    for row in ws[rng]:
        for cell in row:
            cell.fill = fill(bg)
            cell.border = Border(left=line, right=line, top=line, bottom=line)


def data_sheet(wb, name, widths):
    table_name, cols, _ = SHEETS[name]
    ws = wb.create_sheet(name)
    for i, c in enumerate(cols, 1):
        ws.cell(1, i, c)
        ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 14)
    t = Table(displayName=table_name, ref=f"A1:{get_column_letter(len(cols))}2")  # header + 1 blank row
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = "A2"
    return ws


# ---------- chart styling (dark) ----------
def rich(color=MUTED, size=900, bold=False):
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DFont(typeface=FONT))
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def title_obj(text, size=1100, color=TEXT, bold=True):
    cp = CharacterProperties(sz=size, b=bold, solidFill=color, latin=DFont(typeface=FONT))
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def axis_line():
    gp = GraphicalProperties()
    gp.line = LineProperties(solidFill=BORDER)
    return gp


def base_style(ch, title, grid=False):
    ch.title = title_obj(title)
    ch.legend = None
    ch.roundedCorners = False
    ch.width, ch.height = 15, 7.5
    gp = GraphicalProperties(solidFill=CARD)
    gp.line = LineProperties(solidFill=BORDER)
    ch.graphical_properties = gp
    ch.plot_area.graphicalProperties = GraphicalProperties(noFill=True)
    for ax in (ch.x_axis, ch.y_axis):
        ax.delete = False
        ax.txPr = rich(MUTED)
        ax.graphicalProperties = axis_line()
    if grid:
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="21262D"))
    else:
        ch.y_axis.majorGridlines = None


def bar_chart(ws, src, anchor, title, color, cat_col, val_col, last_row, horizontal, fmt="#,##0.##"):
    ch = BarChart()
    ch.add_data(Reference(src, min_col=val_col, min_row=1, max_row=last_row), titles_from_data=True)
    ch.set_categories(Reference(src, min_col=cat_col, min_row=2, max_row=last_row))
    base_style(ch, title)
    ch.gapWidth = 55
    ch.y_axis.numFmt = NumFmt(formatCode=fmt, sourceLinked=False)
    if horizontal:
        ch.type = "bar"
        ch.x_axis.scaling.orientation = "maxMin"  # rank 1 on top
        ch.y_axis.crosses = "max"
    else:
        ch.type = "col"
    s = ch.series[0]
    s.graphicalProperties.solidFill = color
    s.graphicalProperties.line.solidFill = color
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    for a in ("showSerName", "showCatName", "showLegendKey", "showPercent"):
        setattr(ch.dataLabels, a, False)
    ch.dataLabels.numFmt = fmt
    ch.dataLabels.txPr = rich(TEXT, 800)
    ws.add_chart(ch, anchor)


def scatter_chart(ws, src, anchor, title, color):
    ch = ScatterChart()
    ch.style = 13
    xs = Reference(src, min_col=21, min_row=2, max_row=201)   # age_days
    ys = Reference(src, min_col=5, min_row=2, max_row=201)    # stars
    s = Series(ys, xs, title="Repos")
    s.marker = Marker(symbol="circle", size=7)
    s.marker.graphicalProperties = GraphicalProperties(solidFill=color)
    s.marker.graphicalProperties.line.solidFill = color
    s.graphicalProperties = GraphicalProperties()
    s.graphicalProperties.line.noFill = True
    ch.series.append(s)
    base_style(ch, title, grid=True)
    ch.x_axis.title = title_obj("Repo age (days)", 900, MUTED, False)
    ch.y_axis.title = title_obj("Stars", 900, MUTED, False)
    ch.x_axis.scaling.min = 0
    ch.y_axis.scaling.min = 0
    ch.x_axis.majorGridlines = None
    ch.y_axis.numFmt = NumFmt(formatCode="#,##0", sourceLinked=False)
    ws.add_chart(ch, anchor)


def main():
    if os.path.exists(TEMPLATE_PATH) and "--force" not in sys.argv:
        raise SystemExit(f"{TEMPLATE_PATH} already exists. Not overwriting your design. "
                         "Use: python build_template.py --force")
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"

    repo = data_sheet(wb, "Repo_Data", {"name": 28, "description": 50, "topics": 30, "url": 40, "owner": 20, "license": 14})
    lang = data_sheet(wb, "Language_Data", {"language": 18, "total_stars": 14, "avg_stars_per_hour": 20})
    cat = data_sheet(wb, "Category_Data", {"viral_level": 18})
    vel = data_sheet(wb, "Velocity_Data", {"name": 28, "stars_per_hour": 16})
    top = data_sheet(wb, "Topic_Data", {"topic": 24, "total_stars": 14})
    act = data_sheet(wb, "Activity_Data", {"last_push": 16})
    cfg = wb.create_sheet("Config")
    cfg.column_dimensions["A"].width = 28
    cfg.column_dimensions["B"].width = 24
    for i, (k, v) in enumerate([("Last refreshed (IST)", ""), ("Window (days)", 30), ("Rows loaded", 0)], 1):
        cfg.cell(i, 1, k).font = Font(name=FONT, bold=True)
        cfg.cell(i, 2, v)
    cfg["A5"] = "Updated by update_excel.py. Dashboard reads these cells + data sheets. Do not rename sheets or columns."

    # ---------- Dashboard ----------
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.sheet_properties.tabColor = GREEN
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["N"].width = 2
    for c in range(2, 14):
        ws.column_dimensions[get_column_letter(c)].width = 13
    LAST_ROW = 100
    for r in range(1, LAST_ROW + 1):
        for c in range(1, 15):
            ws.cell(r, c).fill = fill(BG)

    # header band
    for row in ws["B2:M4"]:
        for cell in row:
            cell.fill = fill(HEADER)
    for c in range(2, 14):
        ws.cell(4, c).border = Border(bottom=Side(style="medium", color=GREEN))
    ws.merge_cells("C2:H3")
    ws["C2"] = "GITHUB TRENDING INTELLIGENCE"
    ws["C2"].font = font(20, True, TEXT)
    ws["C2"].alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells("I2:M3")
    ws["I2"] = f"Created by {CREATOR}"
    ws["I2"].font = font(11, True, GREEN)
    ws["I2"].alignment = Alignment(horizontal="right", vertical="center", indent=1)
    ws.merge_cells("B4:H4")
    ws["B4"] = '="New repos created in the last "&Config!B2&" days - top "&Config!B3&" by stars"'
    ws["B4"].font = font(9, False, MUTED)
    ws["B4"].alignment = Alignment(indent=1, vertical="center")
    ws.merge_cells("I4:M4")
    ws["I4"] = '="Last refreshed: "&Config!B1'
    ws["I4"].font = font(9, False, MUTED)
    ws["I4"].alignment = Alignment(horizontal="right", indent=1, vertical="center")
    for r, h in ((2, 24), (3, 24), (4, 20)):
        ws.row_dimensions[r].height = h
    if os.path.exists(LOGO):
        img = Image(LOGO)
        img.width = img.height = 46
        ws.add_image(img, "B2")

    # KPI cards: (label, formula, number format, font size, value color)
    kpis = [
        ("REPOS TRACKED", "=COUNTA(Repo_Data!B:B)-1", "#,##0", 22, TEXT),
        ("TOTAL STARS", "=SUM(Repo_Data!E:E)", "#,##0", 22, GREEN),
        ("AVG STARS / HOUR", "=ROUND(AVERAGE(Repo_Data!V:V),1)", "0.0", 22, BLUE),
        ("TOP REPO", '=IF(Repo_Data!B2="","-",Repo_Data!B2)', "@", 14, TEXT),
        ("ORG-BACKED REPOS", '=IFERROR(COUNTIF(Repo_Data!D:D,"Organization")/(COUNTA(Repo_Data!B:B)-1),0)', "0%", 22, PURPLE),
        ("TRENDING TIER+ (10K STARS)", '=COUNTIF(Repo_Data!E:E,">=10000")', "#,##0", 22, ORANGE),
    ]
    for i, (label, formula, nf, size, color) in enumerate(kpis):
        c1, c2 = get_column_letter(2 + 2 * i), get_column_letter(3 + 2 * i)
        box(ws, f"{c1}6:{c2}8")
        ws.merge_cells(f"{c1}6:{c2}6")
        ws.merge_cells(f"{c1}7:{c2}8")
        ws[f"{c1}6"] = label
        ws[f"{c1}6"].font = font(8, True, MUTED)
        ws[f"{c1}6"].alignment = Alignment(horizontal="center", vertical="center")
        ws[f"{c1}7"] = formula
        ws[f"{c1}7"].number_format = nf
        ws[f"{c1}7"].font = font(size, True, color)
        ws[f"{c1}7"].alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
    for r in (6, 7, 8):
        ws.row_dimensions[r].height = 20

    def section(row, text):
        ws.merge_cells(f"B{row}:M{row}")
        ws[f"B{row}"] = text
        ws[f"B{row}"].font = font(11, True, TEXT)
        ws[f"B{row}"].alignment = Alignment(vertical="center")
        for c in range(2, 14):
            ws.cell(row, c).border = Border(bottom=Side(style="medium", color=GREEN))
        ws.row_dimensions[row].height = 22

    # Top 10 table
    section(10, "TOP 10 REPOS BY STARS")
    heads = {"B": "RANK", "C": "REPO", "E": "OWNER", "F": "LANGUAGE", "G": "STARS", "H": "TIER",
             "I": "STARS / HR", "J": "AGE (DAYS)", "K": "FORKS", "L": "OPEN ISSUES", "M": "LICENSE"}
    ws.merge_cells("C11:D11")
    for col in "BCDEFGHIJKLM":
        cell = ws[f"{col}11"]
        cell.fill = fill("21262D")
        if col in heads:
            cell.value = heads[col]
        cell.font = font(9, True, MUTED)
        cell.alignment = Alignment(horizontal="left" if col in "CDEF" else "center", vertical="center", indent=1 if col in "CDEF" else 0)
    ws.row_dimensions[11].height = 22
    src = {"B": ("A", "0"), "C": ("B", "@"), "E": ("C", "@"), "F": ("I", "@"), "G": ("E", "#,##0"),
           "H": ("H", "@"), "I": ("V", "0.00"), "J": ("U", "0.0"), "K": ("F", "#,##0"),
           "L": ("G", "#,##0"), "M": ("K", "@")}
    for i in range(10):
        r, s = 12 + i, 2 + i
        bg = CARD if i % 2 == 0 else ALT
        for col in "BCDEFGHIJKLM":
            ws[f"{col}{r}"].fill = fill(bg)
            ws[f"{col}{r}"].border = Border(bottom=line)
        ws.merge_cells(f"C{r}:D{r}")
        for col, (sc, nf) in src.items():
            cell = ws[f"{col}{r}"]
            cell.value = f'=IF(Repo_Data!$B{s}="","",Repo_Data!{sc}{s})'
            cell.number_format = nf
            cell.font = font(10, col in "CG", BLUE if col == "C" else (GREEN if col == "G" else TEXT))
            cell.alignment = Alignment(horizontal="left" if col in "CEF" else "center", vertical="center", indent=1 if col in "CEF" else 0)
        ws.row_dimensions[r].height = 20
    tiers = {"Trending": ("0C2D6B", BLUE), "Rising Star": ("0F5323", "56D364"), "Viral": ("5C4B00", "E3B341"),
             "Super Viral": ("5A2B00", "FFA657"), "Legendary": ("3B1F7A", "D2A8FF"), "Iconic": ("3B1F7A", "D2A8FF"),
             "Phenomenal": ("67060C", "FF7B72"), "Hall of Fame": ("67060C", "FF7B72")}
    for tier, (bgc, fgc) in tiers.items():
        ws.conditional_formatting.add("H12:H21", CellIsRule(operator="equal", formula=[f'"{tier}"'],
                                                           fill=fill(bgc), font=Font(name=FONT, color=fgc, bold=True)))

    # Insights
    section(23, "KEY INSIGHTS")
    total = "(COUNTA(Repo_Data!B:B)-1)"
    insights = [
        '="Fastest growing: "&Velocity_Data!A2&" at "&TEXT(Velocity_Data!B2,"0.0")&" stars per hour (lifetime average)"',
        '="Most common language: "&Language_Data!A2&" ("&Language_Data!B2&" repos). Fastest language by avg stars/hour: "'
        '&INDEX(Language_Data!A2:A9,MATCH(MAX(Language_Data!E2:E9),Language_Data!E2:E9,0))&" ("&TEXT(MAX(Language_Data!E2:E9),"0.0")&")"',
        f'="Organizations back "&TEXT(COUNTIF(Repo_Data!D:D,"Organization")/{total},"0%")&" of these repos, individual developers "&TEXT(COUNTIF(Repo_Data!D:D,"User")/{total},"0%")',
        f'="No license detected on "&SUMPRODUCT((Repo_Data!B2:B1000<>"")*(Repo_Data!K2:K1000=""))&" of "&{total}&" repos"',
        f'="Pushed in the last 7 days: "&(Activity_Data!B2+Activity_Data!B3+Activity_Data!B4)&" of "&{total}&" repos still active"',
    ]
    for i, f in enumerate(insights):
        r = 24 + i
        ws.merge_cells(f"B{r}:M{r}")
        box(ws, f"B{r}:M{r}")
        ws[f"B{r}"] = f
        ws[f"B{r}"].font = font(10, False, TEXT)
        ws[f"B{r}"].alignment = Alignment(vertical="center", indent=1)
        ws.row_dimensions[r].height = 20

    # Charts: fixed top-N rows of data sheets (Python writes them pre-sorted)
    section(30, "TRENDS & BREAKDOWNS")
    bar_chart(ws, repo, "B32", "Top 10 repos by stars", GREEN, 2, 5, 11, True, "#,##0")
    bar_chart(ws, vel, "H32", "Fastest growing - stars per hour (top 10)", ORANGE, 1, 2, 11, True)
    bar_chart(ws, lang, "B49", "Repos by language (top 8)", BLUE, 1, 2, 9, False, "0")
    bar_chart(ws, lang, "H49", "Language momentum - avg stars per hour", TEAL, 1, 5, 9, False)
    bar_chart(ws, top, "B66", "Hottest topics (top 10 tags)", PURPLE, 1, 2, 11, True, "0")
    bar_chart(ws, cat, "H66", "Repos by popularity tier", RED, 1, 2, 10, False, "0")
    scatter_chart(ws, repo, "B83", "Age vs stars - who grew fastest?", GREEN)
    bar_chart(ws, act, "H83", "Maintenance - days since last push", BLUE, 1, 2, 5, False, "0")

    ws.merge_cells(f"B{LAST_ROW - 1}:M{LAST_ROW - 1}")
    ws[f"B{LAST_ROW - 1}"] = f"Created by {CREATOR}  |  Data: GitHub Search API  |  Auto-refreshed daily by GitHub Actions"
    ws[f"B{LAST_ROW - 1}"].font = font(8, False, MUTED)
    ws[f"B{LAST_ROW - 1}"].alignment = Alignment(horizontal="center")

    ws.print_area = f"A1:N{LAST_ROW}"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    os.makedirs(os.path.dirname(TEMPLATE_PATH) or ".", exist_ok=True)
    wb.save(TEMPLATE_PATH)
    print(f"Template saved: {TEMPLATE_PATH}")


if __name__ == "__main__":
    main()
