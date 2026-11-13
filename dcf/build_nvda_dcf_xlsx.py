"""Builds NVDA_DCF_Model_2026-11-13.xlsx (sheets: DCF, WACC, Comps).
All derived cells are live Excel formulas; only raw inputs/assumptions are hardcoded (blue, with source comments).
Units: $ millions unless stated; per-share in $."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

OUT = "dcf/NVDA_DCF_Model_2026-11-13.xlsx"

# ---------- styles ----------
F_IN = Font(name="Arial", size=10, color="0000FF")
F_FX = Font(name="Arial", size=10, color="000000")
F_LINK = Font(name="Arial", size=10, color="008000")
F_B = Font(name="Arial", size=10, bold=True)
F_HDR = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name="Arial", size=14, bold=True)
FILL_HDR = PatternFill("solid", fgColor="1F4E79")
FILL_SUB = PatternFill("solid", fgColor="D9E1F2")
FILL_IN = PatternFill("solid", fgColor="F2F2F2")
FILL_OUT = PatternFill("solid", fgColor="BDD7EE")
NUM = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
PX = '$#,##0.00'
MULT = '0.0"x"'
CENTER = Alignment(horizontal="center", vertical="center")
THICK, THIN = Side(style="medium"), Side(style="thin")


def put(ws, ref, val, kind="fx", fmt=None, note=None, bold=False):
    c = ws[ref]
    c.value = val
    c.font = {"in": F_IN, "fx": F_FX, "link": F_LINK}[kind]
    if bold:
        c.font = Font(name="Arial", size=10, bold=True, color=c.font.color.rgb if c.font.color else "000000")
    if kind == "in":
        c.fill = FILL_IN
    if fmt:
        c.number_format = fmt
    if note:
        c.comment = Comment(note, "Model")
        c.comment.width, c.comment.height = 320, 110
    if ref[0] != "A":
        c.alignment = CENTER
    return c


def section(ws, row, text, c1=1, c2=14):
    ws.cell(row=row, column=c1, value=text)
    for col in range(c1, c2 + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill, cell.font = FILL_HDR, F_HDR
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)


def subhdr(ws, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.fill, c.font, c.alignment = FILL_SUB, F_B, CENTER


def box(ws, r1, r2, c1=1, c2=14):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(row=r, column=c)
            b = cell.border
            cell.border = Border(
                left=THICK if c == c1 else b.left, right=THICK if c == c2 else b.right,
                top=THICK if r == r1 else b.top, bottom=THICK if r == r2 else b.bottom)


def label(ws, row, text, bold=False, indent=0):
    c = ws.cell(row=row, column=1, value=text)
    c.font = F_B if bold else F_FX
    c.alignment = Alignment(indent=indent)


wb = Workbook()
ws = wb.active
ws.title = "DCF"
wacc = wb.create_sheet("WACC")
comps = wb.create_sheet("Comps")

HIST = ["B", "C", "D"]                      # FY24A..FY26A
PROJ = [L(c) for c in range(5, 15)]         # E..N = FY27E..FY36E
YEARS_H = ["FY2024A", "FY2025A", "FY2026A"]
YEARS_P = [f"FY20{y}E" for y in range(27, 37)]
SRC_FY26 = "Source: NVIDIA Q4/FY2026 press release (25 Feb 2026) & FY2026 10-K, https://investor.nvidia.com/news/press-release-details/2026/NVIDIA-Announces-Financial-Results-for-Fourth-Quarter-and-Fiscal-2026"
SRC_Q2 = "Source: NVIDIA Q2 FY2027 press release / CFO commentary (26 Aug 2026), https://nvidianews.nvidia.com/_gallery/download_pdf/6a8f4ad73d6332b41f76ab70/"

# ---------- header ----------
ws["A1"] = "NVIDIA Corporation (NVDA) - DCF Model"
ws["A1"].font = F_TITLE
ws["A2"] = "Ticker: NVDA | Valuation date: 13-Nov-2026 | Fiscal year ends late January | $ millions except per share"
ws["A2"].font = Font(name="Arial", size=9, italic=True)
label(ws, 4, "Case selector (1=Bear, 2=Base, 3=Bull)", bold=True)
put(ws, "C4", 2, "in", note="Assumption: change to 1/2/3 to switch scenario. All projections pull from the selected block via CHOOSE.")
label(ws, 5, "Selected case")
put(ws, "C5", '=CHOOSE($C$4,"Bear","Base","Bull")', bold=True)

# ---------- market data ----------
section(ws, 7, "MARKET DATA & BALANCE SHEET (not case dependent)")
md = [
    (8, "Current share price ($)", 166.19, PX, "Source: CNN Markets / Motley Fool quote, close 12-Nov-2026 ($166.19). Sources disagreed ($159-$174); verify live quote."),
    (9, "Diluted shares outstanding (M)", 24263, NUM, "Source: derived from Q2 FY27 GAAP net income $59,688M / diluted EPS $2.46 = ~24,263M diluted shares. " + SRC_Q2),
    (11, "Cash, equivalents & marketable debt securities", 56600, NUM, "Source: Q2 FY27 (26-Jul-2026) balance sheet: $22.4B cash + $34.1B marketable debt securities. " + SRC_Q2),
    (12, "Marketable equity securities", 42800, NUM, "Source: Q2 FY27 balance sheet, marketable equity securities $42.8B (strategic stakes) - treated as non-operating asset at 100%. Assumption: no haircut."),
    (13, "Total debt", 33400, NUM, "Source: Q2 FY27 balance sheet: $1.0B short-term + $32.4B long-term debt (incl. $25B notes issued in Q2). Excludes ~$5B operating leases."),
    (15, "Valuation date", dt.date(2026, 11, 13), "dd-mmm-yyyy", "Assumption: analysis date."),
    (16, "FY2027 year end", dt.date(2027, 1, 31), "dd-mmm-yyyy", "Assumption: NVIDIA FY ends last Sunday of January; FY27 ends ~31-Jan-2027."),
    (18, "Tax rate (all cases)", 0.17, PCT, "Assumption: NVIDIA effective tax rate has run ~15-17% (FDII benefit). 17% used; statutory 21% would lower value ~5%."),
]
for r, lab, v, f, n in md:
    label(ws, r, lab)
    put(ws, f"C{r}", v, "in", f, n)
label(ws, 10, "Market capitalization")
put(ws, "C10", "=C8*C9", fmt=NUM)
label(ws, 14, "Net cash / (net debt)", bold=True)
put(ws, "C14", "=C11+C12-C13", fmt=NUM, bold=True)
label(ws, 17, "Stub fraction of FY27 remaining")
put(ws, "C17", "=(C16-C15)/365", fmt="0.000")
# FY27E quarterly build
put(ws, "E8", "FY27E revenue build", bold=True)
q = [
    (9, "Q1 FY27A", 81500, "Source: implied from Q2 FY27 revenue $96.2B being +18% q/q => ~$81.5B. " + SRC_Q2),
    (10, "Q2 FY27A", 96200, SRC_Q2),
    (11, "Q3 FY27E (guidance midpoint)", 108000, "Source: company Q3 FY27 outlook $108.0B +/-2% (26 Aug 2026). Report due ~17-Nov-2026; NVDA has beaten guide midpoint ~5% the last 4 qtrs."),
    (12, "Q4 FY27E", 117000, "Assumption: ~8% q/q growth on Q3 guide; analyst estimate, not company guidance."),
]
for r, lab, v, n in q:
    ws[f"E{r}"] = lab
    ws[f"E{r}"].font = F_FX
    put(ws, f"H{r}", v, "in", NUM, n)
ws["E13"] = "FY27E total revenue"
ws["E13"].font = F_B
put(ws, "H13", "=SUM(H9:H12)", fmt=NUM, bold=True)
box(ws, 7, 18)

# ---------- scenario blocks ----------
ASSUMP = ["Revenue growth (%)", "Gross margin (%)", "R&D (% of revenue)", "SG&A (% of revenue)",
          "D&A (% of revenue)", "CapEx (% of revenue)", "Change in NWC (% of change in revenue)",
          "Terminal growth rate", "WACC adjustment vs WACC sheet"]
SCEN = {
    "BEAR": dict(
        g=[None, .12, .00, -.08, .03, .04, .04, .035, .03, .03],
        gm=[.73, .70, .65, .62, .60, .59, .58, .57, .56, .55],
        rd=[.07, .08, .095, .105, .11, .115, .12, .12, .12, .12],
        sga=[.02, .022, .024, .025, .025, .025, .025, .025, .025, .025],
        da=[.012, .014, .016, .018, .02, .02, .02, .02, .02, .02],
        capex=[.025, .03, .035, .035, .035, .035, .035, .035, .035, .035],
        nwc=[.25] * 10, tg=.025, wadj=.01,
        why="Bear: AI capex digestion in FY29-30, custom ASICs (TPU/Trainium) and AMD take share, gross margin compresses to mid-50s."),
    "BASE": dict(
        g=[None, .35, .18, .10, .07, .06, .05, .045, .04, .035],
        gm=[.735, .73, .72, .71, .705, .70, .695, .69, .685, .68],
        rd=[.07, .07, .075, .08, .085, .09, .09, .095, .095, .10],
        sga=[.02, .02, .021, .022, .023, .024, .025, .025, .025, .025],
        da=[.012, .013, .014, .015, .016, .017, .018, .019, .02, .02],
        capex=[.025, .025, .025, .028, .03, .03, .03, .03, .03, .03],
        nwc=[.20] * 10, tg=.03, wadj=0.0,
        why="Base: FY28 +35% on Blackwell/Rubin ramp (Q3 guide $108B implies strong exit rate), growth fades to ~3.5% by FY36; gross margin drifts from 73.5% to 68%."),
    "BULL": dict(
        g=[None, .50, .28, .18, .13, .10, .08, .06, .05, .04],
        gm=[.74, .745, .745, .74, .74, .735, .735, .73, .73, .73],
        rd=[.07, .065, .065, .07, .07, .075, .075, .08, .08, .08],
        sga=[.02, .018, .018, .019, .02, .02, .02, .02, .02, .02],
        da=[.012, .012, .013, .014, .015, .016, .017, .018, .019, .02],
        capex=[.025, .022, .022, .025, .025, .025, .025, .025, .025, .025],
        nwc=[.15] * 10, tg=.035, wadj=-.015,
        why="Bull: AI infrastructure build-out sustains (sovereign AI, inference, robotics); NVDA holds ~74% gross margin via software/systems moat."),
}
BLOCK_START = {"BEAR": 25, "BASE": 37, "BULL": 49}
KEYS = ["g", "gm", "rd", "sga", "da", "capex", "nwc"]
for name, r0 in BLOCK_START.items():
    s = SCEN[name]
    section(ws, r0, f"{name} CASE ASSUMPTIONS")
    subhdr(ws, r0 + 1, ["Assumption", "", "", ""] + YEARS_P, 1)
    for i, k in enumerate(KEYS):
        r = r0 + 2 + i
        label(ws, r, ASSUMP[i], indent=1)
        for j, col in enumerate(PROJ):
            v = s[k][j]
            if v is None:
                put(ws, f"{col}{r}", "=H13/D75-1", fmt=PCT, note="FY27E growth comes from the quarterly revenue build (H13), not an input.")
                continue
            note = s["why"] if (i == 0 and j == 1) else None
            put(ws, f"{col}{r}", v, "in", PCT, note or (f"Assumption ({name.title()} case)." if j == 1 else None))
    label(ws, r0 + 9, ASSUMP[7], indent=1)
    put(ws, f"C{r0 + 9}", s["tg"], "in", PCT, f"Assumption ({name.title()}): long-run nominal growth; must stay below WACC and ~US nominal GDP.")
    label(ws, r0 + 10, ASSUMP[8], indent=1)
    put(ws, f"C{r0 + 10}", s["wadj"], "in", PCT, f"Assumption ({name.title()}): spread added to CAPM WACC on WACC sheet to reflect scenario risk.")
    box(ws, r0, r0 + 10)

SEL = 61
section(ws, SEL, "SELECTED CASE (consolidation - pulls from block chosen in C4)")
subhdr(ws, SEL + 1, ["Assumption", "", "", ""] + YEARS_P, 1)
for i in range(7):
    r = SEL + 2 + i
    label(ws, r, ASSUMP[i], indent=1)
    for col in PROJ:
        a, b, c = (f"{col}{BLOCK_START[n] + 2 + i}" for n in ("BEAR", "BASE", "BULL"))
        put(ws, f"{col}{r}", f"=CHOOSE($C$4,{a},{b},{c})", fmt=PCT)
for k, i in ((9, 7), (10, 8)):
    r = SEL + k
    label(ws, r, ASSUMP[i], indent=1)
    a, b, c = (f"C{BLOCK_START[n] + k}" for n in ("BEAR", "BASE", "BULL"))
    put(ws, f"C{r}", f"=CHOOSE($C$4,{a},{b},{c})", fmt=PCT)
box(ws, SEL, SEL + 10)
R_G, R_GM, R_RD, R_SGA, R_DA, R_CX, R_NWC = range(SEL + 2, SEL + 9)
R_TG, R_WADJ = SEL + 9, SEL + 10   # 70, 71

# ---------- projections ----------
P = 73
section(ws, P, "INCOME STATEMENT & UNLEVERED FREE CASH FLOW ($M)")
subhdr(ws, P + 1, ["", *YEARS_H, *YEARS_P], 1)
REV, GRW, GP, GMP, RD, SGA, EBIT, EBITM, TAX, NOPAT, DA, CX, NWC, FCF, EBITDA = range(75, 90)
rows_lab = {REV: "Revenue", GRW: "  % growth", GP: "Gross profit", GMP: "  % gross margin", RD: "R&D",
            SGA: "SG&A", EBIT: "EBIT (operating income)", EBITM: "  % EBIT margin", TAX: "(-) Taxes on EBIT",
            NOPAT: "NOPAT", DA: "(+) D&A", CX: "(-) CapEx", NWC: "(-) Increase in NWC",
            FCF: "Unlevered free cash flow", EBITDA: "EBITDA (EBIT + D&A)"}
for r, t in rows_lab.items():
    label(ws, r, t, bold=r in (REV, EBIT, FCF))
hist = {
    REV: ([60922, 130497, 215938], "Source: NVIDIA 10-K FY2024/FY2025 and FY2026 press release, total revenue."),
    GP: ([44301, 97858, 153532], "Source: 10-K FY24/FY25 gross profit; FY26 = 71.1% GAAP gross margin x revenue (FY26 press release)."),
    RD: ([8675, 12914, 18500], "Source: 10-K FY24/FY25 R&D. FY26 is an estimate: total opex derived as GP - GAAP EBIT ($130,387M), split ~80/20 R&D/SG&A."),
    SGA: ([2654, 3491, 4645], "Source: 10-K FY24/FY25 SG&A. FY26 = GP - EBIT - R&D estimate (plug)."),
    DA: ([1508, 1864, 2900], "Source: 10-K FY24/FY25 cash-flow statement D&A. FY26 is an estimate (not found in sources searched) - verify vs 10-K."),
    CX: ([1069, 3236, 6000], "Source: 10-K FY24/FY25 purchases of PP&E and intangibles. FY26 is an estimate - verify vs 10-K."),
}
for r, (vals, note) in hist.items():
    for j, col in enumerate(HIST):
        put(ws, f"{col}{r}", vals[j], "in", NUM, note if j == 0 else None)
for j, col in enumerate(HIST):
    put(ws, f"{col}{EBIT}", f"={col}{GP}-{col}{RD}-{col}{SGA}", fmt=NUM, bold=True)
    put(ws, f"{col}{GMP}", f"={col}{GP}/{col}{REV}", fmt=PCT)
    put(ws, f"{col}{EBITM}", f"={col}{EBIT}/{col}{REV}", fmt=PCT)
    put(ws, f"{col}{EBITDA}", f"={col}{EBIT}+{col}{DA}", fmt=NUM)
    if j:
        put(ws, f"{col}{GRW}", f"={col}{REV}/{HIST[j-1]}{REV}-1", fmt=PCT)
for j, col in enumerate(PROJ):
    prev = "D" if j == 0 else PROJ[j - 1]
    put(ws, f"{col}{REV}", "=$H$13" if j == 0 else f"={prev}{REV}*(1+{col}${R_G})", fmt=NUM, bold=True)
    put(ws, f"{col}{GRW}", f"={col}{REV}/{prev}{REV}-1", fmt=PCT)
    put(ws, f"{col}{GP}", f"={col}{REV}*{col}${R_GM}", fmt=NUM)
    put(ws, f"{col}{GMP}", f"={col}{GP}/{col}{REV}", fmt=PCT)
    put(ws, f"{col}{RD}", f"={col}{REV}*{col}${R_RD}", fmt=NUM)
    put(ws, f"{col}{SGA}", f"={col}{REV}*{col}${R_SGA}", fmt=NUM)
    put(ws, f"{col}{EBIT}", f"={col}{GP}-{col}{RD}-{col}{SGA}", fmt=NUM, bold=True)
    put(ws, f"{col}{EBITM}", f"={col}{EBIT}/{col}{REV}", fmt=PCT)
    put(ws, f"{col}{TAX}", f"={col}{EBIT}*$C$18", fmt=NUM)
    put(ws, f"{col}{NOPAT}", f"={col}{EBIT}-{col}{TAX}", fmt=NUM)
    put(ws, f"{col}{DA}", f"={col}{REV}*{col}${R_DA}", fmt=NUM)
    put(ws, f"{col}{CX}", f"={col}{REV}*{col}${R_CX}", fmt=NUM)
    put(ws, f"{col}{NWC}", f"=({col}{REV}-{prev}{REV})*{col}${R_NWC}", fmt=NUM)
    put(ws, f"{col}{FCF}", f"={col}{NOPAT}+{col}{DA}-{col}{CX}-{col}{NWC}", fmt=NUM, bold=True)
    put(ws, f"{col}{EBITDA}", f"={col}{EBIT}+{col}{DA}", fmt=NUM)
box(ws, P, EBITDA)

# ---------- discounting ----------
D0 = 91
section(ws, D0, "DISCOUNTING (mid-period convention; FY27E counts only the stub to 31-Jan-2027)")
subhdr(ws, D0 + 1, ["", "", "", "", *YEARS_P], 1)
FRAC, PER, FCFC, DF, PV = range(93, 98)
for r, t in {FRAC: "Fraction of year counted", PER: "Discount period (years)", FCFC: "FCF counted",
             DF: "Discount factor", PV: "PV of FCF"}.items():
    label(ws, r, t)
for j, col in enumerate(PROJ):
    prev = PROJ[j - 1] if j else None
    put(ws, f"{col}{FRAC}", "=$C$17" if j == 0 else "=1", fmt="0.000")
    put(ws, f"{col}{PER}", f"={col}{FRAC}/2" if j == 0 else (f"=$C$17+0.5" if j == 1 else f"={prev}{PER}+1"), fmt="0.000")
    put(ws, f"{col}{FCFC}", f"={col}{FCF}*{col}{FRAC}", fmt=NUM)
    put(ws, f"{col}{DF}", f"=1/(1+$C$101)^{col}{PER}", fmt="0.0000")
    put(ws, f"{col}{PV}", f"={col}{FCFC}*{col}{DF}", fmt=NUM)
box(ws, D0, PV)

# ---------- terminal value ----------
T0 = 99
section(ws, T0, "TERMINAL VALUE (Gordon growth on FY36E FCF)")
tv = [
    (101, "WACC (selected case)", f"=WACC!C25+C{R_WADJ}", PCT, "link"),
    (102, "Terminal growth rate", f"=C{R_TG}", PCT, "fx"),
    (103, "Terminal-year FCF (FY37)", f"=N{FCF}*(1+C102)", NUM, "fx"),
    (104, "Terminal value", "=C103/(C101-C102)", NUM, "fx"),
    (105, "PV of terminal value", f"=C104*N{DF}", NUM, "fx"),
    (106, "Implied terminal EV / NTM EBITDA", f"=C104/(N{EBITDA}*(1+C102))", MULT, "fx"),
]
for r, lab, f, fmt, k in tv:
    label(ws, r, lab)
    put(ws, f"C{r}", f, k, fmt)
box(ws, T0, 106)

# ---------- valuation summary ----------
V0 = 108
section(ws, V0, "VALUATION SUMMARY")
vs = [
    (109, "Sum of PV of FCF", f"=SUM(E{PV}:N{PV})", NUM),
    (110, "PV of terminal value", "=C105", NUM),
    (111, "Enterprise value", "=C109+C110", NUM),
    (112, "(+) Net cash incl. equity securities", "=C14", NUM),
    (113, "Equity value", "=C111+C112", NUM),
    (114, "Diluted shares (M)", "=C9", NUM),
    (115, "IMPLIED VALUE PER SHARE", "=C113/C114", PX),
    (116, "Current share price", "=C8", PX),
    (117, "Implied upside / (downside)", "=C115/C116-1", PCT),
    (118, "Terminal value % of EV", "=C110/C111", PCT),
]
for r, lab, f, fmt in vs:
    label(ws, r, lab, bold=r in (111, 113, 115, 117))
    put(ws, f"C{r}", f, fmt=fmt, bold=r in (111, 113, 115, 117))
for r in (115, 117):
    for col in ("A", "B", "C"):
        ws[f"{col}{r}"].fill = FILL_OUT
box(ws, V0, 118)

# ---------- cross-check ----------
X0 = 120
section(ws, X0, "CROSS-CHECK VS PEERS (see Comps sheet)")
xc = [
    (121, "DCF EV / FY27E EBITDA", f"=C111/E{EBITDA}", MULT, "fx"),
    (122, "DCF EV / FY28E EBITDA", f"=C111/F{EBITDA}", MULT, "fx"),
    (123, "Market EV / FY28E EBITDA", f"=(C10-C14)/F{EBITDA}", MULT, "fx"),
    (124, "Peer median EV / LTM EBITDA", "=Comps!D14", MULT, "link"),
    (125, "DCF equity value / FY28E NOPAT (P/E proxy)", f"=C113/F{NOPAT}", MULT, "fx"),
    (126, "Peer median forward P/E", "=Comps!F14", MULT, "link"),
    (127, "Implied terminal EV/EBITDA vs peer 25th pct", "=Comps!D15", MULT, "link"),
]
for r, lab, f, fmt, k in xc:
    label(ws, r, lab)
    put(ws, f"C{r}", f, k, fmt)
box(ws, X0, 127)

# ---------- sensitivity tables ----------
def price_formula(w, g, fcf_row="$E$88:$N$88", last="$N$88"):
    return (f"(SUMPRODUCT({fcf_row},$E${FRAC}:$N${FRAC},1/(1+{w})^$E${PER}:$N${PER})"
            f"+{last}*(1+{g})/({w}-{g})/(1+{w})^$N${PER}+$C$14)/$C$9")


def sens_table(r0, title, row_lab, col_lab, row_axis, col_axis, cell_fn, rfmt, cfmt):
    section(ws, r0, title, 1, 7)
    ws.cell(row=r0 + 1, column=1, value=f"Rows: {row_lab}  |  Columns: {col_lab}").font = Font(name="Arial", size=9, italic=True)
    hr = r0 + 2
    subhdr(ws, hr, [""], 1)
    for j, f in enumerate(col_axis):
        put(ws, f"{L(2 + j)}{hr}", f, fmt=cfmt, bold=True)
        ws[f"{L(2 + j)}{hr}"].fill = FILL_SUB
    for i, f in enumerate(row_axis):
        r = hr + 1 + i
        put(ws, f"A{r}", f, fmt=rfmt, bold=True)
        ws[f"A{r}"].fill = FILL_SUB
        ws[f"A{r}"].alignment = CENTER
        for j in range(5):
            c = put(ws, f"{L(2 + j)}{r}", "=" + cell_fn(i, j, r, L(2 + j), hr), fmt=PX)
            if i == 2 and j == 2:
                c.fill, c.font = FILL_OUT, Font(name="Arial", size=10, bold=True)
    box(ws, r0, hr + 5, 1, 7)


steps = [-2, -1, 0, 1, 2]
S1 = 130
sens_table(S1, "SENSITIVITY 1: WACC vs TERMINAL GROWTH ($ per share)", "WACC", "Terminal growth",
           [f"=$C$101{s*0.005:+.3f}" if s else "=$C$101" for s in steps],
           [f"=$C$102{s*0.005:+.3f}" if s else "=$C$102" for s in steps],
           lambda i, j, r, c, hr: price_formula(f"$A{r}", f"{c}${hr}"), PCT, PCT)

S2 = 139
HELP = 160   # helper block for sensitivity 2
sens_table(S2, "SENSITIVITY 2: REVENUE GROWTH SHIFT vs EBIT MARGIN SHIFT ($ per share)",
           "Shift to annual revenue growth FY28-FY36 (pp)", "Shift to EBIT margin, all years (pp)",
           [s * 0.02 for s in steps], [s * 0.05 for s in steps],
           lambda i, j, r, c, hr: f"$P${HELP + 2 + 2 * (i * 5 + j) + 1}", '+0.0%;-0.0%;0.0%', '+0.0%;-0.0%;0.0%')
for i in range(5):  # axis values are inputs here
    ws[f"A{S2 + 3 + i}"].font = Font(name="Arial", size=10, bold=True, color="0000FF")
for j in range(5):
    ws[f"{L(2 + j)}{S2 + 2}"].font = Font(name="Arial", size=10, bold=True, color="0000FF")
ws[f"A{S2 + 3}"].comment = Comment("Assumption: sensitivity steps. Center row/column = 0 shift = selected case.", "Model")

S3 = 148
WE = ("(WACC!$C$22*({rf}+{b}*WACC!$C$7)+WACC!$C$23*WACC!$C$13+$C$" + str(R_WADJ) + ")")
sens_table(S3, "SENSITIVITY 3: BETA vs RISK-FREE RATE ($ per share)", "Beta", "Risk-free rate",
           [f"=WACC!$C$6{s*0.25:+.2f}" if s else "=WACC!$C$6" for s in steps],
           [f"=WACC!$C$4{s*0.005:+.3f}" if s else "=WACC!$C$4" for s in steps],
           lambda i, j, r, c, hr: price_formula(WE.format(rf=f"{c}${hr}", b=f"$A{r}"), "$C$102"), "0.00", PCT)
for i in range(5):
    ws[f"A{S3 + 3 + i}"].font = Font(name="Arial", size=10, bold=True, color="008000")
for j in range(5):
    ws[f"{L(2 + j)}{S3 + 2}"].font = Font(name="Arial", size=10, bold=True, color="008000")

# helper block: full re-projection for each of 25 growth/margin combinations
section(ws, HELP, "HELPER: Sensitivity 2 re-projections (revenue and FCF for each combination; do not edit)", 1, 16)
subhdr(ws, HELP + 1, ["Combination", "Growth shift", "Margin shift", "Line", *YEARS_P, "", "Value/share"], 1)
for i in range(5):
    for j in range(5):
        rr = HELP + 2 + 2 * (i * 5 + j)
        fr = rr + 1
        dg, dm = f"$A${S2 + 3 + i}", f"${L(2 + j)}${S2 + 2}"
        ws[f"A{rr}"] = f"Row {i+1} / Col {j+1}"
        put(ws, f"B{rr}", f"={dg}", fmt='+0.0%;-0.0%;0.0%')
        put(ws, f"C{rr}", f"={dm}", fmt='+0.0%;-0.0%;0.0%')
        ws[f"D{rr}"], ws[f"D{fr}"] = "Revenue", "FCF"
        for k, col in enumerate(PROJ):
            prev = PROJ[k - 1] if k else None
            put(ws, f"{col}{rr}", f"=${col}${REV}" if k == 0 else f"={prev}{rr}*(1+{col}${R_G}+$B{rr})", fmt=NUM)
            prev_rev = "$D$75" if k == 0 else f"{prev}{rr}"
            put(ws, f"{col}{fr}",
                f"={col}{rr}*({col}${EBITM}+$C{rr})*(1-$C$18)+{col}{rr}*({col}${R_DA}-{col}${R_CX})-({col}{rr}-{prev_rev})*{col}${R_NWC}",
                fmt=NUM)
        put(ws, f"P{fr}", "=" + price_formula("$C$101", "$C$102", f"$E{fr}:$N{fr}", f"$N{fr}"), fmt=PX)

# column widths / freeze
ws.column_dimensions["A"].width = 44
for c in range(2, 17):
    ws.column_dimensions[L(c)].width = 13
ws.freeze_panes = "B3"

# ---------- WACC sheet ----------
w = wacc
w["A1"] = "NVIDIA - Weighted Average Cost of Capital"
w["A1"].font = F_TITLE
section(w, 3, "COST OF EQUITY (CAPM)", 1, 4)
rows = [
    (4, "Risk-free rate (10Y UST)", 0.0475, PCT, "in", "Source: Trading Economics / Morningstar, 10Y UST ~4.7-5.0% in Oct-Nov 2026 (5.0% on 21-Oct-2026). 4.75% used; verify FRED DGS10."),
    (5, "Raw 5Y monthly beta (reference)", 2.09, "0.00", "in", "Source: Intrinio 5Y beta 2.09 (Fool 2.16, Stock Rover 1Y 2.19). Shown for reference only."),
    (6, "Beta used", 1.50, "0.00", "in", "Assumption: raw beta ~2.1 reflects NVDA's momentum-driven volatility; 1.50 = blend of Blume-adjusted beta (~1.73) and semiconductor bottom-up beta (~1.3). Raw beta would give CoE ~15%."),
    (7, "Equity risk premium", 0.05, PCT, "in", "Assumption: 5.0% mature-market ERP (Damodaran implied ERP range 4.3-5.5%)."),
]
for r, lab, v, fmt, k, n in rows:
    label(w, r, lab)
    put(w, f"C{r}", v, k, fmt, n)
label(w, 8, "Cost of equity", bold=True)
put(w, "C8", "=C4+C6*C7", fmt=PCT, bold=True)
box(w, 3, 8, 1, 4)
section(w, 10, "COST OF DEBT", 1, 4)
label(w, 11, "Pre-tax cost of debt")
put(w, "C11", 0.05, "in", PCT, "Assumption: AA-/Aa2-rated issuer; ~10Y UST + 25-50bp => ~5.0% on recently issued $25B senior notes.")
label(w, 12, "Tax rate")
put(w, "C12", "=DCF!C18", "link", PCT)
label(w, 13, "After-tax cost of debt", bold=True)
put(w, "C13", "=C11*(1-C12)", fmt=PCT, bold=True)
box(w, 10, 13, 1, 4)
section(w, 15, "CAPITAL STRUCTURE (market values, gross debt)", 1, 4)
for r, lab, f, fmt, k in [
    (16, "Share price", "=DCF!C8", PX, "link"), (17, "Diluted shares (M)", "=DCF!C9", NUM, "link"),
    (18, "Market capitalization ($M)", "=C16*C17", NUM, "fx"), (19, "Total debt ($M)", "=DCF!C13", NUM, "link"),
    (20, "Total capital ($M)", "=C18+C19", NUM, "fx"),
]:
    label(w, r, lab)
    put(w, f"C{r}", f, k, fmt)
label(w, 22, "Equity weight"); put(w, "C22", "=C18/C20", fmt=PCT)
label(w, 23, "Debt weight"); put(w, "C23", "=C19/C20", fmt=PCT)
box(w, 15, 23, 1, 4)
section(w, 24, "WACC", 1, 4)
label(w, 25, "WACC (before scenario adjustment)", bold=True)
put(w, "C25", "=C22*C8+C23*C13", fmt=PCT, bold=True)
w["C25"].fill = FILL_OUT
w["A27"] = "Note: gross (not net) debt weights used; NVDA is net cash, so net-debt weights would be negative. Cash is added back in the equity bridge."
w["A27"].font = Font(name="Arial", size=9, italic=True)
box(w, 24, 25, 1, 4)
w.column_dimensions["A"].width = 40
w.column_dimensions["C"].width = 14

# ---------- Comps sheet ----------
c = comps
c["A1"] = "AI / SEMICONDUCTOR PEERS - COMPARABLE COMPANY ANALYSIS"
c["A1"].font = F_TITLE
c["A2"] = "Broadcom (AVGO) • AMD (AMD) • TSMC (TSM) • Marvell (MRVL) • Qualcomm (QCOM)"
c["A3"] = "As of early/mid Nov-2026 | Market cap in $B | multiples from data aggregators (S&P/FactSet MCP and SEC EDGAR were not reachable from this environment)"
for r in (2, 3):
    c[f"A{r}"].font = Font(name="Arial", size=9, italic=True)
section(c, 5, "VALUATION MULTIPLES", 1, 6)
subhdr(c, 6, ["Company", "Ticker", "Market cap ($B)", "EV / LTM EBITDA", "P/E (LTM)", "P/E (NTM)"], 1)
peers = [
    ("Broadcom", "AVGO", 1625, 23.0, 38.0, 17.9, "Mkt cap: Macrotrends ($1.60-1.65T). EV/EBITDA: Multiples.vc (LTM 23x). P/E LTM: Macrotrends; NTM: InvestingLive."),
    ("AMD", "AMD", 900, 73.0, None, 31.0, "Mkt cap: Motley Fool (~$900B). EV/EBITDA: Multiples.vc (73x LTM; depressed EBITDA from Xilinx amortization). NTM P/E: InvestingLive. LTM P/E not found."),
    ("TSMC", "TSM", 1975, 16.0, 21.0, 17.9, "Mkt cap: Motley Fool / Multiples.vc ($1.9-2.05T). EV/EBITDA: Multiples.vc. P/E: MarketScreener (LTM), InvestingLive (NTM)."),
    ("Marvell", "MRVL", 149, 54.0, 56.0, 23.0, "Mkt cap: Macrotrends ($148-150B). EV/EBITDA: Multiples.vc. P/E: Macrotrends (LTM), InvestingLive (NTM)."),
    ("Qualcomm", "QCOM", 183, 13.0, 14.0, None, "Mkt cap: Macrotrends ($177-190B). EV/EBITDA: Multiples.vc. P/E LTM: Macrotrends. NTM P/E not found."),
]
for i, (n, t, mc, ev, pe, fpe, note) in enumerate(peers):
    r = 7 + i
    c[f"A{r}"] = n
    put(c, f"B{r}", t)
    put(c, f"C{r}", mc, "in", NUM, "Source: " + note)
    for col, v in (("D", ev), ("E", pe), ("F", fpe)):
        if v is None:
            put(c, f"{col}{r}", "n/a")
        else:
            put(c, f"{col}{r}", v, "in", MULT)
stats = [("Maximum", "MAX"), ("75th percentile", "Q3"), ("Median", "MEDIAN"), ("25th percentile", "Q1"), ("Minimum", "MIN")]
for i, (lab, fn) in enumerate(stats):
    r = 12 + i
    c[f"A{r}"] = lab
    c[f"A{r}"].font = F_B
    for col in "DEF":
        rng = f"{col}7:{col}11"
        f = {"Q3": f"=QUARTILE({rng},3)", "Q1": f"=QUARTILE({rng},1)"}.get(fn, f"={fn}({rng})")
        put(c, f"{col}{r}", f, fmt=MULT)
    for col in "ABCDEF":
        c[f"{col}{r}"].fill = PatternFill("solid", fgColor="F2F2F2")
c["A18"] = "NVIDIA (market price)"
c["A18"].font = F_B
put(c, "B18", "NVDA")
put(c, "C18", "=DCF!C10/1000", "link", NUM)
put(c, "D18", "=(DCF!C10-DCF!C14)/DCF!E89", "link", MULT)
put(c, "E18", "=DCF!C10/DCF!E84", "link", MULT)
put(c, "F18", "=DCF!C10/DCF!F84", "link", MULT)
c["A19"] = "NVIDIA (DCF value)"
c["A19"].font = F_B
put(c, "C19", "=DCF!C113/1000", "link", NUM)
put(c, "D19", "=DCF!C111/DCF!E89", "link", MULT)
put(c, "E19", "=DCF!C113/DCF!E84", "link", MULT)
put(c, "F19", "=DCF!C113/DCF!F84", "link", MULT)
c["A20"] = "NVDA columns D-F use FY27E (Jan-27) EBITDA/NOPAT for LTM-ish and FY28E for NTM; NOPAT used as net income proxy (ignores interest/other income)."
c["A20"].font = Font(name="Arial", size=9, italic=True)
box(c, 5, 19, 1, 6)
section(c, 22, "NOTES & METHODOLOGY", 1, 6)
notes = [
    "Peer set: AI compute/networking silicon (AVGO, AMD, MRVL), foundry (TSM) and diversified fabless (QCOM). Intel excluded (depressed earnings).",
    "Multiples are as reported by aggregators with mixed as-of dates; treat as indicative. EV = market cap + net debt per provider.",
    "Medians use MEDIAN/QUARTILE over available values (n/a excluded). AMD/MRVL LTM EBITDA multiples inflated by acquisition amortization/low trough EBITDA.",
    "Use in DCF: peer median EV/EBITDA cross-checks DCF-implied entry and terminal multiples; peer NTM P/E cross-checks DCF equity / FY28E NOPAT.",
]
for i, t in enumerate(notes):
    c[f"A{23 + i}"] = "• " + t
    c[f"A{23 + i}"].font = Font(name="Arial", size=9)
c.column_dimensions["A"].width = 26
for col in "BCDEF":
    c.column_dimensions[col].width = 16

wb.save(OUT)
print("saved", OUT)
