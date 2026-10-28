#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
prisma_flow.py - PRISMA 2020 flow diagram generated from a table of numbers
===========================================================================
Usage:   python3 prisma_flow.py prisma_flow_numbers.csv output_base
         (writes output_base.png, output_base.svg and output_base.pdf)

WHY A SCRIPT?
  The numbers in the diagram are ALWAYS the numbers of the tables in the manuscript: change one value in
  the CSV, run the script, and the whole diagram is redrawn. The script also CHECKS the arithmetic
  (screened = identified - duplicates, etc.) and refuses to draw if the numbers do not close, so a
  counting error cannot reach a reviewer.

CSV FORMAT (comma-separated, 3 columns: field,value,note). Fields:
  database:<name>           records retrieved from each database (one line per database)
  other:<name>              records from other methods (citation searching, grey literature, portals...)
  duplicates                records removed as duplicates before screening
  screening_excluded:<why>  records excluded at title/abstract screening, one reason per line
  not_retrieved             reports sought but not retrieved
  fulltext_excluded:<why>   reports excluded after full-text reading, one reason per line
  previous_base             (optional) reports of the INITIAL DATABASE (compilations made before the systematic
                            search); they enter directly at the inclusion stage ("previous studies" box of PRISMA 2020)
  previous_excluded:<why>   (optional) reports excluded from the initial database, one reason per line
  included_reports          number of reports included (= distinct sources in the master spreadsheet); must match the sum
  primary_studies           (optional) number of distinct primary studies (PRISMA separates "studies" from "reports")
  included_via_secondary    (optional) reports read only through a secondary citation (apud)
  draft                     (optional) 1 = stamps DRAFT on the diagram
  show_marks                (optional) 1 = mark each number with its origin (dagger = reported, double dagger = balance)
Column 3 (note) says where each number comes from; it is only used for the optional marks.
Lines that begin with # are ignored.
"""
import csv, sys, textwrap
import matplotlib
matplotlib.use("Agg")                                   # no-window backend (runs on a server)
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch

# ---------------------------------------------------------------- 1. read the CSV
src, out = sys.argv[1], sys.argv[2]
dbs, others, ft_reasons, scr_reasons, prev_ex = [], [], [], [], []
sc = {}                                                  # scalar fields
mark = {}                                                # field -> dagger / double dagger / "" (only if show_marks = 1)
for row in csv.reader(open(src, encoding="utf-8")):
    if not row or row[0].strip().startswith("#") or row[0].strip() == "field":
        continue
    field, value = row[0].strip(), row[1].strip()
    n = int(value.replace(",", "").replace(" ", ""))
    note = row[2].strip().lower() if len(row) > 2 else ""
    mark[field] = "†" if note.startswith("reported") else ("‡" if note.startswith("balance") else "")
    if field.startswith("database:"):             dbs.append((field[9:], n))
    elif field.startswith("other:"):              others.append((field[6:], n))
    elif field.startswith("fulltext_excluded:"):  ft_reasons.append((field[18:], n))
    elif field.startswith("screening_excluded:"): scr_reasons.append((field[19:], n))
    elif field.startswith("previous_excluded:"):  prev_ex.append((field[18:], n))
    else:                                         sc[field] = n

# ---------------------------------------------------------------- 2. PRISMA arithmetic
n_db, n_oth = sum(n for _, n in dbs), sum(n for _, n in others)
n_ident = n_db + n_oth                                   # identified (databases + other methods)
n_screened = n_ident - sc["duplicates"]                  # records entering title/abstract screening
n_scr_ex = sum(n for _, n in scr_reasons)                # excluded at screening
n_sought = n_screened - n_scr_ex                         # reports sought for retrieval
n_assessed = n_sought - sc["not_retrieved"]              # reports read in full
n_ft_ex = sum(n for _, n in ft_reasons)                  # excluded at full text
n_new = n_assessed - n_ft_ex                             # reports included from the search
n_prev = sc.get("previous_base", 0)                      # initial database (optional)
n_prev_ex = sum(n for _, n in prev_ex)
n_prev_in = n_prev - n_prev_ex
n_incl = n_new + n_prev_in                               # total reports included
# Check: the declared number of included reports (distinct sources of the master spreadsheet) must equal the sum
if n_incl != sc["included_reports"]:
    sys.exit(f"ERROR: arithmetic gives {n_incl} included reports ({n_new} from the search + {n_prev_in} from the "
             f"initial database) but the CSV declares {sc['included_reports']}. Fix the counts before drawing.")
for name, v in [("screened", n_screened), ("sought", n_sought), ("assessed", n_assessed), ("new", n_new), ("prev_in", n_prev_in)]:
    if v < 0: sys.exit(f"ERROR: negative value in {name}")

# ---------------------------------------------------------------- 3. drawing helpers
# Grayscale palette (print-friendly, journal-neutral)
INK, MID, LIGHT, PALE, BAND = "#262626", "#7F7F7F", "#EDEDED", "#F8F8F8", "#404040"
FONT = "Liberation Sans"                                 # Arial-compatible; falls back to DejaVu Sans if missing
plt.rcParams["font.family"] = [FONT, "DejaVu Sans"]

W = 8.0                                                  # figure width in inches
U = W / 100.0                                            # 1 drawing unit (x or y) in inches -> equal aspect
LINE = 2.35                                              # line height of body text in drawing units (8 pt font)
TLINE = 2.9                                              # line height of box titles (9 pt bold)
PAD = 1.6                                                # inner padding of boxes
CHARS = lambda w: int((w - 2 * PAD) / 0.80)              # approx. characters per line for a box of width w (8 pt)
fmt = lambda n: f"{n:,}"                                 # 1234 -> 1,234
mk = lambda field: mark.get(field, "") if sc.get("show_marks") else ""

boxes = []                                               # collected first, drawn after the figure height is known

def wrap(text, w):
    return textwrap.wrap(text, CHARS(w)) or [""]

class Box:
    """A box with a bold title and body lines. Height is computed from the text, so nothing overflows."""
    def __init__(self, x, w, title, lines=(), style="main", top=None):
        self.x, self.w, self.style = x, w, style
        self.title = sum([wrap(t, w - 2) for t in title.split("\n")], [])      # titles may wrap too
        self.lines = [ln for t in lines for ln in wrap(t, w)]
        self.h = 2 * PAD + TLINE * len(self.title) + (0.9 + LINE * len(self.lines) if self.lines else 0)
        self.top = top                                    # y of the top edge (set by the layout below)
    @property
    def bottom(self): return self.top - self.h
    @property
    def cx(self): return self.x + self.w / 2
    @property
    def cy(self): return self.top - self.h / 2

def draw_box(ax, b):
    fc, ec, lw = {"main": ("white", INK, 1.2), "side": (LIGHT, MID, 1.0), "final": ("#D9D9D9", INK, 1.8),
                  "prev": (PALE, INK, 1.2)}[b.style]
    # soft shadow, then the box itself
    ax.add_patch(FancyBboxPatch((b.x + 0.5, b.bottom - 0.5), b.w, b.h, boxstyle="round,pad=0,rounding_size=1.0",
                                fc="#00000018", ec="none", zorder=1))
    ax.add_patch(FancyBboxPatch((b.x, b.bottom), b.w, b.h, boxstyle="round,pad=0,rounding_size=1.0",
                                fc=fc, ec=ec, lw=lw, zorder=2))
    y = b.top - PAD
    for t in b.title:                                     # bold title, centered
        ax.text(b.cx, y, t, ha="center", va="top", fontsize=9, fontweight="bold", color=INK, zorder=3)
        y -= TLINE
    if b.lines:
        ax.plot([b.x + PAD, b.x + b.w - PAD], [y - 0.1, y - 0.1], color=MID, lw=0.5, zorder=3)   # thin rule under the title
        y -= 0.9
        for ln in b.lines:
            ax.text(b.x + PAD, y, ln, ha="left", va="top", fontsize=8, color=INK, zorder=3)
            y -= LINE

def arrow(ax, x1, y1, x2, y2):
    ax.annotate("", xy=(x2, y2), xytext=(x1, y1), zorder=4,
                arrowprops=dict(arrowstyle="-|>", color=INK, lw=1.1, shrinkA=0, shrinkB=0, mutation_scale=9))

def line(ax, xs, ys):
    ax.plot(xs, ys, color=INK, lw=1.1, zorder=4, solid_capstyle="butt")

# ---------------------------------------------------------------- 4. layout (top to bottom)
XL, WL = 8, 40                                           # left column: main flow
XR, WR = 54, 44                                          # right column: exclusions / side flows
GAP = 5.5                                                # vertical gap between rows (room for arrows)
y = 0.0                                                  # current top edge (decreases downwards)
rows = {}

# --- identification
b_db = Box(XL, WL, f"Records identified from\ndatabases (n = {fmt(n_db)})",
           [f"{nm}: n = {fmt(n)}{mk('database:' + nm)}" for nm, n in dbs])
b_ot = Box(XR, WR, f"Records identified from\nother methods (n = {fmt(n_oth)})",
           [f"{nm}: n = {fmt(n)}{mk('other:' + nm)}" for nm, n in others])
h1 = max(b_db.h, b_ot.h)
b_db.top = b_ot.top = y; b_db.h = b_ot.h = h1            # equal height for the two top boxes
y_join = y - h1 - 4.0                                    # horizontal junction line of the two flows
b_dup = Box(XR, WR, "Records removed before screening", [f"Duplicate records removed: n = {fmt(sc['duplicates'])}{mk('duplicates')}"], "side", top=y_join - 3.0)

# --- screening
y_scr = b_dup.bottom - GAP
b_scr = Box(XL, WL, f"Records screened\n(n = {fmt(n_screened)})", ["Title and abstract screening"], top=y_scr)
b_scr_ex = Box(XR, WR, f"Records excluded (n = {fmt(n_scr_ex)})",
               [f"{nm}: n = {fmt(n)}{mk('screening_excluded:' + nm)}" for nm, n in scr_reasons], "side", top=y_scr)
y_sought = min(b_scr.bottom, b_scr_ex.bottom) - GAP
b_sought = Box(XL, WL, f"Reports sought for retrieval\n(n = {fmt(n_sought)})", ["Search for the full text"], top=y_sought)
b_nr = Box(XR, WR, f"Reports not retrieved (n = {fmt(sc['not_retrieved'])}{mk('not_retrieved')})",
           ["Listed in Appendix B"], "side", top=y_sought)
y_ass = min(b_sought.bottom, b_nr.bottom) - GAP
b_ass = Box(XL, WL, f"Reports assessed for\neligibility (n = {fmt(n_assessed)})", ["Full-text reading"], top=y_ass)
b_ftx = Box(XR, WR, f"Reports excluded (n = {fmt(n_ft_ex)})",
            [f"{nm}: n = {fmt(n)}{mk('fulltext_excluded:' + nm)}" for nm, n in ft_reasons], "side", top=y_ass)

# --- included
y_inc = min(b_ass.bottom, b_ftx.bottom) - GAP
b_new = Box(XL, WL, f"Reports included from\nthe search (n = {fmt(n_new)})", ["Identified through databases and other methods"], "main", top=y_inc)
prev_lines = [f"Reports in the initial database: n = {fmt(n_prev)}{mk('previous_base')}"]
prev_lines += [f"Excluded ({nm}): n = {fmt(n)}{mk('previous_excluded:' + nm)}" for nm, n in prev_ex]
prev_lines += [f"Included from the initial database: n = {fmt(n_prev_in)}"]
b_prev = Box(XR, WR, f"Previous studies: initial database\n(compilations before the search)", prev_lines, "prev", top=y_inc)
y_fin = min(b_new.bottom, b_prev.bottom) - GAP
fin_lines = []
if "primary_studies" in sc: fin_lines.append(f"{fmt(sc['primary_studies'])} distinct primary studies{mk('primary_studies')}")
if "included_via_secondary" in sc: fin_lines.append(f"{sc['included_via_secondary']} reports accessed only through secondary citation (apud)")
b_fin = Box(XL, XR + WR - XL, f"Reports included in the review (n = {fmt(n_incl)})", fin_lines, "final", top=y_fin)
# the first-row boxes were created with top=None and set above; make all heights consistent for the combined heights
y_end = b_fin.bottom - 6.0                               # room for the footnote

# ---------------------------------------------------------------- 5. draw
H = -y_end
fig, ax = plt.subplots(figsize=(W, H * U))
ax.set_xlim(0, 100); ax.set_ylim(y_end, 2.0); ax.axis("off")
for b in [b_db, b_ot, b_dup, b_scr, b_scr_ex, b_sought, b_nr, b_ass, b_ftx, b_new, b_prev, b_fin]:
    draw_box(ax, b)

# phase bands on the left (as in PRISMA 2020): y ranges follow the boxes
def band(y_top, y_bot, label):
    ax.add_patch(FancyBboxPatch((0.6, y_bot), 3.6, y_top - y_bot, boxstyle="round,pad=0,rounding_size=1.0", fc=BAND, ec=BAND, zorder=2))
    ax.text(2.4, (y_top + y_bot) / 2, label, rotation=90, ha="center", va="center", color="white", fontsize=9.5, fontweight="bold", zorder=3)
band(b_db.top + 0.8, b_dup.bottom - 1.8, "Identification")
band(b_scr.top + 1.8, b_ftx.bottom - 0.2 if b_ftx.bottom < b_ass.bottom else b_ass.bottom - 0.2, "Screening")
band(b_new.top + 1.8, b_fin.bottom - 0.2, "Included")

# connectors
line(ax, [b_db.cx, b_db.cx], [b_db.bottom, y_join]); line(ax, [b_ot.cx, b_ot.cx], [b_ot.bottom, y_join])
line(ax, [b_db.cx, b_ot.cx], [y_join, y_join])                         # junction of the two flows
arrow(ax, b_ot.cx, y_join, b_ot.cx, b_dup.top)                          # to the duplicates box
arrow(ax, b_db.cx, y_join, b_db.cx, b_scr.top)                          # main flow continues
arrow(ax, b_scr.x + b_scr.w, b_scr.cy, b_scr_ex.x, b_scr.cy)            # screened -> excluded
arrow(ax, b_scr.cx, b_scr.bottom, b_sought.cx, b_sought.top)
arrow(ax, b_sought.x + b_sought.w, b_sought.cy, b_nr.x, b_sought.cy)
arrow(ax, b_sought.cx, b_sought.bottom, b_ass.cx, b_ass.top)
arrow(ax, b_ass.x + b_ass.w, b_ass.cy, b_ftx.x, b_ass.cy)
arrow(ax, b_ass.cx, b_ass.bottom, b_new.cx, b_new.top)
arrow(ax, b_new.cx, b_new.bottom, b_new.cx, b_fin.top)
arrow(ax, b_prev.cx, b_prev.bottom, b_prev.cx, b_fin.top)

if sc.get("draft"):
    ax.text(50, y_end / 2, "DRAFT", rotation=28, ha="center", va="center", fontsize=64, color="#C00000", alpha=0.12, fontweight="bold")
foot = "Adapted from PRISMA 2020 (Page et al. 2021, BMJ 372:n71)."
if sc.get("show_marks") and any(v for v in mark.values()):
    foot += "\n† reported by the searchers     ‡ balancing item (closes the total)     no mark = exact count in the spreadsheets"
ax.text(50, y_end + 4.0, foot, ha="center", va="top", fontsize=7.5, style="italic", color=MID, linespacing=1.5)

for ext in ("png", "svg", "pdf"):
    fig.savefig(f"{out}.{ext}", dpi=300, bbox_inches="tight", facecolor="white")
print(f"identified {n_ident} | screened {n_screened} | sought {n_sought} | assessed {n_assessed} | "
      f"included from search {n_new} | initial database {n_prev_in} | included total {n_incl}")
