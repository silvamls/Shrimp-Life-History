#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_methods_en.py - builds "search_methods_PRISMA_v3_EN.docx" from the English markdown of the methods document
==================================================================================================================
INPUT : en/meth_all_en.md           (English text, pandoc markdown; produced from the six translated chunks)
        en_out/prisma_flow_PRISMA2020.png   (Figure 1)
OUTPUT: en_out/search_methods_PRISMA_v3_EN.docx

STEPS
 1. Pre-process the markdown: real headings (so the Word navigation pane / table of contents work), typographic dashes
    ('---' -> em dash, '--' -> en dash, except in table separator lines), figure path.
 2. pandoc -> docx using a reference document whose styles were rewritten here (grayscale; ~2 cm margins; Arial/Liberation Sans).
 3. python-docx post-processing: grayscale tables (top/bottom rules, light header fill, small font, column widths proportional to
    the text length), captions kept with their tables, quoted search strings in a monospaced font, title block, header/footer
    with page numbers, automatic table of contents field (updated when the file is opened in Word).
"""
import re, subprocess, sys, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

SRC  = '/home/claude/org/en/meth_all_en.md'
FIG  = '/home/claude/org/en_out/prisma_flow_PRISMA2020.png'
OUT  = '/home/claude/org/en_out/search_methods_PRISMA_v3_EN.docx'
TMP  = '/home/claude/org/en_out/_tmp'
os.makedirs(TMP, exist_ok=True)

GRAY_DARK, GRAY_MID, GRAY_LIGHT = RGBColor(0x26, 0x26, 0x26), RGBColor(0x59, 0x59, 0x59), 'E7E6E6'
FONT = 'Liberation Sans'          # metric-compatible with Arial (Word shows Arial if Liberation is missing -> set both below)

# ---------------------------------------------------------------- 1. markdown pre-processing
lines = open(SRC, encoding='utf-8').read().split('\n')
out = []
for ln in lines:
    if re.match(r'^\|[\s\-:|]+\|$', ln):            # table separator line: leave untouched
        out.append(ln); continue
    ln = ln.replace(' --- ', ' — ').replace('---', '—').replace('--', '–')
    # numbered section headings written as bold lines -> real headings
    m = re.match(r'^\*\*(\d+\.\d+\.\s.+?)\*\*$', ln)
    if m: ln = '## ' + m.group(1)
    else:
        m = re.match(r'^\*\*(\d+\.\s.+?)\*\*$', ln)
        if m: ln = '# ' + m.group(1)
        else:
            m = re.match(r'^\*\*(Appendix [A-D]\s—.+?|Methodological references)\*\*$', ln)
            if m: ln = '# ' + m.group(1)
    ln = ln.replace('(PRISMA_FIG)', f'({FIG}){{width=17cm}}')
    out.append(ln)
md = '\n'.join(out)
# the very first bold lines are the title block: remove the bold markers so that they can be styled below
open(f'{TMP}/doc.md', 'w', encoding='utf-8').write(md)

# ---------------------------------------------------------------- 2. reference docx (styles)
ref = f'{TMP}/ref.docx'
with open(ref, 'wb') as f:
    f.write(subprocess.check_output(['pandoc', '-o', '/dev/stdout', '--print-default-data-file', 'reference.docx']))
d = Document(ref)
for s in d.sections:
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)                      # A4
    s.left_margin = s.right_margin = Cm(2.0); s.top_margin = Cm(2.0); s.bottom_margin = Cm(1.8)
def style_font(st, size, bold=None, color=GRAY_DARK, italic=None, name=FONT):
    st.font.name = name; st.font.size = Pt(size); st.font.color.rgb = color
    rpr = st.element.get_or_add_rPr(); rf = rpr.find(qn('w:rFonts'))
    if rf is None: rf = OxmlElement('w:rFonts'); rpr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'): rf.set(qn(a), name)
    for a in ('w:asciiTheme', 'w:hAnsiTheme', 'w:cstheme', 'w:eastAsiaTheme'):
        if rf.get(qn(a)) is not None: del rf.attrib[qn(a)]
    if bold is not None: st.font.bold = bold
    if italic is not None: st.font.italic = italic
def get_style(name):
    """python-docx matches style names strictly; pandoc's names differ in case ('heading 1'), so match case-insensitively."""
    for st_ in d.styles:
        if st_.name.lower() == name.lower(): return st_
    raise KeyError(name)
for name in ('Normal', 'Body Text', 'First Paragraph', 'Compact', 'Block Text', 'Table', 'Footnote Text', 'Image Caption', 'Caption'):
    try: style_font(get_style(name), 9.5)
    except KeyError: pass
for name, size, before, after in (('Heading 1', 13, 14, 4), ('Heading 2', 10.5, 9, 3), ('Heading 3', 10, 6, 2)):
    st = get_style(name); style_font(st, size, bold=True, color=GRAY_DARK)
    st.paragraph_format.space_before, st.paragraph_format.space_after = Pt(before), Pt(after)
    st.paragraph_format.keep_with_next = True
for name in ('Normal', 'Body Text', 'First Paragraph'):
    pf = get_style(name).paragraph_format; pf.space_before = Pt(0); pf.space_after = Pt(5); pf.line_spacing = 1.08
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if name != 'Normal' else None
st = get_style('Compact'); st.paragraph_format.space_after = Pt(1.5)
for name in ('Hyperlink',):
    try: get_style(name).font.color.rgb = GRAY_MID
    except KeyError: pass
d.save(ref)

# ---------------------------------------------------------------- 3. pandoc
raw = f'{TMP}/raw.docx'
subprocess.check_call(['pandoc', f'{TMP}/doc.md', '-f', 'markdown-smart+pipe_tables', '-o', raw,
                       '--reference-doc', ref, '--resource-path', '/home/claude/org/en_out'])

# ---------------------------------------------------------------- 4. post-processing with python-docx
doc = Document(raw)
body = doc.element.body

def set_cell_border(cell, **kw):
    tcPr = cell._tc.get_or_add_tcPr(); b = tcPr.find(qn('w:tcBorders'))
    if b is None: b = OxmlElement('w:tcBorders'); tcPr.append(b)
    for edge, (sz, color) in kw.items():
        e = b.find(qn(f'w:{edge}'))
        if e is None: e = OxmlElement(f'w:{edge}'); b.append(e)
        e.set(qn('w:val'), 'single' if sz else 'nil'); e.set(qn('w:sz'), str(sz)); e.set(qn('w:space'), '0'); e.set(qn('w:color'), color)

def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr(); sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), fill); tcPr.append(sh)

# 4a. tables ---------------------------------------------------------------------------------------
usable = 17.0                                    # cm between the margins
for t in doc.tables:
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    ncol = len(t.columns)
    # fixed layout + widths proportional to sqrt(mean text length) of each column (long text columns get more room)
    tblPr = t._tbl.tblPr
    lay = OxmlElement('w:tblLayout'); lay.set(qn('w:type'), 'fixed'); tblPr.append(lay)
    w = []
    for j in range(ncol):
        lens = [len(r.cells[j].text) for r in t.rows]
        w.append(max(4.0, (sum(lens) / len(lens)) ** 0.62 + 0.6 * max(lens) ** 0.4))
    tot = sum(w); widths = [usable * x / tot for x in w]
    widths = [max(1.3, x) for x in widths]; k = usable / sum(widths); widths = [x * k for x in widths]
    grid = t._tbl.find(qn('w:tblGrid'))
    for j, gc in enumerate(grid.findall(qn('w:gridCol'))): gc.set(qn('w:w'), str(int(widths[j] / 2.54 * 1440)))
    tw = tblPr.find(qn('w:tblW'))
    if tw is None: tw = OxmlElement('w:tblW'); tblPr.append(tw)
    tw.set(qn('w:type'), 'dxa'); tw.set(qn('w:w'), str(int(usable / 2.54 * 1440)))
    for i, row in enumerate(t.rows):
        trPr = row._tr.get_or_add_trPr(); cs = OxmlElement('w:cantSplit'); trPr.append(cs)       # rows are not split across pages
        if i == 0:
            th = OxmlElement('w:tblHeader'); trPr.append(th)                                      # header row repeats on each page
        for j, c in enumerate(row.cells):
            c.width = Cm(widths[j])
            tcPr = c._tc.get_or_add_tcPr(); mar = OxmlElement('w:tcMar')
            for side, v in (('top', 40), ('bottom', 40), ('left', 70), ('right', 70)):
                e = OxmlElement(f'w:{side}'); e.set(qn('w:w'), str(v)); e.set(qn('w:type'), 'dxa'); mar.append(e)
            tcPr.append(mar)
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(0); p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    r.font.size = Pt(8); r.font.name = FONT; r.font.color.rgb = GRAY_DARK
                    if i == 0: r.font.bold = True
            if i == 0:
                shade(c, GRAY_LIGHT); set_cell_border(c, top=(8, '404040'), bottom=(6, '404040'))
            else:
                set_cell_border(c, bottom=(2, 'BFBFBF') if i < len(t.rows) - 1 else (8, '404040'))
    # remove the vertical borders and style name colours that pandoc's default table style adds
    for c_ in [c for r in t.rows for c in r.cells]:
        set_cell_border(c_, left=(0, 'FFFFFF'), right=(0, 'FFFFFF'))

# small tables stay on one page: keep every row (except the last) with the next one
for t_ in doc.tables:
    if len(t_.rows) <= 12:
        for r_ in t_.rows[:-1]:
            for c_ in r_.cells:
                for p_ in c_.paragraphs: p_.paragraph_format.keep_with_next = True

# 4b. paragraphs: captions, block quotes, title block, figure ----------------------------------------
paras = doc.paragraphs
for i, p in enumerate(paras):
    txt = p.text.strip()
    if re.match(r'^(Table [A-D]?\d+\.|Figure \d+\.)', txt) and p.runs and p.runs[0].bold:
        p.paragraph_format.keep_with_next = True; p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs: r.font.size = Pt(9); r.font.color.rgb = GRAY_DARK
    if p.style.name == 'Block Text':                           # search strings: monospaced, left rule, light fill
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0); p.paragraph_format.left_indent = Cm(0.5); p.paragraph_format.line_spacing = 1.0
        for r in p.runs: r.font.name = 'Liberation Mono'; r.font.size = Pt(8)
        pPr = p._p.get_or_add_pPr(); bd = OxmlElement('w:pBdr'); l = OxmlElement('w:left')
        l.set(qn('w:val'), 'single'); l.set(qn('w:sz'), '12'); l.set(qn('w:space'), '6'); l.set(qn('w:color'), '7F7F7F'); bd.append(l); pPr.append(bd)
        sh = OxmlElement('w:shd'); sh.set(qn('w:val'), 'clear'); sh.set(qn('w:color'), 'auto'); sh.set(qn('w:fill'), 'F2F2F2'); pPr.append(sh)
    if p.style.name in ('Image Caption',):
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for r in p.runs: r.font.size = Pt(9)
    if p._p.xpath('.//w:drawing'):                             # the figure: centered, kept with its caption
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next = True

# title block = first 7 non-empty paragraphs (title, subtitle, document title, date, programme, university, italic note)
tb = [p for p in paras if p.text.strip()][:7]
sizes = [20, 11.5, 13, 10.5, 10, 10, 9.5]
for p, sz in zip(tb, sizes):
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT; p.paragraph_format.space_after = Pt(3)
    for r in p.runs: r.font.size = Pt(sz)
tb[0].paragraph_format.space_before = Pt(0); tb[0].paragraph_format.space_after = Pt(6)
# horizontal rule under the title block
pPr = tb[-1]._p.get_or_add_pPr(); bd = OxmlElement('w:pBdr'); b = OxmlElement('w:bottom')
b.set(qn('w:val'), 'single'); b.set(qn('w:sz'), '8'); b.set(qn('w:space'), '8'); b.set(qn('w:color'), '404040'); bd.append(b); pPr.append(bd)

# table of contents: a static list of the main sections (no page numbers, so nothing can go out of date)
anchor = tb[-1]._p
heads = [p_.text.strip() for p_ in doc.paragraphs if p_.style.name == 'Heading 1']
cont = doc.add_paragraph(); rr = cont.add_run('Contents'); rr.bold = True; rr.font.size = Pt(10.5); rr.font.color.rgb = GRAY_DARK
cont.paragraph_format.space_before = Pt(8); cont.paragraph_format.space_after = Pt(3); cont.paragraph_format.keep_with_next = True
items = [cont]
for h_ in heads:
    q = doc.add_paragraph(); r_ = q.add_run(h_); r_.font.size = Pt(9); r_.font.color.rgb = GRAY_DARK
    q.paragraph_format.space_after = Pt(1); q.paragraph_format.left_indent = Cm(0.3); q.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    items.append(q)
pb = doc.add_paragraph(); pb.add_run().add_break(__import__('docx').enum.text.WD_BREAK.PAGE); items.append(pb)
for q in reversed(items): anchor.addnext(q._p)

# spacing: paragraph right after a table, and the bold labels of the search-string blocks
for t_ in doc.tables:
    nxt = t_._tbl.getnext()
    if nxt is not None and nxt.tag == qn('w:p'):
        pPr_ = nxt.get_or_add_pPr() if hasattr(nxt, 'get_or_add_pPr') else None
from docx.text.paragraph import Paragraph
for t_ in doc.tables:
    nxt = t_._tbl.getnext()
    if nxt is not None and nxt.tag == qn('w:p'):
        Paragraph(nxt, doc._body).paragraph_format.space_before = Pt(7)
for p_ in doc.paragraphs:
    if re.match(r'^Block [ABC] ', p_.text.strip()):
        p_.paragraph_format.space_before = Pt(6); p_.paragraph_format.keep_with_next = True; p_.paragraph_format.space_after = Pt(2)

# 4c. header / footer ------------------------------------------------------------------------------------
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Cm(2.0); sec.top_margin = Cm(2.0); sec.bottom_margin = Cm(1.8)
hp = sec.header.paragraphs[0]; hp.text = 'Literature search, screening and source verification — penaeid life-history compilation (v3, English edition)'
for r in hp.runs: r.font.size = Pt(7.5); r.font.color.rgb = GRAY_MID; r.font.name = FONT
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run('Page '); r.font.size = Pt(8); r.font.color.rgb = GRAY_MID
for typ, txt in (('begin', None), (None, ' PAGE '), ('end', None)):
    rr = fp.add_run(); rr.font.size = Pt(8); rr.font.color.rgb = GRAY_MID
    if typ: fc = OxmlElement('w:fldChar'); fc.set(qn('w:fldCharType'), typ); rr._r.append(fc)
    else: it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = txt; rr._r.append(it)

# 4d. schema hygiene: put the children of tcPr / pPr / tblPr in the order required by the OOXML schema, and complete pgMar
ORDER = {
 'tcPr': ['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd','noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark'],
 'pPr': ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs',
         'suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd',
         'snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment',
         'textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'],
 'tblPr': ['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize','tblW','jc','tblCellSpacing',
           'tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription'],
 'tcMar': ['top','start','left','bottom','end','right'],
 'tcBorders': ['top','start','left','bottom','end','right','insideH','insideV','tl2br','tr2bl'],
 'trPr': ['cnfStyle','divId','gridBefore','gridAfter','wBefore','wAfter','cantSplit','trHeight','tblHeader','tblCellSpacing','jc','hidden'],
}
W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def reorder(root):
    for tag, order in ORDER.items():
        for el in root.iter('{%s}%s' % (W_NS, tag)):
            kids = list(el)
            key = lambda k: order.index(k.tag.split('}')[1]) if k.tag.split('}')[1] in order else len(order)
            for k in kids: el.remove(k)
            for k in sorted(kids, key=key): el.append(k)
reorder(doc.element)
for sct in doc.element.iter('{%s}pgMar' % W_NS):
    for a, v in (('gutter', '0'), ('footer', '567'), ('header', '567')):
        if sct.get(qn('w:' + a)) is None: sct.set(qn('w:' + a), v)
for sec_ in doc.sections:
    reorder(sec_.header._element); reorder(sec_.footer._element)

doc.core_properties.title = 'Literature search, screening and source verification (v3, English edition)'
doc.core_properties.author = 'M.S.S.; A.C.A.F.'
doc.save(OUT)
print('saved', OUT)
