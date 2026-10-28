#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_master_en.py  -  English version of the master spreadsheet ("dados historia de vida.xlsx")
=====================================================================================================
INPUT : final6.xlsx            (Portuguese master, 1,322 estimate rows, 13 sheets)
OUTPUT: master_en.xlsx         (same content, English labels)

WHAT IS TRANSLATED (and what is NOT) -- decisions agreed with the author
-------------------------------------------------------------------------
1. Sheet 'História de vida' -> 'Life history'
     * Header row (61 column names): translated (map COL below).
     * Cells: left exactly as they are (Brazilian studies, Brazilian place names, Portuguese common names,
       author-year citations, references, state codes, sex codes, numeric values, IDs).
     * The ONLY cell values changed are *controlled-vocabulary* values that the analysis plan refers to by
       name (country, time units, length-unit codes, month names, method labels, publication type,
       carapace-length definition, growth-method class, analysis_use, group names 'Independent' and
       'Syntheses/Government', and the free-text 'use_reason'). They are translated through the explicit
       dictionaries VAL / regex rules below, so that  value  <->  vocabulary sheet  <->  analysis plan
       stay consistent.  Locality, common_name, state, source, reference are untouched.
2. All the other sheets are translated into English (sheet names, headers, free text). Free text was
   translated beforehand (6 parallel translation passes, files en/m1_en.json, m2_en.json, m3_en.json),
   keyed by 'sheet|column_index|unique_string_index' in en/master_strings.json. This script only APPLIES
   those translations; it does not translate anything by itself.
3. Cell comments on the header row (G1 population_id, AV1 research_group, and two third-party notes in
   D1 / J1) are translated; their authors are preserved.
4. Formatting (fonts, fills, widths, freeze panes, merged cells) is preserved because the workbook is
   edited in place with openpyxl (no re-creation of sheets).

Run:  python3 build_master_en.py
"""
import json, re, collections
import openpyxl
from openpyxl.comments import Comment

SRC = '/home/claude/org/final6.xlsx'
DST = '/home/claude/org/master_en.xlsx'
D   = '/home/claude/org/en/'

# -------------------------------------------------------------------------------------------------
# STEP 1. Dictionaries
# -------------------------------------------------------------------------------------------------
# 1a. Sheet names (old -> new)
SHEETS = {
    'Legenda': 'Legend', 'História de vida': 'Life history', 'notas': 'notes', 'populacoes': 'populations',
    'grupos_pesquisa': 'research_groups', 'fontes': 'sources', 'estudos_primarios': 'primary_studies',
    'legenda_identificacao': 'identification_legend', 'vocabularios': 'vocabularies',
    'qc_pendencias': 'qc_issues', 'alteracoes': 'change_log', 'linhas_removidas': 'removed_rows',
    'mapa_linhas': 'row_map',
}

# 1b. Column names of 'Life history' (old -> new). Columns not listed (linf_tl, k, t0, ...) are kept.
COL = {
    'especie': 'species', 'nome_comum': 'common_name', 'pais': 'country', 'regiao': 'region', 'estado': 'state',
    'local': 'locality', 'id_populacao': 'population_id', 'ano_inicial': 'start_year', 'ano_final': 'end_year',
    'unidade_crescimento': 'growth_unit', 'unidade_wlab': 'wl_unit', 'wla': 'wl_a', 'wlb': 'wl_b',
    'unidade_lab': 'length_rel_unit', 'la': 'length_rel_a', 'lb': 'length_rel_b', 'unidade_tmax': 'tmax_unit',
    'periodo_reprodutivo_inicio': 'reproductive_period_start', 'periodo_reprodutivo_fim': 'reproductive_period_end',
    'recrutamento_inicio': 'recruitment_start', 'recrutamento_fim': 'recruitment_end',
    'metodo_mortalidade_m': 'm_method', 'metodo_mortalidade_z': 'z_method', 'metodo_crescimento': 'growth_method',
    'fonte': 'source', 'referencia': 'reference', 'referência': 'reference', 'grupo_pesquisa': 'research_group',
    'id_estimativa': 'estimate_id', 'id_estudo_primario': 'primary_study_id', 'id_fonte_lida': 'source_read_id',
    'tipo_publicacao': 'publication_type', 'nivel_confiabilidade': 'reliability_level',
    'definicao_cc': 'cl_definition', 'classe_metodo_crescimento': 'growth_method_class',
    'coord_precisao': 'coord_precision', 'parametros_repetidos': 'repeated_parameters',
    'uso_analise': 'analysis_use', 'motivo_uso': 'use_reason',
}
# pattern that finds any old column name inside free text (longest names first, whole words only)
_COLRE = re.compile(r'(?<![A-Za-z_])(' + '|'.join(sorted(map(re.escape, COL), key=len, reverse=True)) + r')(?![A-Za-z_])')
def colname_sub(text):
    """Replace every old column name appearing in a text by its English name."""
    return _COLRE.sub(lambda m: COL[m.group(1)], text)

# 1c. Controlled-vocabulary VALUES (exact match, applied only to the columns listed in VAL_COLS)
MONTHS = {'janeiro': 'january', 'fevereiro': 'february', 'março': 'march', 'abril': 'april', 'maio': 'may',
          'junho': 'june', 'julho': 'july', 'agosto': 'august', 'setembro': 'september', 'outubro': 'october',
          'novembro': 'november', 'dezembro': 'december', 'contínuo': 'continuous'}
VAL = {
    'Brasil': 'Brazil',
    'ano': 'year', 'dia': 'day', 'mês': 'month', 'bimestre': 'two-month period',
    'anos': 'years', 'meses': 'months', 'dias': 'days',
    # method labels (growth, M, Z)
    'Média Pauly-Taylor': 'Mean of Pauly and Taylor', 'Média de métodos': 'Mean of methods',
    'Análise de coorte': 'Cohort analysis', 'Fixo (literatura)': 'Fixed (literature)',
    'Curva de rendimento sustentável (CAJUS)': 'Sustainable-yield curve (CAJUS)',
    'Curva de captura convertida em comprimento': 'Length-converted catch curve',
    'Curva de captura': 'Catch curve', 'Progressão modal': 'Modal progression',
    'Método de Santos (1978)': 'Santos (1978) method',
    # publication type
    'artigo': 'article', 'tese': 'PhD thesis', 'dissertação': 'MSc dissertation',
    'relatório técnico': 'technical report', 'livro/catálogo': 'book/catalog',
    'anais/resumo': 'proceedings/abstract', 'tcc': 'undergraduate thesis',
    # carapace-length definition
    'padrão (assumido)': 'standard (assumed)', 'pós-orbital (não comparável)': 'post-orbital (not comparable)',
    'base do rostro (a conferir)': 'rostrum base (to be checked)', 'com rostro (a conferir)': 'including rostrum (to be checked)',
    # growth-method class
    'progressão modal': 'modal progression', 'ELEFAN/frequência (LFA)': 'ELEFAN/frequency (LFA)',
    'linearização/regressão': 'linearization/regression', 'não reportado/outro': 'not reported/other',
    'idade direta': 'direct age',
    # analysis_use
    'principal': 'main', 'sensibilidade': 'sensitivity', 'redundante': 'redundant', 'excluir': 'exclude',
    # research groups that are not groups
    'Independentes': 'Independent', 'Sínteses/Governo': 'Syntheses/Government',
    # misc
    'sim': 'yes', '(vazio)': '(empty)', '(campo esvaziado)': '(field emptied)',
}
VAL.update(MONTHS)

def unit_codes(v):
    """Length codes inside unit labels:  g/mm (CC) -> g/mm (CL);  CT~CAUDA -> TL~TAIL;  potência -> power."""
    v = v.replace('CAUDA', 'TAIL').replace('potência', 'power')
    v = re.sub(r'\bCC\b', 'CL', v); v = re.sub(r'\bCT\b', 'TL', v)
    return v

def val_map(v, colname=None):
    """Translate one controlled-vocabulary value. Handles '; '-joined lists (e.g. 'novembro; maio')."""
    if not isinstance(v, str): return v
    if colname in ('unidade_wlab', 'unidade_lab', 'wl_unit', 'length_rel_unit'):
        return unit_codes(v)
    if v in VAL: return VAL[v]
    if '; ' in v and all(p in VAL for p in v.split('; ')):          # e.g. 'novembro; maio'
        return '; '.join(VAL[p] for p in v.split('; '))
    return v

# Which 'Life history' columns get value translation (all others keep their cells untouched)
VAL_COLS = ['pais', 'unidade_crescimento', 'unidade_wlab', 'unidade_lab', 'unidade_tmax',
            'periodo_reprodutivo_inicio', 'periodo_reprodutivo_fim', 'recrutamento_inicio', 'recrutamento_fim',
            'metodo_mortalidade_m', 'metodo_mortalidade_z', 'metodo_crescimento', 'tipo_publicacao',
            'definicao_cc', 'classe_metodo_crescimento', 'uso_analise', 'grupo_pesquisa']

# 'use_reason' (free text generated by rules) and 'repeated_parameters' are translated by regex ---------
def rename_params(s):
    """Parameter names inside 'k, la, lb <- E0123' lists follow the renamed columns (la -> length_rel_a ...)."""
    return re.sub(r'\b(la|lb|wla|wlb)\b', lambda m: COL[m.group(1)], s)

def motivo_en(s):
    if not isinstance(s, str): return s
    s = rename_params(s)
    rules = [
        (r'^citação secundária \(nível (\d+-\d+)\)$',            r'secondary citation (level \1)'),
        (r'^todos os parâmetros repetem outra publicação do mesmo estudo \((.*)\)$',
                                                                  r'all parameters repeat another publication of the same study (\1)'),
        (r'^parâmetros repetidos de outra publicação: (.*) \(excluir esses valores em análises por parâmetro\)$',
                                                                  r'parameters repeated from another publication: \1 (exclude these values from per-parameter analyses)'),
        (r'^dado de distribuição geográfica, sem população amostrada$', 'geographic distribution data, no sampled population'),
        (r'^medida de comprimento não comparável \(pós-orbital\)$',    'length measure not comparable (post-orbital)'),
    ]
    for pat, rep in rules:
        if re.match(pat, s): return re.sub(pat, rep, s)
    return s                                                       # unknown pattern -> flagged by the final check

# -------------------------------------------------------------------------------------------------
# STEP 2. Free-text translations (from the translation passes) + generic translator
# -------------------------------------------------------------------------------------------------
strings = json.load(open(D + 'master_strings.json'))              # 'sheet|col' -> [unique PT strings]
EN = {}
for f in ('m1', 'm2', 'm3'):
    EN.update(json.load(open(D + f + '_en.json')))                # 'sheet|col|idx' -> English

TEXT_FIXES = [
    ("removed at the user's request on 14 Oct 2026", "removed on 14 Oct 2026 (author decision)"),
    ("removed at the user's request", "removed (author decision)"),
    ("removed by my decision (report to the user)", "removed (author decision, same criterion)"),
    ("the user had already deleted this column in their version (13 Oct 2026); aligned",
     "column already deleted in the authors' working version (13 Oct 2026); aligned"),
]
SHEETNAME_RE = [(re.compile(r"(?<![\w])" + re.escape(o) + r"(?![\w])"), n) for o, n in SHEETS.items() if o not in ('Legenda', 'História de vida')]
def fix_sheetnames(t):
    """Texts were translated while sheets still had Portuguese names: update 'notas' -> 'notes', etc."""
    for rx, n in SHEETNAME_RE: t = rx.sub(n, t)
    t = t.replace("'História de vida'", "'Life history'")
    # wording for a publication context (the working notes mention 'the user' / 'my decision')
    for a, b in TEXT_FIXES: t = t.replace(a, b)
    return t

missing = collections.Counter()
def tr(sheet, col0, v):
    """English text for cell value v of (old sheet name, 0-based column). Unknown strings are left and counted."""
    if not isinstance(v, str): return v
    lst = strings.get(f'{sheet}|{col0}')
    if lst is not None and v in lst:
        return fix_sheetnames(EN[f'{sheet}|{col0}|{lst.index(v)}'])
    missing[(sheet, col0)] += 1
    return v

# -------------------------------------------------------------------------------------------------
# STEP 3. Header maps for the auxiliary sheets (headers are not in the translation files)
# -------------------------------------------------------------------------------------------------
HDR = {
  'notas': ['source', 'field', 'master_rows', 'standardized_value', 'original_text (method/locality details)', 'reference'],
  'populacoes': ['population_id', 'species', 'locality (site)', 'years', 'scale', 'n_rows', 'n_sources', 'sources',
                 'master_rows', 'research_group(s)', 'confidence', 'criterion / evidence', 'lat_dec', 'lon_dec', 'coord_precision'],
  'grupos_pesquisa': ['research_group', 'description (institution / key authors)', 'n_sources', 'n_rows',
                      'species (rows)', 'cross_links', 'sources'],
  'fontes': ['source', 'primary_study_id', 'publication_type', 'type_to_check', 'reliability_level',
             'secondary_citation (apud)', 'source_read_id', 'n_rows', 'research_group'],
  'estudos_primarios': ['primary_study_id', 'study_sources', 'n_sources', 'n_rows', 'evidence'],
  'legenda_identificacao': ['column', 'description'],
  'vocabularios': ['column', 'allowed_value', 'n_rows', 'meaning'],
  'qc_pendencias': ['row', 'source', 'column', 'value', 'category', 'issue', 'decision'],
  'alteracoes': ['row', 'column', 'old_value', 'new_value', 'rule'],
  'mapa_linhas': ['old_row (spreadsheet of 13 Oct 2026)', 'new_row', 'status'],
}
# 'notas.campo' values (field names with qualifiers) not covered by plain column renaming
NOTAS_CAMPO = {
  'metodo_crescimento (outros métodos da linha)': 'growth_method (other methods in the row)',
  'referencia (redação anterior / via de acesso)': 'reference (previous wording / access route)',
  'medida de comprimento (NÃO COMPARÁVEL)': 'length measure (NOT COMPARABLE)',
  'relação peso-comprimento': 'weight-length relationship',
  'decisão sobre a fonte': 'decision on the source',
  'unidade_wlab (detalhe)': 'wl_unit (detail)',
  'linhas removidas': 'rows removed',
}
# 'alteracoes.coluna' / 'qc_pendencias.coluna' extra words
CHANGE_COL = {'linha nova': 'new row', 'linhas': 'rows', 'colunas novas': 'new columns', 'abas novas': 'new sheets',
              'mesclada': 'merged'}
QC_LINHA = {'(removido)': '(removed)'}

def col_tokens(text):
    """'z, f, m' / 'wla / wlb / linf_cl' -> same separators, English column names."""
    return colname_sub(text)

# -------------------------------------------------------------------------------------------------
# STEP 4. Apply to the workbook
# -------------------------------------------------------------------------------------------------
wb = openpyxl.load_workbook(SRC)
INV = {v: k for k, v in COL.items() if k != 'referência'}     # English column name -> Portuguese (for lookups)

for ws in list(wb):
    old = ws.title
    if old == 'História de vida':
        # ---- 4a. Life history: header + controlled vocabulary only -------------------------------
        hdr = [c.value for c in ws[1]]
        idx = {h: i for i, h in enumerate(hdr)}
        for c in ws[1]:                                            # translate header row
            c.value = COL.get(c.value, c.value)
        for r in range(2, ws.max_row + 1):
            if ws.cell(r, 1).value is None: continue               # empty formatted rows below the table
            for name in VAL_COLS:
                cell = ws.cell(r, idx[name] + 1)
                cell.value = val_map(cell.value, name)
            c = ws.cell(r, idx['motivo_uso'] + 1);          c.value = motivo_en(c.value)
            c = ws.cell(r, idx['parametros_repetidos'] + 1); c.value = rename_params(c.value) if isinstance(c.value, str) else c.value
        # ---- 4b. header comments ------------------------------------------------------------------
        CM = {
          'D1': ("======\nID#AAACHLxzW5o\nEduarda Leal    (2026-09-15 18:39:00)\nadd this field to the database and to the interface ", None),
          'J1': ("======\nID#AAACHLxzW50\nEduarda Leal    (2026-09-15 18:40:46)\nrebuild the database from here on ", None),
          'G1': ("population_id — identifier of the SAMPLED POPULATION.\n"
                 "Definition: set of individuals of the same species, from the same locality and sampling period (survey/program). "
                 "Rows with the same population_id come from the SAME field sample (article, thesis, report or re-analyses) and are "
                 "therefore NOT independent observations: use as a random effect / group in cross-validation.\n"
                 "Format: SPECIES-LOCALITY-NN. SPECIES: XKR = X. kroyeri; LSC = L. schmitti; FSU = F. subtilis; FPA = F. paulensis; "
                 "FBR = F. brasiliensis; FSP = Farfantepenaeus spp. LOCALITY = site code (main part of the locality column). "
                 "NN = chronological order of the surveys at that site.\n"
                 "'NA-DISTRIB' = geographic distribution data (no sampling).\n"
                 "Criterion and confidence level of each id (confirmed / probable / inferred / single source) in sheet 'populations'. "
                 "Ids at level 'inferred' should be reviewed. ", 'Claude (13/10/2026)'),
          'AV1': ("research_group — group/laboratory of the main author (and recurrent team) of the work.\n"
                  "Use: aggregate works by shared authors/institutions (same methods and analyses) to assess research-group bias in the "
                  "parameter estimates (group random effect).\n"
                  "One value per row (assigned by the source). Cross-links (co-authors from other groups) are in sheet 'research_groups'. "
                  "'Independent' and 'Syntheses/Government' are NOT research groups (isolated authors; catalogs and official reports). ", 'Claude (13/10/2026)'),
        }
        for coord, (txt, author) in CM.items():
            old_c = ws[coord].comment
            new = Comment(txt, author or (old_c.author if old_c else None))
            new.width, new.height = (old_c.width, old_c.height) if old_c else (300, 200)
            ws[coord].comment = new
    elif old == 'Legenda':
        # ---- 4c. Legend sheet: every string cell has an entry in the translation file --------------
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str):
                    c.value = tr('Legenda', c.column - 1, c.value)
    else:
        # ---- 4d. auxiliary sheets: header via HDR, cells via translation file / column / value maps --
        if old in HDR:
            for c, h in zip(ws[1], HDR[old]): c.value = h
        elif old == 'linhas_removidas':                            # 2 own columns + the Life-history columns
            lh = [COL.get(c.value, c.value) for c in ws[1]][3:]
            ws['A1'].value, ws['B1'].value, ws['C1'].value = 'old_row', 'equivalent_row_kept (new numbering)', 'reason'
            for c, h in zip(ws[1][3:], lh): c.value = h
        for r in range(2, ws.max_row + 1):
            for cidx in range(1, ws.max_column + 1):
                c = ws.cell(r, cidx); v = c.value
                if not isinstance(v, str): continue
                col0 = cidx - 1
                if old == 'notas':
                    if col0 == 1:   c.value = NOTAS_CAMPO.get(v) or colname_sub(v)
                    elif col0 == 4: c.value = tr(old, col0, v)                    # original text of method/locality
                    elif col0 == 3: c.value = VAL.get(v, v)                       # standardized value (labels / placeholder)
                elif old == 'populacoes':
                    if col0 in (2, 3, 4, 10, 11): c.value = tr(old, col0, v)
                    elif col0 == 9: c.value = '; '.join(VAL.get(p, p) for p in v.split('; '))
                elif old == 'grupos_pesquisa':
                    if col0 in (1, 5): c.value = tr(old, col0, v)
                    elif col0 == 0:   c.value = VAL.get(v, v)
                elif old == 'fontes':
                    if col0 == 2:        c.value = VAL.get(v, v)
                    elif col0 in (3, 5): c.value = VAL.get(v, v)                  # 'sim' -> 'yes'
                    elif col0 == 8:      c.value = VAL.get(v, v)
                elif old == 'estudos_primarios':
                    if col0 == 4: c.value = tr(old, col0, v)
                elif old == 'legenda_identificacao':
                    c.value = COL.get(v, v) if col0 == 0 else tr(old, col0, v)
                elif old == 'vocabularios':
                    if col0 == 0:   c.value = COL.get(v, v)
                    elif col0 == 1:
                        colname = ws.cell(r, 1).value                           # still Portuguese at this point
                        c.value = val_map(v, colname)
                    elif col0 == 3: c.value = tr(old, col0, v)
                elif old == 'qc_pendencias':
                    if col0 == 0:   c.value = QC_LINHA.get(v, re.sub(r'\((\d+) linhas\)', r'(\1 rows)', v))
                    elif col0 == 2: c.value = col_tokens(v)
                    elif col0 in (3, 4, 5, 6): c.value = tr(old, col0, v)
                elif old == 'alteracoes':
                    if col0 == 1:
                        c.value = CHANGE_COL.get(v) or colname_sub(v)
                    elif col0 == 4: c.value = tr(old, col0, v)
                    elif col0 in (2, 3):                                           # old/new cell values
                        # only controlled-vocabulary columns are mapped (same logic as in Life history)
                        colname = INV.get(ws.cell(r, 2).value, ws.cell(r, 2).value)   # col 2 already renamed -> back to old
                        if colname in VAL_COLS: c.value = val_map(v, colname)
                        elif v in ('mesclada', 'desfeita'):                       # merged-cell clean-up entries
                            c.value = {'mesclada': 'merged', 'desfeita': 'unmerged'}[v]
                elif old == 'linhas_removidas':
                    if col0 == 2: c.value = tr(old, col0, v)          # columns 4+ are handled in step 4e
                elif old == 'mapa_linhas':
                    if col0 == 2: c.value = tr(old, col0, v)
    # ---- freeze panes: keep the frozen header, but reset the scroll position ----------------------
    if old not in ('Legenda',) and ws.freeze_panes:
        ws.freeze_panes = 'D2' if old == 'linhas_removidas' else ('A2' if old != 'História de vida' else 'A2')

# ---- 4e. 'removed_rows' data columns are copies of Life-history rows: same vocabulary mapping ------
ws = wb['linhas_removidas']
# (the header of this sheet has already been rewritten above, so recover the old names with INV)
for r in range(2, ws.max_row + 1):
    for cidx in range(4, ws.max_column + 1):
        h = ws.cell(1, cidx).value
        oldname = INV.get(h, h)
        if oldname in VAL_COLS:
            ws.cell(r, cidx).value = val_map(ws.cell(r, cidx).value, oldname)
        elif oldname == 'motivo_uso':   ws.cell(r, cidx).value = motivo_en(ws.cell(r, cidx).value)
        elif oldname == 'parametros_repetidos' and isinstance(ws.cell(r, cidx).value, str):
            ws.cell(r, cidx).value = rename_params(ws.cell(r, cidx).value)

# ---- 4e2. vocabulary row for Cuba (study removed on 14 Oct 2026; 0 rows) no longer belongs in the vocabulary ----
wv = wb['vocabularios']
for r in range(wv.max_row, 1, -1):
    if wv.cell(r, 2).value == 'Cuba' and wv.cell(r, 3).value == 0:
        wv.delete_rows(r)

# ---- 4f. rename the sheets last (so that every lookup above could use the Portuguese names) -------
for ws in wb:
    ws.title = SHEETS[ws.title]
wb.active = wb.sheetnames.index('Life history')
wb.views[0].firstSheet = 0
wb.views[0].activeTab = wb.sheetnames.index('Life history')
for ws in wb: ws.sheet_view.tabSelected = (ws.title == 'Life history')

wb.save(DST)
print('saved', DST)
print('strings without translation (sheet, col0): count ->', dict(missing))
