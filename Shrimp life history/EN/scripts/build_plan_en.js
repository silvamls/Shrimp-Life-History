// build_plan_en.js - builds "analysis_plan_OSF_EN.docx" (pre-specified statistical analysis plan, English, for OSF registration)
// Each helper below creates one block type (heading, paragraph, bullet, code line, table). Grayscale palette, A4, 2 cm margins.
// Usage: node build_plan_en.js output.docx   (run where the 'docx' npm package is installed)
const fs = require('fs');
const {Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
       BorderStyle, AlignmentType, LevelFormat, HeadingLevel, Footer, Header, PageNumber} = require('docx');

const FONT = 'Arial';              // same family as the methods document
const INK = '262626', MID = '595959', RULE = '404040', LIGHT = 'E7E6E6', PALE = 'F2F2F2';

// ---- helpers ---------------------------------------------------------------
// text with inline **bold** (same convention as the methods document)
function runs(text, size = 19, extra = {}) {
  return text.split('**').filter(s => s !== '').map((s, i) => {
    const bold = text.startsWith('**') ? i % 2 === 0 : i % 2 === 1;
    return new TextRun({text: s, font: FONT, size, bold: bold || extra.bold, italics: extra.italics, color: extra.color || INK});
  });
}
const H1 = t => new Paragraph({heading: HeadingLevel.HEADING_1, keepNext: true, spacing: {before: 260, after: 100},
  children: [new TextRun({text: t, font: FONT, size: 26, bold: true, color: INK})]});
const H2 = t => new Paragraph({heading: HeadingLevel.HEADING_2, keepNext: true, spacing: {before: 160, after: 70},
  children: [new TextRun({text: t, font: FONT, size: 21, bold: true, color: INK})]});
const P = t => new Paragraph({alignment: AlignmentType.JUSTIFIED, spacing: {after: 100, line: 264}, children: runs(t)});
const NOTE = t => new Paragraph({alignment: AlignmentType.JUSTIFIED, spacing: {before: 60, after: 160}, children: runs(t, 17, {italics: true, color: MID})});
const B = t => new Paragraph({numbering: {reference: 'bul', level: 0}, alignment: AlignmentType.JUSTIFIED,
  spacing: {after: 50, line: 264}, children: runs(t, 19)});
// code block (model formulas), monospaced, light gray fill and a left rule
const CODE = t => new Paragraph({spacing: {after: 0}, indent: {left: 300},
  shading: {type: ShadingType.CLEAR, fill: PALE},
  border: {left: {style: BorderStyle.SINGLE, size: 12, space: 6, color: '7F7F7F'}},
  children: [new TextRun({text: t, font: 'Courier New', size: 14, color: INK})]});
const CAP = t => new Paragraph({keepNext: true, spacing: {before: 160, after: 70}, children: runs(t, 17, {bold: true})});

// table: top/bottom rules and light rules between rows (no vertical lines), light-gray header
const none = {style: BorderStyle.NONE, size: 0, color: 'FFFFFF'};
function table(widths, header, rows) {
  const mk = (txt, w, hdr, first, last) => new TableCell({
    width: {size: w, type: WidthType.DXA},
    borders: {left: none, right: none,
      top: hdr ? {style: BorderStyle.SINGLE, size: 8, color: RULE} : none,
      bottom: hdr ? {style: BorderStyle.SINGLE, size: 6, color: RULE}
                  : (last ? {style: BorderStyle.SINGLE, size: 8, color: RULE} : {style: BorderStyle.SINGLE, size: 2, color: 'BFBFBF'})},
    shading: {type: ShadingType.CLEAR, fill: hdr ? LIGHT : 'FFFFFF'},
    margins: {top: 50, bottom: 50, left: 80, right: 80},
    children: [new Paragraph({spacing: {after: 0, line: 228}, children: runs(txt, 16, {bold: hdr})})]});
  return new Table({width: {size: widths.reduce((a, b) => a + b), type: WidthType.DXA}, columnWidths: widths,
    rows: [new TableRow({tableHeader: true, cantSplit: true, children: header.map((h, i) => mk(h, widths[i], true))}),
      ...rows.map((r, n) => new TableRow({cantSplit: true, children: r.map((c, i) => mk(c, widths[i], false, false, n === rows.length - 1))}))]});
}

// ---- content -----------------------------------------------------------------
const C = [];
C.push(new Paragraph({spacing: {before: 0, after: 100},
  children: [new TextRun({text: 'Pre-specified statistical analysis plan', font: FONT, size: 36, bold: true, color: INK})]}));
C.push(new Paragraph({spacing: {after: 100}, children: runs('Variability of life-history parameters of penaeid shrimps exploited in Brazil: among species, regions, research groups and estimation methods', 22)}));
C.push(new Paragraph({spacing: {after: 60}, children: runs('**Version for registration at OSF** — 14 October 2026', 19)}));
C.push(new Paragraph({spacing: {after: 120}, border: {bottom: {style: BorderStyle.SINGLE, size: 8, space: 8, color: RULE}},
  children: runs('Responsible author: [FILL IN full name], PGRPAq/UFRPE; ORCID: [FILL IN]. Extraction validation: Anna Carolina Arnaldo de Figueiredo (A.C.A.F.), ORCID https://orcid.org/0000-0003-2569-972X.', 17, {italics: true, color: MID})}));
C.push(NOTE('Passages in square brackets ([FILL IN] and [CONFIRM]) depend on information held only by the author and are resolved before deposit. This document complements the methods document (search strategy, screening and extraction) and is deposited together with it.'));

C.push(H1('1. Statement on the timing of registration'));
C.push(P('The literature search and the data extraction were completed **before** this registration (extraction began in [FILL IN month/year]). The eligibility criteria were not changed after extraction began [CONFIRM]. Before registration the data were examined only descriptively (counts, cross-tabulations and quality control); no inferential model was fitted with the aim of testing the hypotheses below [CONFIRM]. The hypotheses, models and decision rules of this version were fixed on the basis of the structure of the database (number of estimates, studies and groups per stratum), which is known, and not on the results.'));
C.push(P('This is a **post-collection, pre-analysis** registration. OSF time-stamps it, and any later change is made through a new registered and justified version (Section 10), so that the reader can distinguish what was planned from what was explored.'));

C.push(H1('2. Objective and unit of analysis'));
C.push(P('To quantify the variability of life-history parameters (growth, maturity) of five penaeid shrimp species of Brazil and to **partition that variability** among species, regions, research groups, estimation methods and studies. The unit of analysis is the **estimate** (one row of the master spreadsheet: a value as published, for one sex, method and locality). The database has 1,322 estimates from 162 sources, 150 primary studies and 184 populations, each identified by estimate_id, primary_study_id, population_id and research_group. Rows are labeled main, sensitivity, redundant or exclude (column analysis_use) by mechanical rules described in the methods document.'));

C.push(H1('3. Hypotheses'));
C.push(P('The six confirmatory hypotheses below are tested separately in the two strata of length measure (carapace length, CL, and total length, TL), without conversion between them. Any analysis not listed here will be declared exploratory.'));
C.push(CAP('Table 1. Confirmatory hypotheses, responses and effects structure.'));
C.push(table([2900, 1500, 2600, 2038], ['Hypothesis', 'Response', 'Fixed effects', 'Random effects'], [
  ['**H1.** Species differ in growth. Expected direction: [FILL IN before registration].', 'log L∞; log k (bivariate model)', 'species; sex', 'study, population, group'],
  ['**H2.** Growth varies with latitude: higher k and lower L∞ at low (warmer) latitudes, after accounting for group and method. [CONFIRM direction]', 'log L∞; log k; φ′', 'species; latitude (per 10°, positive southward); sex', 'study, population, group'],
  ['**H3.** The estimation method changes the estimated k (modal progression vs. ELEFAN vs. linearization), within group.', 'log k; log L∞', 'method class; species; sex', 'study, population, group'],
  ['**H4.** There is variance attributable to study, population and research group beyond the residual variance (variance partitioning, not a test).', 'log L∞; log k; φ′', 'species', 'study, population, group'],
  ['**H5.** Parameters change with the year of data collection; within- and between-population change are separated.', 'log L∞; log k', 'species; population year (between); year − population year (within)', 'population (intercept and slope), study, group'],
  ['**H6.** L50 differs between sexes (females larger) and among species.', 'log L50 (CL)', 'species; sex', 'study, population, group'],
]));
C.push(NOTE('Exploratory analyses declared in advance: (E1) species × latitude interaction; (E2) log(Z/k), the only mortality contrast identifiable from length-based methods; (E3) region as a categorical factor in place of latitude. Natural mortality M is not used as a response (Section 6).'));

C.push(H1('4. Data available per analysis'));
C.push(P('The table below shows how many estimates, studies, populations and research groups each analysis will have, considering only rows with analysis_use = main. It was obtained from the master spreadsheet and precedes the running of any model.'));
C.push(CAP('Table 2. Sample size per analysis (main rows).'));
C.push(table([2700, 1200, 1000, 1200, 900, 2038], ['Analysis', 'Estimates', 'Studies', 'Populations', 'Groups', 'Note'], [
  ['L∞ in CL (standard definition)', '217', '31', '50', '12', 'North: 1 estimate'],
  ['L∞ in TL', '156', '17', '22', '12', 'covers the North (F. subtilis)'],
  ['k (all measures)', '324', '38', '64', '16', ''],
  ['L∞ in CL and k, simultaneous', '191', '29', '49', '11', 'F. subtilis: 4'],
  ['L50 in CL', '96', '39', '55', '14', 'X. kroyeri: 65'],
  ['Z and k, simultaneous', '51', '9', '16', '7', 'exploratory (E2)'],
  ['M estimated directly', '11', '3', '3', '3', 'not viable as a response'],
]));
C.push(new Paragraph({spacing: {before: 120, after: 0}, children: []}));
C.push(P('Two consequences of this table are decisions of the plan, made before seeing results. **First**: the L∞ analyses are separated into CL and TL, and the effect of the North is estimable only in the TL stratum, because almost all of the North is in TL and almost all of the rest of the database is in CL. **Second**: M is not analyzed as a response, because almost all M values in the database come from empirical equations (Pauly, Taylor and others) that are functions of k, L∞ and temperature; treating them as independent information would create pseudoreplication.'));

C.push(H1('5. Models'));
C.push(H2('5.1. Main model (multivariate Bayesian)'));
C.push(P('For H1 to H5, L∞ and k are modeled jointly, with an estimated residual correlation, on the log scale, by length-measure stratum. Random intercepts for primary study, population and research group are crossed, and the correlations between L∞ and k are also estimated at the level of each random intercept. The growth performance index φ′ = log10 k + 2 log10 L∞ is obtained as a derived contrast of the posterior distributions.'));
C.push(CODE('brm(bf(log_linf ~ species + sex + lat10 + (1|p|primary_study_id) + (1|q|population_id) + (1|r|research_group)) +'));
C.push(CODE('    bf(log_k    ~ species + sex + lat10 + (1|p|primary_study_id) + (1|q|population_id) + (1|r|research_group)) +'));
C.push(CODE('    set_rescor(TRUE), prior = c(prior(normal(0, 1), class = b),'));
C.push(CODE('    prior(student_t(3, 0, 1), class = sd), prior(lkj(2), class = cor)),'));
C.push(CODE('    chains = 4, iter = 4000, control = list(adapt_delta = 0.95))'));
C.push(new Paragraph({spacing: {before: 80, after: 0}, children: []}));
C.push(P('Sex enters with levels F, M and pooled; the 3 rows of indeterminate sex are excluded. For H6 the equivalent univariate model is used, with log L50 as the response. Convergence criteria: R-hat ≤ 1.01, effective sample size ≥ 400 and no divergent transitions; posterior predictive checks for each model.'));
C.push(H2('5.2. Verification model (frequentist)'));
C.push(P('Each equivalent univariate model is fitted with lme4 or glmmTMB (for example, lmer(log_linf ~ species + sex + lat10 + (1|primary_study_id) + (1|population_id) + (1|research_group))) as a check of the direction and order of magnitude of the effects. Singular fits are reported and do not replace the Bayesian model. P values are not reported for random effects.'));
C.push(H2('5.3. Inference'));
C.push(P('Fixed effects are reported as posterior estimates with 95% credible intervals and probability of direction; contrasts among species, with all pairs within H1, and confidence intervals adjusted by the Holm method in the frequentist check. Variance is reported as the standard deviation of each random effect and as a proportion of the total variance (variance-partition coefficients). No conclusion is drawn from the absence of an effect without reference to the interval.'));

C.push(H1('6. Exclusions and data rules'));
C.push(B('Only rows with analysis_use = main enter the main analyses; redundant, sensitivity and exclude rows are left out (Section 8).'));
C.push(B('Parameters flagged in repeated_parameters are omitted from the per-parameter analyses; the cited row keeps the value.'));
C.push(B('L∞ fixed by convention, M calculated by empirical equations and longevity calculated as 4.605/k are not used as independent responses.'));
C.push(B('L∞ in CL uses only cl_definition = standard; Ferreira (2013), with a post-orbital measure, is kept for the sensitivity analysis.'));
C.push(B('Missing data: complete-case analysis for each model; there is no imputation.'));

C.push(H1('7. Identifiability diagnostics and simplification rules'));
C.push(P('Research group, region and method are confounded in the database (13 of the 20 groups operate in a single macro-region). Before each model is interpreted, the following will be reported: cross-tabulations of group × region × species × method, the rank of the design matrix, and the sensitivity of the estimates to alternative priors (sd ~ exponential(2) and half-normal(0, 0.5)). If a fixed effect is not identifiable with the random effects of the model, it will be declared **not estimable**, and the group random effect will not be removed to make it estimable. Pre-specified order of simplification, applied only to the frequentist model when there is a singular fit: (1) remove random slopes; (2) remove the correlation between intercepts; (3) keep study and group and remove population.'));

C.push(H1('8. Sensitivity analyses'));
C.push(B('Reliability levels 1 to 3 (main) versus inclusion of levels 4 and 5.'));
C.push(B('One estimate per population (the one with the largest n) versus all.'));
C.push(B('Inclusion of the rows flagged as sensitivity and exclusion of rows with a quality-control flag.'));
C.push(B('Journal articles versus grey literature (theses, dissertations, reports).'));
C.push(B('Removal, one at a time, of each research group (leave-one-group-out).'));
C.push(B('Estimation-method class as a stratum.'));
C.push(B('For Xiphopenaeus, analysis by latitude band (taxonomic uncertainty before 2019).'));
C.push(B('Inverse-variance weighting in the subset with a reported standard error or confidence interval.'));

C.push(H1('9. Validation of the extraction, use of AI and availability'));
C.push(P('Screening and extraction were carried out by two human researchers (M.S.S. and A.C.A.F.), independently: each located and read the works and transcribed the values; each validated their own extraction with a large language model (Claude, Anthropic, in separate accounts); and the two sets of results were then compared independently, with the article re-read whenever the values did not match. The comparison covered all sources (162), not a sample; there were one or two divergent sources, in both cases because of internal inconsistency of the work itself (estimates that do not hold across text, tables and figures), handled by the rule of keeping the published value and flagging the inconsistency. Agreement is reported as the number of divergent sources, not as Cohen’s κ. The declaration of the use of computational assistance will accompany the manuscript according to the journal policy.'));
C.push(P('The database (with the change-log, removed-rows and identifier sheets), the organization scripts, the analysis plan and the model code will be deposited with a DOI and a CC-BY license.'));

C.push(H1('10. Changes to this plan'));
C.push(P('Any deviation (model, exclusion, hypothesis) will be recorded as a new version of this document, with date and justification, and the manuscript will indicate which results are confirmatory, which are exploratory and which result from deviations.'));

C.push(H1('References'));
['Bates, D.; Mächler, M.; Bolker, B.; Walker, S. 2015. Fitting linear mixed-effects models using lme4. Journal of Statistical Software 67(1): 1-48.',
 'Brooks, M.E. et al. 2017. glmmTMB balances speed and flexibility among packages for zero-inflated generalized linear mixed modeling. The R Journal 9(2): 378-400.',
 'Bürkner, P.-C. 2017. brms: an R package for Bayesian multilevel models using Stan. Journal of Statistical Software 80(1): 1-28.',
 'O’Dea, R.E. et al. 2021. Preferred reporting items for systematic reviews and meta-analyses in ecology and evolutionary biology: a PRISMA extension. Biological Reviews 96: 1695-1722.',
 'Page, M.J. et al. 2021. The PRISMA 2020 statement: an updated guideline for reporting systematic reviews. BMJ 372: n71.',
 'Pauly, D.; Munro, J.L. 1984. Once more on the comparison of growth in fish and invertebrates. Fishbyte 2(1): 1-21.'
].forEach(t => C.push(new Paragraph({alignment: AlignmentType.LEFT, spacing: {after: 80, line: 252},
  indent: {left: 340, hanging: 340}, children: runs(t, 18)})));

const doc = new Document({
  styles: {default: {document: {run: {font: FONT, size: 19}}},
    paragraphStyles: [
      {id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: {size: 26, bold: true, font: FONT, color: INK}, paragraph: {spacing: {before: 260, after: 100}, outlineLevel: 0}},
      {id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: {size: 21, bold: true, font: FONT, color: INK}, paragraph: {spacing: {before: 160, after: 70}, outlineLevel: 1}}]},
  numbering: {config: [{reference: 'bul', levels: [{level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT,
    style: {paragraph: {indent: {left: 560, hanging: 280}}}}]}]},
  sections: [{properties: {page: {size: {width: 11906, height: 16838}, margin: {top: 1134, right: 1134, bottom: 1021, left: 1134}}},
    headers: {default: new Header({children: [new Paragraph({children: [new TextRun({text: 'Pre-specified statistical analysis plan — penaeid life-history variability (OSF version, 14 Oct 2026)', font: FONT, size: 15, color: MID})]})]})},
    footers: {default: new Footer({children: [new Paragraph({alignment: AlignmentType.CENTER,
      children: [new TextRun({text: 'Page ', font: FONT, size: 16, color: MID}), new TextRun({children: [PageNumber.CURRENT], font: FONT, size: 16, color: MID})]})]})},
    children: C}]});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('ok'); });
