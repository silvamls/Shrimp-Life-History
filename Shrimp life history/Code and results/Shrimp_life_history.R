# ======================================================================================================================
#  Shrimp_life_history.R
# ======================================================================================================================
#  TITLE     Variability of life-history parameters of penaeid shrimps exploited in Brazil: among species, latitude,
#            research groups and estimation methods
#  AUTHOR    Silva, M.L.S. (Matheus Lourenco Soares da Silva) - PGRPAq/UFRPE, Brazil - ORCID 0000-0001-6382-8698
#  CO-AUTHOR Figueiredo, A.C.A. (Anna Carolina Arnaldo de Figueiredo) - ORCID 0000-0003-2569-972X
#  CREATED   2026-10-16          LAST MODIFIED  2026-10-20          VERSION  0.3 (data checks, tables, figures, models)
#  LICENSE   CC-BY 4.0
#  Runs the analyses of the pre-specified statistical analysis plan (OSF: https://osf.io/vq2rx/). The goal is to quantify and
#  partition the variability of growth and size at maturity of five penaeid shrimp species of Brazil among species,
#  latitude, research groups, estimation methods, studies and populations. Analyses are observational: they describe
#  conditional associations and variance components, not causal effects.
#  INPUT
#  life_history_data_EN.xlsx, sheet "Life history" (unit of analysis = one estimate as published; 1,322 rows).
#  Data rules: keep analysis_use == "main"; set parameters listed in repeated_parameters to NA; convert k to year^-1
#  (month x 12, day x 365); phi' in CL only with cl_definition == "standard (assumed)"; drop growth_method_class ==
#  "not reported/other" from models with the method term; complete cases per model, no imputation.
#  RESPONSES
#  log k (primary; CL and TL rows together) | phi' = log10(k) + 2 log10(Linf) (secondary; CL and TL fitted separately)
#  | log L50 (CL and TL fitted separately). Linf is described but not modeled. M and Z are not modeled.
#  MODEL (brms, Gaussian, log scale; same structure for all responses)
#  response ~ species + sex + method + lat10 + (1 | research_group) + (1 | primary_study_id) + (1 | population_id)
#  lat10 = -lat_dec / 10 (positive southward, centered); method reference = modal progression (omitted for L50).
#  Priors: normal(0, 1) for fixed effects; student_t(3, 0, 1) for random-effect SDs and sigma.
#  Sampling: 4 chains, 4000 iterations, adapt_delta = 0.95. Convergence: R-hat <= 1.01, ESS >= 400, no divergences.
#  HYPOTHESES
#  H1 species differ in growth | H2 no relevant latitude effect on growth after group and method | H3 method changes k
#  H4 variance partitioning among group, study and population | H5 L50 differs between sexes (females larger)
#  H2 rule (ROPE = +/- 10% in k per 10 degrees = +/- 0.095 in ln k = +/- 0.041 in log10 phi'): 95% credible interval
#  inside the ROPE = practically null; crossing it = inconclusive; outside it = different from zero.
#  Exploratory: E1 year of data collection | E2 latitude slope for Xiphopenaeus kroyeri alone.
#  Sensitivity: sensitivity-flagged rows, one estimate per population, journals vs grey literature, leave-one-group-out,
#  alternative priors for the random-effect SDs.
#  SOFTWARE   R | readxl, dplyr, tidyr | brms (Stan) | posterior, bayesplot, emmeans, loo | ggplot2
#             Versions are recorded with sessionInfo() at the end of each run.
#  STRUCTURE  1 Data import and first look | 2 Data preparation | 3 Sample sizes | 4 Descriptive statistics
#             5 Imbalance and gaps | 6 Summary tables (xlsx) | 7 Figures | 8 Bayesian models (plan 5.1)
#             9 Variance partitioning (plan 5.2, H4) | 10 Inference and H2 equivalence test (plan 5.3)
#             11 Model checks | 12 Model tables (xlsx) | 13 Model figures
#  NOTE       Any deviation from the registered plan is recorded as a new dated version of the plan.
# ======================================================================================================================


#set work directory
setwd('C:/Matheus/Universidade/Posdoc/Camarada/Shrimp life history/Code and results')

#@ packages....
library(glmmTMB)      # fit models
library(readxl)       #read xlsx spreadsheet
library(ggplot2)      # figures
library(openxlsx)     # formatted Excel tables
library(patchwork)    # combines several ggplots into one figure
library(brms)         # Bayesian multilevel models (Stan back end)
library(posterior)    # summaries of posterior draws: R-hat, effective sample size
library(emmeans)      # contrasts between species / method classes / sexes



#reading data
lh<-read_xlsx('life_history_data_EN.xlsx',sheet = 2)


# ======================================================================================================================
#  1b. FIRST LOOK AT THE SPREADSHEET (structure, types, missing values, unique values)
# ======================================================================================================================
# Goal: confirm that every column was read with the expected type and see how complete each column is.
# Note: read_xlsx() guesses a column type from its first 1000 rows (guess_max = 1000). If a numeric column is empty in
# those rows, it can be read as logical/text. If the check in 1b.1 warns, re-read with guess_max = Inf (see below).
#   lh <- read_xlsx('life_history_data_EN.xlsx', sheet = 2, guess_max = Inf)

# 1b.1 Dimensions and a check of the types of the numeric columns -------------------------------------------------------
dim(lh)                                   # number of rows (estimates) and columns
numeric_cols <- c("k", "linf_cl", "linf_tl", "l50_cl", "l50_tl", "lat_dec", "start_year", "end_year", "n")
not_numeric  <- numeric_cols[!sapply(lh[numeric_cols], is.numeric)]    # columns that should be numbers but are not
if (length(not_numeric) > 0) warning("Not read as numeric: ", paste(not_numeric, collapse = ", "))

# 1b.2 One line per column: type, number of unique values, number and % of missing values -------------------------------
overview <- data.frame(
  column      = names(lh),
  class       = sapply(lh, function(x) class(x)[1]),                   # type assigned by read_xlsx
  n_unique    = sapply(lh, function(x) length(unique(x[!is.na(x)]))),  # distinct values, ignoring NA
  n_missing   = sapply(lh, function(x) sum(is.na(x))),                 # empty cells
  row.names   = NULL
)
overview$pct_missing <- round(100 * overview$n_missing / nrow(lh), 1)
print(overview, row.names = FALSE)

# 1b.3 Unique values and their frequencies for the categorical columns --------------------------------------------------
# table() counts how many rows each value has; useNA = "ifany" also shows empty cells; sort() puts the most frequent first.
categorical_cols <- c("species", "region", "state", "sex", "growth_unit", "growth_method_class", "research_group",
                      "publication_type", "reliability_level", "cl_definition", "analysis_use")
for (v in categorical_cols) {
  cat("\n---", v, "---\n")
  print(sort(table(lh[[v]], useNA = "ifany"), decreasing = TRUE))
}

# 1b.4 How many distinct studies, populations and research groups are there? -------------------------------------------
sapply(lh[c("primary_study_id", "population_id", "research_group", "source")], function(x) length(unique(x)))


# ======================================================================================================================
#  2. DATA PREPARATION
# ======================================================================================================================
# The raw object (lh) is never modified. All changes are made on a copy (d), so the original data stay intact.
d <- as.data.frame(lh)                    # plain data.frame: simpler to use with table(), tapply() and split()

# 2.1 Standardize k to year^-1 -------------------------------------------------------------------------------------------
# growth_unit says whether k was published per year, per month or per day. Multiplying converts it to year^-1.
unit_factor <- c(year = 1, month = 12, day = 365)
d$k_year <- as.numeric(d$k * unit_factor[d$growth_unit])
sum(!is.na(d$k) & is.na(d$k_year))        # should be 0: k values with a missing or unknown unit

# 2.2 Remove values that repeat another publication ---------------------------------------------------------------------
# repeated_parameters lists the parameters of a row that were already reported in another row (for example "k, linf_tl").
# Those values must not be counted twice, so they are set to NA ONLY for that parameter (the row itself is kept).
is_repeated <- function(param) {
  grepl(paste0("\\b", param, "\\b"), d$repeated_parameters, perl = TRUE)   # TRUE if the parameter name is in the text
}
d$k_use       <- ifelse(is_repeated("k"),       NA, d$k_year)
d$linf_cl_use <- ifelse(is_repeated("linf_cl"), NA, d$linf_cl)
d$linf_tl_use <- ifelse(is_repeated("linf_tl"), NA, d$linf_tl)
d$l50_cl_use  <- ifelse(is_repeated("l50_cl"),  NA, d$l50_cl)
d$l50_tl_use  <- ifelse(is_repeated("l50_tl"),  NA, d$l50_tl)
d$m_use       <- ifelse(is_repeated("m"),       NA, d$m)      # natural mortality (year^-1)
d$z_use       <- ifelse(is_repeated("z"),       NA, d$z)      # total mortality (year^-1)

# 2.3 Growth performance index (phi') -------------------------------------------------------------------------------------
# phi' = log10(k) + 2 * log10(Linf). It needs k and Linf from the SAME row. In CL it is computed only where the CL
# definition is the standard one; in TL there is no such restriction.
standard_cl <- d$cl_definition %in% "standard (assumed)"
d$phi_cl <- ifelse(standard_cl, log10(d$k_use) + 2 * log10(d$linf_cl_use), NA)
d$phi_tl <- log10(d$k_use) + 2 * log10(d$linf_tl_use)

# 2.4 Rows used in the main analyses --------------------------------------------------------------------------------------
# analysis_use == "main" keeps the rows of the main analyses; "redundant", "sensitivity" and "exclude" are left out here.
# The single row identified only as "Farfantepenaeus spp." is also left out: it is not one of the five species of the
# plan and it has no k, Linf or L50 value, so it only adds an empty line to every table.
m <- d[d$analysis_use == "main" & d$species != "Farfantepenaeus spp.", ]
nrow(d); nrow(m)                          # all rows vs. main rows

# Parameters that will be inspected below (names of the columns created above)
params <- c("k_use", "linf_cl_use", "linf_tl_use", "phi_cl", "phi_tl", "l50_cl_use", "l50_tl_use", "m_use", "z_use")


# ======================================================================================================================
#  3. SAMPLE SIZES: HOW MUCH DATA DOES EACH PARAMETER HAVE?
# ======================================================================================================================
# 3.1 Overall: estimates, studies, populations and research groups per parameter (main rows only) -----------------------
# A value is "available" when the cell is not NA. Studies, populations and groups count the DISTINCT IDs that contribute.
availability <- t(sapply(params, function(p) {
  has <- m[!is.na(m[[p]]), ]
  c(estimates   = nrow(has),
    studies     = length(unique(has$primary_study_id)),
    populations = length(unique(has$population_id)),
    groups      = length(unique(has$research_group)))
}))
availability

# 3.2 Helper functions for cross-tabulation -------------------------------------------------------------------------------
# count_table(): number of estimates of one parameter, cross-tabulated by one or two variables.
# The grouping variables are turned into factors using ALL the levels present in the main data, so that categories with
# zero estimates of the parameter still appear as 0 (otherwise table() would silently drop them).
count_table <- function(data, param, row_var, col_var = NULL) {
  has <- data[!is.na(data[[param]]), ]                       # keep only rows where the parameter exists
  lv  <- function(v) sort(unique(data[[v]][!is.na(data[[v]])]))   # all categories of a variable
  if (is.null(col_var)) {
    table(factor(has[[row_var]], levels = lv(row_var)))
  } else {
    table(factor(has[[row_var]], levels = lv(row_var)), factor(has[[col_var]], levels = lv(col_var)),
          dnn = c(row_var, col_var))
  }
}

# unique_table(): same layout, but counts DISTINCT studies, populations or groups (id_var) instead of estimates.
# This shows the real replication: 20 estimates from one study are not 20 independent pieces of evidence.
unique_table <- function(data, param, id_var, row_var, col_var) {
  has <- data[!is.na(data[[param]]), ]
  lv  <- function(v) sort(unique(data[[v]][!is.na(data[[v]])]))
  out <- tapply(has[[id_var]],
                list(factor(has[[row_var]], levels = lv(row_var)), factor(has[[col_var]], levels = lv(col_var))),
                function(x) length(unique(x)))
  out[is.na(out)] <- 0                                       # empty cells come out as NA; show them as 0
  names(dimnames(out)) <- c(row_var, col_var)
  out
}

# 3.3 Estimates per parameter by species, region and research group --------------------------------------------------------
count_table(m, "k_use", "species")                           # k per species
count_table(m, "k_use", "species", "region")                 # k per species x region
count_table(m, "k_use", "research_group", "species")         # k per research group x species
count_table(m, "k_use", "species", "growth_method_class")    # k per species x estimation method
count_table(m, "linf_cl_use", "species", "region")           # Linf (CL) per species x region
count_table(m, "linf_tl_use", "species", "region")           # Linf (TL) per species x region
count_table(m, "l50_cl_use", "species", "sex")               # L50 (CL) per species x sex
count_table(m, "phi_cl", "species", "region")                # phi' (CL) per species x region

# 3.4 Example: how many Linf are there for Xiphopenaeus kroyeri? ----------------------------------------------------------
kro <- m[m$species == "Xiphopenaeus kroyeri", ]              # subset with only this species
count_table(kro, "linf_cl_use", "region")                    # Linf (CL) by region
count_table(kro, "linf_cl_use", "research_group")            # Linf (CL) by research group
count_table(kro, "linf_tl_use", "region")                    # Linf (TL) by region
count_table(kro, "linf_tl_use", "research_group")            # Linf (TL) by research group

# 3.5 Distinct studies, populations and groups (real replication) ---------------------------------------------------------
unique_table(m, "k_use", "primary_study_id", "species", "region")   # studies with k per species x region
unique_table(m, "k_use", "population_id",    "species", "region")   # populations with k per species x region
unique_table(m, "k_use", "research_group",   "species", "region")   # groups with k per species x region


# ======================================================================================================================
#  4. DESCRIPTIVE STATISTICS
# ======================================================================================================================
# desc_stats(): n, mean, sd, coefficient of variation (CV = sd / mean), median, min and max of a parameter by one or two
# grouping variables. Rows with NA in the parameter are ignored; groups with no data are returned with n = 0.
desc_stats <- function(data, param, by) {
  data   <- data[complete.cases(data[by]), ]                     # drop rows with a missing grouping value (e.g. no region)
  label  <- do.call(paste, c(data[by], sep = " | "))             # one label per row (e.g. "species | region")
  groups <- split(data[[param]], factor(label, levels = sort(unique(label))))
  out <- t(sapply(groups, function(x) {
    x <- x[!is.na(x)]
    if (length(x) == 0) return(c(n = 0, mean = NA, sd = NA, cv = NA, median = NA, min = NA, max = NA))
    c(n = length(x), mean = mean(x), sd = sd(x), cv = sd(x) / mean(x),   # sd() is NA when n = 1
      median = median(x), min = min(x), max = max(x))
  }))
  round(out, 3)
}

# 4.1 Growth coefficient k (year^-1) ---------------------------------------------------------------------------------------
desc_stats(m, "k_use", "species")
desc_stats(m, "k_use", "region")
desc_stats(m, "k_use", "growth_method_class")
desc_stats(m, "k_use", "research_group")
desc_stats(m, "k_use", c("species", "growth_method_class"))
desc_stats(m, "k_use", c("species", "sex"))

# 4.2 Asymptotic length (Linf), CL and TL are never mixed ---------------------------------------------------------------
desc_stats(m, "linf_cl_use", "species")
desc_stats(m, "linf_tl_use", "species")

# 4.3 Growth performance index (phi') ------------------------------------------------------------------------------------
desc_stats(m, "phi_cl", "species")
desc_stats(m, "phi_tl", "species")

# 4.4 Length at 50% maturity (L50) ---------------------------------------------------------------------------------------
desc_stats(m, "l50_cl_use", c("species", "sex"))
desc_stats(m, "l50_tl_use", c("species", "sex"))

# 4.5 Latitude covered by each species (only rows with k) ----------------------------------------------------------------
desc_stats(m[!is.na(m$k_use), ], "lat_dec", "species")       # decimal latitude (negative = south)


# ======================================================================================================================
#  5. IMBALANCE AND GAPS IN THE DATA
# ======================================================================================================================
# 5.1 Empty cells: species x region combinations with no k ------------------------------------------------------------
# empty_cells() lists the cells of a two-way table that are 0, i.e. combinations with no data at all.
empty_cells <- function(tab) {
  idx <- which(tab == 0, arr.ind = TRUE)                       # positions (row, column) of the zero cells
  data.frame(row = rownames(tab)[idx[, 1]], col = colnames(tab)[idx[, 2]])
}
empty_cells(count_table(m, "k_use", "species", "region"))            # species without k in a region
empty_cells(count_table(m, "k_use", "species", "growth_method_class"))   # species without k by a given method
empty_cells(count_table(m, "linf_cl_use", "species", "region"))      # species without Linf (CL) in a region

# 5.2 Research groups ------------------------------------------------------------------------------------------------------
# Share of the k estimates contributed by each group (a very large share means the group can dominate the results)
round(100 * prop.table(sort(table(m$research_group[!is.na(m$k_use)]), decreasing = TRUE)), 1)

# Number of distinct studies per research group (groups with a single study cannot be separated from that study)
sort(tapply(m$primary_study_id[!is.na(m$k_use)], m$research_group[!is.na(m$k_use)],
            function(x) length(unique(x))), decreasing = TRUE)

# Groups with no k at all in the main data
tab_group_k <- table(m$research_group, !is.na(m$k_use))
names(which(tab_group_k[, "TRUE"] == 0))

# 5.3 Confounding between group, region and method ------------------------------------------------------------------------
# In how many regions does each research group have k data? (1 = the group is confounded with that region)
# Reading the result: the names are the number of regions (0 = group with no k data) and the values are numbers of groups.
tab_gr <- count_table(m, "k_use", "research_group", "region")
table(rowSums(tab_gr > 0))                                       # number of groups by number of regions
# In how many estimation methods does each research group have k data? (1 = group and method are confounded)
tab_gm <- count_table(m, "k_use", "research_group", "growth_method_class")
table(rowSums(tab_gm > 0))                                       # number of groups by number of methods (0 = no k data)

# 5.4 Studies with more than one method (needed to separate the method effect from the study effect) -------------------
has_k <- m[!is.na(m$k_use), ]
methods_per_study <- tapply(has_k$growth_method_class, has_k$primary_study_id, function(x) length(unique(x)))
table(methods_per_study)                                         # studies with 1, 2, ... methods

# 5.5 Estimates per population (how many rows share the same population) ----------------------------------------------------
table(table(has_k$population_id))                                # number of populations with 1, 2, 3, ... k estimates


# 5.6 Latitude coverage and overlap among species (does the data allow a common latitude effect?) -----------------------
# The latitude effect of the models is a SINGLE slope shared by all species, estimated from the species that span several
# latitudes. Species that never overlap in latitude can only be compared through that shared slope (an extrapolation).
# Rows used: main rows with k (the primary response).
k_rows   <- m[!is.na(m$k_use) & !is.na(m$lat_dec), ]
lat_rng  <- do.call(rbind, lapply(split(k_rows, k_rows$species), function(x) {
  data.frame(estimates = nrow(x), populations = length(unique(x$population_id)),
             regions = length(unique(x$region[!is.na(x$region)])),
             lat_min = min(x$lat_dec), lat_max = max(x$lat_dec))     # latitude is negative in the southern hemisphere
}))
lat_rng$lat_span <- lat_rng$lat_max - lat_rng$lat_min                # degrees of latitude covered by the species
lat_rng
# Overlap (in degrees of latitude) between the ranges of each pair of species; 0 = the ranges do not overlap
sp <- rownames(lat_rng)
overlap <- sapply(sp, function(b) sapply(sp, function(a)
  max(0, min(lat_rng[a, "lat_max"], lat_rng[b, "lat_max"]) - max(lat_rng[a, "lat_min"], lat_rng[b, "lat_min"]))))
round(overlap, 1)


# ======================================================================================================================
#  6. SUMMARY TABLES (written to results/Table_descriptive_summary.xlsx)
# ======================================================================================================================
# Extra packages used by the tables and figures. Missing ones are installed once (requires internet).
for (p in c("ggplot2", "openxlsx", "maps", "patchwork")) if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
library(ggplot2)      # figures
library(openxlsx)     # formatted Excel tables
library(patchwork)    # combines several ggplots into one figure
dir.create("results", showWarnings = FALSE)    # all tables and figures are saved here (created inside the working directory)

# Fixed species order (most to least estimates of k) and short names used in tables and figures
species_order <- c("Xiphopenaeus kroyeri", "Litopenaeus schmitti", "Farfantepenaeus subtilis",
                   "Farfantepenaeus paulensis", "Farfantepenaeus brasiliensis")
short_name <- function(x) sub("^(\\w)\\w+ ", "\\1. ", x)         # "Xiphopenaeus kroyeri" -> "X. kroyeri"

# Special characters are written as unicode escapes so that this script stays plain ASCII on any system
LINF <- "L\u221e"; PHI <- "\u03c6\u2032"; YR1 <- "year\u207b\u00b9"; ENDASH <- "\u2013"   # L-inf, phi-prime, year^-1, en dash

# 6.1 Table 1: median (Q1-Q3) and mean (CV %) of each parameter, by species ----------------------------------------------
# Median (Q1-Q3) is the main summary because the parameters are right-skewed; mean and CV (CV = sd / mean x 100) are shown
# next to it. Only main rows are used, repeated values were already removed (Section 2.2).
table_vars <- data.frame(
  var    = c("k_use", "linf_cl_use", "linf_tl_use", "phi_cl", "phi_tl", "l50_cl_use", "l50_tl_use", "m_use", "z_use"),
  label  = c(paste0("k (", YR1, ")"), paste0(LINF, " carapace length (mm)"), paste0(LINF, " total length (mm)"),
             paste0(PHI, " carapace length"), paste0(PHI, " total length"), "L50 carapace length (mm)",
             "L50 total length (mm)", paste0("M natural mortality (", YR1, ")"), paste0("Z total mortality (", YR1, ")")),
  digits = c(2, 1, 1, 2, 2, 1, 1, 2, 2),
  stringsAsFactors = FALSE)

# summarize_one(): statistics of one parameter for one subset of rows
summarize_one <- function(data, var) {
  x <- data[!is.na(data[[var]]), ]
  v <- x[[var]]
  if (length(v) == 0) return(NULL)
  q <- quantile(v, c(0.25, 0.75), names = FALSE)                  # first and third quartiles
  data.frame(n = length(v), studies = length(unique(x$primary_study_id)),
             median = median(v), q1 = q[1], q3 = q[2], mean = mean(v),
             sd = ifelse(length(v) > 1, sd(v), NA), min = min(v), max = max(v))
}

# Build the numeric table: one row per parameter x species (plus a "All species" row)
num_rows <- list()
for (i in seq_len(nrow(table_vars))) {
  v <- table_vars$var[i]
  for (sp_name in c(species_order, "All species")) {
    sub <- if (sp_name == "All species") m else m[m$species == sp_name, ]
    s <- summarize_one(sub, v)
    if (is.null(s)) next                                           # skip species with no value of this parameter
    num_rows[[length(num_rows) + 1]] <- cbind(parameter = table_vars$label[i], species = sp_name,
                                              digits = table_vars$digits[i], s)
  }
}
table1_num <- do.call(rbind, num_rows)
table1_num$cv_pct <- 100 * table1_num$sd / table1_num$mean         # coefficient of variation in %

# Formatted version for the paper: "median (Q1-Q3)" and "mean (CV %)"
fmt <- function(x, dg) unname(mapply(function(v, g) formatC(v, format = "f", digits = g), x, dg))   # dg can be a vector
table1 <- data.frame(
  Parameter = table1_num$parameter,
  Species   = table1_num$species,
  Estimates = table1_num$n,
  Studies   = table1_num$studies,
  `Median (Q1-Q3)` = paste0(fmt(table1_num$median, table1_num$digits), " (", fmt(table1_num$q1, table1_num$digits), ENDASH,
                            fmt(table1_num$q3, table1_num$digits), ")"),
  `Mean (CV, %)`   = ifelse(is.na(table1_num$cv_pct), paste0(fmt(table1_num$mean, table1_num$digits), " (NA)"),
                            paste0(fmt(table1_num$mean, table1_num$digits), " (", fmt(table1_num$cv_pct, 0), ")")),
  check.names = FALSE, stringsAsFactors = FALSE)
table1$Parameter[duplicated(table1$Parameter)] <- ""                # print each parameter name only once (paper style)
table1$Species <- ifelse(table1$Species %in% species_order, table1$Species, "All species")
table1                                                             # look at it in the console

# 6.2 Table 2: estimation methods most used (growth, natural mortality, total mortality) ------------------------------------
# method_table(): for the rows that have a value of the parameter, counts estimates, studies and research groups per method.
method_table <- function(data, has_var, method_var) {
  x <- data[!is.na(data[[has_var]]), ]
  x$method <- ifelse(is.na(x[[method_var]]), "Not reported", x[[method_var]])
  out <- do.call(rbind, lapply(split(x, x$method), function(g)
    data.frame(Method = g$method[1], Estimates = nrow(g), Studies = length(unique(g$primary_study_id)),
               Groups = length(unique(g$research_group)))))
  out$`% of estimates` <- round(100 * out$Estimates / sum(out$Estimates), 1)
  out <- out[order(-out$Estimates), c("Method", "Estimates", "% of estimates", "Studies", "Groups")]
  rownames(out) <- NULL
  out
}
m$growth_value <- ifelse(!is.na(m$k_use) | !is.na(m$linf_cl_use) | !is.na(m$linf_tl_use), 1, NA)   # rows with a growth value
table2_growth_class <- method_table(m, "growth_value", "growth_method_class")   # grouped classes (used in the models)
table2_growth       <- method_table(m, "growth_value", "growth_method")         # specific labels as published
table2_m            <- method_table(m, "m_use", "m_method")                     # natural mortality (M)
table2_z            <- method_table(m, "z_use", "z_method")                     # total mortality (Z)
table2_growth_class; table2_m

# 6.3 Table 3: research groups ------------------------------------------------------------------------------------------------
# Which groups contribute most (estimates, studies, populations), where they work and which parameters they provide.
group_tab <- do.call(rbind, lapply(split(m, m$research_group), function(g)
  data.frame(`Research group` = g$research_group[1],
             Estimates   = nrow(g),
             Studies     = length(unique(g$primary_study_id)),
             Populations = length(unique(g$population_id)),
             Regions     = paste(sort(unique(g$region[!is.na(g$region)])), collapse = ", "),
             Species     = paste(short_name(sort(unique(g$species))), collapse = "; "),
             `with k`    = sum(!is.na(g$k_use)),
             `with Linf` = sum(!is.na(g$linf_cl_use) | !is.na(g$linf_tl_use)),
             `with L50`  = sum(!is.na(g$l50_cl_use) | !is.na(g$l50_tl_use)),
             `with M`    = sum(!is.na(g$m_use)),
             check.names = FALSE, stringsAsFactors = FALSE)))
group_tab$`% of estimates` <- round(100 * group_tab$Estimates / sum(group_tab$Estimates), 1)
group_tab <- group_tab[order(-group_tab$Estimates), ]
rownames(group_tab) <- NULL
group_tab

# 6.4 Table 4: sample sizes and species x region counts (same numbers as Sections 3.1 and 3.3) -------------------------------
avail_df <- data.frame(Parameter = table_vars$label[match(rownames(availability), table_vars$var)], availability,
                       row.names = NULL, check.names = FALSE)
avail_df <- avail_df[!is.na(avail_df$Parameter), ]
k_by_region <- as.data.frame.matrix(count_table(m, "k_use", "species", "region"))
k_by_region <- cbind(Species = rownames(k_by_region), k_by_region)
k_by_region <- k_by_region[k_by_region$Species %in% species_order, ]
k_by_region <- k_by_region[match(species_order, k_by_region$Species), ]

# 6.5 Write everything to one formatted Excel workbook ---------------------------------------------------------------------
head_style <- createStyle(textDecoration = "bold", fgFill = "#E7E6E6", halign = "center", valign = "center",
                          border = "TopBottom", borderStyle = "thin", wrapText = TRUE)
note_style <- createStyle(textDecoration = "italic", fontColour = "#595959")

# add_table(): one formatted sheet = header style, auto column width, frozen header, optional note under the table
add_table <- function(wb, sheet, df, note = NULL) {
  addWorksheet(wb, sheet)
  writeData(wb, sheet, df, headerStyle = head_style, borders = "rows", borderStyle = "thin", borderColour = "#BFBFBF")
  setColWidths(wb, sheet, cols = seq_len(ncol(df)), widths = "auto")
  freezePane(wb, sheet, firstRow = TRUE)
  if (!is.null(note)) { writeData(wb, sheet, note, startRow = nrow(df) + 3); addStyle(wb, sheet, note_style, nrow(df) + 3, 1) }
}
wb <- createWorkbook()
add_table(wb, "Table1_parameters", table1,
          paste0("Main rows only. Median (Q1", ENDASH, "Q3) and mean (coefficient of variation, %). M is mostly obtained from ",
                 "empirical equations of k, ", LINF, " and temperature, and Z from catch curves: both are described only."))
add_table(wb, "Table1_numeric", table1_num[, c("parameter", "species", "n", "studies", "median", "q1", "q3", "mean", "sd",
                                               "cv_pct", "min", "max")])
add_table(wb, "Table2_growth_class", table2_growth_class)
add_table(wb, "Table2_growth_method", table2_growth)
add_table(wb, "Table2_M_method", table2_m)
add_table(wb, "Table2_Z_method", table2_z)
add_table(wb, "Table3_research_groups", group_tab)
add_table(wb, "Table4_sample_sizes", avail_df)
add_table(wb, "Table4_k_species_region", k_by_region)
saveWorkbook(wb, "results/Table_descriptive_summary.xlsx", overwrite = TRUE)


# ======================================================================================================================
#  7. FIGURES (saved to results/ as PNG, 300 dpi, and PDF)
# ======================================================================================================================
# 7.0 Look and colors ------------------------------------------------------------------------------------------------------
# Region is an ORDERED variable (north to south), so it gets ONE blue ramp (light = north, dark = south) and a different
# point shape per region, so regions are still distinguishable without color. Parameter types use fixed categorical colors.
region_levels <- c("N", "NE", "SE", "SE-S", "S")
region_cols   <- c(N = "#86b6ef", NE = "#5598e7", SE = "#2a78d6", `SE-S` = "#1c5cab", S = "#104281")
region_shapes <- c(N = 21, NE = 24, SE = 22, `SE-S` = 23, S = 25)                  # filled shapes (fill = region color)
topic_cols    <- c(Growth = "#2a78d6", Mortality = "#eb6834", Maturity = "#1baf7a",
                   `Reproduction and recruitment` = "#eda100", `Length-weight` = "#e87ba4")
ink <- "#262626"; ink2 <- "#52514e"; grid_col <- "#E5E5E2"

theme_shrimp <- function() {
  theme_minimal(base_size = 11) +
    theme(text = element_text(colour = ink), axis.text = element_text(colour = ink2),
          panel.grid.minor = element_blank(), panel.grid.major.x = element_blank(),
          panel.grid.major.y = element_line(colour = grid_col, linewidth = 0.3),
          strip.text = element_text(face = "bold", hjust = 0.5, colour = ink), legend.position = "bottom",
          legend.title = element_text(face = "bold", size = 9), plot.title = element_text(face = "bold", size = 11),
          plot.margin = margin(8, 10, 8, 8))
}
save_fig <- function(plot, name, width, height) {                  # saves PNG (300 dpi) and PDF with the same name
  ggsave(file.path("results", paste0(name, ".png")), plot, width = width, height = height, dpi = 300, bg = "white")
  try(ggsave(file.path("results", paste0(name, ".pdf")), plot, width = width, height = height, device = cairo_pdf), silent = TRUE)
}

# 7.1 Figure 1: parameters by species (boxplot + points colored by region) -----------------------------------------------
# Each point is one estimate. Boxes are drawn only for species with 5 or more estimates (a box of 2-4 points is misleading).
# Species stay on the x axis and region is shown by color and shape: this shows the species differences AND the region
# imbalance in the same panel, without cells that are almost empty.
# Figure labels use plain ASCII ("Linf", "1/year") so they print correctly with any font; the Excel tables keep the symbols.
fig_vars <- data.frame(
  var   = c("k_use", "linf_cl_use", "linf_tl_use", "l50_cl_use", "l50_tl_use", "m_use"),
  label = c("k (1/year)", "Linf (CL, mm)", "Linf (TL, mm)", "L50 (CL, mm)", "L50 (TL, mm)", "M (1/year)"),
  stringsAsFactors = FALSE)      # CL = carapace length, TL = total length (explained in the figure caption)
fig1 <- do.call(rbind, lapply(seq_len(nrow(fig_vars)), function(i) {
  x <- m[!is.na(m[[fig_vars$var[i]]]), ]
  data.frame(parameter = fig_vars$label[i], species = x$species, region = x$region, value = x[[fig_vars$var[i]]])
}))
fig1$parameter <- factor(fig1$parameter, levels = fig_vars$label)                    # keeps the panel order
fig1$species   <- factor(fig1$species, levels = species_order, labels = short_name(species_order))
fig1$region    <- factor(fig1$region, levels = region_levels)
n_df   <- as.data.frame(table(fig1$parameter, fig1$species), responseName = "n")     # estimates per panel x species
names(n_df)[1:2] <- c("parameter", "species")
n_df   <- n_df[n_df$n > 0, ]
box_df <- merge(fig1, n_df[n_df$n >= 5, c("parameter", "species")])                  # rows of species with >= 5 estimates

p1 <- ggplot(fig1, aes(x = species, y = value)) +
  geom_boxplot(data = box_df, width = 0.55, outlier.shape = NA, fill = "#D9D8D3", colour = "#6e6d69", linewidth = 0.45,
               alpha = 0.55) +                                          # semi-transparent box (alpha) so grid lines show through
  geom_point(aes(fill = region, shape = region), position = position_jitter(width = 0.16, height = 0, seed = 1),
             size = 2.4, colour = "white", stroke = 0.3, alpha = 0.7) +   # alpha < 1: the box stays visible behind the points
  geom_text(data = n_df, aes(x = species, y = -Inf, label = paste0("n = ", n)), inherit.aes = FALSE,
            vjust = -0.6, size = 2.7, colour = ink2) +
  facet_wrap(~ parameter, scales = "free_y", ncol = 3) +
  scale_fill_manual(values = region_cols, name = "Region", na.value = "#B0B0AC", drop = TRUE) +
  scale_shape_manual(values = region_shapes, name = "Region", drop = TRUE, na.value = 21) +
  scale_y_continuous(expand = expansion(mult = c(0.14, 0.06))) +
  labs(x = NULL, y = NULL, caption = "CL = carapace length; TL = total length. Each point is one estimate; boxes (median and quartiles) are shown for n >= 5.") +
  theme_shrimp() + theme(axis.text.x = element_text(face = "italic", angle = 30, hjust = 1),
                         plot.caption = element_text(colour = ink2, size = 8, hjust = 0))
p1
save_fig(p1, "Fig1_parameters_by_species", 12, 7)

# 7.2 Figure 2: estimates over time, by species (small multiples) ---------------------------------------------------------
# Year = middle of the period of data collection (start_year and end_year; the one available if only one exists).
# Rows are parameters and columns are species, so each panel has a single species and no color is needed for species.
# Change time_vars to add other parameters (for example "k_use", "linf_tl_use", "l50_tl_use").
add_trend <- TRUE          # TRUE draws a dashed linear trend per panel (descriptive only; estimates are not independent)
min_n_trend <- 2         # a trend line is drawn only in panels with at least this many estimates
time_vars <- fig_vars[fig_vars$var %in% c("k_use", "linf_cl_use", "l50_cl_use"), ]
m$year_mid <- rowMeans(m[c("start_year", "end_year")], na.rm = TRUE)                  # NaN when both years are missing
m$year_mid[is.nan(m$year_mid)] <- NA
fig2 <- do.call(rbind, lapply(seq_len(nrow(time_vars)), function(i) {
  x <- m[!is.na(m[[time_vars$var[i]]]) & !is.na(m$year_mid), ]
  data.frame(parameter = time_vars$label[i], species = x$species, region = x$region, year = x$year_mid,
             value = x[[time_vars$var[i]]])
}))
fig2$parameter <- factor(fig2$parameter, levels = time_vars$label)
fig2$species   <- factor(fig2$species, levels = species_order, labels = short_name(species_order))
fig2$region    <- factor(fig2$region, levels = region_levels)
n_panel  <- table(fig2$parameter, fig2$species)                                      # estimates per panel
trend_df <- fig2[n_panel[cbind(as.character(fig2$parameter), as.character(fig2$species))] >= min_n_trend, ]   # big panels only
p2 <- ggplot(fig2, aes(x = year, y = value)) +
  geom_point(aes(fill = region, shape = region), size = 3.2, colour = "white", stroke = 0.3, alpha = 0.85) +
  { if (add_trend) geom_smooth(data = trend_df, method = "lm", formula = y ~ x, se = FALSE, colour = "#262626",
                               linewidth = 0.6, linetype = "dashed") } +
  facet_grid(parameter ~ species, scales = "free_y", switch = "y") +
  scale_fill_manual(values = region_cols, name = "Region", na.value = "#B0B0AC", drop = TRUE) +
  scale_shape_manual(values = region_shapes, name = "Region", drop = TRUE, na.value = 21) +
  scale_x_continuous(breaks = seq(1950, 2020, 20), limits = c(1945, 2027)) +      # data span 1947-2025
  labs(x = "Year of data collection (midpoint)", y = NULL) +
  theme_shrimp() + theme(strip.placement = "outside", strip.text.x = element_text(face = "bold.italic"),
                         strip.text.y.left = element_text(angle = 90), panel.grid.major.x = element_line(colour = grid_col, linewidth = 0.3))
p2
save_fig(p2, "Fig2_estimates_over_time", 13, 7.5)

# 7.3 Figure 3: how the research evolved (studies per decade of first publication) ---------------------------------------
# Unit = primary study (the same study published as thesis and article counts once, in its earliest year).
# The year comes from the source label (e.g. "Rodrigues et al., 1992"), so every study has one. A study can address several
# parameter types, so the bars of panel A add up to more than the number of studies. Rows flagged "exclude" (geographic
# distribution only) are left out; redundant rows are kept because they still belong to a study.
lit <- d[d$analysis_use != "exclude" & d$species != "Farfantepenaeus spp.", ]
pos <- regexpr("(19|20)[0-9]{2}", lit$source)                                       # first 4-digit year in the source label
lit$src_year <- as.integer(ifelse(pos > 0, substr(lit$source, pos, pos + 3), NA))
study_year   <- tapply(lit$src_year, lit$primary_study_id, min, na.rm = TRUE)      # earliest year of each study
all_decades  <- seq(floor(min(study_year) / 10) * 10, floor(max(study_year) / 10) * 10, by = 10)
decade_lab   <- paste0(all_decades, "s")

# Parameter type of each row (what the row reports)
lit$Growth     <- !is.na(lit$k) | !is.na(lit$linf_cl) | !is.na(lit$linf_tl)
lit$Mortality  <- !is.na(lit$m) | !is.na(lit$z) | !is.na(lit$f)
lit$Maturity   <- !is.na(lit$l50_cl) | !is.na(lit$l50_tl)
lit$`Reproduction and recruitment` <- !is.na(lit$reproductive_period_start) | !is.na(lit$recruitment_start)
lit$`Length-weight` <- !is.na(lit$wl_a) | !is.na(lit$wl_b)

# long_pairs(): unique (study, category) pairs -> decade -> number of studies per decade and category
long_pairs <- function(df, cat_col, levels_cat) {
  pr <- unique(data.frame(study = df$primary_study_id, cat = df[[cat_col]], stringsAsFactors = FALSE))
  pr <- pr[!is.na(pr$cat) & pr$cat %in% levels_cat, ]
  pr$decade <- factor(paste0(floor(study_year[pr$study] / 10) * 10, "s"), levels = decade_lab)
  out <- as.data.frame(table(decade = pr$decade, cat = factor(pr$cat, levels = levels_cat)), responseName = "n")
  out
}
topic_names <- names(topic_cols)
topic_long  <- do.call(rbind, lapply(topic_names, function(tp)
  data.frame(primary_study_id = unique(lit$primary_study_id[lit[[tp]]]), topic = tp, stringsAsFactors = FALSE)))
names(topic_long)[1] <- "primary_study_id"
cnt_topic  <- long_pairs(transform(topic_long, cat = topic), "cat", topic_names)
cnt_method <- long_pairs(lit[lit$Growth, ], "growth_method_class",
                         names(sort(table(lit$growth_method_class[lit$Growth]), decreasing = TRUE)))
cnt_region <- long_pairs(lit, "region", region_levels)
lit$species_short <- short_name(lit$species)
cnt_species <- long_pairs(lit, "species_short", short_name(species_order))

# stack_plot(): stacked bars of studies per decade; thin white gaps between segments (as in the figure style guide)
stack_plot <- function(cnt, cols, title, legend_title) {
  ggplot(cnt, aes(x = decade, y = n, fill = cat)) +
    geom_col(width = 0.78, colour = "white", linewidth = 0.5) +
    scale_fill_manual(values = cols, name = legend_title, drop = FALSE) +
    scale_x_discrete(drop = FALSE) +
    scale_y_continuous(expand = expansion(mult = c(0, 0.12))) +          # headroom so the tallest bar is never cut
    labs(title = title, x = NULL, y = "Studies (n)") +
    theme_shrimp() + theme(legend.position = "right", plot.title = element_text(face = "bold", size = 11, hjust = 0.5), axis.text.x = element_text(angle = 45, hjust = 1))
}
method_levels <- levels(cnt_method$cat)
method_cols   <- c("#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4")[seq_along(method_levels)]
names(method_cols) <- method_levels
species_cols  <- setNames(c("#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"), short_name(species_order))
pA <- stack_plot(cnt_topic,   topic_cols,   "Parameter type", "Parameter")
pB <- stack_plot(cnt_method,  method_cols,  "Growth method", "Method")
pC <- stack_plot(cnt_region,  region_cols,  "Region", "Region")
pD <- stack_plot(cnt_species, species_cols, "Species", "Species") +
  theme(legend.text = element_text(face = "italic"))
p3 <- (pA | pB) / (pC | pD) + plot_annotation(tag_levels = "A")
p3
save_fig(p3, "Fig3_research_evolution", 14, 8.5)

# Counts behind Figure 3, saved for the paper (one sheet per panel)
wb3 <- createWorkbook()
for (nm in c("topic", "method", "region", "species")) {
  cn <- get(paste0("cnt_", nm))
  wide <- as.data.frame.matrix(xtabs(n ~ decade + cat, data = cn)); wide <- cbind(Decade = rownames(wide), wide)
  add_table(wb3, paste0("Studies_by_", nm), wide)
}
saveWorkbook(wb3, "results/Table_studies_by_decade.xlsx", overwrite = TRUE)

# 7.4 Figure 4: map of the study sites by parameter type ---------------------------------------------------------------------
# One point per POPULATION (population_id = one species at one site, possibly studied by several studies) and parameter type, so
# a population can appear in several panels. Point area = number of estimates of that population (all studies together) of
# that parameter type. Redundant rows (the same estimate repeated in a thesis and in its article) are NOT counted.
# Coordinates come from lat_dec / lon_dec of the master spreadsheet (each population has a single coordinate; rows without
# coordinates are the geographic distribution records, which are excluded). Because populations of different species can
# share the same site, a small fixed jitter (0.15 degree, about 15 km) is added ONCE per population, so that a population
# stays at the same place in every panel.
pop_xy <- unique(lit[!is.na(lit$lat_dec) & !is.na(lit$lon_dec), c("population_id", "lat_dec", "lon_dec")])
pop_xy <- pop_xy[!duplicated(pop_xy$population_id), ]                              # one coordinate per population
set.seed(1)                                                                        # same jitter every time the script runs
pop_xy$lon_plot <- pop_xy$lon_dec + runif(nrow(pop_xy), -0.15, 0.15)
pop_xy$lat_plot <- pop_xy$lat_dec + runif(nrow(pop_xy), -0.15, 0.15)
lit_map <- lit[lit$analysis_use != "redundant", ]                                  # rows that are counted as estimates
map_long <- do.call(rbind, lapply(topic_names, function(tp) {
  rows <- lit_map[lit_map[[tp]] & lit_map$population_id %in% pop_xy$population_id, ]
  if (nrow(rows) == 0) return(NULL)
  cnt <- as.data.frame(table(population_id = rows$population_id), responseName = "n")   # estimates per population_id
  cbind(cnt, topic = tp)
}))
map_long <- merge(map_long, pop_xy, by = "population_id")
map_long$topic <- factor(map_long$topic, levels = topic_names)
map_long <- map_long[order(-map_long$n), ]                                         # big points first, small ones stay visible on top
world <- ggplot2::map_data("world")                                                # country outlines from the 'maps' package
lat_lab <- function(x) paste0(abs(x), "\u00b0", ifelse(x < 0, "S", "N"))          # -20 -> "20 deg S"
lon_lab <- function(x) paste0(abs(x), "\u00b0W")
# One row of five narrow panels (the coast is much taller than wide); strip titles are wrapped so none is truncated.
p4 <- ggplot() +
  geom_polygon(data = world, aes(x = long, y = lat, group = group), fill = "#E4E3DD", colour = "white", linewidth = 0.2) +
  geom_point(data = map_long, aes(x = lon_plot, y = lat_plot, size = n, fill = topic), shape = 21, colour = "white",
             stroke = 0.4, alpha = 0.85) +
  facet_wrap(~ topic, nrow = 1, labeller = label_wrap_gen(width = 18)) +
  scale_fill_manual(values = topic_cols, guide = "none") +
  scale_size_area(max_size = 9, breaks = c(1, 5, 20, 60, 130), name = "Estimates per population (n)",
                  guide = guide_legend(override.aes = list(fill = "#8a8985"))) +   # grey fill so legend points are visible
  scale_x_continuous(breaks = c(-50, -40), labels = lon_lab) +
  scale_y_continuous(breaks = seq(-30, 0, 10), labels = lat_lab) +
  coord_quickmap(xlim = c(-53, -33), ylim = c(-34, 6)) +
  labs(x = NULL, y = NULL) +
  theme_shrimp() + theme(panel.background = element_rect(fill = "#F2F6FA", colour = NA),
                         panel.grid.major = element_line(colour = "white", linewidth = 0.35),   # faint graticule (lines every 10 deg)
                         axis.text = element_text(size = 7, colour = "#8a8985"), panel.spacing = unit(0.6, "lines"))
p4
save_fig(p4, "Fig4_map_study_sites", 14, 7.5)


# ======================================================================================================================
#  8. BAYESIAN MODELS (analysis plan, Section 5.1)
# ======================================================================================================================
# What this section does: fits the multilevel models of the plan, one set per response, and saves each fitted model to
# the folder "models/" so that it is NEVER fitted twice (if the file exists, it is read back in seconds). To refit a model,
# delete its .rds file. Fitting takes several minutes per model, so the first run of this section is slow (about 1 to 2
# hours for all responses on a normal computer); later runs only read the saved files.
#
# The five responses (all on a log scale, as in the plan) are:
#   log_k      log of the von Bertalanffy growth coefficient k (year^-1); CL and TL rows together
#   phi_cl     growth performance index phi' = log10(k) + 2 log10(Linf), carapace-length (CL) rows only
#   phi_tl     phi', total-length (TL) rows only
#   log_l50_cl log of the size at 50% maturity (L50), CL
#   log_l50_tl log of L50, TL
#
# For every response the SAME structure is used (plan Section 5.1):
#   response ~ species + sex + method + lat10 + (1 | research_group) + (1 | primary_study_id) + (1 | population_id)
# "method" is omitted for L50. The first three terms after "~" are FIXED effects (average differences between categories or
# a slope); the three terms in parentheses are RANDOM intercepts (each research group, study and population is allowed to
# have its own baseline level; only the spread of those baselines, a standard deviation, is estimated).

library(brms)         # Bayesian multilevel models (Stan back end)
library(posterior)    # summaries of posterior draws: R-hat, effective sample size
library(emmeans)      # contrasts between species / method classes / sexes
# (the Stan toolchain must be installed: on Windows this is Rtools; see the brms installation notes)

# 8.0 Settings -------------------------------------------------------------------------------------------------------------
dir.create("models", showWarnings = FALSE)  # folder where the fitted models are saved (one .rds file per model)
n_chains    <- 4                            # number of independent MCMC chains (plan: 4)
n_iter      <- 4000                         # iterations per chain, the first half is warm-up (plan: 4000)
adapt_main  <- 0.95                         # adapt_delta of the plan; higher = smaller steps = fewer divergences, slower
adapt_retry <- 0.99                         # used automatically if the first fit has divergent transitions (plan Section 5.1)
adapt_last  <- 0.999                        # last resort if divergences remain at 0.99 (smaller steps still; sampler setting only)
max_treedepth_retry <- 12                   # longest trajectory allowed in the refits (Stan default is 10); sampler setting only
loo_reloo   <- TRUE                         # TRUE = estimates with Pareto k > 0.7 are recomputed exactly by refitting the model without
                                            # them (a few extra fits per model, roughly 2-5 min each). FALSE = skip (faster, keeps the warnings)
seed_all    <- 2026                         # random seed, so that the results can be reproduced exactly
n_cores     <- min(n_chains, parallel::detectCores())   # chains are run in parallel on this many cores
only_responses <- NULL                      # NULL = all five responses; e.g. c("log_l50_tl") fits only one (useful for tests)
stan_silent <- 0                            # brms 'silent' level: 0 = show the Stan compiler/sampler messages (safest on Windows),
                                            # 1 = hide the sampler progress, 2 = hide everything. Level 2 captures the console output
                                            # with sink(); if the C++ compilation fails the sink is left broken and R only prints
                                            # 'Error in sink(type = "output") : invalid connection', hiding the real compiler error.

stan_backend <- "rstan"                     # Stan interface used by brms: "rstan" (default) or "cmdstanr". If rstan cannot compile on
                                            # your computer (error 'sink(type = "output") : invalid connection'), install cmdstanr
                                            # (see diagnose_stan.R) and set this to "cmdstanr"; the statistical model is the same.

# Clear any sink left open by an earlier failed run.
while (sink.number() > 0) sink()
while (sink.number(type = "message") > 2) sink(type = "message")
# Windows + rstan fix. rstan (2.32.x) hard-codes the compiler flag '-std=c++1y' (C++14) on Windows, and this flag comes AFTER the
# C++17 flag of R. The current Stan headers need C++17, so the compilation stops with "'void_t' is not a member of 'std'", and
# rstan then hides that message behind 'Error in sink(type = "output") : invalid connection'. The lines below replace the flag by
# '-std=c++17' for this R session only (nothing is changed on disk). They do nothing on other systems or with cmdstanr.
if (stan_backend == "rstan" && .Platform$OS.type == "windows") local({
  flags_orig <- get("PKG_CPPFLAGS_env_fun", envir = asNamespace("rstan"))
  utils::assignInNamespace("PKG_CPPFLAGS_env_fun", function() gsub("-std=c++1y", "-std=c++17", flags_orig(), fixed = TRUE), ns = "rstan")
})

# 8.1 Definition of the responses ------------------------------------------------------------------------------------------
# var   : column of the data created in Section 2 (it already excludes the values flagged in repeated_parameters)
# f     : transformation applied to the column (natural log; phi' is already a log10 quantity)
# method: does the model include the estimation-method class? (growth responses only)
# rope  : half-width of the region of practical equivalence for the latitude slope (hypothesis H2, plan Section 5.3):
#         +/- 10% in k per 10 degrees = log(1.1) = 0.0953 on the natural-log scale and log10(1.1) = 0.0414 for phi'.
#         It is NA for L50 because H2 concerns growth only.
resp_def <- list(
  log_k      = list(label = "log k",        var = "k_use",      f = log,      method = TRUE,  rope = log(1.1),   base = "e"),
  phi_cl     = list(label = "phi' (CL)",    var = "phi_cl",     f = identity, method = TRUE,  rope = log10(1.1), base = "10"),
  phi_tl     = list(label = "phi' (TL)",    var = "phi_tl",     f = identity, method = TRUE,  rope = log10(1.1), base = "10"),
  log_l50_cl = list(label = "log L50 (CL)", var = "l50_cl_use", f = log,      method = FALSE, rope = NA_real_,   base = "e"),
  log_l50_tl = list(label = "log L50 (TL)", var = "l50_tl_use", f = log,      method = FALSE, rope = NA_real_,   base = "e"))
responses <- if (is.null(only_responses)) names(resp_def) else only_responses

# 8.2 Data of each model ---------------------------------------------------------------------------------------------------
# Rules of the plan (Section 6): main rows only (object m); complete cases, no imputation; for the growth responses the
# 10 rows whose method class is "not reported/other" are left out because the model has a method term.
# Coding of the fixed effects:
#   species: reference = Xiphopenaeus kroyeri (the species with most data); the coefficients are differences from it
#   sex    : F (females), M (males), pooled (both sexes together); reference = F. Rows with indeterminate sex (I) are dropped
#   method : reference = modal progression (the most frequent class, as in the plan)
#   lat10  : latitude / 10, with positive values SOUTHWARD and centred on its mean; one unit = 10 degrees of latitude
make_model_data <- function(r) {
  def <- resp_def[[r]]
  x <- m[!is.na(m[[def$var]]), ]                                       # rows that have the response
  if (def$method) x <- x[!is.na(x$growth_method_class) & x$growth_method_class != "not reported/other", ]
  x <- x[x$sex %in% c("F", "M", "A"), ]                                # F, M or pooled (A = both sexes together)
  lat10 <- -x$lat_dec / 10                                             # decimal latitude is negative in the south
  dat <- data.frame(
    y                = def$f(x[[def$var]]),
    species          = relevel(factor(x$species), ref = "Xiphopenaeus kroyeri"),
    sex              = factor(x$sex, levels = c("F", "M", "A"), labels = c("F", "M", "pooled")),
    lat10            = lat10 - mean(lat10),
    research_group   = factor(x$research_group),
    primary_study_id = factor(x$primary_study_id),
    population_id    = factor(x$population_id))
  if (def$method) dat$method <- relevel(factor(x$growth_method_class), ref = "modal progression")
  droplevels(dat)                                                      # drop species / methods that have no row in this model
}
model_data <- setNames(lapply(responses, make_model_data), responses)
sapply(model_data, nrow)                                               # rows (estimates) of each model

# 8.3 The set of models fitted for each response ---------------------------------------------------------------------------
# M0_null      : only the three random intercepts. Gives the proportion of variance at each level (variance partition, H4)
# M1_full      : the full model of the plan (fixed effects + the three random intercepts)
# M1_no_<term> : the full model WITHOUT one fixed term (species, sex, method or lat10). Comparing it with M1_full shows how
#                much that term contributes (change in variance components, in R-squared and in predictive performance).
fixed_terms <- function(r) c("species", "sex", if (resp_def[[r]]$method) "method", "lat10")
random_part <- c("(1 | research_group)", "(1 | primary_study_id)", "(1 | population_id)")
make_formula <- function(r, drop = NULL, null = FALSE) {
  rhs <- if (null) "1" else setdiff(fixed_terms(r), drop)              # fixed part of the formula
  bf(as.formula(paste("y ~", paste(c(rhs, random_part), collapse = " + "))))
}
model_specs <- function(r) {
  c(list(M0_null = list(null = TRUE, drop = NULL), M1_full = list(null = FALSE, drop = NULL)),
    setNames(lapply(fixed_terms(r), function(t) list(null = FALSE, drop = t)), paste0("M1_no_", fixed_terms(r))))
}

# 8.4 Fitting function -----------------------------------------------------------------------------------------------------
# Priors of the plan: normal(0, 1) for the fixed effects (class b); student_t(3, 0, 1) for the standard deviations of the
# random intercepts (class sd) and for the residual standard deviation (sigma). They are weakly informative: they only rule
# out absurd values on the log scale. The null model has no fixed effects, so it has no prior of class b.
# After fitting, the model is checked for divergent transitions. If there are any, it is refitted with adapt_delta = 0.99 and,
# if some remain, with 0.999 (the random-effects structure is never simplified; only the step size and the maximum trajectory
# length of the sampler change), and the value used is stored in the model (fit$adapt_delta_used).
# Finally two measures of predictive performance are added and saved inside the model (explained in Section 9):
# "loo" (leave-one-out cross-validation, with moment matching and, if loo_reloo is TRUE, an exact refit for the few estimates
# where the approximation fails) and "waic".
count_divergent <- function(fit) {                                     # works with both back ends (rstan and cmdstanr)
  np <- brms::nuts_params(fit)
  sum(np$Value[np$Parameter == "divergent__"])
}
fit_one <- function(r, name, spec) {
  file <- file.path("models", paste0(r, "__", name, ".rds"))
  if (file.exists(file)) return(readRDS(file))                         # already fitted: read it back
  message(format(Sys.time(), "%H:%M:%S"), "  fitting ", r, " / ", name, " (n = ", nrow(model_data[[r]]), ")")
  pri <- c(prior(student_t(3, 0, 1), class = sd), prior(student_t(3, 0, 1), class = sigma))
  if (!spec$null) pri <- c(prior(normal(0, 1), class = b), pri)
  adapt <- adapt_main
  fit <- brm(make_formula(r, spec$drop, spec$null), data = model_data[[r]], family = gaussian(), prior = pri,
             chains = n_chains, iter = n_iter, cores = n_cores, seed = seed_all, control = list(adapt_delta = adapt),
             save_pars = save_pars(all = TRUE),                        # needed for the moment matching of loo
             backend = stan_backend, refresh = 0, silent = stan_silent)
  n_div <- count_divergent(fit)                                        # number of divergent transitions
  for (adapt_next in c(adapt_retry, adapt_last)) {                     # refit with smaller steps while divergences remain
    if (n_div == 0) break
    adapt <- adapt_next
    fit <- update(fit, control = list(adapt_delta = adapt, max_treedepth = max_treedepth_retry), recompile = FALSE,
                  cores = n_cores, seed = seed_all, refresh = 0, silent = stan_silent)
    n_div <- count_divergent(fit)
  }
  fit$adapt_delta_used <- adapt
  fit$n_divergent      <- n_div
  fit <- tryCatch(add_criterion(fit, "loo", moment_match = TRUE, reloo = loo_reloo, cores = n_cores),   # exact LOO where PSIS fails
                  error = function(e) add_criterion(fit, "loo"))       # if that fails, use plain PSIS-LOO
  fit <- add_criterion(fit, "waic")
  saveRDS(fit, file)
  fit
}

# HOW TO READ THE WARNINGS PRINTED WHILE THE MODELS ARE FITTED (they are expected; what matters is the final state, which is
# saved in the table of Section 11.1 and in sheet S6 of the Excel file):
#   "divergent transitions" / "largest R-hat" / "ESS too low" : refer to the FIRST attempt (adapt_delta 0.95). If divergences occurred, the
#        model is refitted at 0.99 (then 0.999) and the same warnings are printed again only if the problem persists. Count what remains.
#   "p_waic estimates greater than 0.4" : WAIC is unreliable for those estimates. LOO (ELPD) is the criterion used; WAIC is only reported.
#   "Pareto k > 0.7" : a few influential estimates for which the leave-one-out approximation is rough. The number is reported (Table 7).
#   "maximum treedepth" : the sampler needed very long trajectories; it is an efficiency warning, not a bias warning.

# 8.5 Fit (or read) every model --------------------------------------------------------------------------------------------
# fits[[response]][[model]] is a fitted brms model. On the first run this loop takes a long time. Each saved model is about
# 30-50 MB on disk and roughly 100 MB in memory; if the computer has little memory, fit one response per R session
# (set only_responses above, e.g. "log_k", and run the script again for the next response) and then run with NULL.
fits <- list()
for (r in responses) {
  fits[[r]] <- list()
  for (nm in names(model_specs(r))) { fits[[r]][[nm]] <- fit_one(r, nm, model_specs(r)[[nm]]); invisible(gc()) }
}


# ======================================================================================================================
#  9. VARIANCE PARTITIONING AND CONTRIBUTION OF EACH TERM (analysis plan, Section 5.2; hypothesis H4)
# ======================================================================================================================
# THE IDEA. Every model describes an estimate (y, on the log scale) as
#        y  =  fixed part  +  u_group  +  u_study  +  u_population  +  residual
# where "fixed part" is the sum of the effects of species, sex, method and latitude, and each u is the departure of a research
# group, a study or a population from the overall level. A model estimates how much each piece VARIES. Variance is the
# natural measure because variances of independent pieces add up to the total variance:
#        total variance = var(fixed part) + var(group) + var(study) + var(population) + var(residual)
# so each piece can be expressed as a share (%) of the total. This section computes those shares in four ways:
#   9.2  Variance-partition coefficients (VPC) of the NULL model (M0_null, no fixed effects): how much of the variance lies at
#        the level of research group, study, population and residual before any fixed effect is considered.
#   9.3  The same shares in the FULL model (M1_full), plus the Bayesian R-squared of Nakagawa and Schielzeth (2013):
#           marginal R2    = share explained by the fixed effects alone = var(fixed) / total
#           conditional R2 = share explained by fixed effects + random intercepts = 1 - var(residual) / total
#   9.4  Variance explained by EACH fixed term (species, sex, method, latitude) in the full model, and the part that is SHARED
#        among terms. Terms overlap in this database (for example a method is used mostly by certain groups and regions), so
#        the variances of the single terms do not add up to the variance of the fixed part: the difference is the shared part.
#   9.5  What happens when a fixed term is REMOVED (models M1_no_<term>): (a) the loss of conditional R2 (the "unique"
#        contribution of the term), (b) the change in each variance component, and (c) the change in the expected log
#        predictive density (ELPD, explained in 9.5). All three are measured against the full model.
# All quantities are computed for every posterior draw, so each one comes with a 95% credible interval (HDI, see 9.1).

# HOW TO READ THE TABLES OF THIS SECTION (details in each subsection)
#   vpc_null   : where the variation of the parameter comes from BEFORE any explanatory term: research group, study, population or
#                residual (the share of the total variance at each level, null model).
#   share_full : the same shares once species, sex, method and latitude are in the model; "fixed effects" = marginal R2.
#   term_var   : how much variance each fixed term carries inside the full model, plus the part shared among terms.
#   drop_tab   : what is lost when ONE term is removed: unique R2, where its variance goes, and the loss of predictive ability (ELPD).
#   model_tab  : all models side by side with their prediction scores (ELPD, LOOIC, WAIC).

# 9.1 Helper functions ----------------------------------------------------------------------------------------------------
# nobars_f(): formula without the random-effect terms (reformulas::nobars in recent lme4 versions, lme4::nobars in older ones)
nobars_f <- function(f) if (requireNamespace("reformulas", quietly = TRUE)) reformulas::nobars(f) else lme4::nobars(f)
# hdi(): 95% highest-density interval = the shortest interval that contains 95% of the posterior draws.
hdi <- function(x, prob = 0.95) {
  x <- sort(x[is.finite(x)]); n <- length(x); k <- floor(prob * n)       # k = number of draws inside the interval
  width <- x[(k + 1):n] - x[1:(n - k)]                                   # width of every candidate interval
  i <- which.min(width)                                                  # the narrowest one
  c(lower = x[i], upper = x[i + k])
}
# post_sum(): median and 95% HDI of a vector of draws (used for every posterior quantity below)
post_sum <- function(x) { h <- hdi(x); c(median = median(x), lower = h[["lower"]], upper = h[["upper"]]) }

# var_comp(): variances (sd^2) of the three random intercepts and of the residual, one row per posterior draw
var_comp <- function(fit) {
  dr <- as_draws_df(fit)
  data.frame(group      = dr$sd_research_group__Intercept^2,
             study      = dr$sd_primary_study_id__Intercept^2,
             population = dr$sd_population_id__Intercept^2,
             residual   = dr$sigma^2)
}

# fixed_draws(): variance of the fixed part, in total and term by term, for every posterior draw.
# For each draw the fixed part of every estimate is X %*% b (design matrix times coefficients); its variance across the
# estimates of the data set is the "variance explained by the fixed effects" used in the R-squared of Nakagawa & Schielzeth.
fixed_draws <- function(fit) {
  fx <- nobars_f(fit$formula$formula)                                # formula without the random-effect terms
  lab <- attr(terms(fx), "term.labels")                                  # names of the fixed terms (species, sex, ...)
  n_draw <- ndraws(fit)
  if (length(lab) == 0) return(list(var_term = NULL, var_fixed = rep(0, n_draw)))   # null model: no fixed effects
  X  <- model.matrix(fx, data = fit$data)[, -1, drop = FALSE]            # design matrix without the intercept
  asg <- attr(model.matrix(fx, data = fit$data), "assign")[-1]           # which term each column belongs to
  dr <- as_draws_df(fit)
  bn <- setdiff(grep("^b_", names(dr), value = TRUE), "b_Intercept")     # coefficient columns, in design-matrix order
  B  <- as.matrix(as.data.frame(dr)[, bn]); stopifnot(ncol(B) == ncol(X))
  part <- lapply(seq_along(lab), function(j) { cols <- which(asg == j); X[, cols, drop = FALSE] %*% t(B[, cols, drop = FALSE]) })
  names(part) <- lab                                                     # each: estimates x draws
  list(var_term  = sapply(part, function(z) apply(z, 2, var)),           # draws x terms
       var_fixed = apply(Reduce(`+`, part), 2, var))                     # draws
}

# partition(): everything needed from one model, per draw
partition <- function(fit) {
  vc <- var_comp(fit); fd <- fixed_draws(fit)
  list(vc = vc, vf = fd$var_fixed, vterm = fd$var_term, total = fd$var_fixed + rowSums(vc))
}
parts <- lapply(fits, function(f) lapply(f, partition))                  # parts[[response]][[model]]

# 9.2 Variance-partition coefficients of the null model ------------------------------------------------------------------
# VPC of a level = variance of that level / sum of all variances. Example: a group VPC of 20% means that 20% of the variation
# among estimates is due to differences among research groups (some groups tend to report higher or lower values).
# In short: this answers "how clustered are the estimates?". High study/group shares mean that estimates from the same study or
# group are alike (and different from other studies/groups), so they are not independent replicates of the biology; the
# residual share is the variation left between estimates of the same population (biology + measurement noise).
level_names <- c(group = "Research group", study = "Study", population = "Population", residual = "Residual")
vpc_null <- do.call(rbind, lapply(responses, function(r) {
  p <- parts[[r]]$M0_null
  do.call(rbind, lapply(names(level_names), function(l)
    data.frame(response = r, component = level_names[[l]], t(post_sum(100 * p$vc[[l]] / p$total)), row.names = NULL)))
}))

# 9.3 Full model: shares of the total variance and Bayesian R-squared ----------------------------------------------------
# In short: same idea as 9.2, now with the fixed effects in the model. Marginal R2 = what species, sex, method and latitude explain
# together. Conditional R2 = what they explain together with the group/study/population differences (1 - residual share).
# When a fixed term "absorbs" differences that were attributed to groups or studies in 9.2, those shares shrink here.
share_full <- do.call(rbind, lapply(responses, function(r) {
  p <- parts[[r]]$M1_full
  items <- c(list(`Fixed effects (marginal R2)` = p$vf / p$total),
             lapply(setNames(names(level_names), level_names), function(l) p$vc[[l]] / p$total),
             list(`Fixed + random (conditional R2)` = 1 - p$vc$residual / p$total))
  do.call(rbind, lapply(names(items), function(i) data.frame(response = r, component = i, t(post_sum(100 * items[[i]])), row.names = NULL)))
}))

# 9.4 Variance explained by each fixed term, and the shared part ---------------------------------------------------------
# The variance of a term is the variance of its own contribution to the fixed part (for example, how much the estimates would
# vary if only the species effect were present). "Shared" = var(fixed part) - sum of the single-term variances; it comes from the
# covariance among terms (positive when terms carry overlapping information, negative when they partly cancel).
# In short: each row is the variance of that term's contribution, as % of the total variance, taken from the full model (all
# terms fitted together, so each term is already adjusted for the others). They need not add up to the marginal R2 because terms
# overlap; a large "shared" part means the data cannot tell the terms apart (for instance a method used by one group in one region).
term_labels <- c(species = "Species", sex = "Sex", method = "Method class", lat10 = "Latitude")
term_var <- do.call(rbind, lapply(responses, function(r) {
  p <- parts[[r]]$M1_full
  out <- lapply(colnames(p$vterm), function(t) data.frame(response = r, term = term_labels[[t]], t(post_sum(100 * p$vterm[, t] / p$total)), row.names = NULL))
  shared <- (p$vf - rowSums(p$vterm)) / p$total
  rbind(do.call(rbind, out), data.frame(response = r, term = "Shared among terms", t(post_sum(100 * shared)), row.names = NULL))
}))

# 9.5 Removing one fixed term at a time ---------------------------------------------------------------------------------
# For each model M1_no_<term> the following are compared with the full model (M1_full):
#  (a) unique R2 = conditional R2 of the full model - conditional R2 of the model without the term (percentage points).
#      It is the part of the variance that ONLY this term explains; what the term shares with the other terms is not counted.
#  (b) change in each variance component = (variance in the reduced model - variance in the full model), in percentage points
#      of the total variance of the full model. Positive = that level absorbs variance that the removed term used to explain.
#      For example, if removing "method" raises the study variance, the method effect was explaining differences among studies.
#  (c) change in ELPD (expected log predictive density). ELPD measures how well the model predicts an estimate that was NOT
#      used to fit it: for each estimate, the model is refitted (approximately, by importance sampling, "PSIS-LOO") without that
#      estimate, and the log of the probability it assigns to the observed value is summed over estimates. Higher ELPD = better
#      out-of-sample prediction. Delta ELPD = ELPD(full) - ELPD(without the term); POSITIVE means the term improves prediction.
#      Its standard error (SE) comes from the estimate-by-estimate differences. As a rule of thumb, a difference larger than
#      about 2 SE is a clear difference; smaller than that is not distinguishable from noise. LOOIC = -2 x ELPD is the same
#      quantity on the "deviance" scale of AIC (lower = better).
# Draws of two different models are paired by their order in the output; because the models are fitted independently this is
# an approximation (it ignores the correlation between models), so intervals of (a) and (b) are slightly conservative or
# optimistic and should be read as approximate. ELPD differences do not have that problem (they are paired by estimate).
# Columns of drop_tab (one row per removed term):
#   unique_r2_*         : percentage points of conditional R2 lost when the term is removed (what ONLY this term explains).
#   d_group/study/pop/resid_* : change in the share of each variance level (pp). Shows WHERE the term's variance goes when removed.
#   delta_elpd, se_elpd : loss of out-of-sample prediction when removed. Positive = the term helps; clear if larger than ~2 SE.
#   delta_looic         : the same on the LOOIC scale (-2 x delta_elpd; positive = removing the term makes the model worse).
elpd_diff <- function(f_full, f_red) {
  d <- f_full$criteria$loo$pointwise[, "elpd_loo"] - f_red$criteria$loo$pointwise[, "elpd_loo"]   # per estimate
  c(diff = sum(d), se = sqrt(length(d) * var(d)))
}
drop_tab <- do.call(rbind, lapply(responses, function(r) {
  full <- parts[[r]]$M1_full
  do.call(rbind, lapply(fixed_terms(r), function(t) {
    red <- parts[[r]][[paste0("M1_no_", t)]]
    r2c_full <- 1 - full$vc$residual / full$total
    r2c_red  <- 1 - red$vc$residual / red$total
    dv <- lapply(names(level_names), function(l) post_sum(100 * (red$vc[[l]] - full$vc[[l]]) / full$total))
    names(dv) <- names(level_names)
    e <- elpd_diff(fits[[r]]$M1_full, fits[[r]][[paste0("M1_no_", t)]])
    data.frame(response = r, term = term_labels[[t]],
               unique_r2_median = 100 * median(r2c_full - r2c_red), unique_r2_lower = 100 * hdi(r2c_full - r2c_red)[[1]],
               unique_r2_upper = 100 * hdi(r2c_full - r2c_red)[[2]],
               d_group_med = dv$group[["median"]], d_group_lo = dv$group[["lower"]], d_group_hi = dv$group[["upper"]],
               d_study_med = dv$study[["median"]], d_study_lo = dv$study[["lower"]], d_study_hi = dv$study[["upper"]],
               d_pop_med = dv$population[["median"]], d_pop_lo = dv$population[["lower"]], d_pop_hi = dv$population[["upper"]],
               d_resid_med = dv$residual[["median"]], d_resid_lo = dv$residual[["lower"]], d_resid_hi = dv$residual[["upper"]],
               delta_elpd = e[["diff"]], se_elpd = e[["se"]], delta_looic = -2 * e[["diff"]], row.names = NULL)
  }))
}))

# 9.6 Model comparison table: ELPD, LOOIC and WAIC of every model -----------------------------------------------------------
# AIC and BIC are not used: they require a maximum-likelihood fit and a count of free parameters, which is ambiguous for models
# with random effects, and they are not defined for Bayesian fits. Their Bayesian counterparts are LOOIC (= -2 ELPD, from
# cross-validation) and WAIC. p_loo is the "effective number of parameters" (how flexible the model is in practice).
# Delta ELPD here is relative to the FULL model of the same response (positive = better than the full model).
# In short: compare models of the SAME response only, and look at the difference to the full model, not at the absolute value.
# ELPD (higher is better) and LOOIC (lower is better) are the main scores; WAIC is shown for completeness. p_loo much larger than
# the number of parameters, or many Pareto k > 0.7, warns that the score of that model is less reliable.
model_desc <- c(M0_null = "Random intercepts only (variance partition)", M1_full = "Full model of the plan",
                M1_no_species = "Full model without species", M1_no_sex = "Full model without sex",
                M1_no_method = "Full model without method class", M1_no_lat10 = "Full model without latitude")
model_tab <- do.call(rbind, lapply(responses, function(r) {
  do.call(rbind, lapply(names(fits[[r]]), function(nm) {
    f <- fits[[r]][[nm]]; e <- elpd_diff(f, fits[[r]]$M1_full)             # ELPD of this model minus ELPD of the full model
    data.frame(response = r, model = nm, description = model_desc[[nm]],
               formula = paste(deparse(formula(f)$formula), collapse = ""), n = nobs(f),
               elpd_loo = f$criteria$loo$estimates["elpd_loo", "Estimate"], se_elpd = f$criteria$loo$estimates["elpd_loo", "SE"],
               p_loo = f$criteria$loo$estimates["p_loo", "Estimate"], looic = f$criteria$loo$estimates["looic", "Estimate"],
               waic = f$criteria$waic$estimates["waic", "Estimate"],
               delta_elpd_vs_full = if (nm == "M1_full") 0 else e[["diff"]], se_delta = if (nm == "M1_full") NA else e[["se"]],
               row.names = NULL)
  }))
}))


# ======================================================================================================================
#  10. INFERENCE ON THE FIXED EFFECTS AND THE EQUIVALENCE TEST FOR LATITUDE (analysis plan, Section 5.3)
# ======================================================================================================================
# All results below come from the FULL model (M1_full) of each response. Estimates are posterior medians with 95% highest-
# density intervals (HDI) and the probability of direction (pd = the share of the posterior on the same side of zero as the
# median; 50% = no direction at all, 100% = every draw has the same sign).
# Percent effects: on the natural-log scale (k, L50) an effect b means a change of 100 x (exp(b) - 1) percent. For phi'
# (log10 scale) b is a difference in log10 phi'; 100 x (10^b - 1) is the percent change in phi'.

# 10.1 Fixed effects ----------------------------------------------------------------------------------------------------------
# In short: one row per coefficient = difference to the REFERENCE level (species: X. kroyeri; sex: F; method: modal progression;
# latitude: per 10 degrees southward). "estimate" is on the model scale, "pct_*" is the same effect in percent, "pd" the
# probability of direction. For comparisons between non-reference levels use the contrasts of 10.2.
# coef_labels(): readable names of the coefficients, in the same order as the "b_" columns of the model
coef_labels <- function(fit) {
  fx <- nobars_f(fit$formula$formula); X <- model.matrix(fx, data = fit$data)[, -1, drop = FALSE]
  asg <- attr(model.matrix(fx, data = fit$data), "assign")[-1]; lab <- attr(terms(fx), "term.labels")
  data.frame(term = term_labels[lab[asg]], level = ifelse(lab[asg] == "lat10", "per 10 degrees southward",
             mapply(function(cn, tl) sub(paste0("^", tl), "", cn), colnames(X), lab[asg])),
             stringsAsFactors = FALSE, row.names = NULL)
}
pd_of <- function(x) max(mean(x > 0), mean(x < 0))
fixed_tab <- do.call(rbind, lapply(responses, function(r) {
  fit <- fits[[r]]$M1_full; dr <- as_draws_df(fit)
  bn <- setdiff(grep("^b_", names(dr), value = TRUE), "b_Intercept"); lab <- coef_labels(fit)
  mult <- if (resp_def[[r]]$base == "e") exp(1) else 10                    # base of the logarithm of the response
  do.call(rbind, lapply(seq_along(bn), function(i) {
    b <- dr[[bn[i]]]; s <- post_sum(b); sp <- post_sum(100 * (mult^b - 1))
    data.frame(response = r, term = lab$term[i], level = lab$level[i], estimate = s[["median"]], lower = s[["lower"]],
               upper = s[["upper"]], pct_median = sp[["median"]], pct_lower = sp[["lower"]], pct_upper = sp[["upper"]],
               pd = 100 * pd_of(b), row.names = NULL)
  }))
}))

# 10.2 Contrasts among species (H1), method classes (H3) and sexes (H5) ---------------------------------------------------
# emmeans computes the "estimated marginal mean" of each category (model prediction for that category with the other terms held
# at their average) from every posterior draw; a contrast is the difference between two categories. All pairs are reported.
# Contrasts are given as "A - B": positive means A is larger than B.
# In short: this gives every pairwise comparison with its own interval (H1 species, H3 method classes, H5 sex), adjusted for the
# other terms. Read the percent column ("A is x% larger/smaller than B") and the probability of direction.
contrast_tab <- function(r, factor_name, label) {
  fit <- fits[[r]]$M1_full
  if (!factor_name %in% attr(terms(nobars_f(fit$formula$formula)), "term.labels")) return(NULL)
  em  <- emmeans(fit, specs = factor_name)                               # marginal means (posterior draws)
  ct  <- contrast(em, method = "pairwise")                               # all pairs
  dr  <- as.matrix(coda::as.mcmc(ct))                                    # draws of every contrast (draws x pairs)
  mult <- if (resp_def[[r]]$base == "e") exp(1) else 10
  do.call(rbind, lapply(colnames(dr), function(cn) {
    b <- dr[, cn]; s <- post_sum(b); sp <- post_sum(100 * (mult^b - 1))
    data.frame(response = r, factor = label, contrast = sub("^contrast ", "", gsub("[()]", "", cn)), estimate = s[["median"]], lower = s[["lower"]],
               upper = s[["upper"]], pct_median = sp[["median"]], pct_lower = sp[["lower"]], pct_upper = sp[["upper"]],
               pd = 100 * pd_of(b), row.names = NULL)
  }))
}
contrast_species <- do.call(rbind, lapply(responses, contrast_tab, factor_name = "species", label = "Species"))
contrast_method  <- do.call(rbind, lapply(responses, contrast_tab, factor_name = "method",  label = "Method class"))
contrast_sex     <- do.call(rbind, lapply(responses, contrast_tab, factor_name = "sex",     label = "Sex"))
# H5 (L50 larger in females) is read from the row "F - M" of the sex contrasts of the two L50 models: a positive estimate with a
# high probability of direction supports it.

# 10.3 H2: is the latitude slope practically zero? (equivalence test) ----------------------------------------------------
# The slope of lat10 is the change in the response for each 10 degrees SOUTHWARD, after accounting for species, sex, method
# and the three random intercepts. The plan fixes, BEFORE looking at the results, a "region of practical equivalence" (ROPE):
# a change smaller than +/- 10% in k per 10 degrees is considered biologically irrelevant. In the model scale this is
# +/- 0.0953 (= ln 1.1) for log k and +/- 0.0414 (= log10 1.1) for phi' (a 10% change in k at constant Linf).
# Decision rule (Kruschke 2018), applied to the 95% HDI of the slope:
#   HDI entirely inside the ROPE      -> "practically null"      (H2 supported: no relevant latitude effect)
#   HDI partly inside, partly outside -> "inconclusive"          (the data cannot exclude a relevant effect)
#   HDI entirely outside the ROPE     -> "different from zero"   (a relevant latitude effect)
# An interval that contains zero but crosses the ROPE is NOT evidence of absence. "% in ROPE" is the share of the posterior
# inside the margin. For comparison among responses the slope is also expressed as the percent change in k per 10 degrees.
h2_tab <- do.call(rbind, lapply(responses, function(r) {
  def <- resp_def[[r]]; fit <- fits[[r]]$M1_full; b <- as_draws_df(fit)$b_lat10
  mult <- if (def$base == "e") exp(1) else 10
  s <- post_sum(b); sp <- post_sum(100 * (mult^b - 1))
  rope <- def$rope
  decision <- if (is.na(rope)) "not tested (H2 concerns growth)" else if (s[["lower"]] >= -rope && s[["upper"]] <= rope) "practically null"
              else if (s[["upper"]] < -rope || s[["lower"]] > rope) "different from zero" else "inconclusive"
  data.frame(response = r, n = nobs(fit), groups = length(unique(fit$data$research_group)),
             studies = length(unique(fit$data$primary_study_id)), populations = length(unique(fit$data$population_id)),
             slope = s[["median"]], lower = s[["lower"]], upper = s[["upper"]], rope = rope,
             pct_k_median = sp[["median"]], pct_k_lower = sp[["lower"]], pct_k_upper = sp[["upper"]],
             pct_in_rope = if (is.na(rope)) NA else 100 * mean(abs(b) < rope), pd = 100 * pd_of(b), decision = decision,
             row.names = NULL)
}))
h2_tab


# ======================================================================================================================
#  11. MODEL CHECKS: CONVERGENCE, IDENTIFIABILITY, RESIDUALS AND FIT (analysis plan, Sections 5.1 and 7)
# ======================================================================================================================
# A model is only interpreted if the MCMC converged and its predictions are compatible with the data. Checks done here:
#  11.1 Convergence (plan Section 5.1): R-hat <= 1.01 (the chains agree), effective sample size >= 400 (enough independent
#       information in the draws; both "bulk" and "tail" are reported) and zero divergent transitions.
#  11.2 Identifiability (plan Section 7): rank of the design matrix of the fixed effects. If rank < number of columns, some
#       fixed effect is not estimable (it is a combination of the others); the group random effect is NOT removed to fix this.
#  11.3 Residuals and fit of the full model: residual = observed - predicted. "Standardized" residuals are divided by the
#       residual standard deviation (sigma) and should look like a standard normal variable (no pattern against the fitted values;
#       points on the diagonal of the QQ plot; very few beyond +/-3).
#  11.4 Posterior predictive check: new data sets are simulated from the fitted model; if the model is adequate, the observed
#       data should look like one of them. Four summaries (mean, sd, minimum, maximum) are compared and a posterior predictive
#       p-value = share of simulated data sets with a value >= the observed one is computed (values near 0 or 1, below
#       0.05 or above 0.95, flag a feature of the data that the model does not reproduce).

# 11.1 Convergence of every model -------------------------------------------------------------------------------------------
conv_tab <- do.call(rbind, lapply(responses, function(r) do.call(rbind, lapply(names(fits[[r]]), function(nm) {
  f <- fits[[r]][[nm]]
  sm <- summarise_draws(subset_draws(as_draws_df(f), variable = c("^b_", "^sd_", "^sigma"), regex = TRUE), "rhat", "ess_bulk", "ess_tail")
  data.frame(response = r, model = nm, max_rhat = max(sm$rhat), min_ess_bulk = min(sm$ess_bulk), min_ess_tail = min(sm$ess_tail),
             divergent = f$n_divergent, adapt_delta = f$adapt_delta_used,
             passed = max(sm$rhat) <= 1.01 && min(sm$ess_bulk) >= 400 && min(sm$ess_tail) >= 400 && f$n_divergent == 0,
             row.names = NULL)
}))))

# 11.2 Rank of the design matrix of the fixed effects (full models) --------------------------------------------------------
rank_tab <- do.call(rbind, lapply(responses, function(r) {
  fit <- fits[[r]]$M1_full; X <- model.matrix(nobars_f(fit$formula$formula), data = fit$data)
  data.frame(response = r, columns = ncol(X), rank = qr(X)$rank, full_rank = qr(X)$rank == ncol(X), row.names = NULL)
}))

# 11.3 Residuals and fit of the full models --------------------------------------------------------------------------------
# fitted_cond : prediction that includes the random intercepts (what the model predicts for an estimate of that group/study/pop.)
# fitted_fix  : prediction from the fixed effects alone (what a NEW group, study and population would be predicted to give)
# The RMSE (root mean squared error) and MAE (mean absolute error) are in the units of the response (log k, log10 phi', log L50).
# How to read fit_tab: sigma = typical error left after all terms; rmse_cond (with random intercepts) vs rmse_fixed (fixed effects
# only) shows how much the group/study/population intercepts help; sd_z close to 1, few |z| > 3 and a large shapiro_p mean
# approximately normal residuals; p_mean/p_sd/p_min/p_max near 0 or 1 (below 0.05 / above 0.95) flag features of the data
# (spread, extremes) that the model does not reproduce; pareto_k_gt07 counts influential estimates (see Section 8.5).
set.seed(seed_all)                                                       # simulations below are random
check_data <- list(); ppc_p <- list()
fit_tab <- do.call(rbind, lapply(responses, function(r) {
  fit <- fits[[r]]$M1_full; y <- fit$data$y
  fitted_cond <- fitted(fit)[, "Estimate"]
  fitted_fix  <- fitted(fit, re_formula = NA)[, "Estimate"]
  sigma_hat   <- median(as_draws_df(fit)$sigma)
  res         <- y - fitted_cond                                          # residuals (observed - predicted)
  z           <- res / sigma_hat                                          # standardized residuals
  # 11.4 posterior predictive check: 1000 simulated data sets (rows) x n estimates (columns)
  yrep <- posterior_predict(fit, ndraws = 1000)
  stats <- list(mean = mean, sd = sd, min = min, max = max)
  p_ppc <- sapply(stats, function(fn) mean(apply(yrep, 1, fn) >= fn(y)))
  check_data[[r]] <<- list(y = y, yrep = yrep[sample(nrow(yrep), 50), ], fitted = fitted_cond, z = z)
  lo <- fit$criteria$loo
  data.frame(response = r, n = length(y), sigma = sigma_hat,
             rmse_cond = sqrt(mean(res^2)), rmse_fixed = sqrt(mean((y - fitted_fix)^2)), mae_cond = mean(abs(res)),
             sd_z = sd(z), n_abs_z_gt3 = sum(abs(z) > 3), shapiro_p = shapiro.test(z)$p.value,
             p_mean = p_ppc[["mean"]], p_sd = p_ppc[["sd"]], p_min = p_ppc[["min"]], p_max = p_ppc[["max"]],
             elpd_loo = lo$estimates["elpd_loo", "Estimate"], se_elpd = lo$estimates["elpd_loo", "SE"],
             p_loo = lo$estimates["p_loo", "Estimate"], pareto_k_gt07 = sum(lo$diagnostics$pareto_k > 0.7),
             row.names = NULL)
}))
fit_tab


# ======================================================================================================================
#  12. TABLES OF THE MODELS (written to results/Table_models_H1-H5.xlsx)
# ======================================================================================================================
# Paper-ready tables (one sheet each, numbers already formatted as text "median [95% HDI]"):
#   Table5_variance_partition : variance partition and contribution of each term (H4; plan Section 5.2)
#   Table6_H2_latitude        : equivalence test for the latitude slope (H2; plan Section 5.3)
#   Table7_model_fit          : convergence, residuals and fit of the full models (plan Sections 5.1 and 7)
# Supplementary sheets:
#   S1_models_ELPD_LOOIC_WAIC : every model with its prediction scores (ELPD, LOOIC, WAIC, p_loo) and the difference to the full model.
#   S2_fixed_effects          : coefficient of every level against its reference level (model scale and percent change).
#   S3_species_contrasts_H1   : all pairwise species comparisons.      S4_method_contrasts_H3 : all pairwise method comparisons.
#   S5_sex_contrast_H5        : female - male (read the F - M row for H5).
#   S6_convergence_all_models : R-hat, effective sample size and divergences of EVERY model (check this before interpreting).
#   S7_numbers_variance       : the unformatted numbers behind Table 5 (Section 9).
#   S8_term_removal_detail    : the 95% intervals of the changes shown as medians in Table 5.
DELTA <- "\u0394"; R2 <- "R\u00b2"; PM <- "\u00b1"       # Delta, R-squared, plus-minus (unicode escapes keep the file ASCII)
resp_lab <- sapply(resp_def, function(d) d$label)                          # response labels used in the tables
fmt_ci <- function(med, lo, hi, d = 1) sprintf("%.*f [%.*f, %.*f]", d, med, d, lo, d, hi)   # "median [lower, upper]"
fmt_se <- function(est, se, d = 1) sprintf("%.*f (%.*f)", d, est, d, se)                       # "estimate (SE)"

# 12.1 Table 5: variance partition and contribution of each fixed term ------------------------------------------------------
t5_names <- c("Response", "Component or term", "Null model (% of total variance)", "Full model (% of total variance)",
              paste0("Unique ", R2, " lost when removed (pp)"), paste0(DELTA, "ELPD when removed (SE)"),
              paste0(DELTA, " research-group variance (pp)"), paste0(DELTA, " study variance (pp)"),
              paste0(DELTA, " population variance (pp)"), paste0(DELTA, " residual variance (pp)"))
t5_row <- function(resp, item, null = "", full = "", ...) {          # one table row = character vector of 10 cells
  v <- c(resp, item, null, full, rep("", 6)); extra <- list(...)
  for (nm in names(extra)) v[match(nm, c("r2", "elpd", "dg", "ds", "dp", "dr")) + 4] <- extra[[nm]]
  v
}
table5_rows <- do.call(rbind, lapply(responses, function(r) {
  rl <- resp_lab[[r]]; vn <- vpc_null[vpc_null$response == r, ]; sf <- share_full[share_full$response == r, ]
  dt <- drop_tab[drop_tab$response == r, ]; tv <- term_var[term_var$response == r, ]
  rows <- lapply(unname(level_names), function(l) {
    a <- vn[vn$component == l, ]; b <- sf[sf$component == l, ]
    t5_row(rl, paste("Random:", tolower(l)), fmt_ci(a$median, a$lower, a$upper), fmt_ci(b$median, b$lower, b$upper))
  })
  rows <- c(rows, lapply(c("Fixed effects (marginal R2)", "Fixed + random (conditional R2)"), function(i) {
    b <- sf[sf$component == i, ]; t5_row(rl, sub("R2", R2, i, fixed = TRUE), "", fmt_ci(b$median, b$lower, b$upper))
  }))
  rows <- c(rows, lapply(seq_len(nrow(dt)), function(i) {
    d <- dt[i, ]; b <- tv[tv$term == d$term, ]
    t5_row(rl, paste("Fixed:", tolower(d$term)), "", fmt_ci(b$median, b$lower, b$upper),
           r2 = sprintf("%.1f", d$unique_r2_median), elpd = fmt_se(d$delta_elpd, d$se_elpd),
           dg = sprintf("%.1f", d$d_group_med), ds = sprintf("%.1f", d$d_study_med),
           dp = sprintf("%.1f", d$d_pop_med), dr = sprintf("%.1f", d$d_resid_med))
  }))
  b <- tv[tv$term == "Shared among terms", ]
  rows <- c(rows, list(t5_row(rl, "Fixed: shared among terms", "", fmt_ci(b$median, b$lower, b$upper))))
  do.call(rbind, rows)
}))
table5 <- as.data.frame(table5_rows, stringsAsFactors = FALSE); names(table5) <- t5_names   # names set last (unicode-safe)
note5 <- paste0("Random and fixed-effect shares: posterior median [95% HDI] of the share of total variance. Null model = random intercepts only; ",
  "full model = species + sex + method + latitude + random intercepts (method omitted for L50). 'Fixed' rows: variance explained by the ",
  "term in the full model; 'shared' = overlap among terms (can be negative). The last six columns compare the full model with the model ",
  "WITHOUT that term: unique ", R2, " = loss of conditional ", R2, " (percentage points, pp); ", DELTA, "ELPD = ELPD(full) - ELPD(without the term), ",
  "positive = the term improves out-of-sample prediction, clear when larger than about 2 SE; ", DELTA, " variance = change of each variance ",
  "component when the term is removed (pp of the total variance of the full model; positive = that level absorbs variance the term used to explain). ",
  "Medians are shown for the differences between models; their 95% HDI are in sheet S8 (they combine draws of independent fits, so they are ",
  "approximate and conservative).")

# 12.2 Table 6: equivalence test of the latitude slope (H2) -----------------------------------------------------------------
table6 <- do.call(rbind, lapply(seq_len(nrow(h2_tab)), function(i) {
  h <- h2_tab[i, ]
  data.frame(Response = resp_lab[[h$response]], `Estimates / groups / studies / populations` = paste(h$n, h$groups, h$studies, h$populations, sep = " / "),
             `Slope per 10 deg southward [95% HDI]` = fmt_ci(h$slope, h$lower, h$upper, 3),
             `Change in k per 10 deg, % [95% HDI]` = fmt_ci(h$pct_k_median, h$pct_k_lower, h$pct_k_upper, 1),
             `ROPE (model scale)` = if (is.na(h$rope)) "-" else sprintf("%s%.3f", PM, h$rope),
             `Posterior inside ROPE (%)` = if (is.na(h$pct_in_rope)) "-" else sprintf("%.1f", h$pct_in_rope),
             `Probability of direction (%)` = sprintf("%.1f", h$pd), Decision = h$decision, check.names = FALSE, stringsAsFactors = FALSE)
}))
note6 <- paste0("Slope of latitude (lat10: positive = southward) in the full model of each response. ROPE = region of practical equivalence fixed in the ",
  "plan: ", PM, "10% change in k per 10 degrees (", PM, "0.0953 on the ln scale for log k; ", PM, "0.0414 on the log10 scale for phi'). 'Practically null' = the ",
  "whole 95% HDI is inside the ROPE; 'inconclusive' = the HDI crosses the ROPE limit; 'different from zero' = the HDI is entirely outside. ",
  "For phi' the change in k is that at constant Linf. L50 rows are descriptive: H2 concerns growth.")

# 12.3 Table 7: convergence, residuals and fit of the full models ------------------------------------------------------------
t7_row <- function(r) {
  f  <- fit_tab[fit_tab$response == r, ]
  cv <- conv_tab[conv_tab$response == r & conv_tab$model == "M1_full", ]; rk <- rank_tab[rank_tab$response == r, ]
  sg <- post_sum(as_draws_df(fits[[r]]$M1_full)$sigma)
  sh <- share_full[share_full$response == r, ]
  pick <- function(comp) sh[sh$component == comp, ]
  mg <- pick("Fixed effects (marginal R2)"); cn <- pick("Fixed + random (conditional R2)")
  v <- c(resp_lab[[r]], f$n, sprintf("%.3f", cv$max_rhat), paste(round(cv$min_ess_bulk), round(cv$min_ess_tail), sep = " / "),
         sprintf("%d (%.2f)", cv$divergent, cv$adapt_delta), paste(rk$rank, rk$columns, sep = " / "),
         fmt_ci(sg[["median"]], sg[["lower"]], sg[["upper"]], 3),
         fmt_ci(mg$median, mg$lower, mg$upper, 0), fmt_ci(cn$median, cn$lower, cn$upper, 0),
         sprintf("%.3f", f$rmse_cond), sprintf("%.3f", f$rmse_fixed), sprintf("%.3f", f$mae_cond),
         sprintf("%.2f", f$sd_z), f$n_abs_z_gt3, sprintf("%.3f", f$shapiro_p),
         sprintf("%.2f / %.2f / %.2f / %.2f", f$p_mean, f$p_sd, f$p_min, f$p_max),
         fmt_se(f$elpd_loo, f$se_elpd), sprintf("%.1f", f$p_loo), f$pareto_k_gt07)
  v
}
t7_names <- c("Response", "Estimates", "Max R-hat", "Min ESS bulk / tail", "Divergent (adapt_delta)", "Design matrix rank / columns",
              "Residual SD (sigma) [95% HDI]", paste0("Marginal ", R2, " (%)"), paste0("Conditional ", R2, " (%)"),
              "RMSE, fixed + random", "RMSE, fixed only", "MAE", "SD of standardized residuals", "|z| > 3 (n)", "Shapiro-Wilk p",
              "Predictive p: mean / sd / min / max", "ELPD loo (SE)", "p_loo", "Pareto k > 0.7 (n)")
table7 <- as.data.frame(do.call(rbind, lapply(responses, t7_row)), stringsAsFactors = FALSE); names(table7) <- t7_names
note7 <- paste0("Full models of the plan. Criteria: R-hat <= 1.01, effective sample size (ESS) >= 400, no divergent transitions (if any occurred, the model was ",
  "refitted with adapt_delta = 0.99). Design matrix rank equal to the number of columns = all fixed effects estimable. RMSE/MAE in units of the response. ",
  "Standardized residuals = residual / sigma. Predictive p = share of 1000 simulated data sets with a statistic >= the observed one (values < 0.05 or > 0.95 flag misfit). ",
  "Pareto k > 0.7 = estimates for which the leave-one-out approximation is unreliable.")

# 12.4 Supplementary tables --------------------------------------------------------------------------------------------------
s1 <- data.frame(Response = resp_lab[model_tab$response], Model = model_tab$model, Description = model_tab$description, Formula = model_tab$formula,
                 Estimates = model_tab$n, `ELPD loo (SE)` = fmt_se(model_tab$elpd_loo, model_tab$se_elpd), p_loo = round(model_tab$p_loo, 1),
                 LOOIC = round(model_tab$looic, 1), WAIC = round(model_tab$waic, 1),
                 `ELPD vs full model (SE)` = ifelse(model_tab$model == "M1_full", "-", fmt_se(model_tab$delta_elpd_vs_full, model_tab$se_delta)),
                 check.names = FALSE, stringsAsFactors = FALSE)
fmt_eff <- function(d) data.frame(Response = resp_lab[d$response], d[, intersect(c("factor", "term", "level", "contrast"), names(d)), drop = FALSE],
                 `Estimate [95% HDI]` = fmt_ci(d$estimate, d$lower, d$upper, 3), `Percent change [95% HDI]` = fmt_ci(d$pct_median, d$pct_lower, d$pct_upper, 1),
                 `Probability of direction (%)` = sprintf("%.1f", d$pd), check.names = FALSE, stringsAsFactors = FALSE, row.names = NULL)
s2 <- fmt_eff(fixed_tab); s3 <- fmt_eff(contrast_species); s4 <- fmt_eff(contrast_method); s5 <- fmt_eff(contrast_sex)
s6 <- data.frame(Response = resp_lab[conv_tab$response], Model = conv_tab$model, `Max R-hat` = round(conv_tab$max_rhat, 3), `Min ESS bulk` = round(conv_tab$min_ess_bulk),
                 `Min ESS tail` = round(conv_tab$min_ess_tail), Divergent = conv_tab$divergent, adapt_delta = conv_tab$adapt_delta,
                 Passed = ifelse(conv_tab$passed, "yes", "NO"), check.names = FALSE, stringsAsFactors = FALSE)
s8 <- data.frame(Response = resp_lab[drop_tab$response], `Term removed` = drop_tab$term,
                 `Unique R2 lost, pp [95% HDI]` = fmt_ci(drop_tab$unique_r2_median, drop_tab$unique_r2_lower, drop_tab$unique_r2_upper),
                 `Delta ELPD (SE)` = fmt_se(drop_tab$delta_elpd, drop_tab$se_elpd), `Delta LOOIC` = round(drop_tab$delta_looic, 1),
                 `Delta group variance, pp [95% HDI]` = fmt_ci(drop_tab$d_group_med, drop_tab$d_group_lo, drop_tab$d_group_hi),
                 `Delta study variance, pp [95% HDI]` = fmt_ci(drop_tab$d_study_med, drop_tab$d_study_lo, drop_tab$d_study_hi),
                 `Delta population variance, pp [95% HDI]` = fmt_ci(drop_tab$d_pop_med, drop_tab$d_pop_lo, drop_tab$d_pop_hi),
                 `Delta residual variance, pp [95% HDI]` = fmt_ci(drop_tab$d_resid_med, drop_tab$d_resid_lo, drop_tab$d_resid_hi),
                 check.names = FALSE, stringsAsFactors = FALSE)
s7 <- rbind(cbind(table = "VPC null model (%)", vpc_null), cbind(table = "Full model share (%)", share_full), cbind(table = "Term variance (%)", setNames(term_var, c("response", "component", "median", "lower", "upper"))))
noteS <- "Percent change = 100 x (exp(b) - 1) for ln responses (k, L50) and 100 x (10^b - 1) for phi'. Contrasts are 'A - B'. Intervals are 95% HDI."

wb5 <- createWorkbook()
add_table(wb5, "Table5_variance_partition", table5, note5)
add_table(wb5, "Table6_H2_latitude", table6, note6)
add_table(wb5, "Table7_model_fit", table7, note7)
add_table(wb5, "S1_models_ELPD_LOOIC_WAIC", s1, "ELPD, LOOIC and WAIC replace AIC/BIC in Bayesian models: higher ELPD (lower LOOIC/WAIC) = better out-of-sample prediction.")
add_table(wb5, "S2_fixed_effects", s2, noteS)
add_table(wb5, "S3_species_contrasts_H1", s3, noteS)
add_table(wb5, "S4_method_contrasts_H3", s4, noteS)
add_table(wb5, "S5_sex_contrast_H5", s5, paste(noteS, "H5: 'F - M' > 0 means females larger (L50 models)."))
add_table(wb5, "S6_convergence_all_models", s6)
add_table(wb5, "S7_numbers_variance", s7)
add_table(wb5, "S8_term_removal_detail", s8, "Intervals combine draws of independent fits (full vs reduced model): approximate and conservative.")
saveWorkbook(wb5, "results/Table_models_H1-H5.xlsx", overwrite = TRUE)


# ======================================================================================================================
#  13. FIGURES OF THE MODELS (saved to results/ as PNG and PDF)
# ======================================================================================================================
# Fig 5 variance partition | Fig 6 equivalence test of the latitude slope (H2) | Fig 7 residual and fit checks of the full models
resp_f <- factor(resp_lab, levels = resp_lab)                              # keeps the order of the responses in the panels

# 13.1 Figure 5: variance partition, null versus full model -------------------------------------------------------------------
# Each bar is 100% of the variance of the response. Null model: only random intercepts. Full model: the fixed effects take a
# share (marginal R2) and the random levels and the residual share the rest. Posterior MEANS are used so each bar adds to 100%.
comp_cols <- c(`Fixed effects` = "#2a78d6", `Research group` = "#eb6834", Study = "#1baf7a", Population = "#eda100", Residual = "#B9B8B3")
fig5 <- do.call(rbind, lapply(responses, function(r) {
  rbind(data.frame(response = resp_lab[[r]], model = "Null", component = names(comp_cols)[2:5],
                   pct = sapply(c("group", "study", "population", "residual"), function(l) 100 * mean(parts[[r]]$M0_null$vc[[l]] / parts[[r]]$M0_null$total))),
        data.frame(response = resp_lab[[r]], model = "Full", component = names(comp_cols),
                   pct = c(100 * mean(parts[[r]]$M1_full$vf / parts[[r]]$M1_full$total),
                           sapply(c("group", "study", "population", "residual"), function(l) 100 * mean(parts[[r]]$M1_full$vc[[l]] / parts[[r]]$M1_full$total)))))
}))
fig5$response  <- factor(fig5$response, levels = resp_lab)
fig5$model     <- factor(fig5$model, levels = c("Null", "Full"))
fig5$component <- factor(fig5$component, levels = rev(names(comp_cols)))   # stacking order: fixed effects at the bottom
dark_text <- c(`Fixed effects` = "white", `Research group` = "white", Study = "white", Population = "#262626", Residual = "#262626")
p5 <- ggplot(fig5, aes(x = model, y = pct, fill = component)) +
  geom_col(width = 0.72, colour = "white", linewidth = 0.6) +
  geom_text(aes(label = ifelse(pct >= 5, sprintf("%.0f", pct), ""), colour = component), position = position_stack(vjust = 0.5), size = 3) +
  facet_wrap(~ response, nrow = 1) +
  scale_fill_manual(values = comp_cols, name = NULL, breaks = names(comp_cols)) +
  scale_colour_manual(values = dark_text, guide = "none") +
  scale_y_continuous(expand = expansion(mult = c(0, 0.02))) +
  labs(x = NULL, y = "Share of total variance (%)",
       caption = "Posterior means (each bar adds to 100%). Null = random intercepts only; Full = fixed effects + random intercepts.") +
  theme_shrimp() + theme(panel.grid.major.x = element_blank(), plot.caption = element_text(colour = ink2, size = 8, hjust = 0))
p5
save_fig(p5, "Fig5_variance_partition", 13, 5)

# 13.2 Figure 6: equivalence test of the latitude slope (H2) ----------------------------------------------------------------
# Posterior distribution of the RATIO of k for each 10 degrees southward (1 = no change; 0.9 = k 10% lower), on a log axis so that
# increases and decreases are symmetric. Log k and phi' (which uses k at constant Linf) are shown on the same scale. The grey band is
# the region of practical equivalence fixed in the plan (ratios between 1/1.1 and 1.1); the thick bar is the 95% HDI and the circle
# the posterior median. HDI inside the band = practically null; crossing the band edge = inconclusive; outside = different from zero.
growth_resp <- responses[responses %in% c("log_k", "phi_cl", "phi_tl")]
if (length(growth_resp) > 0) {                                        # (skipped if only L50 responses are fitted)
  ratio_of <- function(r, b) (if (resp_def[[r]]$base == "e") exp(1) else 10)^b      # ratio of k implied by a slope b
  fig6 <- do.call(rbind, lapply(growth_resp, function(r) {
    lr <- log(ratio_of(r, as_draws_df(fits[[r]]$M1_full)$b_lat10))                   # log ratio (symmetric scale)
    dens <- density(lr, from = log(0.4), to = log(2.5), n = 512)
    data.frame(response = resp_lab[[r]], x = exp(dens$x), y = dens$y / max(dens$y))
  }))
  fig6$response <- factor(fig6$response, levels = resp_lab[growth_resp])
  hdi6 <- h2_tab[h2_tab$response %in% growth_resp, ]
  hdi6$response <- factor(resp_lab[hdi6$response], levels = resp_lab[growth_resp])
  hdi6$lo <- 1 + hdi6$pct_k_lower / 100; hdi6$hi <- 1 + hdi6$pct_k_upper / 100; hdi6$med <- 1 + hdi6$pct_k_median / 100
  p6 <- ggplot(fig6, aes(x = x, y = y)) +
    annotate("rect", xmin = 1 / 1.1, xmax = 1.1, ymin = -Inf, ymax = Inf, fill = "#D9D8D3", alpha = 0.6) +
    geom_vline(xintercept = 1, colour = "#6e6d69", linewidth = 0.4) +
    geom_area(fill = "#2a78d6", alpha = 0.45) + geom_line(colour = "#2a78d6", linewidth = 0.6) +
    geom_segment(data = hdi6, aes(x = lo, xend = hi, y = -0.06, yend = -0.06), inherit.aes = FALSE, linewidth = 1.6,
                 colour = "#262626", lineend = "round") +
    geom_point(data = hdi6, aes(x = med, y = -0.06), inherit.aes = FALSE, size = 2.6, shape = 21, fill = "white",
               colour = "#262626", stroke = 0.8) +
    geom_text(data = hdi6, aes(x = 2.4, y = 0.95, label = decision), inherit.aes = FALSE, hjust = 1, size = 3.4,
              fontface = "bold", colour = "#262626") +
    facet_wrap(~ response, ncol = 1) +
    scale_x_log10(limits = c(0.4, 2.5), breaks = c(0.5, 0.67, 0.8, 0.9, 1, 1.1, 1.25, 1.5, 2), labels = function(x) format(x, digits = 2)) +
    scale_y_continuous(breaks = NULL, expand = expansion(mult = c(0.02, 0.05))) +
    labs(x = "Ratio of k per 10 degrees southward (log axis; 1 = no change)", y = "Posterior density (scaled)",
         caption = "Grey band = region of practical equivalence (ratio 0.91 to 1.10); thick bar = 95% HDI; circle = posterior median.") +
    theme_shrimp() + theme(plot.caption = element_text(colour = ink2, size = 8, hjust = 0))
  p6
  save_fig(p6, "Fig6_H2_latitude_equivalence", 7.5, 6.5)
}

# 13.3 Figure 7: checks of the full models (rows: observed vs simulated data, residuals vs fitted, QQ plot) ------------------
# Row 1 (posterior predictive check): thick line = observed distribution of the response; thin grey lines = 50 simulated data
#   sets from the model. They should surround the observed line.
# Row 2: standardized residuals against fitted values; a flat cloud around 0 is expected; the line is a smooth trend.
#   The trend is a loess curve with span = 1 (it uses all the points for each local fit, so the curve is smooth and shows only
#   the broad shape; the default span of 0.75 follows small clusters of points and wiggles). Lower the span (e.g. 0.75) for a more
#   flexible curve or raise it (e.g. 1.5) for an even smoother one.
# Row 3: normal QQ plot of the standardized residuals; points on the diagonal = residuals as expected from a normal distribution.
chk_theme <- theme_shrimp() + theme(legend.position = "none", panel.grid.major.x = element_line(colour = grid_col, linewidth = 0.3))
chk_plots <- lapply(responses, function(r) {
  cd <- check_data[[r]]
  yr <- data.frame(value = as.vector(t(cd$yrep)), draw = rep(seq_len(nrow(cd$yrep)), each = ncol(cd$yrep)))
  a <- ggplot() + geom_density(data = yr, aes(x = value, group = draw), colour = "#B9B8B3", linewidth = 0.3) +
    geom_density(data = data.frame(value = cd$y), aes(x = value), colour = "#2a78d6", linewidth = 1) +
    labs(title = resp_lab[[r]], x = "Response", y = if (r == responses[1]) "Density" else NULL) + chk_theme
  rf <- data.frame(fitted = cd$fitted, z = cd$z)
  b <- ggplot(rf, aes(fitted, z)) + geom_hline(yintercept = 0, colour = "#6e6d69", linewidth = 0.4) +
    geom_point(colour = "#2a78d6", alpha = 0.6, size = 1.6) + geom_smooth(method = "loess", formula = y ~ x, span = 1, se = FALSE, colour = "#eb6834", linewidth = 0.6) +
    labs(x = "Fitted value", y = if (r == responses[1]) "Standardized residual" else NULL) + chk_theme
  q <- ggplot(rf, aes(sample = z)) + stat_qq(colour = "#2a78d6", alpha = 0.6, size = 1.6) + stat_qq_line(colour = "#6e6d69", linewidth = 0.4) +
    labs(x = "Theoretical quantile", y = if (r == responses[1]) "Sample quantile" else NULL) + chk_theme
  list(a, b, q)
})
p7 <- wrap_plots(c(lapply(chk_plots, `[[`, 1), lapply(chk_plots, `[[`, 2), lapply(chk_plots, `[[`, 3)), nrow = 3, byrow = TRUE)
p7
save_fig(p7, "Fig7_model_checks", 15, 8.5)

# 13.4 Record of the software versions (plan Section 9) ---------------------------------------------------------------------
writeLines(capture.output(sessionInfo()), "results/sessionInfo.txt")
