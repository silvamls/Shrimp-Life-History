# Sources of variability in the life-history parameters of Brazilian penaeid shrimps

Data and R code for a systematic map and Bayesian synthesis of life-history parameters (growth, mortality and maturity) of the five main penaeid shrimps exploited in Brazil: *Xiphopenaeus kroyeri*, *Litopenaeus schmitti*, *Farfantepenaeus subtilis*, *F. paulensis* and *F. brasiliensis*.

Manuscript in preparation: "Beyond biology: sources of variability in the life-history parameters of Brazilian penaeid shrimps" (Silva, Figueiredo and Frédou, Federal Rural University of Pernambuco). Pre-registered analysis plan: https://osf.io/vq2rx/

## About

Shrimp stock assessments in Brazil depend on life-history parameters that are scarce and obtained by different methods. This repository contains the database compiled from the published literature (1,322 estimates from 162 reports, 150 primary studies and 184 populations) and the code used to test whether the differences among published values reflect the biology of the shrimps (species, sex, latitude), the growth-estimation method, or who produced the values (research group, study and population).

## Contents

- `data/life_history_data_EN.xlsx`: the database, where each row is one estimate as published, with sheets describing fields, sources, studies, research groups, populations and quality-control flags.
- `code/Shrimp_life_history.R`: the full analysis, from data preparation to Bayesian models, variance partitioning and model checks, with every step commented.
- `results/`: figures, tables and session information produced by the script.
- `search/`: PRISMA flow diagram and the list of references from the search.
- `analysis_plan/`: the analysis plan registered on OSF.

## Methods in brief

Bayesian multilevel (mixed-effects) models were fitted with `brms` for the growth coefficient (*k*), growth performance (φ′) and length at maturity (*L*50), with species, sex, growth method and latitude as fixed effects and research group, study and population as random effects. Variance was partitioned among these levels and the contribution of each term was assessed by R² and leave-one-out predictive performance.

## How to run

The analysis was run in R 4.6.1 with `brms`, `posterior`, `emmeans`, `readxl`, `openxlsx`, `ggplot2`, `patchwork` and `glmmTMB` (versions in `results/sessionInfo.txt`). Stan needs a C++ toolchain (Rtools on Windows). Set the working directory at the top of the script and run it from top to bottom. Fitting the 28 models takes a few hours; the fitted models are not included because of their size.

## Citation and license

Silva MLS, Figueiredo ACA, Frédou FL (2026) Beyond biology: sources of variability in the life-history parameters of Brazilian penaeid shrimps. 

Data and code are released under CC BY 4.0.

## Contact

Matheus L. S. Silva, PGRPAq/UFRPE, matheus.ssilva@ufrpe.br
