[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.186202190.svg)](https://doi.org/10.5281/zenodo.186202190)

# FES-MADM II — Case Study 2 (Replication Package, v1.0.1)
## Shock-aware supplier portfolio selection using open contracting data (UK Contracts Finder)

This repository provides the **replication package and computational artifacts** for **Case Study 2** of the  
**FES-MADM II (Fuzzy Entropy–Synergy Multi-Attribute Decision-Making)** framework.

It supports full **auditability** and **reproducibility** of the empirical analysis reported in the manuscript:

**“Shock-aware supplier portfolio selection using open contracting data:  
An auditable fuzzy entropy–synergy decision-analytics framework (FES-MADM II)”**

---

## Zenodo archival record (recommended citation)
This repository is archived on Zenodo (concept DOI): **10.5281/zenodo.186202190**  
Use the DOI link: https://doi.org/10.5281/zenodo.186202190

> Note: Zenodo also assigns version-specific DOIs. The concept DOI always resolves to the latest version.

---

## Scope and DOI separation (important)
This repository is a **case-study-level replication package** (Case Study 2).  
It **does not** represent the core FES-MADM II methodological platform.

- Cite the **platform DOI** for the general framework/methodology (archived separately).
- Cite **this Zenodo DOI** for the *data, results, figures, and scripts* associated with Case Study 2.

This separation prevents DOI duplication and ensures clear methodological provenance.

---

## Repository contents
- `data/` — Scenario-specific input datasets (S0–S4), platform-ready.
- `results/` — Computational outputs across the complete α-cut grid, audit indices, tables, and figures.
- `code/` — R implementation used to run FES-MADM II (v1.0.1) and generate outputs/plots.

---

## Scenarios (S0–S4)
The case study evaluates five analytically distinct scenarios derived from UK Contracts Finder open procurement data:

- **S0** — Baseline  
- **S1** — Inflation & market tightening  
- **S2** — Cyber procurement disruption  
- **S3** — Market consolidation  
- **S4** — Budget tightening  

Each scenario is evaluated over the **complete α-cut grid**, enabling comparative assessment of:
ranking stability, criterion discrimination power, and information-supported rank separation.

---

## Framework versioning
- **v1.0.0** — baseline implementation.
- **v1.0.1** — upgraded implementation supporting **automatic computation across all α-cut levels**.

All outputs provided in this repository correspond to **FES-MADM II v1.0.1**.

---

## Reproducibility (minimal workflow)
1. Ensure the scenario input files are located in `data/`.
2. Run the R script in `code/` (v1.0.1).
3. Outputs (rankings, indices, figures) are written to `results/`.

The results and figures in this repository correspond to those reported in the associated manuscript.

---

## License
Released under the **MIT License**.

---

## How to cite
If you use this repository, please cite the Zenodo record:

Kiratsoudis, S. (2026). *FES-MADM II — Case Study 2 (Replication Package, v1.0.1)*. Zenodo. https://doi.org/10.5281/zenodo.186202190
