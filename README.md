# FES-MADM-II — Case Study 2

This repository provides the replication package and computational artifacts for **Case Study 2** of the  
**FES-MADM II (Fuzzy Entropy–Synergy Multi-Attribute Decision-Making)** framework.

The case study corresponds to the article:

**“Shock-aware supplier portfolio selection using open contracting data:  
An auditable fuzzy entropy–synergy decision-analytics framework (FES-MADM II)”**

---

## Scope of this repository

This repository is **not** the core FES-MADM II platform.  
It documents a **fully reproducible application (Case Study 2)** and includes:

- scenario-specific input datasets (S0–S4),
- full computational outputs across all α-cut levels,
- result tables and figures,
- R scripts used for analysis and visualization.

The methodological platform **FES-MADM II** is referenced separately and should be cited via its dedicated DOI.

---

## Framework version

- **FES-MADM II v1.0.0**: baseline implementation.
- **FES-MADM II v1.0.1**: upgraded version supporting **automatic computation across all α-cut levels**.

All results in this repository are produced using **v1.0.1**.

---

## Repository structure
code/ R scripts (model execution and figure generation)
data/ Scenario input files (S0–S4)
results/ Computational outputs and figures


---

## Scenarios

The case study evaluates five platform-ready scenarios derived from UK Contracts Finder data:

- **S0** — Baseline
- **S1** — Inflation & market tightening
- **S2** — Cyber procurement disruption
- **S3** — Market consolidation
- **S4** — Budget tightening

Each scenario is evaluated over the full α-cut grid to assess robustness, discrimination, and auditability.

---

## Reproducibility

All results and figures reported in the associated manuscript can be reproduced using the material provided in this repository.

1. Place the scenario input files in `data/`
2. Execute the FES-MADM II R script located in `code/`
3. Generated outputs will be written to `results/`

---

## License

This repository is released under the **MIT License**.

---

## Citation

If you use this repository, please cite it as a **replication package for Case Study 2** of the FES-MADM II framework.  
The Zenodo DOI will be linked upon archival.
