# R scripts — FES-MADM II (v1.0.1)

This folder contains the **R implementation** used to execute the FES-MADM II framework
and to generate all results and figures for Case Study 2.

## Framework version
- **v1.0.0** — baseline implementation.
- **v1.0.1** — upgraded version supporting automatic computation across all α-cut levels.

All analyses in this repository are based on **v1.0.1**.

## Functionality
The script performs:
- loading of scenario-specific input datasets,
- execution of the FES-MADM II decision-analytics pipeline,
- computation of rankings, weights, and audit indices,
- generation of figures and tables reported in the manuscript.

## Execution
The script is designed to run end-to-end without manual intervention,
assuming the input data are placed in the `data/` directory.
