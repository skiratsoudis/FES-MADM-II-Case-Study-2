###############################################################
# FES-MADM II — Robust Shiny App (All α-cut Scenarios + Insights)
# - Clean, source()-safe script (no stray text)
# - Supports: main α-run, sensitivity run, and batch run across α-cuts
# - Robust Excel import (sheet-name selection) + optional Labels sheet
# - Exports: Main results + All-α results + Detailed assessment to Excel
#
# Plotting refinement (ONLY change requested):
#   - All plots show ONLY nomenclature IDs:
#       Criteria: X1, X2, ..., Xμ
#       Alternatives: Y1, Y2, ..., Yν
#   - Added a "Nomenclature Map" tab (Tab 5) to display ID ↔ full name mapping.
#   - Rotated x-axis labels on crowded category plots for readability.
#
# Usage:
#   1) Save this file as app.R
#   2) In R/RStudio: shiny::runApp("app.R")
#
# Expected Excel structure (default sheet names):
#   1) Data_Center_xi   : MxN numeric matrix (ξ_{μν} center)
#   2) Data_Delta       : MxN numeric matrix (Δ_{μν})
#   3) SBJ_WCenter      : Mx1 numeric vector (x^SBJ_μ center)
#   4) SBJ_WDelta       : Mx1 numeric vector (Δx^SBJ_μ)
#   5) Criterion_Type   : Mx1 text vector, each 'B' or 'C'
#   6) AlphaCuts        : Kx1 numeric vector in [0,1] (optional)
#   7) Labels           : Optional (see template)
###############################################################


# ----------------------------
# Packages (auto-install)
# ----------------------------
install_if_missing <- function(pkgs) {
  for (p in pkgs) {
    if (!requireNamespace(p, quietly = TRUE)) {
      install.packages(p, dependencies = TRUE)
    }
  }
}

pkgs <- c("shiny","shinythemes","readxl","writexl","ggplot2","dplyr","DT","gridExtra","reshape2")
install_if_missing(pkgs)

suppressPackageStartupMessages({
  library(shiny)
  library(shinythemes)
  library(readxl)
  library(writexl)
  library(ggplot2)
  library(dplyr)
  library(DT)
  library(gridExtra)
  library(reshape2)
})

# ----------------------------
# Helpers (robust & consistent)
# ----------------------------
round3 <- function(x) round(x, 3)

clamp01 <- function(x) pmin(pmax(x, 0), 1)

log2safe <- function(x, eps = 1e-15) {
  ifelse(x > eps, log2(x), 0)
}

fixFuzzyPair <- function(L, U) {
  for (i in seq_along(L)) {
    if (L[i] > U[i]) {
      tmp <- L[i]; L[i] <- U[i]; U[i] <- tmp
    }
  }
  list(L = L, U = U)
}

as_num_df <- function(df) {
  # Coerce a data.frame read from Excel (possibly characters) to numeric, NAs -> 0.
  mat <- suppressWarnings(apply(df, 2, function(x) as.numeric(as.character(x))))
  mat[is.na(mat)] <- 0
  as.data.frame(mat)
}

as_chr_df <- function(df) {
  as.data.frame(lapply(df, function(x) toupper(trimws(as.character(x)))))
}

safe_alpha_vec <- function(v) {
  v <- suppressWarnings(as.numeric(v))
  v <- v[is.finite(v)]
  v <- unique(v)
  v <- v[v >= 0 & v <= 1]
  v <- sort(v)
  if (length(v) == 0) v <- seq(0, 1, by = 0.1)
  v
}

# ----------------------------
# Entropy / information measures
# ----------------------------
measure_entropy_alt <- function(vec) {
  val <- 0
  for (a in vec) {
    if (a > 0) val <- val - a * log2safe(a)
  }
  val
}

measure_jointEntropy <- function(pXY_vec) {
  sVal <- 0
  for (px in pXY_vec) {
    if (px > 0) sVal <- sVal - px * log2safe(px)
  }
  sVal
}

measure_condEntropy_altGivenCrit <- function(pY_given_X, pX, eps = 1e-15) {
  # H(Y|X) = Σ_x p(x) H(Y|x), where pY_given_X is MxN with rows summing to 1
  M <- nrow(pY_given_X); N <- ncol(pY_given_X)
  sumVal <- 0
  for (i in 1:M) {
    rowSum <- sum(pY_given_X[i, ])
    if (rowSum > eps) {
      tmpH <- 0
      for (j in 1:N) {
        pYx <- pY_given_X[i, j] / rowSum
        if (pYx > 0) tmpH <- tmpH - pYx * log2safe(pYx)
      }
      sumVal <- sumVal + pX[i] * tmpH
    }
  }
  sumVal
}

# ----------------------------
# Core FES-MADM II computation
# ----------------------------
fesMadmFull <- function(alpha, DC, DD, WC, WD, CT, eps = 1e-15, shiftVal = 1) {
  
  # --- dimensional checks
  if (nrow(DC) != nrow(DD) || ncol(DC) != ncol(DD)) stop("DC and DD must have identical dimensions (MxN).")
  if (nrow(WC) != nrow(WD)) stop("WC and WD must have same number of rows (M).")
  if (ncol(WC) != 1 || ncol(WD) != 1) stop("WC and WD must be Mx1 vectors.")
  if (nrow(DC) != nrow(WC)) stop("DC and WC must have same number of rows (M).")
  if (nrow(CT) != nrow(DC)) stop("CT must have M rows (one per criterion).")
  
  M <- nrow(DC)
  N <- ncol(DC)
  
  # 1) α-cut fuzzy performance intervals
  perfLower <- DC - (1 - alpha) * DD
  perfUpper <- DC + alpha * DD
  
  # 2) Normalize (Benefit/Cost), row-wise
  normLower <- matrix(0, M, N)
  normUpper <- matrix(0, M, N)
  
  for (i in 1:M) {
    rowL <- perfLower[i, ]
    rowU <- perfUpper[i, ]
    ctype <- toupper(as.character(CT[i, 1]))
    
    if (ctype == "B") {
      theMax <- max(rowU)
      if (theMax <= 0) {
        shiftNeeded <- abs(theMax) + shiftVal
        rowL <- rowL + shiftNeeded
        rowU <- rowU + shiftNeeded
        theMax <- max(rowU)
      }
      theMax <- max(theMax, eps)
      normLower[i, ] <- rowL / theMax
      normUpper[i, ] <- rowU / theMax
      
    } else if (ctype == "C") {
      theMin <- min(rowL)
      if (theMin <= 0) {
        shiftNeeded <- abs(theMin) + shiftVal
        rowL <- rowL + shiftNeeded
        rowU <- rowU + shiftNeeded
        theMin <- min(rowL)
      }
      theMin <- max(theMin, eps)
      for (j in 1:N) {
        normLower[i, j] <- theMin / (rowU[j] + eps)
        normUpper[i, j] <- theMin / (rowL[j] + eps)
      }
      fixNorm <- fixFuzzyPair(normLower[i, ], normUpper[i, ])
      normLower[i, ] <- fixNorm$L
      normUpper[i, ] <- fixNorm$U
    } else {
      stop(paste("Invalid Criterion Type at row", i, "- must be 'B' or 'C'."))
    }
  }
  
  # 3) Probabilistic representation p~_{μν} (row-wise)
  probLower <- matrix(0, M, N)
  probUpper <- matrix(0, M, N)
  for (i in 1:M) {
    sL <- sum(normLower[i, ]) + eps
    sU <- sum(normUpper[i, ]) + eps
    probLower[i, ] <- normLower[i, ] / sL
    probUpper[i, ] <- normUpper[i, ] / sU
  }
  
  # 4) α-cut subjective weights
  wL <- WC - (1 - alpha) * WD
  wU <- WC + alpha * WD
  wL[wL < 0] <- 0
  wU[wU < 0] <- 0
  fixSBJ <- fixFuzzyPair(wL, wU)
  wL <- as.numeric(fixSBJ$L)
  wU <- as.numeric(fixSBJ$U)
  wDefSBJ <- (wL + wU) / 2
  
  # 5) Criterion entropy (normalized), diversification d, objective weights x~^OBJ
  hLower <- numeric(M)
  hUpper <- numeric(M)
  for (i in 1:M) {
    hLower[i] <- measure_entropy_alt(probLower[i, ]) / log2(N)
    hUpper[i] <- measure_entropy_alt(probUpper[i, ]) / log2(N)
  }
  hLower[hLower < 0] <- 0
  hUpper[hUpper < 0] <- 0
  fixh <- fixFuzzyPair(hLower, hUpper)
  hLower <- clamp01(fixh$L)
  hUpper <- clamp01(fixh$U)
  
  dLower <- 1 - hUpper
  dUpper <- 1 - hLower
  fixd <- fixFuzzyPair(dLower, dUpper)
  dLower <- clamp01(fixd$L)
  dUpper <- clamp01(fixd$U)
  
  sumDL <- sum(dLower) + eps
  sumDU <- sum(dUpper) + eps
  xObjL <- dLower / sumDL
  xObjU <- dUpper / sumDU
  fixObj <- fixFuzzyPair(xObjL, xObjU)
  xObjL <- clamp01(fixObj$L)
  xObjU <- clamp01(fixObj$U)
  xObjMid <- (xObjL + xObjU) / 2
  xObjMid <- xObjMid / (sum(xObjMid) + eps)
  
  # 6) Integrated weights x~^INT
  denomL <- sum(wL * xObjL) + eps
  denomU <- sum(wU * xObjU) + eps
  xIntL <- (wL * xObjL) / denomL
  xIntU <- (wU * xObjU) / denomU
  fixInt <- fixFuzzyPair(xIntL, xIntU)
  xIntL <- clamp01(as.numeric(fixInt$L))
  xIntU <- clamp01(as.numeric(fixInt$U))
  # enforce exact normalization (numerical safety)
  xIntL <- xIntL / (sum(xIntL) + eps)
  xIntU <- xIntU / (sum(xIntU) + eps)
  xIntMid <- (xIntL + xIntU) / 2
  xIntMid <- xIntMid / (sum(xIntMid) + eps)
  
  # 7) Alternative scores p* (should already sum to 1; re-normalize L/U separately for safety)
  altScoreLower_raw <- numeric(N)
  altScoreUpper_raw <- numeric(N)
  for (j in 1:N) {
    altScoreLower_raw[j] <- sum(probLower[, j] * xIntL)
    altScoreUpper_raw[j] <- sum(probUpper[, j] * xIntU)
  }
  altScoreLower <- pmin(altScoreLower_raw, altScoreUpper_raw)
  altScoreUpper <- pmax(altScoreLower_raw, altScoreUpper_raw)
  
  denL <- sum(altScoreLower) + eps
  denU <- sum(altScoreUpper) + eps
  altScoreLower <- altScoreLower / denL
  altScoreUpper <- altScoreUpper / denU
  altMid <- (altScoreLower + altScoreUpper) / 2
  altMid <- altMid / (sum(altMid) + eps)
  
  # ----------------------------
  # Entropy measures (bits)
  # ----------------------------
  SY_L_raw <- measure_entropy_alt(altScoreLower)
  SY_U_raw <- measure_entropy_alt(altScoreUpper)
  fixSY <- fixFuzzyPair(max(SY_L_raw, 0), max(SY_U_raw, 0))
  SY_L <- fixSY$L; SY_U <- fixSY$U; SY_C <- (SY_L + SY_U) / 2
  
  SX_L_raw <- measure_entropy_alt(xIntL)
  SX_U_raw <- measure_entropy_alt(xIntU)
  fixSX <- fixFuzzyPair(max(SX_L_raw, 0), max(SX_U_raw, 0))
  SX_L <- fixSX$L; SX_U <- fixSX$U; SX_C <- (SX_L + SX_U) / 2
  
  pXY_L <- probLower * xIntL
  pXY_U <- probUpper * xIntU
  
  SXY_L_raw <- measure_jointEntropy(as.vector(pXY_L))
  SXY_U_raw <- measure_jointEntropy(as.vector(pXY_U))
  fixSXY <- fixFuzzyPair(max(SXY_L_raw, 0), max(SXY_U_raw, 0))
  SXY_L <- fixSXY$L; SXY_U <- fixSXY$U; SXY_C <- (SXY_L + SXY_U) / 2
  
  SYX_L_raw <- measure_condEntropy_altGivenCrit(probLower, xIntL)
  SYX_U_raw <- measure_condEntropy_altGivenCrit(probUpper, xIntU)
  fixSYX <- fixFuzzyPair(max(SYX_L_raw, 0), max(SYX_U_raw, 0))
  SYX_L <- fixSYX$L; SYX_U <- fixSYX$U; SYX_C <- (SYX_L + SYX_U) / 2
  
  IXY_L_raw <- SX_L + SY_L - SXY_L
  IXY_U_raw <- SX_U + SY_U - SXY_U
  fixI <- fixFuzzyPair(max(IXY_L_raw, 0), max(IXY_U_raw, 0))
  IXY_L <- fixI$L; IXY_U <- fixI$U; IXY_C <- (IXY_L + IXY_U) / 2
  
  # entropy per criterion of (row) distributions
  SYX_mu_L <- numeric(M)
  SYX_mu_U <- numeric(M)
  for (i in 1:M) {
    rowL <- probLower[i, ]; sL <- sum(rowL)
    rowU <- probUpper[i, ]; sU <- sum(rowU)
    SYX_mu_L[i] <- if (sL > eps) measure_entropy_alt(rowL / sL) else 0
    SYX_mu_U[i] <- if (sU > eps) measure_entropy_alt(rowU / sU) else 0
  }
  fixSYXmu <- fixFuzzyPair(SYX_mu_L, SYX_mu_U)
  SYX_mu_L <- fixSYXmu$L
  SYX_mu_U <- fixSYXmu$U
  SYX_mu_C <- (SYX_mu_L + SYX_mu_U) / 2
  
  # ----------------------------
  # Indices (all clamped to [0,1])
  # ----------------------------
  CES_L <- clamp01(max(SY_L - SYX_U, 0))
  CES_U <- clamp01(max(SY_U - SYX_L, 0))
  fixCES <- fixFuzzyPair(CES_L, CES_U)
  CES_L <- fixCES$L; CES_U <- fixCES$U; CES_C <- (CES_L + CES_U) / 2
  
  NMI_L <- if (SY_L > 0) clamp01(IXY_L / SY_L) else 0
  NMI_U <- if (SY_U > 0) clamp01(IXY_U / SY_U) else 0
  fixNMI <- fixFuzzyPair(NMI_L, NMI_U)
  NMI_L <- fixNMI$L; NMI_U <- fixNMI$U; NMI_C <- (NMI_L + NMI_U) / 2
  
  NlogN <- log2(N)
  ADI_L <- clamp01(1 - (SY_L / NlogN))
  ADI_U <- clamp01(1 - (SY_U / NlogN))
  fixADI <- fixFuzzyPair(ADI_L, ADI_U)
  ADI_L <- fixADI$L; ADI_U <- fixADI$U; ADI_C <- (ADI_L + ADI_U) / 2
  
  CSF_L <- if (SY_L > 0) clamp01(1 - (SYX_L / SY_L)) else 0
  CSF_U <- if (SY_U > 0) clamp01(1 - (SYX_U / SY_U)) else 0
  fixCSF <- fixFuzzyPair(CSF_L, CSF_U)
  CSF_L <- fixCSF$L; CSF_U <- fixCSF$U; CSF_C <- (CSF_L + CSF_U) / 2
  
  # NMGI: entropy-based weights across indices (base-2 consistent)
  idxL <- c(NMI_L, CES_L, CSF_L, ADI_L)
  idxU <- c(NMI_U, CES_U, CSF_U, ADI_U)
  
  p_j_L <- idxL / (sum(idxL) + eps)
  p_j_U <- idxU / (sum(idxU) + eps)
  
  k <- 1 / log2(4)  # base-2 consistency
  # element-wise "entropy contributions" (kept consistent with your current implementation)
  E_j_L <- -k * p_j_U * log2safe(p_j_U)
  E_j_U <- -k * p_j_L * log2safe(p_j_L)
  E_j_L[!is.finite(E_j_L)] <- 0
  E_j_U[!is.finite(E_j_U)] <- 0
  
  d_j_L <- clamp01(1 - E_j_U)
  d_j_U <- clamp01(1 - E_j_L)
  w_j_L <- d_j_L / (sum(d_j_L) + eps)
  w_j_U <- d_j_U / (sum(d_j_U) + eps)
  
  NMGI_L_raw <- clamp01(sum(w_j_L * idxL))
  NMGI_U_raw <- clamp01(sum(w_j_U * idxU))
  fixNMGI <- fixFuzzyPair(NMGI_L_raw, NMGI_U_raw)
  NMGI_L <- fixNMGI$L; NMGI_U <- fixNMGI$U; NMGI_C <- (NMGI_L + NMGI_U) / 2
  
  # Integrated Criteria Importance (ICI)
  ICI_L <- xIntL * SYX_mu_L
  ICI_U <- xIntU * SYX_mu_U
  fixICI <- fixFuzzyPair(ICI_L, ICI_U)
  ICI_L <- fixICI$L; ICI_U <- fixICI$U
  ICI_Def <- (ICI_L + ICI_U) / 2
  
  # Diagnostics (internal consistency checks)
  diag <- data.frame(
    Check = c(
      "sum_xObjL", "sum_xObjU",
      "sum_xIntL", "sum_xIntU",
      "sum_altLower", "sum_altUpper",
      "min_probLower", "min_probUpper",
      "max_probLower", "max_probUpper",
      "I_leq_min_SX_SY", "SXY_identity"
    ),
    Value = c(
      round3(sum(xObjL)), round3(sum(xObjU)),
      round3(sum(xIntL)), round3(sum(xIntU)),
      round3(sum(altScoreLower)), round3(sum(altScoreUpper)),
      round3(min(probLower)), round3(min(probUpper)),
      round3(max(probLower)), round3(max(probUpper)),
      as.character(IXY_C <= min(SX_C, SY_C) + 1e-9),
      as.character(abs(SXY_C - (SX_C + SY_C - IXY_C)) < 1e-6)
    ),
    stringsAsFactors = FALSE
  )
  
  list(
    alpha = alpha,
    M = M, N = N,
    
    altScoreLower = round3(altScoreLower),
    altScoreUpper = round3(altScoreUpper),
    altDef        = round3(altMid),
    
    wLower = round3(wL),
    wUpper = round3(wU),
    wDefSBJ = round3(wDefSBJ),
    
    xObjLower = round3(xObjL),
    xObjUpper = round3(xObjU),
    xObjDef   = round3(xObjMid),
    
    xIntLower = round3(xIntL),
    xIntUpper = round3(xIntU),
    xIntDef   = round3(xIntMid),
    
    hLower = round3(hLower),
    hUpper = round3(hUpper),
    dLower = round3(dLower),
    dUpper = round3(dUpper),
    
    SXY_L = round3(SXY_L), SXY_U = round3(SXY_U), SXY_C = round3(SXY_C),
    SX_L  = round3(SX_L),  SX_U  = round3(SX_U),  SX_C  = round3(SX_C),
    SY_L  = round3(SY_L),  SY_U  = round3(SY_U),  SY_C  = round3(SY_C),
    IXY_L = round3(IXY_L), IXY_U = round3(IXY_U), IXY_C = round3(IXY_C),
    SYX_L = round3(SYX_L), SYX_U = round3(SYX_U), SYX_C = round3(SYX_C),
    
    CES_L = round3(CES_L), CES_U = round3(CES_U), CES_C = round3(CES_C),
    CSF_L = round3(CSF_L), CSF_U = round3(CSF_U), CSF_C = round3(CSF_C),
    ADI_L = round3(ADI_L), ADI_U = round3(ADI_U), ADI_C = round3(ADI_C),
    NMI_L = round3(NMI_L), NMI_U = round3(NMI_U), NMI_C = round3(NMI_C),
    NMGI_L = round3(NMGI_L), NMGI_U = round3(NMGI_U), NMGI_C = round3(NMGI_C),
    
    ICI_L   = round3(ICI_L),
    ICI_U   = round3(ICI_U),
    ICI_Def = round3(ICI_Def),
    
    diagnostics = diag
  )
}

# ---------------------------------------
# Insights (per α-cut and aggregated)
# ---------------------------------------
build_entropy_insights <- function(r) {
  # Robust, scale-aware interpretation using normalized entropy ratios
  M <- r$M; N <- r$N
  Hxy_max <- log2(M * N)
  Hx_max  <- log2(M)
  Hy_max  <- log2(N)
  
  sxy_rel <- if (Hxy_max > 0) r$SXY_C / Hxy_max else NA
  sx_rel  <- if (Hx_max  > 0) r$SX_C  / Hx_max  else NA
  sy_rel  <- if (Hy_max  > 0) r$SY_C  / Hy_max  else NA
  
  insights <- c()
  
  if (!is.na(sxy_rel) && sxy_rel >= 0.85) {
    insights <- c(insights, "High joint entropy S(X,Y) suggests rich informational diversity across criteria and alternatives.")
  } else if (!is.na(sxy_rel) && sxy_rel <= 0.55) {
    insights <- c(insights, "Lower joint entropy S(X,Y) indicates a concentrated information structure with fewer dominant patterns.")
  } else {
    insights <- c(insights, "Intermediate joint entropy S(X,Y) indicates moderate informational diversity across the decision system.")
  }
  
  if (!is.na(sx_rel) && sx_rel >= 0.75) {
    insights <- c(insights, "Criteria entropy S(X) points to a relatively balanced importance distribution, supporting robustness.")
  } else {
    insights <- c(insights, "Lower S(X) indicates that a subset of criteria dominates the weighting structure.")
  }
  
  if (!is.na(sy_rel) && sy_rel >= 0.85) {
    insights <- c(insights, "Alternatives entropy S(Y) is close to maximum, indicating weak discrimination (near-uniform alternative scores).")
  } else if (!is.na(sy_rel) && sy_rel <= 0.65) {
    insights <- c(insights, "Lower S(Y) indicates stronger discrimination among alternatives and clearer ranking separation.")
  } else {
    insights <- c(insights, "Moderate S(Y) indicates partial discrimination among alternatives.")
  }
  
  if (r$IXY_C >= 0.25) {
    insights <- c(insights, "Mutual information I(X;Y) is substantive: criteria explain a meaningful share of alternative variability.")
  } else {
    insights <- c(insights, "Mutual information I(X;Y) is low: criteria explain only a small fraction of alternative variability (decision mapping is weak).")
  }
  
  if (r$SYX_C <= 0.75) {
    insights <- c(insights, "Conditional entropy S(Y|X) is moderate-to-low, so residual uncertainty after conditioning on criteria is limited.")
  } else {
    insights <- c(insights, "Higher S(Y|X) implies notable residual uncertainty in alternative ranking even after incorporating criteria.")
  }
  
  # indices
  if (r$NMI_C >= 0.70) {
    insights <- c(insights, "NMI close to 1 implies a very strong criteria→ranking mapping.")
  } else if (r$NMI_C <= 0.20) {
    insights <- c(insights, "Low NMI indicates weak explanatory power of the criteria set (rank relations are not information-supported).")
  } else {
    insights <- c(insights, "Moderate NMI indicates partial explanatory power.")
  }
  
  if (r$ADI_C >= 0.70) {
    insights <- c(insights, "High ADI indicates clear separation among alternatives.")
  } else if (r$ADI_C <= 0.20) {
    insights <- c(insights, "Low ADI indicates limited distinction among alternatives (scores overlap substantially).")
  } else {
    insights <- c(insights, "Moderate ADI indicates partial separation among alternatives.")
  }
  
  if (r$CES_C >= 0.70) {
    insights <- c(insights, "High CES indicates highly effective criteria in reducing ranking uncertainty.")
  } else if (r$CES_C <= 0.25) {
    insights <- c(insights, "Low CES suggests limited effectiveness of the criteria set; refinement or augmentation may be warranted.")
  } else {
    insights <- c(insights, "Moderate CES suggests acceptable but improvable criteria effectiveness.")
  }
  
  if (r$CSF_C >= 0.75) {
    insights <- c(insights, "High CSF suggests a stable configuration (lower sensitivity to local perturbations).")
  } else if (r$CSF_C <= 0.30) {
    insights <- c(insights, "Low CSF points to a fragile configuration: small perturbations may alter ranking.")
  } else {
    insights <- c(insights, "Moderate CSF suggests intermediate stability.")
  }
  
  if (r$NMGI_C >= 0.80) {
    insights <- c(insights, "High NMGI indicates a coherent, information-efficient decision system.")
  } else if (r$NMGI_C <= 0.30) {
    insights <- c(insights, "Low NMGI indicates limited net informational growth; improvements in criteria selection/data quality are recommended.")
  } else {
    insights <- c(insights, "Moderate NMGI indicates functional but improvable decision coherence.")
  }
  
  insights
}

aggregate_alpha_insights <- function(batch) {
  # batch: named list of results keyed by alpha label
  alphas <- sapply(batch, function(x) x$alpha)
  idx_df <- data.frame(
    alpha = alphas,
    NMI = sapply(batch, function(x) x$NMI_C),
    ADI = sapply(batch, function(x) x$ADI_C),
    CES = sapply(batch, function(x) x$CES_C),
    CSF = sapply(batch, function(x) x$CSF_C),
    NMGI = sapply(batch, function(x) x$NMGI_C)
  )
  
  alt_best <- sapply(batch, function(x) which.max(x$altDef))
  best_tab <- sort(table(alt_best), decreasing = TRUE)
  
  overall <- c()
  overall <- c(overall, sprintf("Computed %d α-cut subscenarios: [%s].",
                                length(alphas),
                                paste0(format(round(alphas, 2), nsmall = 2), collapse = ", ")))
  
  if (length(best_tab) > 0) {
    top_alt <- as.integer(names(best_tab)[1])
    freq <- as.integer(best_tab[1])
    overall <- c(overall, sprintf("Most frequently best alternative: Y%d (wins %d/%d α-cuts).",
                                  top_alt, freq, length(alphas)))
  }
  
  # index stability summary
  for (nm in c("NMI","ADI","CES","CSF","NMGI")) {
    v <- idx_df[[nm]]
    overall <- c(overall,
                 sprintf("%s across α: mean=%.3f, min=%.3f, max=%.3f.",
                         nm, mean(v), min(v), max(v)))
  }
  
  overall
}

# ---------------------------------------
# Labels parsing (optional Labels sheet)
# ---------------------------------------
parse_labels_sheet <- function(path, sheet = "Labels") {
  out <- list(
    alt_ids = NULL, alt_names = NULL,
    crit_ids = NULL, crit_names = NULL,
    crit_types_from_labels = NULL
  )
  if (!(sheet %in% excel_sheets(path))) return(out)
  
  df <- read_excel(path, sheet = sheet, col_names = FALSE)
  df <- as.data.frame(df, stringsAsFactors = FALSE)
  df[is.na(df)] <- ""
  
  # Find blocks by keywords
  row_alts <- which(toupper(trimws(df[[1]])) == "ALTERNATIVES")
  row_crits <- which(toupper(trimws(df[[1]])) == "CRITERIA")
  
  if (length(row_alts) == 1) {
    start <- row_alts + 2
    # until blank row or until "CRITERIA"
    end <- nrow(df)
    if (length(row_crits) == 1) end <- row_crits - 1
    block <- df[start:end, 1:2, drop = FALSE]
    block <- block[trimws(block[[1]]) != "", , drop = FALSE]
    if (nrow(block) > 0) {
      out$alt_ids <- as.character(block[[1]])
      out$alt_names <- as.character(block[[2]])
    }
  }
  
  if (length(row_crits) == 1) {
    start <- row_crits + 2
    end <- nrow(df)
    block <- df[start:end, 1:3, drop = FALSE]
    block <- block[trimws(block[[1]]) != "", , drop = FALSE]
    # stop at first all-empty line after the block
    # (safe: keep only rows with a non-empty CritID)
    if (nrow(block) > 0) {
      out$crit_ids <- as.character(block[[1]])
      out$crit_names <- as.character(block[[2]])
      out$crit_types_from_labels <- toupper(as.character(block[[3]]))
    }
  }
  
  out
}

# ---------------------------------------
# UI
# ---------------------------------------
ui <- fluidPage(
  theme = shinytheme("slate"),
  tags$head(
    tags$style(HTML("
      table.dataTable tbody td { color: #FFFFFF !important; }
      table.dataTable thead th { color: #FFFF00 !important; }
      .dataTables_wrapper .dataTables_length label,
      .dataTables_wrapper .dataTables_filter label,
      .dataTables_wrapper .dataTables_info { color: #FFFF00 !important; }
      .dataTables_wrapper .dataTables_paginate .paginate_button {
        color: #FFFFFF !important; background-color: #222222 !important;
      }
      .dataTables_wrapper .dataTables_length select,
      .dataTables_wrapper .dataTables_filter input {
        background-color: #FFFFFF !important; color: #000000 !important;
      }
      .shiny-text-output, .shiny-html-output,
      h4, label, .control-label { color: #FFFF00 !important; }
      .btn { color: #FFFFFF !important; }
    "))
  ),
  
  navbarPage(
    title = "FES-MADM II",
    id = "tabs",
    
    tabPanel("1. Instructions",
             fluidRow(
               column(12,
                      h3("How to Use"),
                      tags$ul(
                        tags$li("Step 1: Import the Excel file and verify all sheets."),
                        tags$li("Step 2: Run the main α-cut computation and inspect results."),
                        tags$li("Step 3: Run the batch α-cut computations for robustness/stability."),
                        tags$li("Step 4: Export results (main α or all α-cuts) to Excel.")
                      ),
                      hr(),
                      h4("Notation Table (FES-MADM II)"),
                      DTOutput("tableNotation")
               )
             )
    ),
    
    tabPanel("2. Data Import",
             sidebarLayout(
               sidebarPanel(
                 fileInput("fileExcel", "Upload Excel:", accept = c(".xls", ".xlsx")),
                 uiOutput("sheetSelectors"),
                 checkboxInput("useLabels", "Use Labels sheet (if present)", value = TRUE),
                 actionButton("loadData", "Load Data", class = "btn-primary"),
                 hr(),
                 helpText("Tip: If your Excel uses the default sheet names, the selectors will auto-fill.")
               ),
               mainPanel(
                 h4("Preview: Data (Center ξμν)"), DTOutput("tableCenter"),
                 h4("Preview: Data (Δμν)"), DTOutput("tableDelta"),
                 h4("Preview: SBJ Weights (Center xμ^SBJ)"), DTOutput("tableWCenter"),
                 h4("Preview: SBJ Weights (Δxμ^SBJ)"), DTOutput("tableWDelta"),
                 h4("Preview: Criteria Type"), DTOutput("tableType"),
                 h4("Preview: α-cuts (if provided)"), DTOutput("tableAlpha"),
                 h4("Preview: Labels (if provided)"), DTOutput("tableLabels")
               )
             )
    ),
    
    tabPanel("3. Main α-cut Results",
             sidebarLayout(
               sidebarPanel(
                 sliderInput("alphaSliderMain", "Select α-cut (main run):",
                             min = 0, max = 1, value = 0.5, step = 0.1),
                 actionButton("computeFESMADM", "Compute (Main α)", class = "btn-success"),
                 hr(),
                 helpText("Main α-cut results are used by the single-α plots and the main export.")
               ),
               mainPanel(
                 tabsetPanel(
                   tabPanel("Alternatives", DTOutput("tableAltResults")),
                   tabPanel("Criteria Weights",
                            h4("Subjective Weights x̃μ^SBJ,a"), DTOutput("tableSBJ"),
                            hr(),
                            h4("Objective Weights x̃μ^OBJ,a"), DTOutput("tableOBJ"),
                            hr(),
                            h4("Integrated Weights x̃μ^INT,a"), DTOutput("tableINT")
                   ),
                   tabPanel("Integrated Criteria Importance (ICI)", DTOutput("tableICI")),
                   tabPanel("Entropy Measures & Indices",
                            fluidRow(
                              column(6, h4("Entropy Measures"), DTOutput("tableEntropy")),
                              column(6, h4("Entropy-based Indices"), DTOutput("tableIndices"))
                            ),
                            hr(),
                            h4("Diagnostics (consistency checks)"),
                            DTOutput("tableDiagnostics"),
                            hr(),
                            h4("Detailed Assessment (Main α-cut)"),
                            uiOutput("uiDetailedAssessmentMain")
                   )
                 )
               )
             )
    ),
    
    tabPanel("4. All α-cut Scenarios",
             sidebarLayout(
               sidebarPanel(
                 checkboxInput("useAlphaSheet", "Use AlphaCuts sheet (if present)", value = TRUE),
                 textInput("alphaManual", "Or provide α-cuts (comma-separated):", value = "0,0.1,0.2,0.3,0.4,0.5,0.6,0.7,0.8,0.9,1.0"),
                 actionButton("computeBatch", "Compute (All α-cuts)", class = "btn-warning"),
                 downloadButton("downloadBatch", "Download (All α-cuts)", class = "btn-primary"),
                 hr(),
                 helpText("Batch run computes results for every α-cut, then provides stability plots & summary insights.")
               ),
               mainPanel(
                 tabsetPanel(
                   tabPanel("Summary", DTOutput("tableBatchSummary"), hr(), uiOutput("uiBatchInsights")),
                   tabPanel("Alternative Scores vs α", plotOutput("plotAltVsAlpha"), DTOutput("tableAltVsAlpha")),
                   tabPanel("Integrated Weights vs α", plotOutput("plotWIntVsAlpha"), DTOutput("tableWIntVsAlpha")),
                   tabPanel("ICI vs α", plotOutput("plotICIVsAlpha"), DTOutput("tableICIVsAlpha")),
                   tabPanel("Entropy Measures vs α", plotOutput("plotEntropyVsAlpha"), DTOutput("tableEntropyVsAlpha")),
                   tabPanel("Indices vs α", plotOutput("plotIdxVsAlpha"), DTOutput("tableIdxVsAlpha"))
                 )
               )
             )
    ),
    
    tabPanel("5. Plots (Main α)",
             sidebarLayout(
               sidebarPanel(
                 helpText("Crisp plots for the last main α-cut computation."),
                 hr(),
                 checkboxInput("showSBJ", "Show Subjective (SBJ)", value = TRUE),
                 checkboxInput("showOBJ", "Show Objective (OBJ)", value = TRUE),
                 checkboxInput("showINT", "Show Integrated (INT)", value = TRUE)
               ),
               mainPanel(
                 tabsetPanel(
                   tabPanel("Alternative Scores", plotOutput("plotAltScores")),
                   tabPanel("Criteria Weights",
                            h4("a) Crisp xμ^SBJ"), plotOutput("plotWeightsSBJ"),
                            hr(),
                            h4("b) Crisp xμ^OBJ"), plotOutput("plotWeightsOBJ"),
                            hr(),
                            h4("c) Crisp xμ^INT"), plotOutput("plotWeightsINT"),
                            hr(),
                            h4("d) SBJ vs OBJ vs INT"), plotOutput("plotWeightsCompare")
                   ),
                   tabPanel("ICI", plotOutput("plotICIbar")),
                   tabPanel("Entropy Measures", plotOutput("plotEntropyMeasures")),
                   tabPanel("Entropy Indices", plotOutput("plotEntropyIndices")),
                   # --- Added (as requested): mapping table for full names
                   tabPanel("Nomenclature Map",
                            fluidRow(
                              column(6, h4("Alternatives (Yk ↔ Name)"), DTOutput("tableAltMap")),
                              column(6, h4("Criteria (Xμ ↔ Name/Type)"), DTOutput("tableCritMap"))
                            )
                   )
                 )
               )
             )
    ),
    
    tabPanel("6. Sensitivity Analysis",
             sidebarLayout(
               sidebarPanel(
                 sliderInput("alphaSliderSens", "Select α for Sensitivity:",
                             min = 0, max = 1, value = 0.5, step = 0.1),
                 numericInput("rowDelta", "Row # for Δμν:", 1, min = 1),
                 numericInput("colDelta", "Col # for Δμν:", 1, min = 1),
                 numericInput("valDelta", "New Δμν:", 0, step = 0.1),
                 actionButton("applyDelta", "Apply Δμν"),
                 hr(),
                 numericInput("rowWDelta", "Criterion # for Δxμ^SBJ:", 1, min = 1),
                 numericInput("valWDelta", "New Δxμ^SBJ:", 0, step = 0.1),
                 actionButton("applyWDelta", "Apply Δxμ^SBJ"),
                 hr(),
                 numericInput("critTypeIndex", "Criterion # toggle B/C:", 1, min = 1),
                 actionButton("toggleBC", "Toggle B/C"),
                 hr(),
                 actionButton("runSensitivity", "Re-run (Sensitivity)", class = "btn-success")
               ),
               mainPanel(
                 h4("Sensitivity Results"),
                 DTOutput("tableSensitivityAlt"),
                 plotOutput("plotSensitivityAll")
               )
             )
    ),
    
    tabPanel("7. Download Results",
             fluidRow(
               column(12,
                      h3("Export to Excel"),
                      tags$ul(
                        tags$li("Main α-cut export: one run (the last main computation)."),
                        tags$li("All α-cuts export: includes batch tables, stability summaries and per-α detailed assessment.")
                      ),
                      downloadButton("downloadMain", "Download MAIN results (Excel)", class = "btn-primary"),
                      tags$span(" "),
                      downloadButton("downloadAllAlpha", "Download ALL α-cuts (Excel)", class = "btn-warning"),
                      hr(),
                      uiOutput("summaryResults")
               )
             )
    )
  )
)

# ---------------------------------------
# Server
# ---------------------------------------
server <- function(input, output, session) {
  
  dataCenter <- reactiveVal(NULL)
  dataDelta  <- reactiveVal(NULL)
  wCenter    <- reactiveVal(NULL)
  wDelta     <- reactiveVal(NULL)
  critType   <- reactiveVal(NULL)
  alphaCuts  <- reactiveVal(NULL)
  
  labels_raw <- reactiveVal(NULL)
  alt_names  <- reactiveVal(NULL)
  crit_names <- reactiveVal(NULL)
  
  alpha_main_used <- reactiveVal(0.5)
  alpha_sens_used <- reactiveVal(0.5)
  
  # Notation table
  output$tableNotation <- renderDT({
    df <- data.frame(
      Symbol = c("ξμν","Δμν","xμ^SBJ","Δxμ^SBJ","xμ^OBJ","xμ^INT",
                 "S(X,Y)","S(X)","S(Y)","I(X;Y)","S(Y|X)",
                 "NMI","ADI","CES","CSF","NMGI","ICI"),
      Meaning = c(
        "Central performance of criterion μ on alternative ν",
        "Fuzzy deviation (uncertainty) of that performance",
        "Subjective weight (central value)",
        "Fuzzy deviation of the subjective weight",
        "Objective weight from fuzzy diversification degree",
        "Integrated weight combining subjective & objective views",
        "Joint entropy of criteria and alternatives (bits)",
        "Entropy of the criteria distribution (bits)",
        "Entropy of the alternatives distribution (bits)",
        "Mutual information between criteria and alternatives (bits)",
        "Conditional entropy of alternatives given criteria (bits)",
        "Normalized Mutual Information",
        "Alternatives Distinction Index",
        "Criteria Effectiveness Score",
        "Conditional Stability Factor",
        "Net Mutual Growth Index",
        "Integrated Criteria Importance"
      )
    )
    datatable(df, options = list(pageLength = 20), rownames = FALSE)
  })
  
  # Sheet selectors (by name, robust to ordering)
  output$sheetSelectors <- renderUI({
    req(input$fileExcel)
    sheets <- excel_sheets(input$fileExcel$datapath)
    
    # helper for default selection
    pick_default <- function(target) {
      if (target %in% sheets) target else sheets[1]
    }
    
    tagList(
      selectInput("sheetCenter", "Sheet: Data (Center ξμν):", choices = sheets, selected = pick_default("Data_Center_xi")),
      selectInput("sheetDelta",  "Sheet: Data (Δμν):", choices = sheets, selected = pick_default("Data_Delta")),
      selectInput("sheetWCenter","Sheet: SBJ Weights (Center xμ^SBJ):", choices = sheets, selected = pick_default("SBJ_WCenter")),
      selectInput("sheetWDelta", "Sheet: SBJ Weights (Δxμ^SBJ):", choices = sheets, selected = pick_default("SBJ_WDelta")),
      selectInput("sheetType",   "Sheet: Crit.Type (B/C):", choices = sheets, selected = pick_default("Criterion_Type")),
      selectInput("sheetAlpha",  "Sheet: α-cuts (optional):", choices = c("<<none>>", sheets), selected = if ("AlphaCuts" %in% sheets) "AlphaCuts" else "<<none>>"),
      selectInput("sheetLabels", "Sheet: Labels (optional):", choices = c("<<none>>", sheets), selected = if ("Labels" %in% sheets) "Labels" else "<<none>>")
    )
  })
  
  # Load data
  observeEvent(input$loadData, {
    req(input$fileExcel)
    
    path <- input$fileExcel$datapath
    
    read_sheet <- function(sheet) {
      as.data.frame(read_excel(path, sheet = sheet, col_names = FALSE))
    }
    
    # core sheets
    dc <- tryCatch(read_sheet(input$sheetCenter), error = function(e) NULL)
    dd <- tryCatch(read_sheet(input$sheetDelta),  error = function(e) NULL)
    wc <- tryCatch(read_sheet(input$sheetWCenter), error = function(e) NULL)
    wd <- tryCatch(read_sheet(input$sheetWDelta),  error = function(e) NULL)
    ct <- tryCatch(read_sheet(input$sheetType),    error = function(e) NULL)
    
    if (is.null(dc) || is.null(dd) || is.null(wc) || is.null(wd) || is.null(ct)) {
      showNotification("Error reading one or more required sheets. Check sheet selections.", type = "error")
      return()
    }
    
    dataCenter(as_num_df(dc))
    dataDelta(as_num_df(dd))
    wCenter(as_num_df(wc))
    wDelta(as_num_df(wd))
    critType(as_chr_df(ct))
    
    # alpha sheet
    if (!is.null(input$sheetAlpha) && input$sheetAlpha != "<<none>>") {
      ac <- tryCatch(read_sheet(input$sheetAlpha), error = function(e) NULL)
      if (!is.null(ac)) alphaCuts(as_num_df(ac))
    } else {
      alphaCuts(NULL)
    }
    
    # labels
    labels_raw(NULL); alt_names(NULL); crit_names(NULL)
    if (input$useLabels && !is.null(input$sheetLabels) && input$sheetLabels != "<<none>>") {
      labs <- tryCatch(parse_labels_sheet(path, sheet = input$sheetLabels), error = function(e) NULL)
      labels_raw(labs)
      
      # only accept if sizes match loaded matrices
      if (!is.null(labs) && !is.null(labs$alt_names) && ncol(dataCenter()) == length(labs$alt_names)) {
        alt_names(labs$alt_names)
      }
      if (!is.null(labs) && !is.null(labs$crit_names) && nrow(dataCenter()) == length(labs$crit_names)) {
        crit_names(labs$crit_names)
      }
      # optionally override CT if the labels provide types and size matches
      if (!is.null(labs) && !is.null(labs$crit_types_from_labels) && nrow(critType()) == length(labs$crit_types_from_labels)) {
        ct2 <- data.frame(labs$crit_types_from_labels, stringsAsFactors = FALSE)
        names(ct2) <- 1
        critType(as_chr_df(ct2))
      }
    }
    
    showNotification("Data loaded successfully!", type = "message")
  })
  
  # Previews
  output$tableCenter  <- renderDT({ req(dataCenter()); datatable(dataCenter(), options = list(pageLength=8)) })
  output$tableDelta   <- renderDT({ req(dataDelta());  datatable(dataDelta(),  options = list(pageLength=8)) })
  output$tableWCenter <- renderDT({ req(wCenter());    datatable(wCenter(),    options = list(pageLength=8)) })
  output$tableWDelta  <- renderDT({ req(wDelta());     datatable(wDelta(),     options = list(pageLength=8)) })
  output$tableType    <- renderDT({ req(critType());   datatable(critType(),   options = list(pageLength=8)) })
  output$tableAlpha   <- renderDT({ req(alphaCuts());  datatable(alphaCuts(),  options = list(pageLength=8)) })
  
  output$tableLabels <- renderDT({
    req(labels_raw())
    labs <- labels_raw()
    if (is.null(labs) || (is.null(labs$alt_names) && is.null(labs$crit_names))) {
      return(datatable(data.frame(Message = "No readable labels found."), options = list(dom='t')))
    }
    df1 <- data.frame(ID = ifelse(is.null(labs$alt_ids), NA, labs$alt_ids),
                      Alternative = ifelse(is.null(labs$alt_names), NA, labs$alt_names))
    df2 <- data.frame(ID = ifelse(is.null(labs$crit_ids), NA, labs$crit_ids),
                      Criterion = ifelse(is.null(labs$crit_names), NA, labs$crit_names),
                      Type = ifelse(is.null(labs$crit_types_from_labels), NA, labs$crit_types_from_labels))
    datatable(list(Alternatives = df1, Criteria = df2)[[1]], options = list(pageLength=8))
  })
  
  # Main α computation
  fesResMain <- eventReactive(input$computeFESMADM, {
    req(dataCenter(), dataDelta(), wCenter(), wDelta(), critType())
    DC <- as.matrix(dataCenter())
    DD <- as.matrix(dataDelta())
    WC <- as.matrix(wCenter())
    WD <- as.matrix(wDelta())
    CT <- as.matrix(critType())
    alpha_main_used(input$alphaSliderMain)
    fesMadmFull(alpha_main_used(), DC, DD, WC, WD, CT)
  })
  
  # Helpers for labels (tables keep full names if provided)
  alt_label_vec <- reactive({
    r <- fesResMain()
    if (!is.null(alt_names())) return(alt_names())
    if (!is.null(r)) return(paste0("Y", seq_len(r$N)))
    NULL
  })
  crit_label_vec <- reactive({
    r <- fesResMain()
    if (!is.null(crit_names())) return(crit_names())
    if (!is.null(r)) return(paste0("X", seq_len(r$M)))
    NULL
  })
  
  # --- Plot-safe nomenclature (ALWAYS IDs, never full names)
  alt_id_vec <- reactive({
    r <- fesResMain()
    if (is.null(r)) return(NULL)
    paste0("Y", seq_len(r$N))
  })
  
  crit_id_vec <- reactive({
    r <- fesResMain()
    if (is.null(r)) return(NULL)
    paste0("X", seq_len(r$M))
  })
  
  # --- Nomenclature mapping tables (IDs ↔ Full names)
  output$tableAltMap <- renderDT({
    r <- fesResMain(); req(r)
    ids <- paste0("Y", seq_len(r$N))
    nm  <- if (!is.null(alt_names())) alt_names() else rep(NA_character_, r$N)
    
    df <- data.frame(
      ID = ids,
      Name = nm,
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    
    datatable(df, options = list(pageLength = 20, scrollX = TRUE), rownames = FALSE)
  })
  
  output$tableCritMap <- renderDT({
    r <- fesResMain(); req(r)
    ids <- paste0("X", seq_len(r$M))
    nm  <- if (!is.null(crit_names())) crit_names() else rep(NA_character_, r$M)
    tp  <- as.character(critType()[,1])
    
    df <- data.frame(
      ID = ids,
      Name = nm,
      Type = tp,
      check.names = FALSE,
      stringsAsFactors = FALSE
    )
    
    datatable(df, options = list(pageLength = 20, scrollX = TRUE), rownames = FALSE)
  })
  
  # Main α tables
  output$tableAltResults <- renderDT({
    r <- fesResMain(); req(r)
    labels <- alt_label_vec()
    df <- data.frame(
      Alternative = labels,
      Lower = r$altScoreLower,
      Upper = r$altScoreUpper,
      Crisp = r$altDef,
      check.names = FALSE
    ) %>% arrange(desc(Crisp))
    datatable(df, options = list(pageLength = 15), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableSBJ <- renderDT({
    r <- fesResMain(); req(r)
    labels <- crit_label_vec()
    df <- data.frame(
      Criterion = labels,
      Lower = r$wLower,
      Upper = r$wUpper,
      Crisp = r$wDefSBJ,
      check.names = FALSE
    )
    datatable(df, options = list(pageLength = 15), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableOBJ <- renderDT({
    r <- fesResMain(); req(r)
    labels <- crit_label_vec()
    df <- data.frame(
      Criterion = labels,
      Lower = r$xObjLower,
      Upper = r$xObjUpper,
      Crisp = r$xObjDef,
      check.names = FALSE
    )
    datatable(df, options = list(pageLength = 15), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableINT <- renderDT({
    r <- fesResMain(); req(r)
    labels <- crit_label_vec()
    df <- data.frame(
      Criterion = labels,
      Lower = r$xIntLower,
      Upper = r$xIntUpper,
      Crisp = r$xIntDef,
      check.names = FALSE
    )
    datatable(df, options = list(pageLength = 15), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableICI <- renderDT({
    r <- fesResMain(); req(r)
    labels <- crit_label_vec()
    df <- data.frame(
      Criterion = labels,
      Lower = r$ICI_L,
      Upper = r$ICI_U,
      Crisp = r$ICI_Def,
      check.names = FALSE
    )
    datatable(df, options = list(pageLength = 15), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableEntropy <- renderDT({
    r <- fesResMain(); req(r)
    infoProp_L <- ifelse(r$SY_L > 0, clamp01(1 - r$SYX_L / r$SY_L), NA)
    infoProp_U <- ifelse(r$SY_U > 0, clamp01(1 - r$SYX_U / r$SY_U), NA)
    infoProp_C <- ifelse(!is.na(infoProp_L) & !is.na(infoProp_U),
                         (infoProp_L + infoProp_U) / 2, NA)
    
    df <- data.frame(
      Measure = c("S(X,Y) [bits]", "S(X) [bits]", "S(Y) [bits]", "I(X;Y) [bits]", "S(Y|X) [bits]", "I(Y|X)"),
      Lower   = c(r$SXY_L, r$SX_L, r$SY_L, r$IXY_L, r$SYX_L, infoProp_L),
      Upper   = c(r$SXY_U, r$SX_U, r$SY_U, r$IXY_U, r$SYX_U, infoProp_U),
      Crisp   = c(r$SXY_C, r$SX_C, r$SY_C, r$IXY_C, r$SYX_C, infoProp_C),
      check.names = FALSE
    )
    datatable(df, options = list(pageLength = 10), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableIndices <- renderDT({
    r <- fesResMain(); req(r)
    df <- data.frame(
      Index = c("NMI", "ADI", "CES", "CSF", "NMGI"),
      Lower = c(r$NMI_L, r$ADI_L, r$CES_L, r$CSF_L, r$NMGI_L),
      Upper = c(r$NMI_U, r$ADI_U, r$CES_U, r$CSF_U, r$NMGI_U),
      Crisp = c(r$NMI_C, r$ADI_C, r$CES_C, r$CSF_C, r$NMGI_C),
      check.names = FALSE
    )
    datatable(df, options = list(pageLength = 10), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$tableDiagnostics <- renderDT({
    r <- fesResMain(); req(r)
    datatable(r$diagnostics, options = list(pageLength = 20), rownames = FALSE)
  })
  
  output$uiDetailedAssessmentMain <- renderUI({
    r <- fesResMain(); req(r)
    items <- build_entropy_insights(r)
    tags$div(
      tags$h4(sprintf("Detailed Insights @ α = %.2f", r$alpha)),
      tags$ul(lapply(items, tags$li))
    )
  })
  
  # ----------------------------
  # Main α plots (NOW: IDs only)
  # ----------------------------
  output$plotAltScores <- renderPlot({
    r <- fesResMain(); req(r)
    labels <- alt_id_vec()
    df <- data.frame(Alternative = labels, Score = as.numeric(r$altDef)) %>% arrange(desc(Score))
    ggplot(df, aes(x = reorder(Alternative, -Score), y = Score)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(Score)), vjust = -0.5, size = 3) +
      labs(title = sprintf("Alternative Scores p*ν (Crisp) @ α = %.2f", alpha_main_used()),
           x = "Alternative", y = "Score") +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))
  })
  
  output$plotWeightsSBJ <- renderPlot({
    r <- fesResMain(); req(r)
    labels <- crit_id_vec()
    df <- data.frame(Criterion = labels, Weight = as.numeric(r$wDefSBJ)) %>% arrange(desc(Weight))
    ggplot(df, aes(x = reorder(Criterion, -Weight), y = Weight)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(Weight)), vjust = -0.5, size = 3) +
      labs(title = "Subjective Weights xμ^SBJ (Crisp)",
           x = "Criterion", y = "Weight") +
      theme_minimal()
  })
  
  output$plotWeightsOBJ <- renderPlot({
    r <- fesResMain(); req(r)
    labels <- crit_id_vec()
    df <- data.frame(Criterion = labels, Weight = as.numeric(r$xObjDef)) %>% arrange(desc(Weight))
    ggplot(df, aes(x = reorder(Criterion, -Weight), y = Weight)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(Weight)), vjust = -0.5, size = 3) +
      labs(title = "Objective Weights xμ^OBJ (Crisp)",
           x = "Criterion", y = "Weight") +
      theme_minimal()
  })
  
  output$plotWeightsINT <- renderPlot({
    r <- fesResMain(); req(r)
    labels <- crit_id_vec()
    df <- data.frame(Criterion = labels, Weight = as.numeric(r$xIntDef)) %>% arrange(desc(Weight))
    ggplot(df, aes(x = reorder(Criterion, -Weight), y = Weight)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(Weight)), vjust = -0.5, size = 3) +
      labs(title = "Integrated Weights xμ^INT (Crisp)",
           x = "Criterion", y = "Weight") +
      theme_minimal()
  })
  
  output$plotWeightsCompare <- renderPlot({
    r <- fesResMain(); req(r)
    labels <- crit_id_vec()
    
    SBJ <- as.numeric(r$wDefSBJ); SBJ[is.na(SBJ)] <- 0
    OBJ <- as.numeric(r$xObjDef); OBJ[is.na(OBJ)] <- 0
    INT <- as.numeric(r$xIntDef); INT[is.na(INT)] <- 0
    
    df <- data.frame(Criterion = labels, SBJ = SBJ, OBJ = OBJ, INT = INT)
    df_melt <- melt(df, id.vars = "Criterion", variable.name = "Type", value.name = "Weight")
    
    df_melt_filtered <- df_melt %>%
      filter((Type == "SBJ" & input$showSBJ) |
               (Type == "OBJ" & input$showOBJ) |
               (Type == "INT" & input$showINT))
    
    if (nrow(df_melt_filtered) == 0) {
      ggplot() + annotate("text", x = 0.5, y = 0.5, label = "No Weights Selected", size = 6) + theme_void()
    } else {
      ggplot(df_melt_filtered, aes(x = Criterion, y = Weight, colour = Type, group = Type)) +
        geom_line(size = 1) +
        geom_point(size = 3) +
        geom_text(aes(label = round3(Weight)), vjust = -0.5, size = 3, colour = "black") +
        labs(title = sprintf("xμ^SBJ, xμ^OBJ, xμ^INT (Crisp) @ α = %.2f", alpha_main_used()),
             x = "Criterion", y = "Weight", colour = "Type") +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, vjust = 1, hjust = 1))
    }
  })
  
  output$plotICIbar <- renderPlot({
    r <- fesResMain(); req(r)
    labels <- crit_id_vec()
    df <- data.frame(Criterion = labels, ICI = as.numeric(r$ICI_Def)) %>% arrange(desc(ICI))
    ggplot(df, aes(x = reorder(Criterion, -ICI), y = ICI)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(ICI)), vjust = -0.5, size = 3) +
      labs(title = sprintf("Integrated Criteria Importance Iμ^a (Crisp) @ α = %.2f", alpha_main_used()),
           x = "Criterion", y = "ICI") +
      theme_minimal()
  })
  
  output$plotEntropyMeasures <- renderPlot({
    r <- fesResMain(); req(r)
    infoProp_C <- ifelse(r$SY_C > 0, clamp01(1 - r$SYX_C / r$SY_C), NA)
    df <- data.frame(
      Measure = c("I(X;Y)", "S(X)", "S(X,Y)", "S(Y)", "S(Y|X)"),
      Value   = round3(c(r$IXY_C, r$SX_C, r$SXY_C, r$SY_C, r$SYX_C))
    )
    ggplot(df, aes(x = Measure, y = Value)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = Value), vjust = -0.5, size = 3) +
      labs(title = sprintf("Entropy Measures (Crisp, bits) @ α = %.2f", alpha_main_used()),
           x = "Measure", y = "Bits") +
      theme_minimal()
  })
  
  output$plotEntropyIndices <- renderPlot({
    r <- fesResMain(); req(r)
    df <- data.frame(
      Index = c("NMI", "ADI", "CES", "CSF", "NMGI"),
      Value = round3(c(r$NMI_C, r$ADI_C, r$CES_C, r$CSF_C, r$NMGI_C))
    ) %>% arrange(desc(Value))
    ggplot(df, aes(x = reorder(Index, -Value), y = Value)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = Value), vjust = -0.5, size = 3) +
      labs(title = sprintf("Entropy Indices (Crisp) @ α = %.2f", alpha_main_used()),
           x = "Index", y = "[0,1]") +
      theme_minimal()
  })
  
  # ----------------------------
  # Sensitivity operations (in-place edits)
  # ----------------------------
  observeEvent(input$applyDelta, {
    dd <- dataDelta(); req(dd)
    i <- input$rowDelta; j <- input$colDelta
    if (i <= nrow(dd) && j <= ncol(dd)) {
      dd[i, j] <- as.numeric(input$valDelta)
      dataDelta(dd)
      showNotification(paste("Updated Δ[", i, ",", j, "] =", round3(dd[i, j])), type = "message")
    } else {
      showNotification("Invalid row/column for Δμν!", type = "error")
    }
  })
  
  observeEvent(input$applyWDelta, {
    wd <- wDelta(); req(wd)
    i <- input$rowWDelta
    if (i <= nrow(wd)) {
      wd[i, 1] <- as.numeric(input$valWDelta)
      wDelta(wd)
      showNotification(paste("Updated Δxμ^SBJ for criterion", i, "=", round3(wd[i, 1])),
                       type = "message")
    } else {
      showNotification("Invalid row for Δxμ^SBJ!", type = "error")
    }
  })
  
  observeEvent(input$toggleBC, {
    ct <- critType(); req(ct)
    i <- input$critTypeIndex
    if (i <= nrow(ct)) {
      oldVal <- toupper(as.character(ct[i, 1]))
      newVal <- ifelse(oldVal == "B", "C", "B")
      ct[i, 1] <- newVal
      critType(ct)
      showNotification(paste("Toggled criterion", i, "from", oldVal, "to", newVal),
                       type = "message")
    } else {
      showNotification("Invalid row for toggling B/C!", type = "error")
    }
  })
  
  fesResSens <- eventReactive(input$runSensitivity, {
    req(dataCenter(), dataDelta(), wCenter(), wDelta(), critType())
    DC <- as.matrix(dataCenter())
    DD <- as.matrix(dataDelta())
    WC <- as.matrix(wCenter())
    WD <- as.matrix(wDelta())
    CT <- as.matrix(critType())
    alpha_sens_used(input$alphaSliderSens)
    fesMadmFull(alpha_sens_used(), DC, DD, WC, WD, CT)
  })
  
  output$tableSensitivityAlt <- renderDT({
    r <- fesResSens(); req(r)
    labels <- if (!is.null(alt_names())) alt_names() else paste0("Y", 1:r$N)
    df <- data.frame(
      Alternative = labels,
      Lower = r$altScoreLower,
      Upper = r$altScoreUpper,
      Crisp = r$altDef,
      check.names = FALSE
    ) %>% arrange(desc(Crisp))
    datatable(df, options = list(pageLength = 15), rownames = FALSE) |>
      formatRound(columns = 2:4, digits = 3)
  })
  
  output$plotSensitivityAll <- renderPlot({
    r <- fesResSens(); req(r)
    labelsC <- paste0("X", 1:r$M)
    labelsA <- paste0("Y", 1:r$N)
    
    dfW <- data.frame(Criterion = labelsC, Weight = as.numeric(r$xIntDef)) %>% arrange(desc(Weight))
    p1 <- ggplot(dfW, aes(x = reorder(Criterion, -Weight), y = Weight)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(Weight)), vjust = -0.5, size = 3) +
      labs(title = sprintf("Integrated Weights xμ^INT (Crisp) @ α = %.2f", alpha_sens_used()),
           x = "Criterion", y = "Weight") +
      theme_minimal()
    
    dfA <- data.frame(Alternative = labelsA, Score = as.numeric(r$altDef)) %>% arrange(desc(Score))
    p2 <- ggplot(dfA, aes(x = reorder(Alternative, -Score), y = Score)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = round3(Score)), vjust = -0.5, size = 3) +
      labs(title = sprintf("Alternative Scores p*ν (Crisp) @ α = %.2f", alpha_sens_used()),
           x = "Alternative", y = "Score") +
      theme_minimal() +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust = 1))
    
    dfI <- data.frame(Index = c("NMI","ADI","CES","CSF","NMGI"),
                      Value = round3(c(r$NMI_C, r$ADI_C, r$CES_C, r$CSF_C, r$NMGI_C))) %>%
      arrange(desc(Value))
    p3 <- ggplot(dfI, aes(x = reorder(Index, -Value), y = Value)) +
      geom_bar(stat = "identity") +
      geom_text(aes(label = Value), vjust = -0.5, size = 3) +
      labs(title = sprintf("Entropy Indices (Crisp) @ α = %.2f", alpha_sens_used()),
           x = "Index", y = "[0,1]") +
      theme_minimal()
    
    grid.arrange(p1, p2, p3, ncol = 3)
  })
  
  # ----------------------------
  # Batch α computations
  # ----------------------------
  batchRes <- eventReactive(input$computeBatch, {
    req(dataCenter(), dataDelta(), wCenter(), wDelta(), critType())
    
    # determine alpha vector
    alphas <- NULL
    if (input$useAlphaSheet && !is.null(alphaCuts())) {
      alphas <- safe_alpha_vec(alphaCuts()[,1])
    } else {
      alphas <- safe_alpha_vec(strsplit(input$alphaManual, ",")[[1]])
    }
    
    DC <- as.matrix(dataCenter())
    DD <- as.matrix(dataDelta())
    WC <- as.matrix(wCenter())
    WD <- as.matrix(wDelta())
    CT <- as.matrix(critType())
    
    out <- list()
    for (a in alphas) {
      key <- sprintf("a%02d", as.integer(round(a * 100)))
      out[[key]] <- fesMadmFull(a, DC, DD, WC, WD, CT)
    }
    out
  })
  
  output$tableBatchSummary <- renderDT({
    b <- batchRes(); req(b)
    alphas <- sapply(b, function(x) x$alpha)
    best_i <- sapply(b, function(x) which.max(x$altDef))
    best_score <- sapply(b, function(x) max(x$altDef))
    idx <- data.frame(
      alpha = round3(alphas),
      bestAlternative = best_i,
      bestScore = round3(best_score),
      NMI = sapply(b, function(x) x$NMI_C),
      ADI = sapply(b, function(x) x$ADI_C),
      CES = sapply(b, function(x) x$CES_C),
      CSF = sapply(b, function(x) x$CSF_C),
      NMGI = sapply(b, function(x) x$NMGI_C)
    )
    # map bestAlternative to label if exists
    if (!is.null(alt_names())) {
      idx$bestAlternative <- alt_names()[idx$bestAlternative]
    } else {
      idx$bestAlternative <- paste0("Y", idx$bestAlternative)
    }
    datatable(idx, options = list(pageLength = 20), rownames = FALSE) |>
      formatRound(columns = c("alpha","bestScore","NMI","ADI","CES","CSF","NMGI"), digits = 3)
  })
  
  output$uiBatchInsights <- renderUI({
    b <- batchRes(); req(b)
    lines <- aggregate_alpha_insights(b)
    tags$div(
      tags$h4("Integrated Insights (All α-cuts)"),
      tags$ul(lapply(lines, tags$li))
    )
  })
  
  # Export: full batch (all α-cuts) to Excel
  output$downloadBatch <- downloadHandler(
    filename = function() {
      paste0("FESMADMII_AllAlpha_Results_", format(Sys.time(), "%Y%m%d_%H%M%S"), ".xlsx")
    },
    content = function(file) {
      b <- batchRes(); req(b)
      
      # Prepare summary, alternatives, weights, ICI matrices
      alphas <- as.numeric(sapply(b, function(x) x$alpha))
      r0 <- b[[1]]
      N <- r0$N
      M <- r0$M
      alt_labels <- if (!is.null(alt_names())) alt_names() else paste0("Y", 1:N)
      crit_labels <- if (!is.null(crit_names())) crit_names() else paste0("X", 1:M)
      
      # Summary
      df_summary <- data.frame(
        alpha = round3(alphas),
        TopAlternative = sapply(b, function(x) {
          ii <- which.max(x$altDef); if (length(alt_labels) >= ii) alt_labels[ii] else paste0("Y", ii)
        }),
        TopScore = sapply(b, function(x) max(as.numeric(x$altDef))),
        TopCriterion_ICI = sapply(b, function(x) {
          ii <- which.max(x$ICI_Def); if (length(crit_labels) >= ii) crit_labels[ii] else paste0("X", ii)
        }),
        MaxICI = sapply(b, function(x) max(as.numeric(x$ICI_Def))),
        NMI  = sapply(b, function(x) x$NMI_C),
        ADI  = sapply(b, function(x) x$ADI_C),
        CES  = sapply(b, function(x) x$CES_C),
        CSF  = sapply(b, function(x) x$CSF_C),
        NMGI = sapply(b, function(x) x$NMGI_C),
        check.names = FALSE
      )
      
      # Matrices
      mat_alt <- sapply(b, function(x) as.numeric(x$altDef))
      df_alt <- data.frame(Alternative = alt_labels, mat_alt, check.names = FALSE)
      colnames(df_alt)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
      
      mat_wint <- sapply(b, function(x) as.numeric(x$xIntDef))
      df_wint <- data.frame(Criterion = crit_labels, mat_wint, check.names = FALSE)
      colnames(df_wint)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
      
      mat_ici <- sapply(b, function(x) as.numeric(x$ICI_Def))
      df_ici <- data.frame(Criterion = crit_labels, mat_ici, check.names = FALSE)
      colnames(df_ici)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
      
      # Entropy measures
      infoProp <- sapply(b, function(x) {
        if (is.finite(x$SY_C) && x$SY_C > 0) clamp01(1 - (x$SYX_C / x$SY_C)) else NA_real_
      })
      df_entropy <- data.frame(
        alpha = round3(alphas),
        SXY = sapply(b, function(x) x$SXY_C),
        SX  = sapply(b, function(x) x$SX_C),
        SY  = sapply(b, function(x) x$SY_C),
        IXY = sapply(b, function(x) x$IXY_C),
        SYX = sapply(b, function(x) x$SYX_C),
        IYgX = infoProp,
        check.names = FALSE
      )
      
      # Indices
      df_indices <- data.frame(
        alpha = round3(alphas),
        NMI = sapply(b, function(x) x$NMI_C),
        ADI = sapply(b, function(x) x$ADI_C),
        CES = sapply(b, function(x) x$CES_C),
        CSF = sapply(b, function(x) x$CSF_C),
        NMGI = sapply(b, function(x) x$NMGI_C),
        check.names = FALSE
      )
      
      writexl::write_xlsx(list(
        Batch_Summary = df_summary,
        AltScores_Crisp = df_alt,
        WeightsINT_Crisp = df_wint,
        ICI_Crisp = df_ici,
        EntropyMeasures_Crisp = df_entropy,
        Indices_Crisp = df_indices
      ), path = file)
    }
  )
  
  # Alt scores vs alpha (table)
  output$tableAltVsAlpha <- renderDT({
    b <- batchRes(); req(b)
    r0 <- b[[1]]
    N <- r0$N
    labels <- if (!is.null(alt_names())) alt_names() else paste0("Y", 1:N)
    
    alphas <- sapply(b, function(x) x$alpha)
    mat <- sapply(b, function(x) as.numeric(x$altDef))
    df <- data.frame(Alternative = labels, mat, check.names = FALSE)
    colnames(df)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
    datatable(df, options = list(pageLength = 15, scrollX = TRUE), rownames = FALSE) |>
      formatRound(columns = 2:ncol(df), digits = 3)
  })
  
  output$plotAltVsAlpha <- renderPlot({
    b <- batchRes(); req(b)
    r0 <- b[[1]]
    N <- r0$N
    labels <- paste0("Y", 1:N)  # IDs only
    alphas <- as.numeric(sapply(b, function(x) x$alpha))
    
    # scores: N x K (columns correspond to alphas)
    scoreMat <- do.call(cbind, lapply(b, function(x) as.numeric(x$altDef)))
    K <- length(alphas)
    
    # Long format for score trajectories
    df_scores <- data.frame(
      alpha = rep(alphas, each = N),
      Alternative = rep(labels, times = K),
      Score = as.vector(scoreMat)
    )
    
    p_scores <- ggplot(df_scores, aes(x = alpha, y = Score, group = Alternative, colour = Alternative)) +
      geom_line(size = 1) + geom_point(size = 2) +
      labs(title = "Crisp Alternative Scores across α-cuts",
           x = "α", y = "Crisp score p*ν") +
      theme_minimal()
    
    # Rank trajectories (1 = best)
    rankMat <- apply(scoreMat, 2, function(v) rank(-v, ties.method = "min"))
    df_rank <- data.frame(
      alpha = rep(alphas, each = N),
      Alternative = rep(labels, times = K),
      Rank = as.vector(rankMat)
    )
    
    p_rank <- ggplot(df_rank, aes(x = alpha, y = Rank, group = Alternative, colour = Alternative)) +
      geom_line(size = 1) + geom_point(size = 2) +
      scale_y_reverse(breaks = sort(unique(df_rank$Rank))) +
      labs(title = "Alternative Rank Stability across α-cuts (1 = best)",
           x = "α", y = "Rank") +
      theme_minimal()
    
    gridExtra::grid.arrange(p_scores, p_rank, ncol = 2)
  })
  
  # Integrated weights vs alpha (table + plot)
  output$tableWIntVsAlpha <- renderDT({
    b <- batchRes(); req(b)
    r0 <- b[[1]]
    M <- r0$M
    labels <- if (!is.null(crit_names())) crit_names() else paste0("X", 1:M)
    alphas <- sapply(b, function(x) x$alpha)
    
    mat <- sapply(b, function(x) as.numeric(x$xIntDef))
    df <- data.frame(Criterion = labels, mat, check.names = FALSE)
    colnames(df)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
    datatable(df, options = list(pageLength = 20, scrollX = TRUE), rownames = FALSE) |>
      formatRound(columns = 2:ncol(df), digits = 3)
  })
  
  output$plotWIntVsAlpha <- renderPlot({
    b <- batchRes(); req(b)
    r0 <- b[[1]]
    M <- r0$M
    labels <- paste0("X", 1:M)  # IDs only
    alphas <- as.numeric(sapply(b, function(x) x$alpha))
    
    # weights: M x K
    wMat <- do.call(cbind, lapply(b, function(x) as.numeric(x$xIntDef)))
    K <- length(alphas)
    
    df_w <- data.frame(
      alpha = rep(alphas, each = M),
      Criterion = rep(labels, times = K),
      Weight = as.vector(wMat)
    )
    
    p_lines <- ggplot(df_w, aes(x = alpha, y = Weight, group = Criterion, colour = Criterion)) +
      geom_line(size = 1) + geom_point(size = 2) +
      labs(title = "Crisp Integrated Weights across α-cuts",
           x = "α", y = "xμ^INT (crisp)") +
      theme_minimal()
    
    # Heatmap for quick sensitivity inspection
    df_hm <- df_w
    p_hm <- ggplot(df_hm, aes(x = alpha, y = Criterion, fill = Weight)) +
      geom_tile() +
      labs(title = "Integrated Weight Heatmap across α-cuts",
           x = "α", y = "Criterion") +
      theme_minimal() +
      theme(legend.position = "right")
    
    gridExtra::grid.arrange(p_lines, p_hm, ncol = 2)
  })
  
  # ICI vs alpha (table + plot)
  output$tableICIVsAlpha <- renderDT({
    b <- batchRes(); req(b)
    r0 <- b[[1]]
    M <- r0$M
    labels <- if (!is.null(crit_names())) crit_names() else paste0("X", 1:M)
    alphas <- sapply(b, function(x) x$alpha)
    
    mat <- sapply(b, function(x) as.numeric(x$ICI_Def))
    df <- data.frame(Criterion = labels, mat, check.names = FALSE)
    colnames(df)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
    datatable(df, options = list(pageLength = 20, scrollX = TRUE), rownames = FALSE) |>
      formatRound(columns = 2:ncol(df), digits = 3)
  })
  
  output$plotICIVsAlpha <- renderPlot({
    b <- batchRes(); req(b)
    r0 <- b[[1]]
    M <- r0$M
    labels <- paste0("X", 1:M)  # IDs only
    alphas <- as.numeric(sapply(b, function(x) x$alpha))
    K <- length(alphas)
    
    iciMat <- do.call(cbind, lapply(b, function(x) as.numeric(x$ICI_Def)))
    
    df_ici <- data.frame(
      alpha = rep(alphas, each = M),
      Criterion = rep(labels, times = K),
      ICI = as.vector(iciMat)
    )
    
    p_lines <- ggplot(df_ici, aes(x = alpha, y = ICI, group = Criterion, colour = Criterion)) +
      geom_line(size = 1) + geom_point(size = 2) +
      labs(title = "Integrated Criteria Importance (ICI) across α-cuts",
           x = "α", y = "ICI (crisp)") +
      theme_minimal()
    
    p_hm <- ggplot(df_ici, aes(x = alpha, y = Criterion, fill = ICI)) +
      geom_tile() +
      labs(title = "ICI Heatmap across α-cuts",
           x = "α", y = "Criterion") +
      theme_minimal()
    
    gridExtra::grid.arrange(p_lines, p_hm, ncol = 2)
  })
  
  # Entropy measures vs alpha (table + plot)
  output$tableEntropyVsAlpha <- renderDT({
    b <- batchRes(); req(b)
    alphas <- as.numeric(sapply(b, function(x) x$alpha))
    
    infoProp <- sapply(b, function(x) {
      if (is.finite(x$SY_C) && x$SY_C > 0) clamp01(1 - (x$SYX_C / x$SY_C)) else NA_real_
    })
    
    df <- data.frame(
      alpha = round3(alphas),
      SXY = sapply(b, function(x) x$SXY_C),
      SX  = sapply(b, function(x) x$SX_C),
      SY  = sapply(b, function(x) x$SY_C),
      IXY = sapply(b, function(x) x$IXY_C),
      SYX = sapply(b, function(x) x$SYX_C),
      IYgX = infoProp,
      check.names = FALSE
    )
    
    datatable(df, options = list(pageLength = 20, scrollX = TRUE), rownames = FALSE) |>
      formatRound(columns = 1:ncol(df), digits = 3)
  })
  
  output$plotEntropyVsAlpha <- renderPlot({
    b <- batchRes(); req(b)
    alphas <- as.numeric(sapply(b, function(x) x$alpha))
    
    infoProp <- sapply(b, function(x) {
      if (is.finite(x$SY_C) && x$SY_C > 0) clamp01(1 - (x$SYX_C / x$SY_C)) else NA_real_
    })
    
    df <- data.frame(
      alpha = alphas,
      `S(X,Y)` = sapply(b, function(x) x$SXY_C),
      `S(X)`   = sapply(b, function(x) x$SX_C),
      `S(Y)`   = sapply(b, function(x) x$SY_C),
      `I(X;Y)` = sapply(b, function(x) x$IXY_C),
      `S(Y|X)` = sapply(b, function(x) x$SYX_C),
      `I(Y|X)` = infoProp,
      check.names = FALSE
    )
    
    df_m <- reshape2::melt(df, id.vars = "alpha", variable.name = "Measure", value.name = "Value")
    
    p1 <- ggplot(df_m, aes(x = alpha, y = Value, group = Measure, colour = Measure)) +
      geom_line(size = 1) + geom_point(size = 2) +
      labs(title = "Entropy Measures across α-cuts (Crisp)",
           x = "α", y = "Value (bits / proportion)") +
      theme_minimal()
    
    # Optional zoom for mutual-information-related signals
    df_m2 <- df_m %>% filter(Measure %in% c("I(X;Y)","I(Y|X)"))
    p2 <- ggplot(df_m2, aes(x = alpha, y = Value, group = Measure, colour = Measure)) +
      geom_line(size = 1) + geom_point(size = 2) +
      labs(title = "Mutual Information & Explained-Proportion across α-cuts",
           x = "α", y = "Value") +
      theme_minimal()
    
    gridExtra::grid.arrange(p1, p2, ncol = 2)
  })
  
  # Indices vs alpha
  output$tableIdxVsAlpha <- renderDT({
    b <- batchRes(); req(b)
    alphas <- sapply(b, function(x) x$alpha)
    df <- data.frame(
      alpha = round3(alphas),
      NMI = sapply(b, function(x) x$NMI_C),
      ADI = sapply(b, function(x) x$ADI_C),
      CES = sapply(b, function(x) x$CES_C),
      CSF = sapply(b, function(x) x$CSF_C),
      NMGI = sapply(b, function(x) x$NMGI_C)
    )
    datatable(df, options = list(pageLength = 20, scrollX = TRUE), rownames = FALSE) |>
      formatRound(columns = 1:ncol(df), digits = 3)
  })
  
  output$plotIdxVsAlpha <- renderPlot({
    b <- batchRes(); req(b)
    alphas <- sapply(b, function(x) x$alpha)
    df <- data.frame(
      alpha = alphas,
      NMI = sapply(b, function(x) x$NMI_C),
      ADI = sapply(b, function(x) x$ADI_C),
      CES = sapply(b, function(x) x$CES_C),
      CSF = sapply(b, function(x) x$CSF_C),
      NMGI = sapply(b, function(x) x$NMGI_C)
    )
    df_m <- melt(df, id.vars = "alpha", variable.name = "Index", value.name = "Value")
    ggplot(df_m, aes(x = alpha, y = Value, group = Index, colour = Index)) +
      geom_line(size = 1) + geom_point(size = 2) +
      labs(title = "Entropy-based Indices across α-cuts",
           x = "α", y = "Index value") +
      theme_minimal()
  })
  
  # ----------------------------
  # Downloads
  # ----------------------------
  output$downloadMain <- downloadHandler(
    filename = function() sprintf("FESMADM2_MAIN_alpha_%s.xlsx", format(Sys.Date(), "%Y%m%d")),
    content = function(file) {
      r <- fesResMain(); req(r)
      M <- r$M; N <- r$N
      labelsA <- if (!is.null(alt_names())) alt_names() else paste0("Y", 1:N)
      labelsC <- if (!is.null(crit_names())) crit_names() else paste0("X", 1:M)
      
      infoProp_L <- ifelse(r$SY_L > 0, clamp01(1 - r$SYX_L / r$SY_L), NA)
      infoProp_U <- ifelse(r$SY_U > 0, clamp01(1 - r$SYX_U / r$SY_U), NA)
      infoProp_C <- ifelse(!is.na(infoProp_L) & !is.na(infoProp_U), (infoProp_L + infoProp_U)/2, NA)
      
      df_alts <- data.frame(Alternative = labelsA, Lower = r$altScoreLower, Upper = r$altScoreUpper, Crisp = r$altDef)
      df_sbj  <- data.frame(Criterion = labelsC, Lower = r$wLower, Upper = r$wUpper, Crisp = r$wDefSBJ)
      df_obj  <- data.frame(Criterion = labelsC, Lower = r$xObjLower, Upper = r$xObjUpper, Crisp = r$xObjDef)
      df_int  <- data.frame(Criterion = labelsC, Lower = r$xIntLower, Upper = r$xIntUpper, Crisp = r$xIntDef)
      df_ici  <- data.frame(Criterion = labelsC, Lower = r$ICI_L, Upper = r$ICI_U, Crisp = r$ICI_Def)
      
      df_entropy <- data.frame(
        Measure = c("S(X,Y) [bits]", "S(X) [bits]", "S(Y) [bits]", "I(X;Y) [bits]", "S(Y|X) [bits]", "I(Y|X)"),
        Lower = c(r$SXY_L, r$SX_L, r$SY_L, r$IXY_L, r$SYX_L, infoProp_L),
        Upper = c(r$SXY_U, r$SX_U, r$SY_U, r$IXY_U, r$SYX_U, infoProp_U),
        Crisp = c(r$SXY_C, r$SX_C, r$SY_C, r$IXY_C, r$SYX_C, infoProp_C)
      )
      
      df_indices <- data.frame(
        Index = c("NMI","ADI","CES","CSF","NMGI"),
        Lower = c(r$NMI_L, r$ADI_L, r$CES_L, r$CSF_L, r$NMGI_L),
        Upper = c(r$NMI_U, r$ADI_U, r$CES_U, r$CSF_U, r$NMGI_U),
        Crisp = c(r$NMI_C, r$ADI_C, r$CES_C, r$CSF_C, r$NMGI_C)
      )
      
      # Detailed assessment sheet
      assess <- build_entropy_insights(r)
      df_assess <- data.frame(
        Section = c(sprintf("Detailed Insights @ α = %.2f", r$alpha), rep("", length(assess)-1)),
        Bullet = assess,
        stringsAsFactors = FALSE
      )
      
      write_xlsx(list(
        Alternatives = df_alts,
        SubjectiveWeights = df_sbj,
        ObjectiveWeights = df_obj,
        IntegratedWeights = df_int,
        ICI = df_ici,
        Entropy = df_entropy,
        Indices = df_indices,
        Diagnostics = r$diagnostics,
        Detailed_Assessment = df_assess
      ), path = file)
    }
  )
  
  output$downloadAllAlpha <- downloadHandler(
    filename = function() sprintf("FESMADM2_ALL_ALPHA_%s.xlsx", format(Sys.Date(), "%Y%m%d")),
    content = function(file) {
      # If batch is not computed yet, compute it on-the-fly using current settings
      b <- batchRes()
      if (is.null(b)) {
        # emulate button click computation
        req(dataCenter(), dataDelta(), wCenter(), wDelta(), critType())
        alphas <- NULL
        if (input$useAlphaSheet && !is.null(alphaCuts())) {
          alphas <- safe_alpha_vec(alphaCuts()[,1])
        } else {
          alphas <- safe_alpha_vec(strsplit(input$alphaManual, ",")[[1]])
        }
        DC <- as.matrix(dataCenter())
        DD <- as.matrix(dataDelta())
        WC <- as.matrix(wCenter())
        WD <- as.matrix(wDelta())
        CT <- as.matrix(critType())
        b <- list()
        for (a in alphas) {
          key <- sprintf("a%02d", as.integer(round(a * 100)))
          b[[key]] <- fesMadmFull(a, DC, DD, WC, WD, CT)
        }
      }
      
      r0 <- b[[1]]
      M <- r0$M; N <- r0$N
      labelsA <- if (!is.null(alt_names())) alt_names() else paste0("Y", 1:N)
      labelsC <- if (!is.null(crit_names())) crit_names() else paste0("X", 1:M)
      alphas <- sapply(b, function(x) x$alpha)
      
      # Summary sheets
      best_i <- sapply(b, function(x) which.max(x$altDef))
      best_score <- sapply(b, function(x) max(x$altDef))
      best_lbl <- if (!is.null(alt_names())) alt_names()[best_i] else paste0("Y", best_i)
      
      df_summary <- data.frame(
        alpha = round3(alphas),
        bestAlternative = best_lbl,
        bestScore = round3(best_score),
        NMI = sapply(b, function(x) x$NMI_C),
        ADI = sapply(b, function(x) x$ADI_C),
        CES = sapply(b, function(x) x$CES_C),
        CSF = sapply(b, function(x) x$CSF_C),
        NMGI = sapply(b, function(x) x$NMGI_C)
      )
      
      mat_alt <- sapply(b, function(x) as.numeric(x$altDef))
      df_alt_all <- data.frame(Alternative = labelsA, mat_alt, check.names = FALSE)
      colnames(df_alt_all)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
      
      mat_wint <- sapply(b, function(x) as.numeric(x$xIntDef))
      df_wint_all <- data.frame(Criterion = labelsC, mat_wint, check.names = FALSE)
      colnames(df_wint_all)[-1] <- paste0("α=", format(round(alphas, 2), nsmall = 2))
      
      df_idx_all <- data.frame(
        alpha = round3(alphas),
        NMI = sapply(b, function(x) x$NMI_C),
        ADI = sapply(b, function(x) x$ADI_C),
        CES = sapply(b, function(x) x$CES_C),
        CSF = sapply(b, function(x) x$CSF_C),
        NMGI = sapply(b, function(x) x$NMGI_C)
      )
      
      # Overall insights
      overall_lines <- aggregate_alpha_insights(b)
      df_overall <- data.frame(Line = overall_lines, stringsAsFactors = FALSE)
      
      # Per-α detailed assessments (stacked)
      assess_rows <- list()
      k <- 1
      for (key in names(b)) {
        rr <- b[[key]]
        bullets <- build_entropy_insights(rr)
        assess_rows[[k]] <- data.frame(
          alpha = sprintf("%.2f", rr$alpha),
          bullet = bullets,
          stringsAsFactors = FALSE
        )
        k <- k + 1
      }
      df_assess_all <- do.call(rbind, assess_rows)
      
      # Per-α sheets (compact)
      out_list <- list(
        Summary_AllAlpha = df_summary,
        AltScores_AllAlpha = df_alt_all,
        IntegratedWeights_AllAlpha = df_wint_all,
        Indices_AllAlpha = df_idx_all,
        Overall_Assessment = df_overall,
        Detailed_Assessment_AllAlpha = df_assess_all
      )
      
      # Add per-α detailed results in separate sheets (safe sheet names <=31 chars)
      for (i in seq_along(b)) {
        rr <- b[[i]]
        a <- rr$alpha
        tag <- sprintf("a%02d", as.integer(round(a*100)))
        
        df_alts <- data.frame(Alternative = labelsA, Lower = rr$altScoreLower, Upper = rr$altScoreUpper, Crisp = rr$altDef)
        df_int  <- data.frame(Criterion = labelsC, Lower = rr$xIntLower, Upper = rr$xIntUpper, Crisp = rr$xIntDef)
        df_ici  <- data.frame(Criterion = labelsC, Lower = rr$ICI_L, Upper = rr$ICI_U, Crisp = rr$ICI_Def)
        df_entropy <- data.frame(
          Measure = c("S(X,Y)","S(X)","S(Y)","I(X;Y)","S(Y|X)"),
          Crisp = c(rr$SXY_C, rr$SX_C, rr$SY_C, rr$IXY_C, rr$SYX_C)
        )
        df_indices <- data.frame(
          Index = c("NMI","ADI","CES","CSF","NMGI"),
          Crisp = c(rr$NMI_C, rr$ADI_C, rr$CES_C, rr$CSF_C, rr$NMGI_C)
        )
        out_list[[paste0(tag, "_Alt")]] <- df_alts
        out_list[[paste0(tag, "_WInt")]] <- df_int
        out_list[[paste0(tag, "_ICI")]] <- df_ici
        out_list[[paste0(tag, "_Entropy")]] <- df_entropy
        out_list[[paste0(tag, "_Idx")]] <- df_indices
        out_list[[paste0(tag, "_Diag")]] <- rr$diagnostics
      }
      
      # Ensure sheet names are unique and within Excel limits
      nm <- names(out_list)
      nm2 <- substr(nm, 1, 31)
      # make unique if truncation collisions occur
      nm2u <- make.unique(nm2, sep = "_")
      names(out_list) <- nm2u
      
      write_xlsx(out_list, path = file)
    }
  )
  
  # Overall summary (download tab)
  output$summaryResults <- renderUI({
    r <- fesResMain()
    if (is.null(r)) {
      return(tags$div(tags$em("Run a main α-cut computation to populate the summary.")))
    }
    
    labelsA <- alt_label_vec()
    labelsC <- crit_label_vec()
    
    alpha_val <- alpha_main_used()
    best_alt <- labelsA[which.max(r$altDef)]
    best_score <- max(r$altDef)
    best_crit <- labelsC[which.max(r$ICI_Def)]
    best_ici  <- max(r$ICI_Def)
    
    infoProp <- ifelse(r$SY_C > 0, clamp01(1 - (r$SYX_C / r$SY_C)), NA)
    
    tags$div(
      tags$h3("Summary (Main α-cut)"),
      tags$table(
        style = "width:100%; border-collapse:collapse; border: 1px solid black;",
        tags$tr(
          tags$th(colspan = 2,
                  style = "text-align:left; padding:8px; border:1px solid black;
                           background-color:orange; color:#000; font-weight:bold; font-size:18px;",
                  "Metric")
        ),
        tags$tr(tags$td("Selected α-cut"), tags$td(sprintf("%.2f", alpha_val))),
        tags$tr(tags$td("Optimal Alternative"), tags$td(paste0(best_alt, " (Crisp p*: ", sprintf("%.3f", best_score), ")"))),
        tags$tr(tags$td("Most Significant Criterion (ICI)"), tags$td(paste0(best_crit, " (ICI: ", sprintf("%.3f", best_ici), ")"))),
        tags$tr(
          tags$th(colspan = 2,
                  style = "text-align:left; padding:8px; border:1px solid black;
                           background-color:orange; color:#000; font-weight:bold; font-size:18px;",
                  "Key Entropy Measures (Crisp)")
        ),
        tags$tr(tags$td("S(X,Y)"), tags$td(sprintf("%.3f", r$SXY_C))),
        tags$tr(tags$td("S(X)"), tags$td(sprintf("%.3f", r$SX_C))),
        tags$tr(tags$td("S(Y)"), tags$td(sprintf("%.3f", r$SY_C))),
        tags$tr(tags$td("I(X;Y)"), tags$td(sprintf("%.3f", r$IXY_C))),
        tags$tr(tags$td("S(Y|X)"), tags$td(sprintf("%.3f", r$SYX_C))),
        tags$tr(tags$td("I(Y|X)"), tags$td(ifelse(is.na(infoProp), "NA", sprintf("%.3f", infoProp)))),
        tags$tr(
          tags$th(colspan = 2,
                  style = "text-align:left; padding:8px; border:1px solid black;
                           background-color:orange; color:#000; font-weight:bold; font-size:18px;",
                  "Key Indices (Crisp)")
        ),
        tags$tr(tags$td("NMI"), tags$td(sprintf("%.3f", r$NMI_C))),
        tags$tr(tags$td("ADI"), tags$td(sprintf("%.3f", r$ADI_C))),
        tags$tr(tags$td("CES"), tags$td(sprintf("%.3f", r$CES_C))),
        tags$tr(tags$td("CSF"), tags$td(sprintf("%.3f", r$CSF_C))),
        tags$tr(tags$td("NMGI"), tags$td(sprintf("%.3f", r$NMGI_C)))
      )
    )
  })
}

shinyApp(ui = ui, server = server)
