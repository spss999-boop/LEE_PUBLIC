#!/usr/bin/env Rscript
# DEVA_5_NA.R — Minimal / Set-Consistency Guaranteed (many-to-many fixed)
# ASCII-only comments

suppressWarnings(suppressMessages({
  library(dplyr)
  library(tidyr)
  library(openxlsx)
}))

# -------- 0) Params & Paths ---------------------------------------------------
SPAN <- "25100308"               # or set externally then source()
DAY  <- substr(SPAN, 1, 6)
INFILE  <- sprintf("C:/LEE/out/raw/EHIS_RA3_%s.csv", SPAN)
OUTDIR  <- sprintf("C:/LEE/out/SCORE/%s", DAY)
OUT_TRN2 <- sprintf("%s/DEVA_%s_TRN2.csv", OUTDIR, SPAN)
OUT_XLSX <- sprintf("%s/DEVA_%s.xlsx", OUTDIR, SPAN)
if (!dir.exists(OUTDIR)) dir.create(OUTDIR, recursive = TRUE, showWarnings = FALSE)

num <- function(x) suppressWarnings(as.numeric(x))
first_non_na <- function(v){ i <- which(!is.na(v))[1]; if (length(i)) v[i] else NA }

read_csv_win <- function(path){
  x <- tryCatch(read.csv(path, fileEncoding = "CP949", header=TRUE, check.names = FALSE,
                         na.strings = c("","NA"), stringsAsFactors = FALSE), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, fileEncoding = "UTF-8", header=TRUE, check.names = FALSE,
                         na.strings = c("","NA"), stringsAsFactors = FALSE), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, header=TRUE, check.names = FALSE,
                         na.strings = c("","NA"), stringsAsFactors = FALSE), error=function(e) NULL)
  if (is.null(x)) stop("Failed to read CSV: ", path)
  x
}

# add missing columns as NA to avoid dplyr summarise errors
ensure_cols <- function(df, cols){
  miss <- setdiff(cols, names(df))
  for (k in miss) df[[k]] <- NA
  df
}

# -------- 1) Load -------------------------------------------------------------
DAT <- read_csv_win(INFILE)

# normalize header
names(DAT) <- toupper(trimws(names(DAT)))

# fallback: HA_22ND <= HA22_G
if (!("HA_22ND" %in% names(DAT))) {
  if ("HA22_G" %in% names(DAT)) {
    DAT$HA_22ND <- DAT$HA22_G
  } else {
    stop("Missing required column: HA_22ND (or HA22_G)")
  }
}

# fallback: UPPRO_FR <= UPPRO
if (!("UPPRO_FR" %in% names(DAT)) && "UPPRO" %in% names(DAT)) {
  DAT$UPPRO_FR <- DAT$UPPRO
}

# required keys (snapshot + id)
req <- c("REFER1","HNO_FR","HNAME_FR","RANO_FR","CLASS","UPPRO_FR","HA_22ND")
miss <- setdiff(req, names(DAT)); if (length(miss)) stop("Missing keys: ", paste(miss, collapse=", "))

# coerce important fields
DAT$REFER1  <- num(DAT$REFER1)
DAT$ROUND   <- as.integer(DAT$REFER1)   # 0=snapshot, >=1 past (larger=older)
DAT$UPPRO_FR <- num(DAT$UPPRO_FR)       # percent (higher is better)

# columns used later (create if absent to avoid summarise errors)
need_fr <- c(
  "JO_BEFO_FR","JORANK_FR","JO_TREN_FR","JO_DUE_FR",
  "MA_BEFO_FR","MARANK_FR","MA_DUE_FR",
  "OW_BEFO_FR","OWRANK_FR",
  "HO_BEFO_FR","HORANK_FR","HO_DUE_FR",
  "ENHO_FR","FANPRO_FR","MANAGER_FR","OWNER_FR","JOCKEY_FR",
  "JO_GAP_RT","JO_GAP_RK","JO_GAP_DU",
  "MA_GAP_RT","MA_GAP_RK","MA_GAP_DU",
  "OW_GAP_RT","OW_GAP_RK",
  "HO_GAP_RT","HO_GAP_RK","HO_GAP_DU"
)
need_hi <- c(
  "POWER","FANPRO_HI",
  "JO_BEFO_HI","JORANK_HI","JO_TREN_HI","JO_DUE_HI",
  "MA_BEFO_HI","MARANK_HI","MA_DUE_HI",
  "OW_BEFO_HI","OWRANK_HI",
  "HO_BEFO_HI","HORANK_HI","HO_DUE_HI"
)
DAT <- ensure_cols(DAT, c(need_fr, need_hi))

# -------- 2) HIST / SNAP0 (carry HA_22ND explicitly) -------------------------
HIST <- DAT %>% filter(ROUND >= 1) %>% arrange(HNO_FR, ROUND)

SNAP0 <- DAT %>%
  arrange(HNO_FR, ROUND) %>%
  group_by(HNO_FR) %>%
  summarise(
    HNAME_FR = first_non_na(HNAME_FR),
    RANO_FR  = first_non_na(RANO_FR),
    CLASS    = first_non_na(CLASS),
    HA_22ND  = first_non_na(HA_22ND),
    JO_BEFO_FR = first_non_na(JO_BEFO_FR), JORANK_FR = first_non_na(JORANK_FR),
    JO_TREN_FR = first_non_na(JO_TREN_FR), JO_DUE_FR = first_non_na(JO_DUE_FR),
    MA_BEFO_FR = first_non_na(MA_BEFO_FR), MARANK_FR = first_non_na(MARANK_FR), MA_DUE_FR = first_non_na(MA_DUE_FR),
    OW_BEFO_FR = first_non_na(OW_BEFO_FR), OWRANK_FR = first_non_na(OWRANK_FR),
    HO_BEFO_FR = first_non_na(HO_BEFO_FR), HORANK_FR = first_non_na(HORANK_FR), HO_DUE_FR = first_non_na(HO_DUE_FR),
    ENHO_FR = first_non_na(ENHO_FR), FANPRO_FR = first_non_na(FANPRO_FR),
    MANAGER_FR = first_non_na(MANAGER_FR), OWNER_FR = first_non_na(OWNER_FR), JOCKEY_FR = first_non_na(JOCKEY_FR),
    JO_GAP_RT = first_non_na(JO_GAP_RT), JO_GAP_RK = first_non_na(JO_GAP_RK), JO_GAP_DU = first_non_na(JO_GAP_DU),
    MA_GAP_RT = first_non_na(MA_GAP_RT), MA_GAP_RK = first_non_na(MA_GAP_RK), MA_GAP_DU = first_non_na(MA_GAP_DU),
    OW_GAP_RT = first_non_na(OW_GAP_RT), OW_GAP_RK = first_non_na(OW_GAP_RK),
    HO_GAP_RT = first_non_na(HO_GAP_RT), HO_GAP_RK = first_non_na(HO_GAP_RK), HO_GAP_DU = first_non_na(HO_GAP_DU),
    UPPRO_FR = first_non_na(UPPRO_FR),
    .groups = "drop"
  )
SNAP0$ROUND <- 0L

# -------- 3) Feature build (use HIST_R unique per (HNO_FR,ROUND)) -------------
HIST_R <- HIST %>% group_by(HNO_FR, ROUND) %>%
  summarise(
    RANO_FR = first_non_na(RANO_FR),
    CLASS   = first_non_na(CLASS),
    HA_22ND = first_non_na(HA_22ND),            # carry HA_22ND into TRAIN
    UPPRO_FR = first_non_na(UPPRO_FR),          # target per (HNO_FR,ROUND)
    POWERY  = suppressWarnings(median(num(POWER), na.rm=TRUE)),
    FANPRO_HI = first_non_na(FANPRO_HI),
    JO_BEFO_HI = suppressWarnings(median(num(JO_BEFO_HI), na.rm=TRUE)),
    JORANK_HI  = suppressWarnings(median(num(JORANK_HI),  na.rm=TRUE)),
    JO_TREN_HI = first_non_na(JO_TREN_HI),
    JO_DUE_HI  = suppressWarnings(median(num(JO_DUE_HI),  na.rm=TRUE)),
    MA_BEFO_HI = suppressWarnings(median(num(MA_BEFO_HI), na.rm=TRUE)),
    MARANK_HI  = suppressWarnings(median(num(MARANK_HI),  na.rm=TRUE)),
    MA_DUE_HI  = suppressWarnings(median(num(MA_DUE_HI), na.rm=TRUE)),
    OW_BEFO_HI = suppressWarnings(median(num(OW_BEFO_HI), na.rm=TRUE)),
    OWRANK_HI  = suppressWarnings(median(num(OWRANK_HI),  na.rm=TRUE)),
    HO_BEFO_HI = suppressWarnings(median(num(HO_BEFO_HI), na.rm=TRUE)),
    HORANK_HI  = suppressWarnings(median(num(HORANK_HI), na.rm=TRUE)),
    HO_DUE_HI  = suppressWarnings(median(num(HO_DUE_HI), na.rm=TRUE)),
    .groups = "drop"
  ) %>% arrange(HNO_FR, ROUND)

make_windows <- function(df){
  df <- df[order(df$ROUND),]
  n <- nrow(df)
  pw <- num(df$POWERY); fp <- num(df$FANPRO_HI)
  PW_MI_TR <- PW_MID_TR <- PW_ALL_TR <- FP_MI_TR <- FP_MID_TR <- FP_ALL_TR <- rep(NA_real_, n)
  for (i in seq_len(n)){
    if (i < n){
      j2 <- min(i+5L, n)
      PW_MI_TR[i]  <- pw[i+1L]
      PW_MID_TR[i] <- mean(pw[(i+1L):j2], na.rm=TRUE)
      PW_ALL_TR[i] <- mean(pw[(i+1L):n],  na.rm=TRUE)
      FP_MI_TR[i]  <- fp[i+1L]
      FP_MID_TR[i] <- mean(fp[(i+1L):j2], na.rm=TRUE)
      FP_ALL_TR[i] <- mean(fp[(i+1L):n],  na.rm=TRUE)
    }
  }
  cbind(df, PW_MI_TR, PW_MID_TR, PW_ALL_TR, FP_MI_TR, FP_MID_TR, FP_ALL_TR)
}
TRAIN_FULL <- HIST_R %>% group_by(HNO_FR) %>% group_modify(~make_windows(.x)) %>% ungroup()
TRN <- TRAIN_FULL   # one row per (HNO_FR, ROUND)

# r=0 feature map (per HNO_FR)
PW_FP <- HIST_R %>% arrange(HNO_FR, ROUND) %>% group_by(HNO_FR) %>%
  summarise(
    PW_MI_TR  = dplyr::first(POWERY[ROUND==1]),
    PW_MID_TR = mean(num(POWERY)[ROUND >= 1 & ROUND <= 5], na.rm = TRUE),
    PW_ALL_TR = mean(num(POWERY)[ROUND >= 1], na.rm = TRUE),
    FP_MI_TR  = dplyr::first(FANPRO_HI[ROUND==1]),
    FP_MID_TR = mean(num(FANPRO_HI)[ROUND >= 1 & ROUND <= 5], na.rm = TRUE),
    FP_ALL_TR = mean(num(FANPRO_HI)[ROUND >= 1], na.rm = TRUE),
    .groups = "drop"
  )
SNAP0 <- SNAP0 %>% left_join(PW_FP, by="HNO_FR")
SNAP0$FP_MI_TR <- num(SNAP0$FANPRO_FR)

# -------- 4) Assemble matrices (no many-to-many joins) -----------------------
FEATS <- unique(c(
  "PW_MI_TR","PW_MID_TR","PW_ALL_TR","FP_MI_TR","FP_MID_TR","FP_ALL_TR",
  "JO_RT_TR","JO_RK_TR","JO_DU_TR","MA_RT_TR","MA_RK_TR","MA_DU_TR",
  "OW_RT_TR","OW_RK_TR","HO_RT_TR","HO_RK_TR","HO_DU_TR",
  "JO_BEFO_HI","JORANK_HI","JO_TREN_HI","JO_DUE_HI",
  "MA_BEFO_HI","MARANK_HI","MA_DUE_HI",
  "OW_BEFO_HI","OWRANK_HI",
  "HO_BEFO_HI","HORANK_HI","HO_DUE_HI"
))
FEATS_TRAIN <- intersect(FEATS, names(TRN))
FEATS_SCORE <- intersect(FEATS, names(SNAP0))
FEATS_COMMON <- intersect(FEATS_TRAIN, FEATS_SCORE)
FEATS_COMMON <- unique(c(FEATS_COMMON, "HA22_ID"))

# Build TRN2 with HA fields (no many-to-many joins)
ha_levels <- sort(unique(as.character(DAT$HA_22ND)))
TRN2 <- TRN %>%
  mutate(HA22_ID = as.integer(factor(as.character(HA_22ND), levels = ha_levels))) %>%
  select(HNO_FR, HA_22ND, HA22_ID, RANO_FR, ROUND, all_of(FEATS_COMMON)) %>%
  left_join(HIST_R %>% select(HNO_FR, ROUND, UPPRO_FR), by = c("HNO_FR","ROUND")) %>%
  rename(LABEL = UPPRO_FR) %>%
  filter(!is.na(LABEL)) %>%
  arrange(RANO_FR, ROUND, HNO_FR)

# Remove groups with n<2 (pairwise safety)
TRN2_grp <- TRN2 %>% count(RANO_FR, ROUND, name = "n")
TRN2 <- TRN2 %>% inner_join(TRN2_grp %>% filter(n >= 2), by = c("RANO_FR","ROUND"))

grp_train <- TRN2 %>% count(RANO_FR, ROUND, name = "n") %>% pull(n) %>% as.integer()
Xtrain <- as.matrix(TRN2[, FEATS_COMMON, drop=FALSE]); Xtrain[is.na(Xtrain)] <- 0

# SCORE with HA fields (no join)
SCORE <- SNAP0 %>% mutate(HA22_ID = as.integer(factor(as.character(HA_22ND), levels = ha_levels)))
SCORE <- SCORE %>% mutate(UPPRO_FR = num(UPPRO_FR))
Xscore <- as.matrix(SCORE[, FEATS_COMMON, drop=FALSE]); Xscore[is.na(Xscore)] <- 0

# -------- 5) Outputs (no model fit in NA variant) ----------------------------
cov_pct <- function(v) round(100 * mean(!is.na(v)), 2)
if (length(FEATS_COMMON)) {
  MODEL_IMPORTANCE <- data.frame(
    feature = FEATS_COMMON,
    coverage_pct = sapply(TRN2[, FEATS_COMMON, drop=FALSE], cov_pct),
    stringsAsFactors = FALSE
  ) %>% arrange(desc(coverage_pct))
} else {
  MODEL_IMPORTANCE <- data.frame(feature=character(), coverage_pct=numeric())
}

openxlsx::write.xlsx(
  list(
    SNAP0 = SCORE %>% select(HNO_FR, HNAME_FR, RANO_FR, CLASS, HA_22ND, HA22_ID,
                             all_of(setdiff(FEATS_COMMON, "HA22_ID")), UPPRO_FR),
    TRAIN = TRN2 %>% select(HNO_FR, HA_22ND, HA22_ID, RANO_FR, ROUND, LABEL,
                            all_of(setdiff(FEATS_COMMON, "HA22_ID"))),
    SCORE = SCORE %>% select(HNO_FR, HNAME_FR, RANO_FR, CLASS, UPPRO_FR),
    MODEL_IMPORTANCE = MODEL_IMPORTANCE,
    FEATURES_USED = data.frame(feature=FEATS_COMMON, stringsAsFactors = FALSE)
  ),
  file = OUT_XLSX, overwrite = TRUE
)

write.csv(TRN2, OUT_TRN2, row.names = FALSE)
cat(sprintf("DEVA_5_NA done. Written: %s\n", OUT_XLSX))
