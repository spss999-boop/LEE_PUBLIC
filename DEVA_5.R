#!/usr/bin/env Rscript
# DEVA_5.R  - RA3 r=0 ranking with XGBoost (pairwise)
# - Paths: C:/LEE
# - Target for ranking: UPPRO_FR (higher is better)
# - Features: PW_*_TR, FP_*_TR, TR_GAP_11 (+ HI snapshots mirrored FR→HI at r=0)
# - Groups for ranking: race key (RANO_FR) within each ROUND (>=1 for train, 0 for score)
# - Outputs: Excel with SNAP0 / TRAIN / SCORE / MODEL_IMPORTANCE / FEATURES_USED
# NOTE: Requires xgboost. If not installed, stops with an instruction.

# ========== 0) User params (source() friendly) ================================
# Edit SPAN here (YYMMDDRR), then source("C:/LEE/R/DEVA_5.R")
SPAN <- "25100308"
DAY  <- substr(SPAN, 1, 6)

# ========== 1) Paths ==========================================================
infile  <- sprintf("C:/LEE/out/raw/EHIS_RA3_%s.csv", SPAN)
outdir  <- sprintf("C:/LEE/out/SCORE/%s", DAY)
outfile <- sprintf("%s/DEVA_%s.xlsx", outdir, SPAN)
if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE, showWarnings = FALSE)

# ========== 2) Libraries & helpers ===========================================
suppressWarnings({
  library(dplyr)
  library(tidyr)
  library(openxlsx)
})
if (!requireNamespace("xgboost", quietly = TRUE)) stop("Please install.packages('xgboost') and re-run.")
options(stringsAsFactors = FALSE)
num <- function(x) suppressWarnings(as.numeric(x))
first_non_na <- function(v){ i <- which(!is.na(v))[1]; if (length(i)) v[i] else NA }

# ========== 3) Load ===========================================================
read_csv_win <- function(path){
  x <- tryCatch(read.csv(path, fileEncoding = "CP949"), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, fileEncoding = "UTF-8"), error=function(e) NULL)
  if (is.null(x)) stop("Failed to read CSV: ", path)
  x
}
DAT <- read_csv_win(infile)
names(DAT) <- toupper(trimws(names(DAT)))
req <- c("REFER1","HNO_FR","HNAME_FR","RANO_FR","CLASS")
miss <- setdiff(req, names(DAT)); if (length(miss)) stop("Missing keys: ", paste(miss, collapse=", "))
DAT$ROUND <- as.integer(DAT$REFER1)

# HIST/ SNAP0 ---------------------------------------------------------------
HIST <- DAT %>% filter(ROUND >= 1) %>% arrange(HNO_FR, ROUND)
SNAP0 <- DAT %>% arrange(HNO_FR, ROUND) %>% group_by(HNO_FR) %>%
  summarise(
    HNAME_FR = first_non_na(HNAME_FR),
    RANO_FR  = first_non_na(RANO_FR),
    CLASS    = first_non_na(CLASS),
    JO_BEFO_FR = first_non_na(JO_BEFO_FR), JORANK_FR = first_non_na(JORANK_FR), JO_TREN_FR = first_non_na(JO_TREN_FR), JO_DUE_FR = first_non_na(JO_DUE_FR),
    MA_BEFO_FR = first_non_na(MA_BEFO_FR), MARANK_FR = first_non_na(MARANK_FR), MA_DUE_FR = first_non_na(MA_DUE_FR),
    OW_BEFO_FR = first_non_na(OW_BEFO_FR), OWRANK_FR = first_non_na(OWRANK_FR),
    HO_BEFO_FR = first_non_na(HO_BEFO_FR), HORANK_FR = first_non_na(HORANK_FR), HO_DUE_FR = first_non_na(HO_DUE_FR),
    ENHO_FR = first_non_na(ENHO_FR), FANPRO_FR = first_non_na(FANPRO_FR), MANAGER_FR = first_non_na(MANAGER_FR), OWNER_FR = first_non_na(OWNER_FR), JOCKEY_FR = first_non_na(JOCKEY_FR),
    JO_GAP_RT = first_non_na(JO_GAP_RT), JO_GAP_RK = first_non_na(JO_GAP_RK), JO_GAP_DU = first_non_na(JO_GAP_DU),
    MA_GAP_RT = first_non_na(MA_GAP_RT), MA_GAP_RK = first_non_na(MA_GAP_RK), MA_GAP_DU = first_non_na(MA_GAP_DU),
    OW_GAP_RT = first_non_na(OW_GAP_RT), OW_GAP_RK = first_non_na(OW_GAP_RK),
    HO_GAP_RT = first_non_na(HO_GAP_RT), HO_GAP_RK = first_non_na(HO_GAP_RK), HO_GAP_DU = first_non_na(HO_GAP_DU),
    UPPRO_FR = first_non_na(UPPRO_FR),
    .groups = "drop"
  )
SNAP0$ROUND <- 0L

# ========== 4) Feature build (POWER-based windows + FR→HI mirror) ============
# per-round collapse (median) for POWER windows
HIST_R <- HIST %>% group_by(HNO_FR, ROUND) %>%
  summarise(
    RANO_FR = first_non_na(RANO_FR),
    CLASS = first_non_na(CLASS),
    POWERY = median(num(POWER), na.rm=TRUE),
    FANPRO_HI = first_non_na(FANPRO_HI),
    JO_BEFO_HI = median(num(JO_BEFO_HI), na.rm=TRUE), JORANK_HI = median(num(JORANK_HI), na.rm=TRUE), JO_TREN_HI = first_non_na(JO_TREN_HI), JO_DUE_HI = median(num(JO_DUE_HI), na.rm=TRUE),
    MA_BEFO_HI = median(num(MA_BEFO_HI), na.rm=TRUE), MARANK_HI = median(num(MARANK_HI), na.rm=TRUE), MA_DUE_HI = median(num(MA_DUE_HI), na.rm=TRUE),
    OW_BEFO_HI = median(num(OW_BEFO_HI), na.rm=TRUE), OWRANK_HI = median(num(OWRANK_HI), na.rm=TRUE),
    HO_BEFO_HI = median(num(HO_BEFO_HI), na.rm=TRUE), HORANK_HI = median(num(HORANK_HI), na.rm=TRUE), HO_DUE_HI = median(num(HO_DUE_HI), na.rm=TRUE),
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

# r=0: FR→HI mirror & PW/FP windows from HIST_R
TRN <- TRAIN_FULL
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
# r=0 forced mapping: FANPRO_FR -> FP_MI_TR (overwrite exact)
SNAP0$FP_MI_TR <- num(SNAP0$FANPRO_FR)


# TR pairwise (FR-HIavg etc.)
SNAP0$JO_RT_TR <- num(SNAP0$JO_GAP_RT); SNAP0$JO_RK_TR <- num(SNAP0$JO_GAP_RK); SNAP0$JO_DU_TR <- num(SNAP0$JO_GAP_DU)
SNAP0$MA_RT_TR <- num(SNAP0$MA_GAP_RT); SNAP0$MA_RK_TR <- num(SNAP0$MA_GAP_RK); SNAP0$MA_DU_TR <- num(SNAP0$MA_GAP_DU)
SNAP0$OW_RT_TR <- num(SNAP0$OW_GAP_RT); SNAP0$OW_RK_TR <- num(SNAP0$OW_GAP_RK)
SNAP0$HO_RT_TR <- num(SNAP0$HO_GAP_RT); SNAP0$HO_RK_TR <- num(SNAP0$HO_GAP_RK); SNAP0$HO_DU_TR <- num(SNAP0$HO_GAP_DU)
SNAP0$JO_TREN_TR <- num(SNAP0$JO_TREN_FR)

# FR -> HI mirror (r=0 snapshot uses current FR as HI to match TRAIN schema)
SNAP0$JO_BEFO_HI <- SNAP0$JO_BEFO_FR; SNAP0$JORANK_HI <- SNAP0$JORANK_FR; SNAP0$JO_TREN_HI <- SNAP0$JO_TREN_FR; SNAP0$JO_DUE_HI <- SNAP0$JO_DUE_FR
SNAP0$MA_BEFO_HI <- SNAP0$MA_BEFO_FR; SNAP0$MARANK_HI <- SNAP0$MARANK_FR; SNAP0$MA_DUE_HI <- SNAP0$MA_DUE_FR
SNAP0$OW_BEFO_HI <- SNAP0$OW_BEFO_FR; SNAP0$OWRANK_HI <- SNAP0$OWRANK_FR
SNAP0$HO_BEFO_HI <- SNAP0$HO_BEFO_FR; SNAP0$HORANK_HI <- SNAP0$HORANK_FR; SNAP0$HO_DUE_HI <- SNAP0$HO_DUE_FR

# ========== 5) Assemble matrices =============================================
# label for ranking: UPPRO_FR (higher better)
y_col <- "UPPRO_FR"
FEATS <- unique(c("PW_MI_TR","PW_MID_TR","PW_ALL_TR","FP_MI_TR","FP_MID_TR","FP_ALL_TR",
                  "JO_RT_TR","JO_RK_TR","JO_DU_TR","MA_RT_TR","MA_RK_TR","MA_DU_TR",
                  "OW_RT_TR","OW_RK_TR","HO_RT_TR","HO_RK_TR","HO_DU_TR",
                  "JO_BEFO_HI","JORANK_HI","JO_TREN_HI","JO_DUE_HI",
                  "MA_BEFO_HI","MARANK_HI","MA_DUE_HI",
                  "OW_BEFO_HI","OWRANK_HI",
                  "HO_BEFO_HI","HORANK_HI","HO_DUE_HI"))
# Use only features that exist in BOTH TRAIN (TRN) and SCORE (SNAP0)
FEATS_TRAIN <- intersect(FEATS, names(TRN))
FEATS_SCORE <- intersect(FEATS, names(SNAP0))
FEATS_COMMON <- intersect(FEATS_TRAIN, FEATS_SCORE)

# TRAIN matrix (ROUND>=1)
TRN2 <- TRN %>%
  dplyr::select(HNO_FR, RANO_FR, ROUND, all_of(FEATS_COMMON)) %>%
  dplyr::left_join(HIST %>% dplyr::select(HNO_FR, ROUND, UPPRO_FR), by = c("HNO_FR","ROUND")) %>%
  dplyr::rename(LABEL = UPPRO_FR) %>%
  dplyr::filter(!is.na(LABEL))

# Group vector: size per (RANO_FR, ROUND)
TRN2 <- TRN2 %>% arrange(RANO_FR, ROUND, HNO_FR)
grp_train <- TRN2 %>% count(RANO_FR, ROUND, name="n") %>% pull(n) %>% as.integer()
Xtrain <- as.matrix(TRN2[, FEATS_COMMON, drop=FALSE]); Xtrain[is.na(Xtrain)] <- 0

# SCORE matrix (ROUND=0)
SCORE <- SNAP0
SCORE <- SCORE %>% mutate(UPPRO_FR = num(UPPRO_FR))
Xscore <- as.matrix(SCORE[, FEATS_COMMON, drop=FALSE]); Xscore[is.na(Xscore)] <- 0

# ========== 6) Train ranker ==================================================
Dtrain <- xgboost::xgb.DMatrix(data=Xtrain, label=TRN2$LABEL)
xgboost::setinfo(Dtrain, "group", grp_train)
params <- list(objective="rank:pairwise", eval_metric="ndcg@4",
               eta=0.10, max_depth=5, min_child_weight=10, subsample=0.8, colsample_bytree=0.8)

ranker <- xgboost::xgb.train(params=params, data=Dtrain, nrounds=300, verbose=0)

# ========== 7) Predict & rank =================================================
Dscore <- xgboost::xgb.DMatrix(data=Xscore)
SCORE$POWER_HAT <- as.numeric(predict(ranker, Dscore))  # higher better
SCORE$POWER_SCORE <- SCORE$POWER_HAT
SCORE <- SCORE %>% arrange(desc(POWER_HAT)) %>% mutate(RANK = seq_len(n()))

# ========== 8) Feature importance ============================================
IMP <- xgboost::xgb.importance(model=ranker)
IMP <- IMP[,c("Feature","Gain","Cover","Frequency")]
colnames(IMP) <- c("feature","gain","cover","freq")

# ========== 9) Excel ==========================================================
wb <- createWorkbook()
add_sheet <- function(name, df){ addWorksheet(wb, name); writeData(wb, name, df); freezePane(wb, name, firstRow=TRUE); setColWidths(wb, name, 1:max(1,ncol(df)), widths="auto") }
add_sheet("SNAP0", SCORE %>% dplyr::select(HNO_FR, HNAME_FR, RANO_FR, CLASS, all_of(FEATS_COMMON), UPPRO_FR, POWER_HAT, POWER_SCORE, RANK))
add_sheet("TRAIN", TRN2 %>% dplyr::select(HNO_FR, RANO_FR, ROUND, LABEL, all_of(FEATS_COMMON)))
add_sheet("SCORE", SCORE %>% dplyr::select(HNO_FR, HNAME_FR, RANO_FR, CLASS, UPPRO_FR, POWER_HAT, POWER_SCORE, RANK))
add_sheet("MODEL_IMPORTANCE", IMP)
add_sheet("FEATURES_USED", data.frame(feature=FEATS_COMMON, stringsAsFactors = FALSE))

saveWorkbook(wb, outfile, overwrite=TRUE)
cat(sprintf("DEVA_5 done. Written: %s\n", outfile))
