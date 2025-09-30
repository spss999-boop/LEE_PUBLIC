#!/usr/bin/env Rscript
# DEVA_2.R - RA3 r=0 scoring (robust, source-run). ASCII-only comments.

### 001) USER PARAMS ###########################################################
SPAN <- "25071309"   # <- set YYMMDDRR here

### 002) PATHS & LIBS ##########################################################
infile  <- sprintf("C:/LEE/out/raw/EHIS_RA3_%s.csv", SPAN)
outdir  <- sprintf("C:/LEE/out/SCORE/%s", substr(SPAN,1,6))
outfile <- sprintf("%s/DEVA_%s.xlsx", outdir, SPAN)
if (!dir.exists(outdir)) dir.create(outdir, recursive = TRUE, showWarnings = FALSE)

suppressWarnings({
  library(dplyr)
  library(openxlsx)
})
options(stringsAsFactors = FALSE)

### 003) HELPERS ###############################################################
read_csv_win <- function(path){
  x <- tryCatch(read.csv(path, fileEncoding="CP949"), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, fileEncoding="UTF-8"), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, fileEncoding="UTF-16LE", sep="\t"), error=function(e) NULL)
  if (is.null(x)) stop(paste("Failed to read CSV:", path))
  x
}
safe_first <- function(v){ i <- which(!is.na(v))[1]; if (length(i)==0 || is.na(i)) return(NA); v[i] }
to_num <- function(x){ suppressWarnings(as.numeric(x)) }
is_num <- function(x){ is.numeric(x) || is.integer(x) }
nuniq  <- function(v){ length(unique(v[!is.na(v)])) }

gap_long_to_tr <- c(
  "JO_GAP_RT"="JO_RT_TR","JO_GAP_RK"="JO_RK_TR","JO_GAP_DU"="JO_DU_TR",
  "MA_GAP_RT"="MA_RT_TR","MA_GAP_RK"="MA_RK_TR","MA_GAP_DU"="MA_DU_TR",
  "OW_GAP_RT"="OW_RT_TR","OW_GAP_RK"="OW_RK_TR",
  "HO_GAP_RT"="HO_RT_TR","HO_GAP_RK"="HO_RK_TR","HO_GAP_DU"="HO_DU_TR"
)
TR_GAP_11 <- c("JO_RT_TR","JO_RK_TR","JO_DU_TR","MA_RT_TR","MA_RK_TR","MA_DU_TR",
               "OW_RT_TR","OW_RK_TR","HO_RT_TR","HO_RK_TR","HO_DU_TR")

HI_INCLUDE <- c("JO_BEFO_HI","JORANK_HI","JO_TREN_HI","JO_DUE_HI",
                "MA_BEFO_HI","MARANK_HI","MA_DUE_HI",
                "OW_BEFO_HI","OWRANK_HI",
                "HO_BEFO_HI","HORANK_HI","HO_DUE_HI",
                "ENHO_HI","FANPRO_HI","MANAGER_HI","OWNER_HI")
HI_MODEL <- setdiff(HI_INCLUDE, c("MANAGER_HI","OWNER_HI"))

WIN6_TREN <- c("PW_MI_TR","PW_MID_TR","PW_ALL_TR","FP_MI_TR","FP_MID_TR","FP_ALL_TR","JO_TREN_TR")
COVER_MIN <- 0.50

median_impute <- function(df, cols){
  for (cc in cols){
    v <- df[[cc]]
    med <- suppressWarnings(stats::median(v, na.rm=TRUE))
    if (!is.na(med)) { v[is.na(v)] <- med; df[[cc]] <- v }
  }
  df
}
add_sheet_write <- function(wb, sheet, df, shade_cols=NULL, green_cols=NULL){
  addWorksheet(wb, sheet)
  writeData(wb, sheet, df)
  if (ncol(df)>0){
    addStyle(wb, sheet, createStyle(textDecoration="bold"), rows=1, cols=1:ncol(df), gridExpand=TRUE)
    if (nrow(df)>0){
      num_cols <- which(sapply(df, is.numeric))
      if (length(num_cols)>0) addStyle(wb, sheet, createStyle(numFmt="0.###"), rows=2:(nrow(df)+1), cols=num_cols, gridExpand=TRUE, stack=TRUE)
      if (!is.null(shade_cols) && length(shade_cols)>0) addStyle(wb, sheet, createStyle(fgFill="#F2F2F2"), rows=2:(nrow(df)+1), cols=shade_cols, gridExpand=TRUE, stack=TRUE)
      if (!is.null(green_cols) && length(green_cols)>0) addStyle(wb, sheet, createStyle(fgFill="#E6F4EA"), rows=2:(nrow(df)+1), cols=green_cols, gridExpand=TRUE, stack=TRUE)
    }
  }
  freezePane(wb, sheet, firstRow=TRUE)
  setColWidths(wb, sheet, 1:max(1,ncol(df)), widths="auto")
}
ord_last <- function(df){ if (!"HNO_FR1" %in% names(df)) return(df); cbind(df[, setdiff(names(df),"HNO_FR1"), drop=FALSE], HNO_FR1=df$HNO_FR1) }

# 003-ADD) recent helpers (ROUND: 1=most-recent, 2=older, ...)
pick_recent <- function(val, round) {
  val <- suppressWarnings(as.numeric(val))
  round <- as.integer(round)
  if (!length(val)) return(NA_real_)
  i <- which(!is.na(val))
  if (!length(i)) return(NA_real_)
  i <- i[order(round[i])]                 # smaller ROUND is more recent
  as.numeric(val[i[1]])
}
mean_recent_k <- function(val, round, k = 5) {
  val <- suppressWarnings(as.numeric(val))
  round <- as.integer(round)
  idx <- which(round %in% seq_len(k))
  if (!length(idx)) return(NA_real_)
  mean(val[idx], na.rm = TRUE)
}


### 004) LOAD & NORMALIZE #######################################################
DAT <- read_csv_win(infile)
names(DAT) <- trimws(toupper(names(DAT)))
req <- c("REFER1","HNO_FR","HNAME_FR","RANO_FR","CLASS")
miss <- setdiff(req, names(DAT)); if (length(miss)>0) stop(paste("Missing keys:", paste(miss, collapse=", ")))

DAT$ROUND   <- as.integer(DAT$REFER1)
DAT$HNO_FR  <- as.character(DAT$HNO_FR)
DAT$RANO_FR <- as.character(DAT$RANO_FR)
DAT$CLASS   <- as.character(DAT$CLASS)
if ("POWER" %in% names(DAT)) DAT$POWER <- to_num(DAT$POWER)

HIST <- DAT %>% dplyr::filter(ROUND >= 1) %>% dplyr::arrange(HNO_FR, ROUND)
if (nrow(HIST)==0) stop("No historical rows (ROUND>=1).")
HIST <- HIST %>% dplyr::group_by(HNO_FR) %>% dplyr::mutate(IS_LAST = ROUND == max(ROUND)) %>% dplyr::ungroup()

# special: non-last NA -> 0 for JO_TREN_HI, HO_BEFO_HI
for (nm in c("JO_TREN_HI","HO_BEFO_HI")){
  if (nm %in% names(HIST)){ v <- to_num(HIST[[nm]]); v[!HIST$IS_LAST & is.na(v)] <- 0; HIST[[nm]] <- v }
}

### 005) SNAP0 (FR->TR, FR->HI) ################################################
fr_keep <- intersect(c(
  "HNO_FR","HNAME_FR","RANO_FR","CLASS",
  "JO_BEFO_FR","JORANK_FR","JO_TREN_FR","JO_DUE_FR",
  "MA_BEFO_FR","MARANK_FR","MA_DUE_FR",
  "OW_BEFO_FR","OWRANK_FR",
  "HO_BEFO_FR","HORANK_FR","HO_DUE_FR",
  "ENHO_FR","FANPRO_FR","MANAGER_FR","OWNER_FR",
  names(gap_long_to_tr)
), names(DAT))

SNAP0 <- DAT %>% dplyr::arrange(HNO_FR, ROUND) %>% dplyr::group_by(HNO_FR) %>%
  dplyr::summarise(dplyr::across(all_of(setdiff(fr_keep,"HNO_FR")), ~ safe_first(.x)), .groups="drop")

SNAP0$ROUND <- 0L
SNAP0$HNO_FR <- as.character(SNAP0$HNO_FR)
SNAP0$RANO_FR <- as.character(SNAP0$RANO_FR)

# GAP(long) -> TR mirror
for (src in names(gap_long_to_tr)){
  tr <- gap_long_to_tr[[src]]
  SNAP0[[tr]] <- if (src %in% names(SNAP0)) to_num(SNAP0[[src]]) else NA_real_
}

# JO_TREN_TR at r=0: FR > HI > NA (later impute with med if needed)
SNAP0$JO_TREN_TR <- if ("JO_TREN_FR" %in% names(SNAP0)) to_num(SNAP0$JO_TREN_FR) else if ("JO_TREN_HI" %in% names(SNAP0)) to_num(SNAP0$JO_TREN_HI) else NA_real_

# FR -> HI mirror (16 HI)
mirror_pairs <- c(
  "JO_BEFO_HI"="JO_BEFO_FR","JORANK_HI"="JORANK_FR","JO_TREN_HI"="JO_TREN_FR","JO_DUE_HI"="JO_DUE_FR",
  "MA_BEFO_HI"="MA_BEFO_FR","MARANK_HI"="MARANK_FR","MA_DUE_HI"="MA_DUE_FR",
  "OW_BEFO_HI"="OW_BEFO_FR","OWRANK_HI"="OWRANK_FR",
  "HO_BEFO_HI"="HO_BEFO_FR","HORANK_HI"="HORANK_FR","HO_DUE_HI"="HO_DUE_FR",
  "ENHO_HI"="ENHO_FR","FANPRO_HI"="FANPRO_FR","MANAGER_HI"="MANAGER_FR","OWNER_HI"="OWNER_FR"
)
for (hi in names(mirror_pairs)){
  frn <- mirror_pairs[[hi]]
  SNAP0[[hi]] <- if (frn %in% names(SNAP0)) SNAP0[[frn]] else NA
}

# Debut flag
HINFO <- HIST %>% dplyr::group_by(HNO_FR) %>% dplyr::summarise(N_HI = dplyr::n(), .groups="drop")
SNAP0 <- SNAP0 %>% dplyr::left_join(HINFO, by="HNO_FR") %>% dplyr::mutate(IS_DEBUT = is.na(N_HI) | N_HI==0L)

### 006) TRAIN (windows, TR from HI; drop last) ################################
y_col <- if ("POWER" %in% names(HIST)) "POWER" else if ("MED_POWER" %in% names(HIST)) "MED_POWER" else NA_character_
if (is.na(y_col)) stop("POWER or MED_POWER not found.")
if (!"FANPRO_HI" %in% names(HIST)) stop("FANPRO_HI missing in HIST.")

make_train <- function(H){
  sp <- split(H, H$HNO_FR)
  out <- lapply(sp, function(df){
    df <- df[order(df$ROUND), , drop=FALSE]
    n <- nrow(df)
    pw <- to_num(df[[y_col]])
    fp <- to_num(df[["FANPRO_HI"]])
    PW_MI_TR  <- PW_MID_TR <- PW_ALL_TR <- FP_MI_TR <- FP_MID_TR <- FP_ALL_TR <- rep(NA_real_, n)
    add_gap <- function(cur, past) if (all(is.finite(c(cur,past)))) cur - past else NA_real_
    mean_later <- function(v, idx) if (length(idx)>0) mean(v[idx], na.rm=TRUE) else NA_real_
    JO_RT_TR <- JO_RK_TR <- JO_DU_TR <- rep(NA_real_, n)
    MA_RT_TR <- MA_RK_TR <- MA_DU_TR <- rep(NA_real_, n)
    OW_RT_TR <- OW_RK_TR <- rep(NA_real_, n)
    HO_RT_TR <- HO_RK_TR <- HO_DU_TR <- rep(NA_real_, n)
    
    for (i in seq_len(n)){
      idx <- if (i < n) (i+1L):n else integer(0)
      if (i < n){
        j2 <- min(i+5L, n)
        PW_MI_TR[i]  <- pw[i+1L]
        PW_MID_TR[i] <- mean(pw[(i+1L):j2], na.rm=TRUE)
        PW_ALL_TR[i] <- mean(pw[(i+1L):n],  na.rm=TRUE)
        FP_MI_TR[i]  <- fp[i+1L]
        FP_MID_TR[i] <- mean(fp[(i+1L):j2], na.rm=TRUE)
        FP_ALL_TR[i] <- mean(fp[(i+1L):n],  na.rm=TRUE)
      }
      if ("JO_BEFO_HI" %in% names(df)) JO_RT_TR[i] <- add_gap(to_num(df$JO_BEFO_HI[i]), mean_later(to_num(df$JO_BEFO_HI), idx))
      if ("JORANK_HI"  %in% names(df)) JO_RK_TR[i] <- add_gap(mean_later(to_num(df$JORANK_HI), idx), to_num(df$JORANK_HI[i]))
      if ("JO_DUE_HI"  %in% names(df)) JO_DU_TR[i] <- add_gap(to_num(df$JO_DUE_HI[i]),  mean_later(to_num(df$JO_DUE_HI), idx))
      if ("MA_BEFO_HI" %in% names(df)) MA_RT_TR[i] <- add_gap(to_num(df$MA_BEFO_HI[i]), mean_later(to_num(df$MA_BEFO_HI), idx))
      if ("MARANK_HI"  %in% names(df)) MA_RK_TR[i] <- add_gap(mean_later(to_num(df$MARANK_HI), idx), to_num(df$MARANK_HI[i]))
      if ("MA_DUE_HI"  %in% names(df)) MA_DU_TR[i] <- add_gap(to_num(df$MA_DUE_HI[i]),  mean_later(to_num(df$MA_DUE_HI), idx))
      if ("OW_BEFO_HI" %in% names(df)) OW_RT_TR[i] <- add_gap(to_num(df$OW_BEFO_HI[i]), mean_later(to_num(df$OW_BEFO_HI), idx))
      if ("OWRANK_HI"  %in% names(df)) OW_RK_TR[i] <- add_gap(mean_later(to_num(df$OWRANK_HI), idx), to_num(df$OWRANK_HI[i]))
      if ("HO_BEFO_HI" %in% names(df)) HO_RT_TR[i] <- add_gap(to_num(df$HO_BEFO_HI[i]), mean_later(to_num(df$HO_BEFO_HI), idx))
      if ("HORANK_HI"  %in% names(df)) HO_RK_TR[i] <- add_gap(mean_later(to_num(df$HORANK_HI), idx), to_num(df$HORANK_HI[i]))
      if ("HO_DUE_HI"  %in% names(df)) HO_DU_TR[i] <- add_gap(to_num(df$HO_DUE_HI[i]),  mean_later(to_num(df$HO_DUE_HI), idx))
    }
    df$PW_MI_TR <- PW_MI_TR; df$PW_MID_TR <- PW_MID_TR; df$PW_ALL_TR <- PW_ALL_TR
    df$FP_MI_TR <- FP_MI_TR; df$FP_MID_TR <- FP_MID_TR; df$FP_ALL_TR <- FP_ALL_TR
    df$JO_TREN_TR <- to_num(df$JO_TREN_HI)
    df$JO_RT_TR <- JO_RT_TR; df$JO_RK_TR <- JO_RK_TR; df$JO_DU_TR <- JO_DU_TR
    df$MA_RT_TR <- MA_RT_TR; df$MA_RK_TR <- MA_RK_TR; df$MA_DU_TR <- MA_DU_TR
    df$OW_RT_TR <- OW_RT_TR; df$OW_RK_TR <- OW_RK_TR
    df$HO_RT_TR <- HO_RT_TR; df$HO_RK_TR <- HO_RK_TR; df$HO_DU_TR <- HO_DU_TR
    df$Y <- to_num(df$POWER)
    df
  })
  TRAIN <- dplyr::bind_rows(out)
  TRAIN <- TRAIN %>% dplyr::group_by(HNO_FR) %>% dplyr::mutate(IS_LAST = ROUND == max(ROUND)) %>% dplyr::ungroup()
  TRAIN <- TRAIN %>% dplyr::filter(!IS_LAST)
  TRAIN
}
TRAIN_FULL <- make_train(HIST)

### 007) TRAIN FEATURE SELECTION ################################################
ALL_FEATS <- unique(c(WIN6_TREN, TR_GAP_11, HI_INCLUDE))
present_feats <- intersect(ALL_FEATS, names(TRAIN_FULL))
isnum_vec <- sapply(TRAIN_FULL[, present_feats, drop=FALSE], is_num)
num_feats <- present_feats[isnum_vec]
coverage <- if (length(num_feats)) sapply(TRAIN_FULL[, num_feats, drop=FALSE], function(x) mean(!is.na(x))) else numeric(0)
uniq_cnt <- if (length(num_feats)) sapply(TRAIN_FULL[, num_feats, drop=FALSE], function(x) length(unique(x[!is.na(x)]))) else numeric(0)
USED_FEATS <- if (length(num_feats)) num_feats[coverage[num_feats] >= COVER_MIN & uniq_cnt[num_feats] >= 2] else character(0)
if (length(USED_FEATS) < 4L) USED_FEATS <- intersect(WIN6_TREN, names(TRAIN_FULL))

TRAIN <- TRAIN_FULL %>% dplyr::filter(!is.na(Y)) %>% dplyr::select(HNO_FR, ROUND, CLASS, dplyr::all_of(USED_FEATS), Y)
TRAIN_BEFORE_IMPUTE <- TRAIN
TRAIN <- median_impute(TRAIN, USED_FEATS)

is_const <- function(v){
  if (!v %in% names(TRAIN_BEFORE_IMPUTE)) return(TRUE)
  x <- TRAIN_BEFORE_IMPUTE[[v]]; x <- x[!is.na(x)]
  length(x)==0 || length(unique(x))<=1
}
dropped_const <- if (length(USED_FEATS)) USED_FEATS[vapply(USED_FEATS, is_const, logical(1))] else character(0)
USED_FEATS <- setdiff(USED_FEATS, dropped_const)
if (length(USED_FEATS) < 1L) stop("No usable features for LM after constant-drop.")

### 008) SCORE (r=0 windows) — use recent rounds (ROUND=1 is most recent) #####
PW_FP <- HIST %>%
  dplyr::arrange(HNO_FR, ROUND) %>%
  dplyr::group_by(HNO_FR) %>%
  dplyr::summarise(
    # POWER windows: recent(1), recent 1..5 avg, all avg
    PW_MI_TR  = pick_recent(.data[["POWER"]], ROUND),
    PW_MID_TR = mean_recent_k(.data[["POWER"]], ROUND, k = 5),
    PW_ALL_TR = mean(suppressWarnings(as.numeric(.data[["POWER"]])), na.rm = TRUE),
    # FANPRO windows: same rule on FANPRO_HI
    FP_MI_TR  = pick_recent(FANPRO_HI, ROUND),
    FP_MID_TR = mean_recent_k(FANPRO_HI, ROUND, k = 5),
    FP_ALL_TR = mean(suppressWarnings(as.numeric(FANPRO_HI)), na.rm = TRUE),
    .groups = "drop"
  )

SNAP0 <- SNAP0 %>% dplyr::left_join(PW_FP, by = "HNO_FR")
SCORE <- SNAP0 %>% dplyr::filter(!IS_DEBUT)
SCORE_DEBUT <- SNAP0 %>% dplyr::filter(IS_DEBUT)
SCORE_BEFORE_IMPUTE <- SCORE


### 009) FIT & PREDICT #########################################################
form <- stats::as.formula(paste("Y ~", paste(USED_FEATS, collapse=" + ")))
fit  <- stats::lm(form, data=TRAIN)

for (v in USED_FEATS){
  if (!v %in% names(SCORE)) SCORE[[v]] <- NA_real_
  if (any(is.na(SCORE[[v]]))){
    med <- suppressWarnings(stats::median(TRAIN[[v]], na.rm=TRUE))
    if (!is.na(med)) SCORE[[v]][is.na(SCORE[[v]])] <- med
  }
}
SCORE$POWER_HAT   <- as.numeric(stats::predict(fit, newdata=SCORE[, USED_FEATS, drop=FALSE]))
SCORE$POWER_SCORE <- -SCORE$POWER_HAT
SCORE <- SCORE[order(SCORE$POWER_HAT, decreasing=FALSE), ]
SCORE$RANK <- seq_len(nrow(SCORE))

### 010) TABLES ################################################################
TRAIN_OUT <- TRAIN[, c("HNO_FR","ROUND","CLASS", USED_FEATS, "Y")]
MODEL <- data.frame(feature = names(stats::coef(fit)), coef = as.numeric(stats::coef(fit)), row.names=NULL)

qa_feats <- intersect(unique(c(WIN6_TREN, TR_GAP_11)), names(SCORE))
QA_UNIQ_SC <- if (length(qa_feats)) data.frame(feature=qa_feats, uniq_cnt_sc=sapply(SCORE[, qa_feats, drop=FALSE], nuniq)) else data.frame(feature=character(), uniq_cnt_sc=integer())
QA_NA_SC   <- if (length(qa_feats)) data.frame(feature=qa_feats, na_count_sc=sapply(SCORE[, qa_feats, drop=FALSE], function(x) sum(is.na(x)))) else data.frame(feature=character(), na_count_sc=integer())

feat_pool <- intersect(unique(c(WIN6_TREN, TR_GAP_11, HI_MODEL)), names(TRAIN_BEFORE_IMPUTE))
cov2  <- if (length(feat_pool)) sapply(TRAIN_BEFORE_IMPUTE[, feat_pool, drop=FALSE], function(x) mean(!is.na(x))) else numeric(0)
uniq2 <- if (length(feat_pool)) sapply(TRAIN_BEFORE_IMPUTE[, feat_pool, drop=FALSE], function(x) length(unique(x[!is.na(x)]))) else numeric(0)
FEATURES_USED <- data.frame(
  feature = names(cov2),
  coverage = as.numeric(cov2),
  uniq_cnt = as.numeric(uniq2),
  used_in_model = names(cov2) %in% USED_FEATS,
  row.names = NULL
)

SCORE$HNO_FR1 <- SCORE$HNO_FR
SNAP0$HNO_FR1 <- SNAP0$HNO_FR
TRAIN_OUT$HNO_FR1 <- TRAIN_OUT$HNO_FR

### 011) NA REPORTS (ALWAYS) ###################################################
DT_BASE <- DAT %>% dplyr::filter(ROUND >= 1)

cols_hi_train <- intersect(HI_INCLUDE, names(DT_BASE))
NA_HI_TRAIN <- if (length(cols_hi_train)==0) data.frame() else {
  tmp <- DT_BASE[, c("HNO_FR","ROUND", cols_hi_train)]
  tmp[rowSums(is.na(tmp[, cols_hi_train, drop=FALSE])) > 0, , drop=FALSE]
}
cols_hi_score <- intersect(HI_INCLUDE, names(SCORE))
NA_HI_SCORE <- if (length(cols_hi_score)==0) data.frame() else {
  tmp <- SCORE[, c("HNO_FR","ROUND", cols_hi_score), drop=FALSE]
  tmp[rowSums(is.na(tmp[, cols_hi_score, drop=FALSE])) > 0, , drop=FALSE]
}
NA_FEATURES_TRAIN <- if (length(USED_FEATS)==0) data.frame() else {
  tmp <- TRAIN_BEFORE_IMPUTE[, c("HNO_FR","ROUND", USED_FEATS)]
  tmp[rowSums(is.na(tmp[, USED_FEATS, drop=FALSE])) > 0, , drop=FALSE]
}
S_BEFORE <- SCORE_BEFORE_IMPUTE
for (v in USED_FEATS) if (!v %in% names(S_BEFORE)) S_BEFORE[[v]] <- NA
NA_FEATURES_SCORE <- if (length(USED_FEATS)==0) data.frame() else {
  tmp <- S_BEFORE[, c("HNO_FR","ROUND", USED_FEATS)]
  tmp[rowSums(is.na(tmp[, USED_FEATS, drop=FALSE])) > 0, , drop=FALSE]
}
NA_SUMMARY <- {
  a <- if (length(USED_FEATS)) data.frame(FEATURE=USED_FEATS, N_MISSING=colSums(is.na(TRAIN_BEFORE_IMPUTE[, USED_FEATS, drop=FALSE])), WHERE="TRAIN", row.names=NULL) else NULL
  b <- if (length(USED_FEATS)) data.frame(FEATURE=USED_FEATS, N_MISSING=colSums(is.na(S_BEFORE[, USED_FEATS, drop=FALSE])), WHERE="SCORE", row.names=NULL) else NULL
  if (is.null(a) && is.null(b)) data.frame(FEATURE=character(), N_MISSING=integer(), WHERE=character()) else rbind(a,b)
}
DROPPED_CONST_FEATS <- if (length(dropped_const)>0) {
  data.frame(FEATURE=dropped_const, REASON=rep("zero_variance_in_training", length(dropped_const)))
} else {
  data.frame(FEATURE=character(), REASON=character())
}

### 012) EXCEL #################################################################
wb <- createWorkbook()
add_sheet_write(wb, "SNAP0", ord_last(SNAP0))
add_sheet_write(wb, "TRAIN", ord_last(TRAIN_OUT))
add_sheet_write(wb, "MODEL", MODEL)
score_cols <- names(SCORE)
shade_targets <- intersect(c(WIN6_TREN, TR_GAP_11), score_cols)
green_targets <- intersect(c("POWER_HAT","POWER_SCORE","RANK"), score_cols)
shade_cols <- match(shade_targets, score_cols); shade_cols <- shade_cols[!is.na(shade_cols)]
green_cols <- match(green_targets, score_cols); green_cols <- green_cols[!is.na(green_cols)]
add_sheet_write(wb, "SCORE", ord_last(SCORE), shade_cols=shade_cols, green_cols=green_cols)
if (nrow(SCORE_DEBUT)>0) add_sheet_write(wb, "SCORE_DEBUT", ord_last(SCORE_DEBUT))
add_sheet_write(wb, "QA_UNIQ_SC", QA_UNIQ_SC)
add_sheet_write(wb, "QA_NA_SC",   QA_NA_SC)
add_sheet_write(wb, "FEATURES_USED", FEATURES_USED)
add_sheet_write(wb, "NA_HI_TRAIN", NA_HI_TRAIN)
add_sheet_write(wb, "NA_HI_SCORE", NA_HI_SCORE)
add_sheet_write(wb, "NA_FEATURES_TRAIN", NA_FEATURES_TRAIN)
add_sheet_write(wb, "NA_FEATURES_SCORE", NA_FEATURES_SCORE)
add_sheet_write(wb, "NA_SUMMARY", NA_SUMMARY)
add_sheet_write(wb, "DROPPED_CONST_FEATS", DROPPED_CONST_FEATS)
saveWorkbook(wb, outfile, overwrite=TRUE)
cat(sprintf("DEVA done. Written: %s\n", outfile))
