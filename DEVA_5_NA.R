#!/usr/bin/env Rscript
# DEVA_5_NA.R — Full-row TRAIN + minimal SNAP0/SCORE + MODEL_IMPORTANCE + XGBoost
# Usage: source("C:/LEE/R/DEVA_5_NA.R")

# ========== 0) Params ========================================================
SPAN <- "25100308"                      # YYMMDDRR
DAY  <- substr(SPAN, 1, 6)
INFILE  <- sprintf("C:/LEE/out/raw/EHIS_RA3_%s.csv", SPAN)
OUTDIR  <- sprintf("C:/LEE/out/SCORE/%s", DAY)
if (!dir.exists(OUTDIR)) dir.create(OUTDIR, recursive = TRUE, showWarnings = FALSE)
OUT_XLSX <- sprintf("%s/DEVA_%s.xlsx", OUTDIR, SPAN)
OUT_XGB  <- sprintf("%s/xboost-%s.csv", OUTDIR, SPAN)

# ========== 1) Libraries & helpers (EXACTLY as in DEVA_5.R) ==================
suppressWarnings({
  library(dplyr)
  library(tidyr)
  library(openxlsx)
})
if (!requireNamespace("xgboost", quietly = TRUE)) stop("Please install.packages('xgboost') and re-run.")
options(stringsAsFactors = FALSE)
num <- function(x) suppressWarnings(as.numeric(x))
first_non_na <- function(v){ i <- which(!is.na(v))[1]; if (length(i)) v[i] else NA }

# ========== 2) Load (EXACT read_csv_win from DEVA_5.R) =======================
read_csv_win <- function(path){
  x <- tryCatch(read.csv(path, fileEncoding = "CP949"), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, fileEncoding = "UTF-8"), error=function(e) NULL)
  if (is.null(x)) stop("Failed to read CSV: ", path)
  x
}
DF <- read_csv_win(INFILE)
DF <- as.data.frame(DF)
names(DF) <- toupper(trimws(names(DF)))
if (!"HNO_FR" %in% names(DF)) stop("Missing HNO_FR")
if (!"HA_22ND" %in% names(DF)) stop("Missing HA_22ND")
if ("REFER1" %in% names(DF) && !"ROUND" %in% names(DF)) DF$ROUND <- suppressWarnings(as.integer(DF$REFER1))
levels_ha <- sort(unique(as.character(DF$HA_22ND)))
DF$HA22_ID <- as.integer(factor(as.character(DF$HA_22ND), levels = levels_ha))
if (!"LABEL" %in% names(DF) && "UPPRO_FR" %in% names(DF)) DF$LABEL <- DF$UPPRO_FR

# ========== 3) Build 9 features (HA22_ID→HNO_FR 정렬, ROUND 검증용 내림차순) ===
# 규칙: (1) 그룹키는 항상 HA22_ID, (2) 정렬은 HA22_ID→HNO_FR, (3) ROUND는 과→최근(내림차순) 확인용
#       (4) 1행 말은 원시값 그대로 투입, (5) HI 우선/FR 강제치환, (6) r=0에서 FR→HI 미러 + FANPRO_FR→PW_MI_TR

# HI 우선/FR 치환 헬퍼
pick_hi_fr <- function(d, hi, fr){
  h <- if (hi %in% names(d)) d[[hi]] else NULL
  f <- if (fr %in% names(d)) d[[fr]] else NULL
  v <- if (!is.null(h)) h else f
  num(v)
}

# 말(HNO_FR) 단위 윈도우 6개 + 조키 3개(치환 포함) 생성
make_feats_9 <- function(g){
  # 정렬: ROUND 오름차순으로 윈도 계산, 끝에 ROUND 내림차순로 보여줌
  g <- g[order(g$ROUND), , drop = FALSE]
  n <- nrow(g)
  pw <- if ("POWER" %in% names(g)) num(g$POWER) else rep(NA_real_, n)
  fp <- if ("FANPRO_HI" %in% names(g)) num(g$FANPRO_HI) else if ("FANPRO_FR" %in% names(g)) num(g$FANPRO_FR) else rep(NA_real_, n)
  
  # 윈도우: 여러 행일 때만 계산, 1행이면 자기값 유지
  PW_MI_TR  <- rep(NA_real_, n)
  PW_MID_TR <- rep(NA_real_, n)
  PW_ALL_TR <- rep(NA_real_, n)
  FP_MI_TR  <- rep(NA_real_, n)
  FP_MID_TR <- rep(NA_real_, n)
  FP_ALL_TR <- rep(NA_real_, n)
  
  if (n == 1){
    PW_MI_TR  <- pw
    PW_MID_TR <- pw
    PW_ALL_TR <- pw
    FP_MI_TR  <- fp
    FP_MID_TR <- fp
    FP_ALL_TR <- fp
  } else {
    for (i in seq_len(n)){
      j2 <- min(i+5L, n)
      if (i < n){
        PW_MI_TR[i]  <- pw[i+1L]
        PW_MID_TR[i] <- mean(pw[(i+1L):j2], na.rm=TRUE)
        PW_ALL_TR[i] <- mean(pw[(i+1L):n],  na.rm=TRUE)
        FP_MI_TR[i]  <- fp[i+1L]
        FP_MID_TR[i] <- mean(fp[(i+1L):j2], na.rm=TRUE)
        FP_ALL_TR[i] <- mean(fp[(i+1L):n],  na.rm=TRUE)
      } else {
        # 마지막 라운드 보정: 직전 정보가 없으므로 자기값으로 보존
        PW_MI_TR[i]  <- pw[i]
        PW_MID_TR[i] <- pw[i]
        PW_ALL_TR[i] <- pw[i]
        FP_MI_TR[i]  <- fp[i]
        FP_MID_TR[i] <- fp[i]
        FP_ALL_TR[i] <- fp[i]
      }
    }
  }
  
  # r=0 강제 미러: FR→HI, 특히 FANPRO_FR → PW_MI_TR 치환
  if ("ROUND" %in% names(g) && any(g$ROUND == 0)){
    idx0 <- which(g$ROUND == 0)
    if ("FANPRO_FR" %in% names(g)){
      PW_MI_TR[idx0] <- num(g$FANPRO_FR[idx0])  # 사용자 요구: FANPRO_FR -> PW_MI_TR
    }
  }
  
  # 조키 3개: HI 우선, 없으면 FR로 강제 치환
  JO_BEFO <- pick_hi_fr(g, "JO_BEFO_HI", "JO_BEFO_FR")
  JORANK  <- pick_hi_fr(g, "JORANK_HI",  "JORANK_FR")
  JO_TREN <- pick_hi_fr(g, "JO_TREN_HI",  "JO_TREN_FR")
  
  out <- cbind(g,
               PW_MI_TR=PW_MI_TR, PW_MID_TR=PW_MID_TR, PW_ALL_TR=PW_ALL_TR,
               FP_MI_TR=FP_MI_TR, FP_MID_TR=FP_MID_TR, FP_ALL_TR=FP_ALL_TR,
               JO_BEFO=JO_BEFO, JORANK=JORANK, JO_TREN=JO_TREN)
  # 보기용: ROUND 내림차순(과→최근)으로 반환
  out[order(out$ROUND, decreasing = TRUE), , drop = FALSE]
}

# 9개 피처 생성(HA22_ID→HNO_FR 정렬)
DF <- DF %>% arrange(HA22_ID, HNO_FR, ROUND)
HA22_HNO <- DF %>% group_by(HA22_ID, HNO_FR) %>% group_modify(~make_feats_9(.x)) %>% ungroup()

# ========== 4) TRAIN (xboost csv 기반 또는 HA22_HNO에서 키/라벨 선두) =========
lead_keys <- c("HA22_ID","HA_22ND","HNO_FR")
lead_more <- intersect(c("RANO_FR","LABEL","UPPRO_FR","REFER1","ROUND"), names(HA22_HNO))
rest_cols <- setdiff(names(HA22_HNO), c(lead_keys, lead_more))
ord <- c(lead_keys, lead_more, rest_cols)
seen <- character(0); final_cols <- character(0)
for (nm in ord) if (!(nm %in% seen)) { final_cols <- c(final_cols, nm); seen <- c(seen, nm) }
TRAIN <- HA22_HNO[, final_cols, drop = FALSE]
TRAIN <- TRAIN[order(TRAIN$HA22_ID, TRAIN$HNO_FR), , drop = FALSE]

# 실제 모델 투입 피처(9개)
FEATS_9 <- c("PW_MI_TR","PW_MID_TR","PW_ALL_TR","FP_MI_TR","FP_MID_TR","FP_ALL_TR","JO_BEFO","JORANK","JO_TREN")
feats <- intersect(FEATS_9, names(TRAIN))
key_cols <- c("HA22_ID","HA_22ND", intersect("HNO_FR", names(TRAIN)))
lab_col <- intersect(c("LABEL","UPPRO_FR"), names(TRAIN))
if (!length(feats)) stop("No model features found in TRAIN (FEATS_9 missing)")

# ========== 5) Grouping(항상 HA22_ID) & matrices ============================
TRAIN <- TRAIN[order(TRAIN$HA22_ID, TRAIN$HNO_FR), , drop = FALSE]
grp <- as.integer(rle(TRAIN$HA22_ID)$lengths)
Xtrain <- as.matrix(TRAIN[, feats, drop = FALSE]); Xtrain[is.na(Xtrain)] <- 0
label  <- if (length(lab_col)) num(TRAIN[[lab_col[1]]]) else rep(0, nrow(TRAIN))
Dtrain <- xgboost::xgb.DMatrix(data = Xtrain, label = label)
xgboost::setinfo(Dtrain, "group", grp)
params <- list(objective = "rank:pairwise", eval_metric = "ndcg@4",
               eta=0.10, max_depth=5, min_child_weight=10, subsample=0.8, colsample_bytree=0.8)
model <- xgboost::xgb.train(params = params, data = Dtrain, nrounds = 300, verbose = 0)

# ========== 6) 예측 전량 저장: pw_hat_22nd(유닛별 전 출전마×ROUND) ==========
Xall <- as.matrix(TRAIN[, feats, drop = FALSE]); Xall[is.na(Xall)] <- 0
PRED <- TRAIN[, c("HA22_ID","HA_22ND","HNO_FR","ROUND","RANO_FR"), drop = FALSE]
PRED$POWER_HAT <- as.numeric(predict(model, xgboost::xgb.DMatrix(Xall)))
# 유닛 내부 랭킹(참고용)
PRED <- PRED[order(PRED$HA22_ID, -PRED$POWER_HAT, PRED$HNO_FR), , drop = FALSE]
PRED$RANK_IN_HA <- ave(-PRED$POWER_HAT, PRED$HA22_ID, FUN = function(z) rank(z, ties.method = "first"))

# ========== 7) ha22_hno 시트: 9,200행 + 우측 9개 피처(검증용) ================
HA22_HNO_VIEW <- HA22_HNO[order(HA22_HNO$HA22_ID, HA22_HNO$HNO_FR, -HA22_HNO$ROUND), , drop = FALSE]
HA22_COLS <- intersect(c("HA22_ID","HA_22ND","HNO_FR","ROUND","RANO_FR","CLASS","UPPRO_FR"), names(HA22_HNO_VIEW))
ha22_hno_sheet <- HA22_HNO_VIEW[, c(HA22_COLS, feats), drop = FALSE]

# ========== 8) SCORE (HA22_ID 단위 랭킹 표시, 자르지 않음) =================
SCORE <- PRED[, c("HA22_ID","HA_22ND","HNO_FR","RANO_FR","POWER_HAT","RANK_IN_HA"), drop = FALSE]
colnames(SCORE)[colnames(SCORE)=="RANK_IN_HA"] <- "RANK"  # 시각 일관성

# ========== 9) MODEL_IMPORTANCE & FEATURES_USED =============================
IMP <- xgboost::xgb.importance(model = model)
IMP <- IMP[, c("Feature","Gain","Cover","Frequency")]
colnames(IMP) <- c("feature","gain","cover","freq")
FEATURES_USED <- data.frame(feature = feats, stringsAsFactors = FALSE)

# ========== 10) Outputs ======================================================
# XBOOST_DF: 모델 입력(키 + 라벨 + 9피처) — DEVA의 TRAIN 시트에 붙여넣기 용
XBOOST_DF <- TRAIN[, unique(c(key_cols, if ("RANO_FR" %in% names(TRAIN)) "RANO_FR" else NULL, lab_col, feats)), drop = FALSE]
for (cc in feats) if (is.numeric(XBOOST_DF[[cc]])) XBOOST_DF[[cc]][is.na(XBOOST_DF[[cc]])] <- 0
write.csv(XBOOST_DF, OUT_XGB, row.names = FALSE, fileEncoding = "CP949")

wb <- openxlsx::createWorkbook()
add_ws <- function(name, df){ openxlsx::addWorksheet(wb, name); openxlsx::writeData(wb, name, df); openxlsx::freezePane(wb, name, firstRow=TRUE); openxlsx::setColWidths(wb, name, 1:max(1,ncol(df)), widths="auto") }
add_ws("ha22_hno", ha22_hno_sheet)
add_ws("pw_hat_22nd", PRED)
add_ws("TRAIN", XBOOST_DF)
add_ws("SCORE", SCORE)
add_ws("MODEL_IMPORTANCE", IMP)
add_ws("FEATURES_USED", FEATURES_USED)
openxlsx::saveWorkbook(wb, OUT_XLSX, overwrite = TRUE)

cat(sprintf("DEVA_5_NA (HA22-groups) done. Rows: TRAIN=%d, PRED=%d, ha22_hno=%d → %s
",
            nrow(TRAIN), nrow(PRED), nrow(ha22_hno_sheet), OUT_XLSX))
