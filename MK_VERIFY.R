# MK_VERIFY.R (POOR.DBF에 RANO_FR, HNO_FR, RECCO, HA_22ND 포함)
# - RAW, C_TABLE(말×라운드), POOR_LOCAL(AA×라운드), CANDIDATE, HR_RESULT 생성
# - HR_DEL(말 단위)와 POOR_LOCAL(AA 단위) 분리 계산
# - POOR(DBF): 레코드(RECCO) 기준으로 플래그를 되돌려 붙여 저장
# - DBF 컬럼: RANO_FR, HNO_FR, RECCO, HA_22ND, POOR_LOCAL, HR_DEL
#
# 실행 예:
# Rscript R/MK_VERIFY.R "C:/HK/RA3_RANO1.DBF" "C:/HK/OUT" "C:/HK/POOR.DBF" 2.5

suppressWarnings({ options(stringsAsFactors = FALSE) })

# -------------------- 파라미터 --------------------
args <- commandArgs(trailingOnly = TRUE)
in_path   <- if (length(args) >= 1) args[1] else "C:/HK/RA3_RANO1.DBF"   # 입력(RAW)
out_dir   <- if (length(args) >= 2) args[2] else "C:/HK/OUT"              # 엑셀 출력 디렉터리
poor_out  <- if (length(args) >= 3) args[3] else "C:/HK/POOR.DBF"         # POOR DBF 경로
z_cut     <- if (length(args) >= 4) as.numeric(args[4]) else 2.5           # 게이트 Z 컷
eps       <- 1e-12
n_min     <- 3                                                              # N_ROUND 임계
write_dbf <- TRUE                                                           # DBF 저장 고정(On)

# -------------------- 유틸 --------------------
need_pkg <- function(pkg) {
  if (!requireNamespace(pkg, quietly = TRUE)) stop("패키지 설치 필요: ", pkg)
}
ucol <- function(df) { names(df) <- toupper(trimws(names(df))); df }
mad_s <- function(x) {
  x <- as.numeric(x); x <- x[is.finite(x)]
  if (length(x) == 0) return(NA_real_)
  val <- stats::mad(x, constant = 1.4826, na.rm = TRUE)
  if (!is.finite(val)) NA_real_ else val
}
p80 <- function(x) {
  x <- as.numeric(x); x <- x[is.finite(x)]
  if (length(x) == 0) return(NA_real_)
  as.numeric(stats::quantile(x, probs = 0.80, type = 7, na.rm = TRUE))
}
gate_flag <- function(x, med, madv, p80v, zcut, eps=1e-12) {
  z <- if (is.finite(madv) && madv > 0) (x - med)/madv else NA_real_
  gz <- is.finite(z) && (z >= zcut)
  gp <- is.finite(x) && is.finite(p80v) && (x > (p80v + eps))
  gate <- if (gz && gp) "Z|P80" else if (gz) "Z" else if (gp) "P80" else ""
  list(gate=gate, z=z)
}
safe_mean <- function(x) {
  x <- as.numeric(x); x <- x[is.finite(x)]
  if (length(x) == 0) NA_real_ else mean(x)
}

# -------------------- 입력 로드 --------------------
need_pkg("openxlsx")
if (!file.exists(in_path)) stop("입력 파일 없음: ", in_path)

ext <- tolower(tools::file_ext(in_path))
if (ext == "dbf") {
  need_pkg("foreign")
  raw0 <- tryCatch(as.data.frame(foreign::read.dbf(in_path, as.is = TRUE)), error=function(e) NULL)
} else if (ext == "csv") {
  raw0 <- tryCatch(read.csv(in_path, header=TRUE, check.names = FALSE, na.strings = c("","NA")), error=function(e) NULL)
} else if (ext %in% c("xlsx","xls")) {
  raw0 <- tryCatch(openxlsx::read.xlsx(in_path), error=function(e) NULL)
} else {
  stop("지원하지 않는 입력 형식: ", ext, " (DBF/CSV/XLSX)")
}
if (is.null(raw0)) stop("입력 로드 실패: ", in_path)
raw0 <- ucol(raw0)

# 필수 컬럼 확인/정규화
if (!("RECCO"  %in% names(raw0))) stop("RECCO 컬럼이 필요합니다.")
ha_col <- if ("HA22_G" %in% names(raw0)) "HA22_G" else if ("HA_22ND" %in% names(raw0)) "HA_22ND" else NA
if (is.na(ha_col)) stop("HA22_G 또는 HA_22ND 컬럼이 필요합니다.")
if (!("HNO_FR" %in% names(raw0))) stop("HNO_FR 컬럼이 필요합니다.")
if (!("REFER1" %in% names(raw0))) stop("REFER1 컬럼이 필요합니다.")
if (!("POWER"  %in% names(raw0))) stop("POWER 컬럼이 필요합니다.")
has_rano <- "RANO_FR" %in% names(raw0)
if (!has_rano) message("참고: RANO_FR 컬럼이 없어 DBF에는 NA로 기록됩니다.")

# RAW 구성(요청 필드 보존: HA_22ND, RANO_FR, DEL_INTR)
RAW <- data.frame(
  RECCO   = raw0$RECCO,
  HA22_G  = raw0[[ha_col]],
  HNO_FR  = raw0$HNO_FR,
  REFER1  = raw0$REFER1,
  POWER   = raw0$POWER,
  HA_22ND = if ("HA_22ND" %in% names(raw0)) raw0$HA_22ND else raw0[[ha_col]],
  RANO_FR = if (has_rano) raw0$RANO_FR else NA,
  DEL_INTR = if ("DEL_INTR" %in% names(raw0)) raw0$DEL_INTR else 0,
  stringsAsFactors = FALSE
)

# 필터: REFER1>=1, 숫자 변환
RAW$REFER1 <- suppressWarnings(as.numeric(RAW$REFER1))
RAW$HNO_FR <- suppressWarnings(as.numeric(RAW$HNO_FR))
RAW$POWER  <- suppressWarnings(as.numeric(RAW$POWER))
RAW$DEL_INTR <- suppressWarnings(as.numeric(RAW$DEL_INTR))
RAW$DEL_INTR[is.na(RAW$DEL_INTR)] <- 0
RAW <- RAW[is.finite(RAW$REFER1) & RAW$REFER1 >= 1 & is.finite(RAW$HNO_FR) & is.finite(RAW$POWER), ]

ROWCOUNT_IN <- nrow(RAW)

# -------------------- AA 레벨 대표값(X_R) 및 r_x 계산 --------------------
library(stats)

# DEL_INTR != 0인 행을 필터링한 데이터로 r_x 계산 (mean 사용)
RAW_filtered <- RAW[RAW$DEL_INTR != 0, ]
if (nrow(RAW_filtered) == 0) {
  # DEL_INTR != 0인 데이터가 없으면 전체 데이터 사용
  RAW_filtered <- RAW
}

# r_x 계산 (mean 사용)
r_x_tab <- aggregate(POWER ~ HA22_G + HNO_FR + REFER1, data = RAW_filtered, FUN = safe_mean)
names(r_x_tab)[names(r_x_tab)=="POWER"] <- "r_x"

# 기존 X_R 계산 유지 (호환성을 위해)
X_R_tab <- aggregate(POWER ~ HA22_G + HNO_FR + REFER1, data = RAW, FUN = safe_mean)
names(X_R_tab)[names(X_R_tab)=="POWER"] <- "X_R"

# RAW 데이터에 r_x 추가
RAW <- merge(RAW, r_x_tab[, c("HA22_G", "HNO_FR", "REFER1", "r_x")], 
             by = c("HA22_G", "HNO_FR", "REFER1"), all.x = TRUE, sort = FALSE)

aa_stats <- do.call(rbind, lapply(split(X_R_tab, list(X_R_tab$HA22_G, X_R_tab$HNO_FR), drop=TRUE), function(df){
  xr <- as.numeric(df$X_R)
  med <- stats::median(xr, na.rm=TRUE)
  madv<- mad_s(xr)
  p80v<- p80(xr)
  nrd <- sum(is.finite(xr))
  data.frame(HA22_G=df$HA22_G[1], HNO_FR=df$HNO_FR[1],
             MEDIAN_AA=med, MAD_AA=madv, P80_AA=p80v, N_ROUND=nrd, stringsAsFactors = FALSE)
}))
row.names(aa_stats) <- NULL

AA <- merge(X_R_tab, aa_stats, by=c("HA22_G","HNO_FR"), all.x=TRUE, sort=FALSE)

# r_x 데이터를 AA에 병합하여 xr 필드 생성
AA <- merge(AA, r_x_tab[, c("HA22_G", "HNO_FR", "REFER1", "r_x")], 
            by = c("HA22_G", "HNO_FR", "REFER1"), all.x = TRUE, sort = FALSE)
# xr 필드는 r_x를 직접 참조
AA$xr <- AA$r_x
gf <- mapply(function(x, med, madv, p80v){
  g <- gate_flag(x, med, madv, p80v, z_cut, eps)
  c(g$gate, g$z)
}, AA$X_R, AA$MEDIAN_AA, AA$MAD_AA, AA$P80_AA, SIMPLIFY = FALSE)
AA$GATE <- vapply(gf, function(v) as.character(v[[1]]), character(1))
AA$Z_MAD <- as.numeric(vapply(gf, function(v) as.numeric(v[[2]]), numeric(1)))
AA$POOR_LOCAL <- as.integer( (AA$N_ROUND >= n_min) & (AA$GATE != "") )

# CANDIDATE
CANDIDATE <- AA
CANDIDATE$CANDIDATE <- as.integer(CANDIDATE$GATE != "")

# -------------------- HR 레벨 대표값(X_C) --------------------
X_C_tab <- aggregate(POWER ~ HNO_FR + REFER1, data = RAW, FUN = safe_mean)
names(X_C_tab)[names(X_C_tab)=="POWER"] <- "X_C"

hr_stats <- do.call(rbind, lapply(split(X_C_tab, X_C_tab$HNO_FR, drop=TRUE), function(df){
  xc  <- as.numeric(df$X_C)
  med <- stats::median(xc, na.rm=TRUE)
  madv<- mad_s(xc)
  p80v<- p80(xc)
  nrd <- sum(is.finite(xc))
  data.frame(HNO_FR=df$HNO_FR[1],
             MEDIAN=med, MAD=madv, P80=p80v, N_ROUND_HR=nrd, stringsAsFactors = FALSE)
}))
row.names(hr_stats) <- NULL

C_TABLE <- merge(X_C_tab, hr_stats, by="HNO_FR", all.x=TRUE, sort=FALSE)
gf2 <- mapply(function(x, med, madv, p80v){
  g <- gate_flag(x, med, madv, p80v, z_cut, eps)
  c(g$gate, g$z)
}, C_TABLE$X_C, C_TABLE$MEDIAN, C_TABLE$MAD, C_TABLE$P80, SIMPLIFY = FALSE)
C_TABLE$GATE  <- vapply(gf2, function(v) as.character(v[[1]]), character(1))
C_TABLE$Z_MAD <- as.numeric(vapply(gf2, function(v) as.numeric(v[[2]]), numeric(1)))
C_TABLE$HR_DEL <- as.integer( (C_TABLE$N_ROUND_HR >= n_min) & (C_TABLE$GATE != "") )

HR_RESULT <- C_TABLE

# -------------------- 정렬 및 컬럼 정리 --------------------
ord_AA <- AA[order(AA$HA22_G, AA$HNO_FR, AA$REFER1),
             c("HA22_G","HNO_FR","REFER1","X_R","xr","MEDIAN_AA","MAD_AA","Z_MAD","P80_AA","N_ROUND","GATE","POOR_LOCAL")]
ord_CT <- C_TABLE[order(C_TABLE$HNO_FR, C_TABLE$REFER1),
                  c("HNO_FR","REFER1","X_C","MEDIAN","MAD","Z_MAD","P80","N_ROUND_HR","GATE","HR_DEL")]
ord_CA <- CANDIDATE[order(CANDIDATE$HA22_G, CANDIDATE$HNO_FR, CANDIDATE$REFER1),
                    c("HA22_G","HNO_FR","REFER1","X_R","MEDIAN_AA","MAD_AA","Z_MAD","P80_AA","N_ROUND","GATE","CANDIDATE")]
ord_HR <- HR_RESULT[order(HR_RESULT$HNO_FR, HR_RESULT$REFER1),
                    c("HNO_FR","REFER1","X_C","MEDIAN","MAD","Z_MAD","P80","N_ROUND_HR","GATE","HR_DEL")]

# -------------------- 엑셀 저장 --------------------
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
stamp <- format(Sys.time(), "%y%m%d%H%M")
xlsx_path <- file.path(out_dir, sprintf("PR1_%s.xlsx", stamp))
wb <- openxlsx::createWorkbook()
openxlsx::addWorksheet(wb, "RAW");         openxlsx::writeData(wb, "RAW", RAW)
openxlsx::addWorksheet(wb, "C_TABLE");     openxlsx::writeData(wb, "C_TABLE", ord_CT)
openxlsx::addWorksheet(wb, "POOR_LOCAL");  openxlsx::writeData(wb, "POOR_LOCAL", ord_AA)
openxlsx::addWorksheet(wb, "CANDIDATE");   openxlsx::writeData(wb, "CANDIDATE", ord_CA)
openxlsx::addWorksheet(wb, "HR_RESULT");   openxlsx::writeData(wb, "HR_RESULT", ord_HR)

PARAMS <- data.frame(
  INPUT=in_path, OUTPUT_EXCEL=xlsx_path, OUTPUT_POOR=poor_out,
  Z_CUT=z_cut, EPS=eps, N_MIN=n_min,
  ROWS_IN=ROWCOUNT_IN,
  AA_KEYS=nrow(ord_AA), HR_KEYS=nrow(ord_CT),
  stringsAsFactors = FALSE
)
openxlsx::addWorksheet(wb, "PARAMS"); openxlsx::writeData(wb, "PARAMS", PARAMS)

SUMMARY <- data.frame(
  AA_POOR_LOCAL_1 = sum(ord_AA$POOR_LOCAL==1, na.rm=TRUE),
  HR_DEL_1        = sum(ord_CT$HR_DEL==1, na.rm=TRUE),
  AA_GATE_POS     = sum(ord_AA$GATE!="", na.rm=TRUE),
  HR_GATE_POS     = sum(ord_CT$GATE!="",  na.rm=TRUE),
  AA_N_GE_3       = sum(ord_AA$N_ROUND>=n_min, na.rm=TRUE),
  HR_N_GE_3       = sum(ord_CT$N_ROUND_HR>=n_min, na.rm=TRUE)
)
openxlsx::addWorksheet(wb, "SUMMARY"); openxlsx::writeData(wb, "SUMMARY", SUMMARY)
openxlsx::saveWorkbook(wb, xlsx_path, overwrite = TRUE)

# -------------------- POOR 출력(DBF: 레코드(RECCO) 단위) --------------------
# 원자료에 플래그 병합
RAW_FLAGS <- merge(
  RAW[, c("RECCO","HA22_G","HNO_FR","REFER1","HA_22ND","RANO_FR")],
  ord_AA[, c("HA22_G","HNO_FR","REFER1","POOR_LOCAL")],
  by=c("HA22_G","HNO_FR","REFER1"),
  all.x=TRUE, sort=FALSE
)
RAW_FLAGS <- merge(
  RAW_FLAGS,
  ord_CT[, c("HNO_FR","REFER1","HR_DEL")],
  by=c("HNO_FR","REFER1"),
  all.x=TRUE, sort=FALSE
)

# DBF 저장 (요청 컬럼만)
if (isTRUE(write_dbf)) {
  if (!requireNamespace("foreign", quietly = TRUE)) {
    stop("DBF 저장 실패: 'foreign' 패키지 미설치. install.packages('foreign') 실행 후 재시도하세요.")
  }
  POOR_DBF <- RAW_FLAGS[, c("RANO_FR","HNO_FR","RECCO","HA_22ND","POOR_LOCAL","HR_DEL")]
  dbf_path <- {
    p <- poor_out
    if (!grepl("\\.dbf$", tolower(p))) sub("\\.[a-zA-Z0-9]+$", ".dbf", p) else p
  }
  foreign::write.dbf(POOR_DBF, dbf_path, factor2char = TRUE)
  message("[완료] POOR(DBF): ", dbf_path)
}

message("[완료] Excel: ", xlsx_path)