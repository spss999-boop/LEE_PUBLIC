
# ===================== # ha_ud.r =====================
suppressWarnings({ options(stringsAsFactors = FALSE) })

# ===================== CONFIG =====================
BASE         <- "C:/HK"
INDBF        <- file.path(BASE, "RA3_RANO1.DBF")        # 입/출력 DBF
OUT_XLSX_DIR <- file.path(BASE, "OUT")
OUT_XLSX     <- file.path(OUT_XLSX_DIR, "RA3_RANO1_OUT.xlsx")
SPSS_DIR     <- file.path(BASE, "SPSS")                 # vr<SPAN>.xlsx 저장 폴더
Z_CUT        <- 2.5
EPS          <- 1e-12
YCOL_INTRA   <- "POWER"   # 라운드 내부(행 단위) 판정 컬럼
YCOL_FOR_EB  <- "POWER"   # EB 및 POWER_ADJ 산출 컬럼

# ===================== DEPS =====================
need_pkg <- c("data.table", "foreign", "openxlsx")
for (p in need_pkg) {
  if (!requireNamespace(p, quietly = TRUE)) stop("need '", p, "'")
}
library(data.table)
library(foreign)
library(openxlsx)

if (!dir.exists(OUT_XLSX_DIR)) dir.create(OUT_XLSX_DIR, recursive = TRUE, showWarnings = FALSE)
if (!dir.exists(SPSS_DIR))     dir.create(SPSS_DIR,     recursive = TRUE, showWarnings = FALSE)

# ===================== HELPERS =====================
ucol <- function(df) { names(df) <- toupper(trimws(names(df))); df }

snum <- function(v) { suppressWarnings(as.numeric(v)) }

mad_s <- function(x) {
  x <- as.numeric(x); x <- x[is.finite(x)]
  if (!length(x)) return(NA_real_)
  val <- stats::mad(x, constant = 1.4826, na.rm = TRUE)
  if (!is.finite(val) || val == 0) NA_real_ else val
}

p80 <- function(x) {
  x <- as.numeric(x); x <- x[is.finite(x)]
  if (!length(x)) return(NA_real_)
  as.numeric(stats::quantile(x, probs = 0.80, type = 7, na.rm = TRUE))
}

p90 <- function(x) {
  x <- as.numeric(x); x <- x[is.finite(x)]
  if (!length(x)) return(NA_real_)
  as.numeric(stats::quantile(x, probs = 0.90, type = 7, na.rm = TRUE))
}

safe_write_dbf <- function(df, path) {
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  df2 <- copy(df)
  is_num <- vapply(df2, is.numeric,  logical(1))
  is_chr <- vapply(df2, is.character, logical(1))
  if (any(is_num)) for (nm in names(df2)[is_num]) { x <- df2[[nm]]; x[!is.finite(x)] <- 0;  df2[[nm]] <- as.numeric(x) }
  if (any(is_chr)) for (nm in names(df2)[is_chr]) { x <- df2[[nm]]; x[is.na(x)]     <- ""; df2[[nm]] <- as.character(x) }
  tmp <- file.path(dirname(path), paste0("tmp_", basename(path)))
  foreign::write.dbf(as.data.frame(df2), tmp)
  if (file.exists(path)) unlink(path)
  file.rename(tmp, path)
}

# ===================== LOAD =====================
if (!file.exists(INDBF)) stop("Input DBF not found: ", INDBF)
d0_df <- ucol(as.data.frame(foreign::read.dbf(INDBF, as.is = TRUE)))

# --- 안정 조인키: RECNO (원본 로드 직후 부여, 세션 내내 불변) ---
if (!"RECNO" %in% names(d0_df)) d0_df$RECNO <- seq_len(nrow(d0_df))
# RA3_RANO1.DBF에도 RECCO 필드를 두어 POOR.DBF의 RECCO와 직접 조인 가능하게
if (!"RECCO" %in% names(d0_df)) d0_df$RECCO <- as.integer(NA_integer_)

# 필수 컬럼 점검
req <- c("HNO_FR", "REFER1", "HA_22ND", YCOL_INTRA, YCOL_FOR_EB)
miss <- setdiff(req, names(d0_df))
if (length(miss)) stop("Missing cols: ", paste(miss, collapse = ", "))

# cast
d0_df$HNO_FR   <- as.character(d0_df$HNO_FR)
d0_df$REFER1   <- as.integer(snum(d0_df$REFER1))
d0_df$HA_22ND  <- as.character(d0_df$HA_22ND)
if (!"RANO_FR" %in% names(d0_df)) d0_df$RANO_FR <- ""
d0_df[[YCOL_INTRA]]  <- snum(d0_df[[YCOL_INTRA]])
d0_df[[YCOL_FOR_EB]] <- snum(d0_df[[YCOL_FOR_EB]])

DT <- as.data.table(d0_df)

# SPAN (for snapshot naming)
SPAN <- {
  s <- unique(as.character(DT$RANO_FR))
  s <- s[grepl("^[0-9]{8}$", s)]
  if (length(s) >= 1) s[1] else "NA"
}
TEST_XLSX <- file.path(SPSS_DIR, sprintf("vr%s.xlsx", SPAN))  # 스냅샷: vr<SPAN>.xlsx

# ===================== EB 5-bin + POWER_ADJ =====================
DT[, .orig := .I]
DTcalc <- DT[!is.na(REFER1) & REFER1 >= 1 & is.finite(get(YCOL_FOR_EB))]

EBG <- DTcalc[, .(
  N_G    = sum(is.finite(get(YCOL_FOR_EB))),
  MEAN_G = mean(get(YCOL_FOR_EB), na.rm = TRUE),
  SD_G   = stats::sd(get(YCOL_FOR_EB), na.rm = TRUE)
), by = .(HA_22ND)]
EBG[!is.finite(SD_G), SD_G := 0]
EBG[, `:=`(DF_G = pmax(N_G - 1L, 0L), SSW_G = pmax(N_G - 1L, 0L) * (SD_G^2))]

MU   <- stats::weighted.mean(EBG$MEAN_G, w = pmax(EBG$N_G, 1), na.rm = TRUE)
S2W  <- sum(EBG$SSW_G, na.rm = TRUE) / max(sum(EBG$DF_G, na.rm = TRUE), 1)
VM   <- stats::var(EBG$MEAN_G[is.finite(EBG$MEAN_G)], na.rm = TRUE)
AINV <- mean(1 / pmax(EBG$N_G, 1), na.rm = TRUE)
TAU2 <- max(VM - S2W * AINV, 0)

EBG[, W_G := if ((S2W + TAU2) == 0) 0 else (N_G * TAU2) / (N_G * TAU2 + S2W + 1e-12)]
EBG[, EB_MEAN := W_G * MEAN_G + (1 - W_G) * MU]

brks <- unique(stats::quantile(EBG$EB_MEAN, probs = seq(0, 1, length.out = 6), na.rm = TRUE, type = 7))
if (length(brks) < 6 || any(!is.finite(brks))) {
  EBG[, RANK := frank(EB_MEAN, ties.method = "average", na.last = "keep")]
  EBG[, HA22_G := as.integer(ceiling(5 * RANK / max(RANK, na.rm = TRUE)))]
  EBG[, `:=`(RANK = NULL)]
} else {
  EBG[, HA22_G := as.integer(cut(EB_MEAN, breaks = brks, include.lowest = TRUE, right = TRUE))]
}
EBG[, HA22_G5 := HA22_G]

DT <- DT[EBG[, .(HA_22ND, HA22_G5, HA22_G, EB_MEAN)], on = .(HA_22ND)]
setorder(DT, .orig); DT[, .orig := NULL]

DT[, POWER_ADJ := fifelse(is.finite(get(YCOL_FOR_EB)) & is.finite(EB_MEAN),
                          get(YCOL_FOR_EB) - EB_MEAN, NA_real_)]

# ===================== INTRA (행 단위; 보고용) =====================
DT[, `:=`(
  HNO_FR_I = suppressWarnings(as.integer(HNO_FR)),
  HA22_G_I = suppressWarnings(as.integer(HA22_G)),
  REFER1_I = suppressWarnings(as.integer(REFER1))
)]
setorder(DT, HNO_FR_I, HA22_G_I, REFER1_I)

blk_stats <- DT[
  !is.na(REFER1) & is.finite(get(YCOL_INTRA)) & is.finite(HA22_G_I),
  .(
    X_R   = stats::median(get(YCOL_INTRA), na.rm = TRUE),
    MAD_R = mad_s(get(YCOL_INTRA)),
    P80_R = p80(get(YCOL_INTRA)),
    P90_R = p90(get(YCOL_INTRA)),
    N_RG  = .N
  ),
  by = .(HNO_FR, HA22_G, REFER1)
]

DT[blk_stats, `:=`(X_R = i.X_R, MAD_R = i.MAD_R, P80_R = i.P80_R, P90_R = i.P90_R, N_RG = i.N_RG),
   on = .(HNO_FR, HA22_G, REFER1)]

HA22_PER_ROUND <- DT[!is.na(REFER1) & is.finite(HA22_G_I),
                     .(N_HA22_ROUND = uniqueN(HA22_G, na.rm = TRUE)), by = .(REFER1)]
DT[HA22_PER_ROUND, N_HA22_ROUND := i.N_HA22_ROUND, on = .(REFER1)]

DT[, Z_R := fifelse(is.finite(MAD_R) & MAD_R > 0, (get(YCOL_INTRA) - X_R) / MAD_R, NA_real_)]
DT[, GATE_R := fifelse(is.finite(Z_R) & Z_R >= Z_CUT & is.finite(get(YCOL_INTRA)) & is.finite(P90_R) & (get(YCOL_INTRA) > P90_R + EPS), "Z|P90",
                       fifelse(is.finite(Z_R) & Z_R >= Z_CUT, "Z",
                               fifelse(is.finite(get(YCOL_INTRA)) & is.finite(P90_R) & (get(YCOL_INTRA) > P90_R + EPS), "P90", "")))]
DT[, DEL_INTR := as.integer(GATE_R != "")]
DT[, SKIP_INTR := as.integer((!is.na(N_RG) & N_RG < 3L) | (!is.na(N_HA22_ROUND) & N_HA22_ROUND <= 2L))]
DT[SKIP_INTR == 1L, `:=`(Z_R = NA_real_, GATE_R = "", DEL_INTR = 0L)]  # 스킵이면 인트라 판정 무효(보고만)

# ===================== R_X (보고용; 평균, INTRA삭제 제외 기준) =====================
RBLK <- DT[
  !is.na(REFER1) & is.finite(get(YCOL_INTRA)) & (DEL_INTR == 0L),
  .(
    R_X = mean(get(YCOL_INTRA), na.rm = TRUE),
    N_RK = .N
  ),
  by = .(HNO_FR, HA22_G, REFER1)
]
DT[RBLK, `:=`(R_X = i.R_X, N_RK = i.N_RK), on = .(HNO_FR, HA22_G, REFER1)]

# ===================== INTER (POOR_LOCAL) =====================
# 모든 (h,g,r) 조합을 보존하고 xr(R_X)을 left join (NA 허용)
BASE_TAB <- unique(DT[, .(HNO_FR, HA22_G, REFER1)])
XR   <- RBLK[, .(HNO_FR, HA22_G, REFER1, xr = R_X, n_rk = N_RK)]
POOR_LOCAL <- merge(BASE_TAB, XR, by = c("HNO_FR", "HA22_G", "REFER1"), all.x = TRUE)

# (h,g)별 xr 분포 통계
AA <- POOR_LOCAL[, .(
  ma  = stats::median(xr, na.rm = TRUE),
  da  = mad_s(xr),
  p8a = p80(xr),
  na  = .N
), by = .(HNO_FR, HA22_G)]
POOR_LOCAL <- AA[POOR_LOCAL, on = .(HNO_FR, HA22_G)]

POOR_LOCAL[, za := fifelse(is.finite(da) & da > 0, (xr - ma) / da, NA_real_)]
POOR_LOCAL[, ga := fifelse(is.finite(za) & za >= Z_CUT & is.finite(xr) & is.finite(p8a) & (xr > p8a + EPS), "Z|P80",
                           fifelse(is.finite(za) & za >= Z_CUT, "Z",
                                   fifelse(is.finite(xr) & is.finite(p8a) & (xr > p8a + EPS), "P80", "")))]
POOR_LOCAL[, `:=`(di = as.integer((na >= 3) & (ga != "")),
                  si = as.integer(na < 3))]
# del_pw: di==1 제외한 xr 평균 (h,g)
POOR_LOCAL[, del_pw := mean(xr[di == 0L], na.rm = TRUE), by = .(HNO_FR, HA22_G)]

# ===================== DEL_INTER / DEL_ROU / DEL_ANY =====================
POOR_LOCAL[, `:=`(HNO_FR = as.integer(HNO_FR), REFER1 = as.integer(REFER1), HA22_G = as.integer(HA22_G))]
DT[,        `:=`(HNO_FR = as.integer(HNO_FR), REFER1 = as.integer(REFER1), HA22_G = as.integer(HA22_G))]

DEL_INTER <- POOR_LOCAL[, .(HNO_FR, HA22_G, REFER1, DEL_INTER = as.integer(di == 1L))]

PER_ROUND_DI <- POOR_LOCAL[, .(
  N_DI = sum(di == 1L, na.rm = TRUE),
  N_G  = uniqueN(HA22_G, na.rm = TRUE)
), by = .(HNO_FR, REFER1)]
# 존재하는 g 전부가 di==1이면 라운드 전체 삭제
DEL_ROU_TAB <- PER_ROUND_DI[, .(HNO_FR, REFER1, DEL_ROU = as.integer(N_DI == N_G & N_G >= 1L))]

DT2 <- copy(DT)
DT2[DEL_INTER,   DEL_INTER := i.DEL_INTER, on = .(HNO_FR, HA22_G, REFER1)]
DT2[is.na(DEL_INTER), DEL_INTER := 0L]
DT2[DEL_ROU_TAB, DEL_ROU   := i.DEL_ROU,   on = .(HNO_FR, REFER1)]
DT2[is.na(DEL_ROU), DEL_ROU := 0L]
DT2[, DEL_ANY := as.integer((DEL_INTER == 1L) | (DEL_ROU == 1L))]

# ===================== C TABLE (DEL_ANY==0 + POOR1 제외) =====================
D_keep <- DT2[!is.na(REFER1) & (DEL_ANY == 0L)]
POOR1  <- PER_ROUND_DI[N_DI == N_G & N_G >= 1L][order(HNO_FR, REFER1)]
EXCL   <- POOR1[, .(HNO_FR, REFER1)]
if (nrow(EXCL)) D_keep <- D_keep[!EXCL, on = .(HNO_FR, REFER1)]

C_T <- D_keep[, .(X_C = mean(POWER_ADJ, na.rm = TRUE)), by = .(HNO_FR, REFER1)]
CT_stats <- C_T[, .(
  MEDIAN_C   = stats::median(X_C, na.rm = TRUE),
  MAD_C      = mad_s(X_C),
  P80_C      = p80(X_C),
  N_ROUND_HR = .N
), by = .(HNO_FR)]

C_TABLE <- C_T[CT_stats, on = .(HNO_FR)]   # LONG
C_TABLE[, Z_C := fifelse(is.finite(MAD_C) & MAD_C > 0, (X_C - MEDIAN_C) / MAD_C, NA_real_)]
C_TABLE[, HR_GATE_Z := as.integer(is.finite(Z_C) & Z_C >= Z_CUT & N_ROUND_HR >= 3)]
C_TABLE[, HR_GATE_P := as.integer(is.finite(X_C) & is.finite(P80_C) & (X_C > P80_C + EPS) & N_ROUND_HR >= 3)]
C_TABLE[, HR_DEL    := as.integer(N_ROUND_HR >= 3 & (HR_GATE_Z == 1L | HR_GATE_P == 1L))]

# 피벗
ALL_H <- sort(unique(DT2$HNO_FR))
REFS  <- sort(unique(C_TABLE$REFER1))
CT_PIVOT_DT <- if (length(REFS)) dcast(C_TABLE, HNO_FR ~ REFER1, value.var = "X_C") else data.table(HNO_FR = ALL_H)
HR_PIVOT_DT <- if (length(REFS)) dcast(C_TABLE[, .(HNO_FR, REFER1, HR_DEL)], HNO_FR ~ REFER1, value.var = "HR_DEL") else data.table(HNO_FR = ALL_H)
CT_PIVOT_DT <- data.table(HNO_FR = ALL_H)[CT_PIVOT_DT, on = .(HNO_FR)]
HR_PIVOT_DT <- data.table(HNO_FR = ALL_H)[HR_PIVOT_DT, on = .(HNO_FR)]

# ===================== EXCEL =====================
wb <- openxlsx::createWorkbook()

# POOR1 sheet (existence-based all-di==1)
openxlsx::addWorksheet(wb, "POOR1")
if (nrow(POOR1) > 0) {
  openxlsx::writeData(wb, "POOR1", as.data.frame(POOR1))
} else {
  openxlsx::writeData(wb, "POOR1", data.frame(message = "no rounds with all existing g having di==1"))
}

# RAW
RAW <- DT2[, .(
  h    = HNO_FR,
  g    = as.integer(HA22_G),
  r    = REFER1,
  ran  = RANO_FR,
  pow  = get("POWER"),
  padj = POWER_ADJ,
  x_r  = X_R,
  mad_r = MAD_R,
  p80_r = P80_R,
  p90_r = P90_R,
  z     = Z_R,
  gr    = GATE_R,
  din   = DEL_INTR,
  sin   = SKIP_INTR,
  di    = DEL_INTER,
  drou  = DEL_ROU,
  dany  = DEL_ANY,
  r_x   = R_X,
  n_rk  = N_RK
)]
RAW$RECCO <- DT2$RECNO # RECCO(RECNO)도 반드시 붙여줌

openxlsx::addWorksheet(wb, "RAW")
if (all(c("h","g","r") %in% names(RAW))) data.table::setnames(RAW, c("h","g","r"), c("HNO_FR","HA22_G","REFER1"))
openxlsx::writeData(wb, "RAW", as.data.frame(RAW))
.raw_start_col <- ncol(RAW) + 2
raw_summary_top <- data.frame(
  Field = c("din","sin","di","drou","dany"),
  Count = c(sum(RAW$din == 1, na.rm = TRUE),
            sum(RAW$sin == 1, na.rm = TRUE),
            sum(RAW$di  == 1, na.rm = TRUE),
            sum(RAW$drou== 1, na.rm = TRUE),
            sum(RAW$dany== 1, na.rm = TRUE))
)
openxlsx::writeData(wb, "RAW", raw_summary_top, startCol = .raw_start_col, startRow = 1)

# POOR_LOCAL
POOR_LOCAL2 <- copy(POOR_LOCAL)
POOR_LOCAL2[DEL_ROU_TAB, DEL_ROU := i.DEL_ROU, on = .(HNO_FR, REFER1)]
POOR_LOCAL2[is.na(DEL_ROU), DEL_ROU := 0L]
POOR_LOCAL_OUT <- POOR_LOCAL2[, .(
  h = HNO_FR, g = as.integer(HA22_G), r = REFER1,
  xr, del_pw, ma, da, p8a, za, na, ga, di, si, drou = DEL_ROU
)]
openxlsx::addWorksheet(wb, "POOR_LOCAL")
if (all(c("h","g","r") %in% names(POOR_LOCAL_OUT))) data.table::setnames(POOR_LOCAL_OUT, c("h","g","r"), c("HNO_FR","HA22_G","REFER1"))
openxlsx::writeData(wb, "POOR_LOCAL", as.data.frame(POOR_LOCAL_OUT))
.pl_start_col <- ncol(POOR_LOCAL_OUT) + 2
pl_summary_top <- data.frame(
  Field = c("di","si","drou"),
  Count = c(sum(POOR_LOCAL_OUT$di  == 1, na.rm = TRUE),
            sum(POOR_LOCAL_OUT$si  == 1, na.rm = TRUE),
            sum(POOR_LOCAL_OUT$drou== 1, na.rm = TRUE))
)
openxlsx::writeData(wb, "POOR_LOCAL", pl_summary_top, startCol = .pl_start_col, startRow = 1)

# ===================== C TABLE LONG =====================
openxlsx::addWorksheet(wb, "C TABLE LONG")
C_TABLE_LONG <- as.data.frame(C_TABLE)
RAW_DF <- as.data.frame(RAW)  # RAW는 이미 데이터프레임 형태

# 각 행에 대해 RAW에서 같은 HNO_FR, REFER1 값 가진 행의 개수 count로 입력
get_count <- function(h, r) {
  sum(RAW_DF$HNO_FR == h & RAW_DF$REFER1 == r, na.rm = TRUE)
}
C_TABLE_LONG$count <- mapply(get_count, C_TABLE_LONG$HNO_FR, C_TABLE_LONG$REFER1)

openxlsx::writeData(wb, "C TABLE LONG", C_TABLE_LONG)

# 연한 하늘색 스타일 정의
skyblue <- openxlsx::createStyle(fgFill = "#CCFFFF")

# HR_DEL==1인 행(row)만 색칠 (헤더 제외, 데이터 행은 +1)
rows_to_color <- which(C_TABLE_LONG$HR_DEL == 1) + 1
if (length(rows_to_color) > 0) {
  openxlsx::addStyle(
    wb, sheet = "C TABLE LONG", style = skyblue,
    rows = rows_to_color, cols = 1:ncol(C_TABLE_LONG),
    gridExpand = TRUE, stack = TRUE
  )
}

# C TABLE (피벗)
openxlsx::addWorksheet(wb, "C TABLE")
openxlsx::writeData(wb, "C TABLE", as.data.frame(CT_PIVOT_DT))
if (nrow(CT_PIVOT_DT) > 0 && ncol(CT_PIVOT_DT) > 1) {
  pink <- openxlsx::createStyle(fgFill = "#FFCCE5")
  r_vals_chr <- colnames(CT_PIVOT_DT)[-1]              # 열 이름(문자)
  h_vals_chr <- as.character(CT_PIVOT_DT[[1]])         # 행 키(문자)
  
  # 1) HR_DEL 분홍
  if (ncol(HR_PIVOT_DT) > 1) {
    hr_mat <- as.matrix(HR_PIVOT_DT[, -1, with = FALSE])
    idx <- which(hr_mat == 1L, arr.ind = TRUE)
    if (nrow(idx) > 0) {
      rows <- idx[, "row"] + 1   # 데이터 행 → 시트 행(+헤더)
      cols <- idx[, "col"] + 1   # 데이터 열 → 시트 열(+헤더)
      for (k in seq_len(nrow(idx))) {
        openxlsx::addStyle(wb, "C TABLE", style = pink, rows = rows[k], cols = cols[k], gridExpand = FALSE, stack = TRUE)
      }
    }
  }
  
  # 2) POOR1 분홍(라운드 전체 제외 표시)
  if (nrow(EXCL) > 0) {
    for (k in seq_len(nrow(EXCL))) {
      hh <- as.character(EXCL$HNO_FR[k]); rr <- as.character(EXCL$REFER1[k])
      row_i <- match(hh, h_vals_chr)
      col_j <- match(rr, r_vals_chr)
      if (is.na(row_i) || is.na(col_j)) next
      openxlsx::addStyle(wb, "C TABLE", style = pink, rows = row_i + 1L, cols = col_j + 1L, gridExpand = FALSE, stack = TRUE)
    }
  }
}

# ---- C TABLE 시트 밑에 HNO_FR별 삭제 통계 추가 (DEL_ROU==1 또는 DEL_ANY==1 기준) -----
CTABLE_DEL_STATS <- DT2[, .(
  deleted_cells = sum(DEL_ROU == 1L | DEL_ANY == 1L, na.rm = TRUE),
  total_cells   = .N,
  deleted_ratio = ifelse(.N > 0, sum(DEL_ROU == 1L | DEL_ANY == 1L, na.rm = TRUE) / .N, NA_real_)
), by = .(HNO_FR)][order(HNO_FR)]
ctable_nrow <- nrow(CT_PIVOT_DT) + 3
openxlsx::writeData(
  wb, "C TABLE", as.data.frame(CTABLE_DEL_STATS),
  startCol = 1, startRow = ctable_nrow, withFilter = FALSE
)

# ===================== DBF SAVE (원본 + 파생 통합 저장) =====================
# 1) 파생 컬럼만 추출 (RECNO로 조인)
DERIVED_DT <- DT2[, .(
  RECNO,
  HA22_G         = as.integer(HA22_G),
  POWER_ADJ,
  X_R, MAD_R, P80_R, P90_R, Z_R, GATE_R,
  DEL_INTR, SKIP_INTR, N_RG,
  R_X, N_RK,
  DEL_INTER, DEL_ROU, DEL_ANY
)]

# 2) 원본 d0_df(모든 필드 보존)에 파생만 좌조인
OUT_DBF_DT <- as.data.table(d0_df)
OUT_DBF_DT[DERIVED_DT, `:=`(
  HA22_G    = i.HA22_G,
  POWER_ADJ = i.POWER_ADJ,
  X_R       = i.X_R,
  MAD_R     = i.MAD_R,
  P80_R     = i.P80_R,
  P90_R     = i.P90_R,
  Z_R       = i.Z_R,
  GATE_R    = i.GATE_R,
  DEL_INTR  = i.DEL_INTR,
  SKIP_INTR = i.SKIP_INTR,
  N_RG      = i.N_RG,
  R_X       = i.R_X,
  N_RK      = i.N_RK,
  DEL_INTER = i.DEL_INTER,
  DEL_ROU   = i.DEL_ROU,
  DEL_ANY   = i.DEL_ANY
), on = .(RECNO)]

# 3) 조인키 컬럼 동기화: RECCO = RECNO (라벨은 조인 편의용)
OUT_DBF_DT[, RECCO := as.integer(RECNO)]

# 4) 원본 순서로 저장
setorder(OUT_DBF_DT, RECNO)
safe_write_dbf(OUT_DBF_DT, INDBF)

# ===================== POOR.DBF 저장 (RECCO=RECNO로 변경, 소수점 없이) =====================
# ===================== POOR.DBF 저장 (RECCO=RECNO로 변경, 소수점 없이) =====================
POOR_PATH <- file.path(BASE, "POOR.DBF")

# 1. HR_DEL==1 (분홍/하늘색)인 (HNO_FR, REFER1) 조합 추출
HRDEL1_COMBO <- C_TABLE[HR_DEL == 1, .(HNO_FR, REFER1)]    # C_TABLE의 하늘색(hr_del==1)
POOR1_COMBO  <- POOR1[, .(HNO_FR, REFER1)]                # POOR1의 (라운드 전체 제외) 조합

# 2. 하늘색 + 분홍(POOR1) 조합 합집합
ALL_HRDEL_COMBO <- unique(rbind(HRDEL1_COMBO, POOR1_COMBO))

# 3. poor_local: RAW$din==1이면 1, 아니면 0 (RECCO=RECNO로 일치)
poor_local_vec <- ifelse(RAW$din == 1, 1L, 0L)

# 4. hr_del: (HNO_FR, REFER1)이 하늘색+분홍(POOR1) 합집합에 있으면 1, 아니면 0
hr_del_vec <- mapply(
  function(h, r) as.integer(any(ALL_HRDEL_COMBO$HNO_FR == h & ALL_HRDEL_COMBO$REFER1 == r)),
  RAW$HNO_FR, RAW$REFER1)

POOR_DBF_DT <- data.table(
  RECCO      = RAW$RECCO,
  RANO_FR    = RAW$ran,
  HNO_FR     = RAW$HNO_FR,
  HA_22ND    = DT2$HA_22ND,
  REFER1     = RAW$REFER1,
  POOR_LOCAL = poor_local_vec,
  HR_DEL     = hr_del_vec,
  DEL_ANY    = poor_local_vec,     # DEL_ANY=poor_local와 동일하게 세팅
  DEL_ROU    = DT2$DEL_ROU
)

# 5. 모든 컬럼을 정수 또는 문자로 소수점 없이 저장
for (col in names(POOR_DBF_DT)) {
  if (is.numeric(POOR_DBF_DT[[col]]) && !is.integer(POOR_DBF_DT[[col]])) {
    POOR_DBF_DT[[col]] <- as.integer(POOR_DBF_DT[[col]])
  }
  if (is.factor(POOR_DBF_DT[[col]])) {
    POOR_DBF_DT[[col]] <- as.character(POOR_DBF_DT[[col]])
  }
}

safe_write_dbf(POOR_DBF_DT, POOR_PATH)
# ===================== SAVE EXCEL =====================
openxlsx::saveWorkbook(wb, OUT_XLSX, overwrite = TRUE)
openxlsx::saveWorkbook(wb, TEST_XLSX, overwrite = TRUE)

cat(
  "[HA_UD] DONE.\n",
  "DBF updated (all original fields + derived): ", INDBF, "\n",
  "POOR.DBF saved (RECCO=RECNO, 소수점 없이): ", POOR_PATH, "\n",
  "Excel saved: ", OUT_XLSX, " and ", TEST_XLSX, "\n",
  sep = ""
) 