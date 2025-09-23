# C:/HK/R/POOR_EHIS.R — 부진경주(GATE) 식별 및 POOR.DBF 생성(POWER 필드 포함)
# - 라운드별 대표값(중앙값, X_HR) 산출
# - GATE 공식(Z_MAD, P80)로 부진경주 삭제 판정(POOR_LOCAL, HR_DEL)
# - 동일 라운드 HA22_G가 3개 이하면 삭제 안 함
# - POOR.DBF, POOR_check.csv 산출

suppressWarnings({ options(stringsAsFactors = FALSE) })

BASE    <- "C:/HK"
INDBF   <- file.path(BASE, "RA3_RANO1.DBF")
POORDBF <- file.path(BASE, "POOR.DBF")
CHECKCSV<- file.path(BASE, "POOR_check.csv")

if (!requireNamespace("foreign", quietly=TRUE)) stop("need 'foreign'")
library(foreign)

safe_write_dbf <- function(df, path){
  dir.create(dirname(path), showWarnings=FALSE, recursive=TRUE)
  df2 <- df
  is_num <- vapply(df2, is.numeric,  logical(1))
  is_chr <- vapply(df2, is.character, logical(1))
  if (any(is_num)) for (nm in names(df2)[is_num]) { x<-df2[[nm]]; x[is.na(x)]<-0;  df2[[nm]]<-as.numeric(x) }
  if (any(is_chr)) for (nm in names(df2)[is_chr]) { x<-df2[[nm]]; x[is.na(x)]<-""; df2[[nm]]<-as.character(x) }
  tmp <- file.path(dirname(path), paste0("tmp_", basename(path)))
  foreign::write.dbf(df2, tmp); if (file.exists(path)) unlink(path); file.rename(tmp, path)
}

# 1. 데이터 읽기 및 컬럼명 체크
d0 <- as.data.frame(read.dbf(INDBF, as.is=TRUE))
names(d0) <- toupper(trimws(names(d0)))
need <- c("HNO_FR", "REFER1", "HA22_G", "POWER", "RECCO")
miss <- setdiff(need, names(d0))
if (length(miss)) stop("Missing cols: ", paste(miss, collapse=", "))
d0$HNO_FR   <- as.character(d0$HNO_FR)
d0$REFER1   <- as.integer(d0$REFER1)
d0$HA22_G   <- as.character(d0$HA22_G)
d0$POWER    <- as.numeric(d0$POWER)
d0$RECCO    <- as.character(d0$RECCO)

# 2. 라운드별 대표값(중앙값) 산출
med_tab <- aggregate(POWER ~ HNO_FR + REFER1, d0, function(x) median(x, na.rm=TRUE))
names(med_tab)[3] <- "X_HR"

# 3. 라운드별 HA22_G 개수 계산
g_count <- aggregate(HA22_G ~ HNO_FR + REFER1, d0, function(x) length(unique(x)))
names(g_count)[3] <- "N_G_IN_ROUND"

# 4. HA22_G별 라운드 개수(N_ROUND) 계산
n_round <- aggregate(REFER1 ~ HA22_G, d0, function(x) length(unique(x)))
names(n_round)[2] <- "N_ROUND"

# 5. per-horse(마) 통계 & GATE 이상치
by_h <- split(med_tab$X_HR, med_tab$HNO_FR)
HSTAT <- do.call(rbind, lapply(names(by_h), function(h){
  x <- by_h[[h]]
  m <- median(x, na.rm=TRUE)
  mad0 <- median(abs(x - m), na.rm=TRUE)
  mad <- if (is.finite(mad0) && !is.na(mad0)) mad0*1.4826 else NA_real_
  p80 <- as.numeric(quantile(x, 0.80, na.rm=TRUE, type=7))
  data.frame(HNO_FR=h, MEDIAN=m, MAD=mad, P80=p80, N_ROUND_HR=sum(is.finite(x)))
}))
HR0 <- merge(med_tab, HSTAT, by="HNO_FR", all.x=TRUE, sort=FALSE)
HR0$Z_MAD <- ifelse(is.finite(HR0$MAD) & HR0$MAD>0, (HR0$X_HR - HR0$MEDIAN)/HR0$MAD, NA_real_)

# 6. GATE 공식 적용 (Z_MAD >= 2.5 또는 X_HR > P80)
Z_CUT <- 2.5
HR0$GATE_HIT_Z   <- !is.na(HR0$Z_MAD) & (HR0$Z_MAD >= Z_CUT)
HR0$GATE_HIT_P80 <- is.finite(HR0$X_HR) & is.finite(HR0$P80) & (HR0$X_HR > HR0$P80 + 1e-12)
HR0$GATE_HIT     <- HR0$GATE_HIT_Z | HR0$GATE_HIT_P80
HR0$GATE_REASON  <- ifelse(HR0$GATE_HIT_Z & HR0$GATE_HIT_P80, "Z|P80",
                           ifelse(HR0$GATE_HIT_Z, "Z",
                                  ifelse(HR0$GATE_HIT_P80, "P80", "")))

# 7. 라운드/마/등급 결합 및 삭제 플래그
HR1 <- merge(HR0, g_count, by=c("HNO_FR", "REFER1"), all.x=TRUE, sort=FALSE)
HR1 <- merge(HR1, d0[,c("HNO_FR","REFER1","HA22_G","RECCO","POWER")], by=c("HNO_FR","REFER1"), all.x=TRUE, sort=FALSE)
HR1 <- merge(HR1, n_round, by="HA22_G", all.x=TRUE, sort=FALSE)

# 비교/삭제 조건: N_G_IN_ROUND ≤ 3 이면 무조건 존치, 그 외 GATE 공식
HR1$POOR_LOCAL <- ifelse(HR1$N_G_IN_ROUND <= 3, 0L, as.integer(HR1$GATE_HIT))
HR1$HR_DEL     <- HR1$POOR_LOCAL # HR_DEL도 같이 제공

# 8. RECCO별 중복 없이 1행만 추출(중복은 첫번째만)
poor <- HR1[!duplicated(HR1$RECCO), c(
  "RECCO","HNO_FR","REFER1","HA22_G","POWER",
  "POOR_LOCAL","HR_DEL","X_HR","Z_MAD","P80","GATE_REASON","N_G_IN_ROUND","N_ROUND"
)]

# 9. DBF, CSV로 저장
safe_write_dbf(poor, POORDBF)
write.csv(poor, CHECKCSV, row.names=FALSE, fileEncoding="UTF-8")

cat("[POOR_EHIS] DONE: ", POORDBF, " and ", CHECKCSV, "\n")

