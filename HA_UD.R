# C:/HK/R/HA_UD.R — 5-bin grading (HA22_G5, HA22_G=HA22_G5) + call POOR_EHIS.R
# - EB 평균, 5-bin 등급 생성(HA22_G5)
# - 결과를 RA3_RANO1.DBF로 갱신
# - POOR_EHIS.R을 자동 실행(GATE 공식/POOR.DBF/RA3_RANO.DBF 생성)

suppressWarnings({ options(stringsAsFactors = FALSE) })

BASE   <- "C:/HK"
INDBF  <- file.path(BASE, "RA3_RANO1.DBF")   # input/output
OUTDBF <- INDBF
YCOL   <- "POWER"                            # or POWER_ADJ
POOR_SCRIPT <- file.path(BASE, "R", "POOR_EHIS.R")
MAKE_POOR   <- TRUE                          # call POOR_EHIS.R after grading

if (!requireNamespace("foreign", quietly=TRUE)) stop("need 'foreign'")
library(foreign)

# Helper
ucol <- function(df){ names(df) <- toupper(trimws(names(df))); df }
snum <- function(v){ suppressWarnings(as.numeric(v)) }
safe_write_dbf <- function(df, path){
  df2 <- df
  is_num <- vapply(df2, is.numeric,  logical(1))
  is_chr <- vapply(df2, is.character, logical(1))
  if (any(is_num)) for (nm in names(df2)[is_num]) { x<-df2[[nm]]; x[is.na(x)]<-0;  df2[[nm]]<-as.numeric(x) }
  if (any(is_chr)) for (nm in names(df2)[is_chr]) { x<-df2[[nm]]; x[is.na(x)]<-""; df2[[nm]]<-as.character(x) }
  foreign::write.dbf(df2, path)
}

# Load
if (!file.exists(INDBF)) stop("Input DBF not found: ", INDBF)
d0 <- ucol(as.data.frame(foreign::read.dbf(INDBF, as.is=TRUE)))
need <- c("HA_22ND", YCOL)
mis  <- setdiff(need, names(d0)); if (length(mis)) stop("Missing cols: ", paste(mis, collapse=","))

d0$HA_22ND  <- as.character(d0$HA_22ND)
d0[[YCOL]]  <- snum(d0[[YCOL]])

# EB baseline (by HA_22ND)
agg <- aggregate(
  d0[[YCOL]] ~ d0$HA_22ND,
  FUN=function(x) c(N=sum(!is.na(x)), MEAN=mean(x,na.rm=TRUE), SD=stats::sd(x,na.rm=TRUE))
)
colnames(agg) <- c("HA_22ND","PACK")
A <- data.frame(HA_22ND=agg$HA_22ND,
                N_G=agg$PACK[,"N"], MEAN_G=agg$PACK[,"MEAN"], SD_G=agg$PACK[,"SD"], stringsAsFactors=FALSE)
A$SD_G[is.na(A$SD_G)] <- 0
A$DF_G  <- pmax(A$N_G - 1, 0); A$SSW_G <- A$DF_G * (A$SD_G^2)
TOT_DF <- sum(A$DF_G, na.rm=TRUE); S2W <- if (TOT_DF>0) sum(A$SSW_G,na.rm=TRUE)/TOT_DF else 0
MU <- if (sum(A$N_G,na.rm=TRUE)>0) sum(A$N_G*A$MEAN_G,na.rm=TRUE)/sum(A$N_G,na.rm=TRUE) else mean(A$MEAN_G,na.rm=TRUE)
VAR_MEANS <- stats::var(A$MEAN_G, na.rm=TRUE); AVG_INV_N <- mean(1/A$N_G[A$N_G>0])
TAU2 <- max(VAR_MEANS - S2W*AVG_INV_N, 0)
A$W_G <- if (S2W+TAU2==0) 0 else (A$N_G*TAU2)/(A$N_G*TAU2 + S2W + 1e-12)
A$EB_MEAN <- A$W_G*A$MEAN_G + (1-A$W_G)*MU

# 5-bin grading
q5 <- unique(quantile(A$EB_MEAN, probs=seq(0,1,length.out=6), na.rm=TRUE, type=7))
cut5 <- function(x, brks) as.integer(cut(x, breaks=brks, include.lowest=TRUE, right=TRUE))
A$HA22_G5 <- cut5(A$EB_MEAN, q5)
A$HA22_G  <- A$HA22_G5

# Join & write back
df_en <- merge(d0, A[,c("HA_22ND","EB_MEAN","HA22_G5","HA22_G")], by="HA_22ND", all.x=TRUE)
df_en$POWER_ADJ <- df_en[[YCOL]] - df_en$EB_MEAN
df_en$EB_MEAN <- NULL
for (nm in names(df_en)) if (is.numeric(df_en[[nm]])) df_en[[nm]] <- if (nm %in% c("POWER","POWER_ADJ")) round(df_en[[nm]],4) else round(df_en[[nm]],2)
safe_write_dbf(df_en, OUTDBF)
cat("[HA_UD 5-bin] Updated: ", OUTDBF, "\n", sep="")

# Call POOR_EHIS.R to create POOR.DBF & RA3_RANO.DBF
if (MAKE_POOR) {
  if (!file.exists(POOR_SCRIPT)) stop("[HA_UD] POOR_EHIS.R not found at: ", POOR_SCRIPT)
  msg <- try(system2("Rscript", args=c(POOR_SCRIPT), stdout=TRUE, stderr=TRUE), silent=TRUE)
  if (inherits(msg,"try-error")) stop("[HA_UD] failed to run POOR_EHIS.R: ", as.character(msg))
  cat("[HA_UD] POOR_EHIS.R done.\n")
}

# ================== (append) MK_VERIFY 자동 실행 블록 ==================

# null-coalesce
`%||%` <- function(a, b) if (is.null(a)) b else a

# HA_UD.R 파일이 위치한 디렉터리(R/)를 알아내는 함수
.get_script_dir <- function() {
  # Rscript로 실행될 때
  args <- commandArgs(trailingOnly = FALSE)
  ix <- grep("^--file=", args)
  if (length(ix) > 0) {
    p <- sub("^--file=", "", args[ix[1]])
    return(normalizePath(dirname(p), winslash = "/", mustWork = FALSE))
  }
  # RStudio에서 실행될 때
  if (requireNamespace("rstudioapi", quietly = TRUE) &&
      isTRUE(try(rstudioapi::isAvailable(), silent = TRUE))) {
    ctx <- try(rstudioapi::getActiveDocumentContext(), silent = TRUE)
    if (!inherits(ctx, "try-error") && nzchar(ctx$path)) {
      return(normalizePath(dirname(ctx$path), winslash = "/", mustWork = FALSE))
    }
  }
  # source()로 실행될 때
  if (!is.null(sys.frames()[[1]]$ofile)) {
    return(normalizePath(dirname(sys.frames()[[1]]$ofile), winslash = "/", mustWork = FALSE))
  }
  # 최후: 현재 작업 폴더
  getwd()
}

# 경로 계산: HA_UD.R이 R/ 폴더에 있다고 가정 → 프로젝트 루트는 상위 폴더
.script_dir <- .get_script_dir()                # 예: C:/HK/R
.project_dir <- normalizePath(file.path(.script_dir, ".."),
                              winslash = "/", mustWork = FALSE)  # 예: C:/HK

# MK_VERIFY.R 절대 경로
mk_verify_script <- normalizePath(file.path(.script_dir, "MK_VERIFY.R"),
                                  winslash = "/", mustWork = FALSE)

if (!file.exists(mk_verify_script)) {
  stop("MK_VERIFY.R 경로를 찾을 수 없습니다: ", mk_verify_script,
       "\nHA_UD.R과 동일한 R/ 폴더에 MK_VERIFY.R이 있는지 확인하세요.")
}

# 입력/출력 기본값(필요 시 HA_UD.R 위쪽 로직에서 options()로 덮어쓰기 가능)
mk_in_path  <- getOption("mk.in_path",  normalizePath(file.path(.project_dir, "RA3_RANO1.DBF"),
                                                      winslash = "/", mustWork = FALSE))
mk_out_dir  <- getOption("mk.out_dir",  normalizePath(file.path(.project_dir, "OUT"),
                                                      winslash = "/", mustWork = FALSE))
mk_poor_out <- getOption("mk.poor_out", normalizePath(file.path(.project_dir, "POOR.DBF"),
                                                      winslash = "/", mustWork = FALSE))
mk_z_cut    <- as.character(getOption("mk.z_cut", 2.5))

# 경로 준비
if (!file.exists(mk_in_path)) {
  stop("MK_VERIFY 입력 파일을 찾을 수 없습니다: ", mk_in_path,
       "\nHA_UD.R 단계에서 생성/갱신되었는지, 또는 options(mk.in_path='...')로 지정했는지 확인하세요.")
}
if (!dir.exists(mk_out_dir)) {
  ok <- try(dir.create(mk_out_dir, recursive = TRUE), silent = TRUE)
  if (inherits(ok, "try-error") || isFALSE(ok)) {
    stop("출력 디렉터리를 생성할 수 없습니다: ", mk_out_dir)
  }
}

# Rscript 실행 파일
.RS_BIN <- {
  cand <- file.path(R.home("bin"), "Rscript")
  if (file.exists(cand)) cand else "Rscript"
}

message("[HA_UD] 다음 단계로 MK_VERIFY 실행 시작...")
args_vec <- c(mk_verify_script, mk_in_path, mk_out_dir, mk_poor_out, mk_z_cut)

res <- tryCatch(
  {
    out <- suppressWarnings(system2(.RS_BIN, args = args_vec, stdout = TRUE, stderr = TRUE))
    attr(out, "status") <- attr(out, "status") %||% 0L
    out
  },
  error = function(e) {
    attr(msg <- paste0("MK_VERIFY 실행 실패: ", conditionMessage(e)), "status") <- 1L
    msg
  }
)

status <- attr(res, "status")
cat(paste(res, collapse = "\n"), "\n")
if (!is.null(status) && status != 0L) {
  stop("[HA_UD] MK_VERIFY 실행이 오류 상태로 종료되었습니다(status=", status, ").")
} else {
  message("[HA_UD] MK_VERIFY 정상 완료.")
}

# ================= (end append) MK_VERIFY 자동 실행 블록 =================