#!/usr/bin/env Rscript
# train.R — Build full-row TRAIN with derived features (no filtering), key-first order
# Usage: source("C:/LEE/R/train.R"); or Rscript C:/LEE/R/train.R 25100308

suppressWarnings(suppressMessages({
  library(dplyr)
}))

args <- commandArgs(trailingOnly = TRUE)
SPAN <- if (length(args) >= 1) args[1] else "25100308"
DAY  <- substr(SPAN, 1, 6)
INFILE <- sprintf("C:/LEE/out/raw/EHIS_RA3_%s.csv", SPAN)
OUTDIR <- sprintf("C:/LEE/out/SCORE/%s", DAY)
if (!dir.exists(OUTDIR)) dir.create(OUTDIR, recursive = TRUE, showWarnings = FALSE)
OUTFILE <- sprintf("%s/DEVA_TRAIN_fullrows_with_derived_%s.csv", OUTDIR, SPAN)

read_csv_win <- function(path){
  x <- tryCatch(read.csv(path, fileEncoding = "CP949"), error=function(e) NULL)
  if (is.null(x)) x <- tryCatch(read.csv(path, fileEncoding = "UTF-8"), error=function(e) NULL)
  if (is.null(x)) stop("Failed to read CSV: ", path)
  x
}
num <- function(x) suppressWarnings(as.numeric(x))

DF <- read_csv_win(INFILE)
DF <- as.data.frame(DF)
names(DF) <- toupper(trimws(names(DF)))

if (!"HA_22ND" %in% names(DF)) stop("Missing HA_22ND")
levels_ha <- sort(unique(as.character(DF$HA_22ND)))
DF$HA22_ID <- as.integer(factor(as.character(DF$HA_22ND), levels = levels_ha))

if ("REFER1" %in% names(DF) && !"ROUND" %in% names(DF)) DF$ROUND <- suppressWarnings(as.integer(DF$REFER1))
if (!"LABEL" %in% names(DF) && "UPPRO_FR" %in% names(DF)) DF$LABEL <- DF$UPPRO_FR

# Derived features per horse (no row drops)
add_win <- function(g){
  if ("ROUND" %in% names(g)) g <- g[order(g$ROUND), , drop = FALSE]
  pw <- num(g$POWER); fp <- num(g$FANPRO_HI)
  n  <- nrow(g)
  g$PW_MI_TR  <- dplyr::lead(pw, 1)
  g$PW_MID_TR <- vapply(seq_len(n), function(i) mean(pw[(i+1):min(i+5, n)], na.rm = TRUE), numeric(1))
  g$PW_ALL_TR <- vapply(seq_len(n), function(i) if (i < n) mean(pw[(i+1):n], na.rm = TRUE) else NA_real_, numeric(1))
  g$FP_MI_TR  <- dplyr::lead(fp, 1)
  g$FP_MID_TR <- vapply(seq_len(n), function(i) mean(fp[(i+1):min(i+5, n)], na.rm = TRUE), numeric(1))
  g$FP_ALL_TR <- vapply(seq_len(n), function(i) if (i < n) mean(fp[(i+1):n], na.rm = TRUE) else NA_real_, numeric(1))
  g
}
if (!"HNO_FR" %in% names(DF)) stop("Missing HNO_FR")
DF <- DF %>% group_by(HNO_FR) %>% group_modify(~add_win(.x)) %>% ungroup()

# Key-first columns (keep everything)
lead_keys <- c("HA22_ID","HA_22ND","HNO_FR")
lead_more <- intersect(c("RANO_FR","LABEL","UPPRO_FR","REFER1","ROUND"), names(DF))
rest_cols <- setdiff(names(DF), c(lead_keys, lead_more))
ord <- c(lead_keys, lead_more, rest_cols)
seen <- character(0); final_cols <- character(0)
for (nm in ord) if (!(nm %in% seen)) { final_cols <- c(final_cols, nm); seen <- c(seen, nm) }
TRAIN <- DF[, final_cols, drop = FALSE]

# Sort: HA22_ID -> RANO_FR(if exists) -> HNO_FR
ord_cols <- c("HA22_ID", if ("RANO_FR" %in% names(TRAIN)) "RANO_FR" else NULL, if ("HNO_FR" %in% names(TRAIN)) "HNO_FR" else NULL)
if (length(ord_cols)) TRAIN <- TRAIN[do.call(order, TRAIN[ord_cols]), , drop = FALSE]

write.csv(TRAIN, OUTFILE, row.names = FALSE, fileEncoding = "CP949")
cat(sprintf("TRAIN full rows with derived saved: %s (rows=%d, cols=%d)\n", OUTFILE, nrow(TRAIN), ncol(TRAIN)))
