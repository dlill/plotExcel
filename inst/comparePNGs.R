#### HEADER ================================================================
#
# comparePNGs.R
#
# [PURPOSE]
# Develop comparison functionality
#
# [AUTHOR]
# Daniel Lill (daniel.lill@intiquan.com)
#
# [CLEANING]
rm(list = grep("^(\\.input|\\.output)", ls(all.names = TRUE), value = TRUE))
rm(list = ls())
#
# [INPUT]
.input <- "../"
#
# [OUTPUT]
.outputFolder <- "../Output/comparePNGs"
#
# [OTHER]
#
## Preliminaries ====
# Set working directory to script folder
try(setwd(dirname(rstudioapi::getSourceEditorContext()$path)), silent = TRUE)
stopifnot(basename(dirname(getwd())) == "excelPlot")

# Default packages (do not load other packages, use "::" instead)
library(dplyr)
library(ggplot2)


# -------------------------------------------------------------------------#
#  ----
# -------------------------------------------------------------------------#
devtools::load_all("C:/PROJECTS/1_PROJTOOLS/excelPlot/")

l <- paste0(system.file("exampleData/04-IrisMulti.pdf", package = "excelPlot"), "::page 1")
r <- paste0(system.file("exampleData/04-IrisMulti.pdf", package = "excelPlot"), "::page 2")

# fll <- applyPngPipelineOnePage(parsePlotSpec(l))
# flr <- applyPngPipelineOnePage(parsePlotSpec(r))

dInfo <- data.table(left = l, right = r, diff = "diff(left,right)")

# debugonce(parseTable)
# debugonce(plotExcel)
# debugonce(parseDiffSpec)
# debugonce(compareImages)
plotExcel(dInfo, filename = "C:/PROJECTS/tmp.xlsx")

browseURL(.Last.value)
# -------------------------------------------------------------------------#
# Exit ----
# -------------------------------------------------------------------------#
