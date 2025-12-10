#' Compare plots of multiple folders with the same folder structure
#'
#' @param projectMulti named vector to IQRnlmeProject folders
#' @param filename filepath to wriet excel file to
#' @param fileSelection vector of files to restrict selection to
#'
#' @returns Writes a plotExcel file
#' @export
#' @md
#' @family UI
#' @importFrom pdftools pdf_length
#' @importFrom data.table rbindlist
compareProjects_excelPlot <- function(projectMulti, filename, fileSelection = NULL, FLAGreturnDataFrame = FALSE) {

  mf <- missing(filename)

  # Get basic overview of pdf files and pages
  dPdfInfo <- lapply(projectMulti, function(y) {
    pdfFiles <- list.files(y, pattern = ".pdf$", full.names = TRUE, recursive = TRUE)
    pdfFilesWithinProject <- unique(gsub(paste0(y, "/"), "", pdfFiles))

    dPdfInfo <- lapply(setNames(pdfFiles, nm = pdfFilesWithinProject), function(x) {
      data.table(nPages = pdftools::pdf_length(x))
    })
    dPdfInfo <- data.table::rbindlist(dPdfInfo, idcol = "pdfFile")
    dPdfInfo[,`:=`(projectPath = y)]
    dPdfInfo
  })
  dPdfInfo <- rbindlist(dPdfInfo, idcol = "projectComment")

  message(paste0(
    "Use this code to conveniently specify the subset of plots to include in the Excel table.\n",
    "The order of fileSelection will be applied.\n\n",
    "fileSelection <- ", paste0(deparse(sort(unique(dPdfInfo$pdfFile)), width.cutoff = 20), collapse = "\n")))

  if (mf) {return()}


  # Sort According to fileSelection
  if (!is.null(fileSelection)) {
    dPdfInfo <- dPdfInfo[pdfFile %in% fileSelection]
    dPdfInfo[,`:=`(pdfFile = factor(pdfFile, levels = fileSelection))]
    dPdfInfo <- dPdfInfo[order(pdfFile)]
    dPdfInfo[,`:=`(pdfFile = as.character(pdfFile))]
  }

  # Further wrangling of death
  dPdfInfo <- dPdfInfo[,list(page = 1:nPages), by = c("pdfFile", "projectPath", "projectComment")]
  dPdfInfo[,`:=`(plotSpec = paste0(file.path(projectPath, pdfFile), "::page ", page))]
  dPdfInfo[,`:=`(File = paste0(gsub("/", " / ", pdfFile), ", page ", page, "::vcenterSize48"))]
  dPdfInfo[,`:=`(File = factor(File, unique(File)))]
  dPdfInfo <- dcast(dPdfInfo, File ~ projectComment, value.var = "plotSpec")

  # Final assembly of the excelPlot specification table
  # Function from compare_projects, but to be sure this is the right verison, I'm pasting it into here
  getTabHeader <- function(projectMulti) {
    # Shorten paths
    projectPathsShort <- projectMulti
    projectPathsShort <- gsub("../Models/","Models/",projectPathsShort)
    projectPathsShort <- gsub("../Output/FINALMODELS/","Final/",projectPathsShort)
    projectPathsShort <- gsub("MODEL[_ ]*","", projectPathsShort)
    projectPathsShort <- gsub("/$","",projectPathsShort)
    projectPathsShort <- strsplit(projectPathsShort, "/")
    projectPathsShort <- lapply(projectPathsShort, function(x) paste0(gsub("^(\\d+).*","\\1",x), collapse = "/"))
    projectPathsShort <- do.call(c, projectPathsShort)
    projectPathsShort <- setNames(projectPathsShort, names(projectMulti))
    dHeader <- data.table::as.data.table(as.list(c(Parameter = "Short path", projectPathsShort)))
    dHeader
  }
  dTabHeader <- getTabHeader(projectMulti)
  setnames(dTabHeader, old = "Parameter", new = "File")
  dPdfInfo <- data.table::rbindlist(list(dTabHeader, data.table(File = ""), dPdfInfo), fill = TRUE, use.names = TRUE)

  if (FLAGreturnDataFrame) return(dPdfInfo)

  plotExcel(d = dPdfInfo, filename = filename, textColWidth = 10, headerRowStyle = "centerSize48")
}


#' Export all plots in a folder into Excel
#'
#' @param path Path with plots. Will be searched recursively for pdf and png files
#' @param filename File path of the output excel file
#' @param fileSelection Vector of plots to be included. Default: NULL = include all files
#' @param resolution in dpi
#' @param CFLAGLayout Default: "no". What to do with the layout data.table:
#' "no" = Nothing, just continue,
#' "return" = return the data.table instead of writing the excel. Useful if you want to modify this yourself, e.g. to add another plot with same structure as a second column.
#' "insert" = insert the table as a tribble call into your script. Needs package github.com/dlill/RSAddins
#' @param nPagesMax Maximum pages per file.
#' @param FLAGopenExcel Open the excel file right away?
#' @param FLAGtemp Write to temp file?
#' @param filterRegexpRemove regex to remove files
#' @param compareToCommit Commit hash to compare the current HEAD to.
#'
#' @returns Writes the Excel file to filename.
#' @export
#' @md
#' @family UI
#' @importFrom tools file_ext
#' @importFrom pdftools pdf_length
#' @importFrom data.table rbindlist
plotExcelFolder <- function(path, filename, fileSelection = NULL, resolution = 150, CFLAGLayout = c("no", "return", "insert"),
                            nPagesMax = 4, FLAGopenExcel = FALSE, FLAGtemp = FALSE, filterRegexpRemove = NULL,
                            compareToCommit = NULL
                            ) {


  mf <- missing(filename)
  sf <- substitute(filename)
  if (!FLAGtemp & !mf) {
    # If print is wanted
    deparsedfilename <- deparse(sf)
  } else {
    filename <- paste0("C:/PROJECTS/tmp", format(Sys.time(), "--%Y-%m-%d_%H%M"),".xlsx")
    deparsedfilename <- deparse(filename)
  }
  verifyArg(fileSelection   , expectedMode = "character", allowNull = TRUE)
  CFLAGLayout <- match.arg(CFLAGLayout)
  verifyArg(FLAGopenExcel   , expectedMode = "logical")
  verifyArg(FLAGtemp        , expectedMode = "logical")
  verifyArg(nPagesMax       , expectedMode = "numeric")


  # Get basic overview of pdf files and pages
  pdfFiles <- list.files(path, pattern = "\\.(pdf|png)$", full.names = TRUE, recursive = TRUE)
  if (!is.null(filterRegexpRemove)) pdfFiles <- grep(filterRegexpRemove, pdfFiles, value = TRUE, invert = TRUE)
  pdfFilesWithinProject <- gsub("^/", "", gsub(paste0(path), "", pdfFiles))

  dPdfInfo <- lapply(setNames(pdfFiles, nm = pdfFilesWithinProject), function(x) {
    data.table(nPages = ifelse(tools::file_ext(x) == "pdf", pdftools::pdf_length(x), 1))
  })
  dPdfInfo <- data.table::rbindlist(dPdfInfo, idcol = "pdfFile")
  dPdfInfo

  message(
    "Use this code to conveniently specify the subset of plots to include in the Excel table.\n",
    "The order of fileSelection will be applied.\n\n",
    "fileSelection <- ", paste0(gsub("c\\(", "c(\n", deparse(unique(dPdfInfo$pdfFile), width.cutoff = 20)), collapse = "\n"),"\n",
    "\n",
    "Alternatively, use FLAGinsertLayout to insert the full excel spec table into your Rscript")

  if (!FLAGtemp & mf & CFLAGLayout == "no") {
    # This means we are only interested in fileSelection.
    return()
  }

  if (any(dPdfInfo[,nPages > nPagesMax])) {
    message("The following files have more pages than nPagesMax (", nPagesMax, "). Only ", nPagesMax, " pages will be extracted:\n",
            "\n", paste0("  - ", dPdfInfo[nPages > nPagesMax,paste0(pdfFile, " - ", nPages)], collapse = "\n"))
  }
  dPdfInfo[nPages > nPagesMax,`:=`(nPages = min(nPages, nPagesMax))]


  # Sort According to fileSelection
  if (!is.null(fileSelection)) {
    dPdfInfo <- dPdfInfo[pdfFile %in% fileSelection]
    dPdfInfo[,`:=`(pdfFile = factor(pdfFile, levels = fileSelection))]
    dPdfInfo <- dPdfInfo[order(pdfFile)]
    dPdfInfo[,`:=`(pdfFile = as.character(pdfFile))]
  }

  # Further wrangling of death
  dPdfInfo <- dPdfInfo[,list(page = 1:nPages), by = c("pdfFile")]
  dPdfInfo[,`:=`(plotSpec = paste0(file.path(path, pdfFile), "::page ", page, "::resolution ", resolution))]
  dPdfInfo[,`:=`(File = paste0(gsub("/", " / ", pdfFile), ", page ", page, "::vcenter"))]
  dPdfInfo <- dPdfInfo[,list(File, Plot = plotSpec)]

  if (!is.null(compareToCommit)) {
    dPdfInfo[,`:=`(OLD = paste0(Plot,"::commit ", compareToCommit))]
    dPdfInfo[,`:=`(DIFF = "diff(Plot,OLD)")]
  }

  if (CFLAGLayout == "return") {
    return(dPdfInfo)
  }
  if (CFLAGLayout == "insert") {
    e <- rstudioapi::getSourceEditorContext()
    rstudioapi::documentSave(id = e$id)
    row <- e$selection[[1]]$range$end[1]

    codeToInsert <- paste0(c(paste0("dLayout <- ", RSAddins::outputMdTable2(dPdfInfo)),
                             "",
                             paste0("excelPlot::plotExcel(d = dLayout, filename = ", deparsedfilename, ", textColWidth = 10)"),
                             "\n"),
                           collapse = "\n")

    rstudioapi::insertText(location = rstudioapi::document_position(row, 1), text = codeToInsert, e$id)
    rstudioapi::documentSave(id = e$id)
    return(invisible(dPdfInfo))
  }

  # Export
  filename <- plotExcel(d = dPdfInfo, filename = filename, textColWidth = 10)

  if (FLAGopenExcel) {shell.exec(normalizePath(filename))}

  invisible(filename)
}





