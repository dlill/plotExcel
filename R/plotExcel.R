# ------------------------------------------------------------------------- #
# Excel export pipeline ----
# ------------------------------------------------------------------------- #

#' Export plots into an Excel table.
#'
#' @param d \code{data.table("plot.pdf::spec value::spec2 value2...", "text::style", "diff(colname1,colname2)", ...)}.
#' For supported plot spec arguments, see [plotSpec()].
#' Available text styles can be listed by [availableStyles()].
#' Visual diffing of two plots is done by \code{"diff(colname1,colname2)"}, see example.
#' @param filename File path to export excel file to.
#' @param headerRowStyle Default: "center". Style used for header row.
#' @param FLAGaddBorders Default: FALSE. Add borders around all cells?
#' @param FLAGpdf Default: FALSE. Export Excel as pdf and open in pdf viewer? Useful for quick drafting.
#' @param pdfPageSize PDF pagination mode: `"single"` fits the worksheet to one
#'   page; `"A4"` splits it across A4 pages.
#' @param textColWidth Default: 5. Width of pure text columns, in cm.
#' @param FLAGopenExcel Default: FALSE. Open Excel after export?
#' @param FLAGtemp Default: FALSE. Write to a temporary file?
#'
#' @returns `filename`, but is called for its side effect.
#' @export
#' @md
#' @family UI
#' @importFrom data.table as.data.table
#' @importFrom openxlsx saveWorkbook
#' @importFrom tools file_path_sans_ext
#'
#' @examples
#' \dontrun{
#' d <- data.table(tibble::tribble(
#'   ~Description, ~`Plots 1`, ~`Plots 2`, ~Comparison,
#'   "Crop::3",
#'   paste0(system.file("exampleData/01-Iris.pdf", package = "plotExcel"), "::xmax 85"),
#'   paste0(system.file("exampleData/02-Iris-Brewer.pdf", package = "plotExcel")),
#'   "diff(`Plots 1`, `Plots 2`)",
#'   "Text rotated up::4",
#'   paste0(system.file("exampleData/04-IrisMulti.pdf", package = "plotExcel"), "::page 2"),
#'   paste0(system.file("exampleData/04-IrisMulti.pdf", package = "plotExcel"), "::page 1"),
#'   "diff(`Plots 1`, `Plots 2`)"
#' ))
#' filename <- "~/.plotExcel/example.xlsx"
#' plotExcel(d, filename = filename, headerRowStyle = "center",
#'   FLAGaddBorders = FALSE, FLAGpdf = FALSE, textColWidth = 5)
#' # Export to pdf and open in viewer - commented out for CRAN check, run manually
#' # plotExcel(d, filename = filename, headerRowStyle = "center",
#' #   FLAGaddBorders = FALSE, FLAGpdf = TRUE, textColWidth = 5)
#' # After inspection, remove the file
#' unlink(filename)
#' }
plotExcel <- function(d, filename, headerRowStyle = "center", FLAGaddBorders = FALSE, FLAGpdf = FALSE, textColWidth = 5,
                      FLAGopenExcel = FALSE, FLAGtemp = FALSE, pdfPageSize = c("single", "A4")) {

  # ------------------------------------------------------------------------- #
  # Input verification ----
  # ------------------------------------------------------------------------- #
  mf <- missing(filename)
  if (mf & !FLAGtemp) stop("filename can't be missing when FLAGtemp == FALSE")
  pdfPageSize <- match.arg(pdfPageSize)
  if (FLAGtemp) {
    filename <- resolveTempFilename()
  }

  # ------------------------------------------------------------------------- #
  # Crunch ----
  # ------------------------------------------------------------------------- #
  dParsed <- parseTable(d, headerRowStyle)

  # ------------------------------------------------------------------------- #
  # Handle plots ----
  # ------------------------------------------------------------------------- #
  # Apply extraction and cropping pipeline
  lapply(dParsed[ISPLOT == TRUE, SPEC], applyPngPipelineOnePage)

  # Apply diffs, then flag them as plots
  lapply(dParsed[ISDIFF == TRUE, SPEC], applyImageDiff)
  dParsed[ISDIFF == TRUE,`:=`(ISPLOT = TRUE)]

  # Get width and height in cm.
  # It is a weird bug of gs that it does not include the dpi in the metadata, and somehow imagemagick can't overwrite the metadata...
  dParsed[ISPLOT == TRUE,c("WIDTHCM", "HEIGHTCM"):=(data.table::as.data.table(t(sapply(seq_along(FILE), function(i) {
    # Old apporach
    # pxInfo <- system(paste0('identify -format "%w\n%h" ', FILE[[i]]), intern = TRUE) # x pixels, y pixels
    # pxInfo <- as.numeric(pxInfo)

    # New approach
    img <- magick::image_read(FILE[[i]])
    pxInfo <- unlist(magick::image_info(img)[c("width", "height")])
    pxInfo / SPEC[[i]]$resolution * 2.54
  }))))]

  dwidths <- dParsed[,list(WIDTHCM = max(WIDTHCM, na.rm = TRUE)), by = "COLID"]
  dwidths[!is.finite(WIDTHCM),`:=`(WIDTHCM = textColWidth)]
  dheights <- dParsed[,list(HEIGHTCM = max(HEIGHTCM, na.rm = TRUE)), by = "ROWID"]
  dheights[!is.finite(HEIGHTCM),`:=`(HEIGHTCM = 2)]

  # ------------------------------------------------------------------------- #
  # Populate the Excel ----
  # ------------------------------------------------------------------------- #
  wb <- populateExcel(dParsed = dParsed,
                      dwidths = dwidths,
                      dheights = dheights,
                      FLAGaddBorders = FLAGaddBorders)

  # ------------------------------------------------------------------------- #
  # Export ----
  # ------------------------------------------------------------------------- #
  if (!dir.exists(dirname(filename))) {dir.create(dirname(filename), showWarnings = FALSE, recursive = TRUE)}
  filename <- resolveLockedFilePath(filename)
  t0 <- Sys.time()
  openxlsx::saveWorkbook(wb = wb, file = filename, overwrite = TRUE)
  message("Excel file was saved to ", filename, " (", signif(Sys.time() - t0, 2), " s)")
  message("Run `availableStyles()` to see text styles and decorator syntax.")

  if (FLAGpdf) {
    message("Generating a pdf from the Excel file since FLAGpdf = TRUE.")
    pdfFilename <- file.path(tempdir(), paste0(tools::file_path_sans_ext(basename(filename)), ".pdf"))
    convertOfficeToPdf(fileIn = filename, fileOut = pdfFilename, pageSize = pdfPageSize)
    utils::browseURL(pdfFilename)
  }

  if (FLAGopenExcel) {shell.exec(normalizePath(filename))}

  invisible(filename)

}


#' Parse the data.table which specifies the texts and plots to be exported
#'
#' @inheritParams plotExcel
#'
#' @returns Augmented `d` with parsed information
#' @md
#' @importFrom data.table copy melt rbindlist
parseTable <- function(d, headerRowStyle) {

  d <- data.table::copy(d)
  if (anyDuplicated(names(d))) {
    stop("d has duplicated names. This is not supported. Duplicated names: ",
         "\n", paste0("  - ", unique(names(d)[duplicated(names(d))]), collapse = "\n"))
  }

  # Add row and col indices
  d[,`:=`(ROWID = (1:.N) + 1)]
  d <- data.table::melt(d,id.vars = "ROWID", variable.name = "VARIABLE", variable.factor = FALSE, value.name = "VALUE")
  d[,`:=`(COLID = as.numeric(factor(VARIABLE, unique(VARIABLE))))]

  # Use colnames as first row
  d <- data.table::rbindlist(list(
    d[,unique(.SD),.SDcols = c("COLID", "VARIABLE")][,`:=`(ROWID = 1, VALUE = paste0(VARIABLE,"::", headerRowStyle))],
    d
  ), use.names = TRUE)
  d[,`:=`(ID = 1:.N)]

  # Parse specs into lists and get certain variables into the main table
  d[,`:=`(ISPLOT = file.exists(gsub("::.*","", VALUE)))]
  d[,`:=`(ISDIFF = grepl("diff\\(",VALUE))]
  d[ISPLOT == FALSE & ISDIFF == FALSE,`:=`(SPEC = lapply(VALUE, function(x) {parseTextSpec(x)}))]
  d[ISPLOT == TRUE,`:=`(SPEC = lapply(VALUE, function(x) {parsePlotSpec(x)}))]
  d[ISPLOT == TRUE,`:=`(PATHS = lapply(SPEC, function(x) do.call(epFiles, x)))]
  d[,`:=`(FILE = sapply(PATHS, function(x) x$tmpPathCommitPageCrop))]

  d[ISDIFF == TRUE,`:=`(SPEC = list(parseDiffSpec(ROWID = ROWID, VALUE = VALUE, d = d))), by = c("COLID", "ROWID")]
  d[ISDIFF == TRUE,`:=`(FILE = sapply(SPEC, function(x) x$filename))]

  d
}

#' Given all parsed information, construct the Excel workbook
#'
#' @param dParsed Parsed d, data.table
#' @param dwidths data.table with column width information
#' @param dheights data.table with row height information
#' @param FLAGaddBorders Add borders around all cells? Default: FALSE
#'
#' @returns An openxlsx workbook `wb`
#' @md
#' @importFrom openxlsx createWorkbook addWorksheet insertImage writeData addStyle pageSetup freezePane setColWidths setRowHeights createStyle
populateExcel <- function(dParsed, dwidths, dheights, FLAGaddBorders) {
  wb <- openxlsx::createWorkbook()

  sheet <- "Sheet1"
  openxlsx::addWorksheet(wb = wb, sheetName = sheet)

  # Not vectorized, but the tables aren't large anyway.
  i <- (seq_len(nrow(dParsed)))[[1]]
  for (i in seq_len(nrow(dParsed))) {
    content <- dParsed[i]

    if (content$ISPLOT) {
      openxlsx::insertImage(wb = wb,
                            sheet = sheet,
                            file = content$FILE[[1]],
                            startRow = content$ROWID,
                            startCol = content$COLID,
                            width = content$WIDTHCM,
                            height = content$HEIGHTCM,
                            units = "cm",
                            dpi = content$SPEC[[1]]$resolution)
    } else {
      openxlsx::writeData(wb = wb,
                          sheet = sheet,
                          x = content$SPEC[[1]]$text,
                          startRow = content$ROWID,
                          startCol = content$COLID,
                          colNames = FALSE)

      # styleList is defined within this package.
      openxlsx::addStyle(wb = wb,
                         sheet = sheet,
                         style = styleList[[content$SPEC[[1]]$style]],
                         rows = content$ROWID,
                         cols = content$COLID)
    }

  }

  # Finalize layout
  openxlsx::pageSetup(wb = wb, sheet = sheet, fitToWidth = TRUE, fitToHeight = TRUE) # So pdf export on is done on a single page
  openxlsx::freezePane(   wb, sheet = sheet, firstRow = TRUE, firstCol = TRUE)
  openxlsx::setColWidths( wb, sheet = sheet, cols = dwidths$COLID, widths = dwidths$WIDTHCM * 5.3)
  openxlsx::setRowHeights(wb, sheet = sheet, rows = dheights$ROWID, heights = dheights$HEIGHTCM / 2.54 * 72 * 1.05)

  if (FLAGaddBorders) {openxlsx::addStyle(wb = wb, sheet = sheet, style = openxlsx::createStyle(border = "TopBottomLeftRight"),
                                          rows = unique(dParsed$ROWID), cols = unique(dParsed$COLID), gridExpand = TRUE, stack = TRUE)}

  wb
}

# ------------------------------------------------------------------------- #
# Excel cell styles ----
# ------------------------------------------------------------------------- #

styleList <- list(
  left          = openxlsx::createStyle(fontSize = 18, wrapText = TRUE, textDecoration = "bold", halign = NULL, valign = NULL),
  center        = openxlsx::createStyle(fontSize = 18, wrapText = TRUE, textDecoration = "bold", halign = "center"),
  vcenter       = openxlsx::createStyle(fontSize = 18, wrapText = TRUE, textDecoration = "bold", valign = "center"),
  hvcenter      = openxlsx::createStyle(fontSize = 18, wrapText = TRUE, textDecoration = "bold", halign = "center", valign = "center"),
  rotateUp      = openxlsx::createStyle(fontSize = 18, wrapText = TRUE, textDecoration = "bold", halign = "right", valign = "center", textRotation = 90),
  rotateDown    = openxlsx::createStyle(fontSize = 18, wrapText = TRUE, textDecoration = "bold", halign = "left" , valign = "center", textRotation = -90),
  leftSize48    = openxlsx::createStyle(fontSize = 48, wrapText = TRUE, textDecoration = "bold", halign = NULL, valign = NULL),
  centerSize48  = openxlsx::createStyle(fontSize = 48, wrapText = TRUE, textDecoration = "bold", halign = "center"),
  vcenterSize48 = openxlsx::createStyle(fontSize = 48, wrapText = TRUE, textDecoration = "bold", valign = "center"),
  plain         = openxlsx::createStyle()
  )


#' Print available styles and decorator syntax
#'
#' Lists the numbered text styles and the decorator syntax accepted by
#' [parsePlotSpec()] and [parseTextSpec()], along with the image-diff expression
#' syntax accepted in `plotExcel()` input tables. Plot decorator keys are
#' derived from the arguments of [plotSpec()].
#'
#' @returns Prints available styles and decorator syntax.
#' @export
#' @md
#' @family UI
#'
#' @examples
#' availableStyles()
availableStyles <- function() {
  plotDecoratorKeys <- setdiff(names(formals(plotSpec)), "path")

  cat(
    "Available text styles:\n",
    paste0(seq_along(styleList), ": ", names(styleList), collapse = "\n"),
    "\n\nPlot decorator syntax:\n",
    "path::key value::key2 value2\n",
    "Valid keys: ", paste(plotDecoratorKeys, collapse = ", "),
    "\n\nText decorator syntax:\n",
    '"text::style"',
    "\n\nDiff syntax:\n",
    "diff(col1, col2)\n",
    sep = ""
  )
}



