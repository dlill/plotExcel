

#' Resolve a temporary output filename
#'
#' @param ext File extension without a leading dot.
#'
#' @return A timestamped filename in the preferred project directory when it
#'   exists, otherwise in the user's home directory.
#' @md
resolveTempFilename <- function(ext = "xlsx") {
  baseDir <- if (.Platform$OS.type == "windows") "C:/PROJECTS" else path.expand("~/PROJECTS")
  if (!dir.exists(baseDir)) baseDir <- path.expand("~/")

  file.path(baseDir, paste0("tmp", format(Sys.time(), "--%Y-%m-%d_%H%M"), ".", ext))
}


#' Verify that input arguments meet certain conditions
#'
#' @param x Any object
#' @param allowNull TRUE: x can be null. If TRUE and x is null, no additional tests are made.
#' @param expectedClass Missing or character denoting the class of x
#' @param expectedMode Missing or character denoting the mode of x
#' @param expectedLength Missing or integer denoting the length of x
#' @param expectedSign Missing or 1 or -1 denoting the sign of x
#' @param expectedNames Names which must at least be present in x
#' @param expectedTestFun A function for which the call `expectedTestFun(x)` returns TRUE (test is passed) or FALSE (test is failed)
#'
#' @return Called for side-effect. If all tests pass, nothing happens. If errors occur, they are collected in informative error messages.
#' @md
#'
#' @examples
#' \dontrun{
#'
#' # Intended use case: Within a function
#' f <- function(myFancyArgument) {verifyArg(myFancyArgument,
#'   expectedMode = "character", expectedLength = 2)}
#' f(c(2)) # Two errors
#' f(c(2,3)) # One error
#' f(as.character(c(2))) # One error
#' f(as.character(c(2,3))) # No error
#'
#' # Works also with more complex expressions:
#' f2 <- function(myFancyDF) {verifyArg(myFancyDF$TIME, expectedMode = "numeric")}
#' f2(data.frame(TIME = letters))
#'
#' # Checking for integers is a bit tricky and should not necessarily be done with expectedClass
#'
#' # Doubles
#' verifyArg(1, expectedClass = "numeric")    # Gives no error
#' verifyArg(1, expectedClass = "double")     # Gives error
#' verifyArg(1, expectedMode = "numeric")     # Gives no error
#' verifyArg(1, expectedTestFun = is.numeric) # Gives no error, last resort
#'
#' # Integers
#' verifyArg(1L, expectedClass = "numeric") # Gives error
#' verifyArg(1L, expectedMode = "numeric")  # Gives no error
#'
#'
#' }
verifyArg <- function(x, allowNull = FALSE,
                      expectedClass,
                      expectedMode,
                      expectedLength,
                      expectedSign ,
                      expectedNames,
                      expectedTestFun) {
  subx <- substitute(x)
  charx <- deparse(subx)

  if (allowNull & is.null(x)) return()
  if (!allowNull & is.null(x)) stop(charx, " must not be NULL")

  errors <- c()
  if (!missing(expectedClass) && !expectedClass %in% class(x))
    errors <- c(errors,
                paste0("Class of ", charx, " -----------"),
                paste0("* Expected: '", expectedClass, "'"),
                paste0("* Supplied: '", paste0(class(x), collapse = ","), "'")
    )
  if (!missing(expectedMode) && !expectedMode %in% mode(x))
    errors <- c(errors,
                paste0("Mode of ", charx, " -----------"),
                paste0("* Expected: '", expectedMode, "'"),
                paste0("* Supplied: '", paste0(mode(x), collapse = ","), "'")
    )
  if (!missing(expectedLength) && length(x) != expectedLength)
    errors <- c(errors,
                paste0("Length of ", charx, " -----------"),
                paste0("* Expected: '", expectedLength, "'"),
                paste0("* Supplied: '", paste0(length(x), collapse = ","), "'")
    )
  if (!missing(expectedSign) && !expectedSign %in% c(1,-1))
    stop("expectedSign should be one of 1,-1. ",
         "Please update your call to verifyArg. ",
         "(This is an error in the call to verifyArg itself, not an argument check.)")
  if (!missing(expectedSign) && any(sign(x) != expectedSign))
    # [ ] Add support for sign(0)
    errors <- c(errors,
                paste0("Sign of ", charx, " -----------"),
                paste0("* Expected: '", expectedSign, "'"),
                paste0("* Elements with wrong sign: '", paste0(which(sign(x) != expectedSign), collapse = ","), "'")
    )
  if (!missing(expectedNames) && any(! expectedNames %in% names(x)))
    errors <- c(errors,
                paste0("Required names of ", charx, " -----------"),
                paste0("* Expected: '", paste0(expectedNames, collapse = ","), "'"),
                paste0("* Supplied:  '", paste0(intersect(names(x), expectedNames), collapse = ","), "'"),
                paste0("* Missing:  '", paste0(setdiff(expectedNames, names(x)), collapse = ","), "'"),
                paste0("* Additional:  '", paste0(setdiff(names(x), expectedNames), collapse = ","), "'")
    )
  if (!missing(expectedTestFun)) {
    subTest <- substitute(expectedTestFun)
    charTest <- as.character(subTest)
    test <- expectedTestFun(x)
    if (!test)
      errors <- c(errors,
                  paste0("The following test returned FALSE ----"),
                  paste0("  ", charTest, "(", charx, ")")
      )
  }

  if (length(errors))
    stop(paste0(errors, collapse = "\n"))

}


#' Check if file is writable; if not, append a datetime stamp
#'
#' @param path File path to check
#'
#' @returns Original `path` if writable; otherwise, a datetime-stamped path
#' @md
#' @importFrom tools file_path_sans_ext file_ext
resolveLockedFilePath <- function(path) {
  original_dir <- getwd()
  on.exit(setwd(original_dir))

  target_dir <- dirname(path)
  setwd(target_dir)

  file_name <- basename(path)
  time_suffix <- format(Sys.time(), "__%Y-%m-%d_%H%M")
  stamped_name <- paste0(
    tools::file_path_sans_ext(file_name),
    time_suffix,
    ".",
    tools::file_ext(file_name)
  )

  is_locked <- file.exists(file_name) && tryCatch({
    file.rename(file_name, stamped_name)
    file.rename(stamped_name, file_name)
    FALSE
  }, warning = function(w) TRUE)

  if (is_locked) {
    message(
      "File '", file_name, "' appears to be locked (possibly open in another program).\n",
      "--> Output will be saved as: '", stamped_name, "'"
    )
    file.copy(file_name, stamped_name, overwrite = TRUE)
    file_name <- stamped_name
  }

  file.path(target_dir, file_name)
}


#' Number of pages / slides of a plot input file, without conversion
#'
#' Cheap dispatcher used before the plot pipeline runs so that Office files
#' don't have to be converted to PDF just to get a page count.
#'
#' * `.pdf`  -> [pdftools::pdf_length()]
#' * `.png`  -> 1
#' * `.pptx` -> number of `ppt/slides/slide*.xml` entries in the zip container
#' * `.docx` -> `<Pages>` field parsed from `docProps/app.xml` in the zip
#'   container (as saved by Word). Falls back to 1 when no such metadata is
#'   present (e.g. docx files produced by tools other than Word).
#' * anything else (`.doc`, `.ppt`, ...) -> 1
#'
#' @param path Path to a plot input file
#'
#' @return Positive integer scalar
#' @md
#' @importFrom tools file_ext
#' @importFrom pdftools pdf_length
getNPages <- function(path) {
  verifyArg(path, expectedClass = "character", expectedLength = 1)
  ext <- tolower(tools::file_ext(path))
  if (ext == "pdf")  return(pdftools::pdf_length(path))
  if (ext == "png")  return(1L)
  if (ext == "pptx") return(getNSlidesPptx(path))
  if (ext == "docx") return(getNPagesDocx(path))
  1L
}


#' Number of slides in a PPTX file, from its zip container
#'
#' Counts entries matching `ppt/slides/slide*.xml`.
#'
#' @param path Path to a `.pptx` file
#'
#' @return Positive integer scalar
#' @md
getNSlidesPptx <- function(path) {
  entries <- utils::unzip(path, list = TRUE)$Name
  length(grep("^ppt/slides/slide[0-9]+\\.xml$", entries))
}


#' Number of pages in a DOCX file, from Word-written metadata
#'
#' Reads the `<Pages>` value from `docProps/app.xml` in the zip container. That
#' value is only present (and reliable) if the docx was saved by Word after a full
#' rendering. When the metadata is missing or reports `1` (which is often stale for
#' multi-page docs that were never opened in Word), we fall back to a headless
#' Office->PDF conversion and count pages with `pdftools::pdf_length()`.
#'
#' @param path Path to a `.docx` file
#'
#' @return Positive integer scalar
#' @md
getNPagesDocx <- function(path) {
  # Fast path: trust the <Pages> metadata only when Word wrote a plausibly
  # rendered value (>= 2). Values of 1 (or missing) are unreliable for docx
  # files that were generated programmatically or never opened/saved by Word.
  entries <- utils::unzip(path, list = TRUE)$Name
  if ("docProps/app.xml" %in% entries) {
    exdir <- tempfile("docx-app-")
    dir.create(exdir, showWarnings = FALSE, recursive = TRUE)
    utils::unzip(path, files = "docProps/app.xml", exdir = exdir)
    appXml <- paste(readLines(file.path(exdir, "docProps", "app.xml"), warn = FALSE),
                    collapse = "\n")
    unlink(exdir, recursive = TRUE)
    m <- regmatches(appXml, regexpr("<Pages>[0-9]+</Pages>", appXml))
    if (length(m)) {
      n <- as.integer(gsub("</?Pages>", "", m))
      if (!is.na(n) && n >= 2) return(n)
    }
  }

  # Fallback: convert to a temporary PDF and count pages from it
  tmpPdf <- tempfile(fileext = ".pdf")
  on.exit(unlink(tmpPdf), add = TRUE)
  if (Sys.info()["sysname"] == "Windows") {
    convertOfficeToPdfWindows(fileIn = path, fileOut = tmpPdf)
  } else {
    convertOfficeToPdfLinux(fileIn = path, fileOut = tmpPdf)
  }
  pdftools::pdf_length(tmpPdf)
}






#' Print the libre office macro used to put everything on one page with libreoffice headless conversion to pdf
#'
#' This function is more for documentation purposes, until I find a better way to install it.
#'
#' @returns Prints the macro
#' @md
onePageMacro <- function() {

  "Sub MakeOnePagePDF(inputFile As String, outputFile As String)
  Dim document As Object
  Dim sheets As Object
  Dim sheet As Object
  Dim pageStyle As Object
  Dim pageStyles As Object
  Dim exportSettings(1) As New com.sun.star.beans.PropertyValue

  ' Open the spreadsheet
    document = StarDesktop.loadComponentFromURL(ConvertToURL(inputFile), \"_blank\", 0, Array())

    ' Access sheets
  sheets = document.getSheets()
  sheet = sheets.getByIndex(0)  ' Adjust if you want a different sheet

    ' Set print area to all used cells
  Dim usedRange As Object
  usedRange = sheet.getCellRangeByPosition(0, 0, sheet.Columns.Count - 1, sheet.Rows.Count - 1)
  sheet.setPrintAreas(Array(usedRange.getRangeAddress()))

  ' Access page style directly
    pageStyles = document.StyleFamilies.getByName(\"PageStyles\")
    pageStyle = pageStyles.getByName(sheet.PageStyle)

    ' Scale to fit on a single page
  pageStyle.ScaleToPagesX = 1
  pageStyle.ScaleToPagesY = 1

  ' Export to PDF
    exportSettings(0).Name = \"FilterName\"
    exportSettings(0).Value = \"calc_pdf_Export\"
    exportSettings(1).Name = \"FilterData\"
    exportSettings(1).Value = Array()
    document.storeToURL(ConvertToURL(outputFile), exportSettings())

    ' Close without saving
  document.close(True)
  End Sub
  "
}

