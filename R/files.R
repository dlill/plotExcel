
#' Construct all temp plot file names from a plotSpec
#'
#' @inheritParams plotSpec
#' @param ... Sink for unused
#'
#' @return List of paths to the respective files
#' @export
#' @md
#' @family UI
#' @importFrom tools file_path_sans_ext file_ext
#'
#' @examples
#' path <- system.file("exampleData/01-Iris.pdf", package = "plotExcel")
#' epFiles(path, commit = "HEAD", page = 1)
epFiles <- function(path, commit = "HEAD", page = 1, resolution = 100, xmin = 0, xmax = 100, ymin = 0, ymax = 100, ...) {

  path <- normalizePath(path, mustWork = TRUE, winslash = .Platform$file.sep)
  basepath <- basename(path)
  dirpath <- dirname(path)
  dirpathDigest <- substr(digest::digest(dirpath),1,12)

  # Extension for the "converted to pdf" intermediate:
  # - Office files (docx/doc/pptx/ppt/xlsx/xlsm/xls) will be converted to pdf.
  # - pdf/png files are just copied and keep their extension.
  sourceExt <- tolower(tools::file_ext(path))
  pdfExt <- if (sourceExt %in% c("docx", "doc", "pptx", "ppt", "xlsx", "xlsm", "xls")) "pdf" else sourceExt

  path                  = path
  tmpPathCommit         = paste0(file.path(tempdir(), dirpathDigest, paste0(tools::file_path_sans_ext(basepath), "-commit-", commit, ".", tools::file_ext(path))))
  tmpPathCommitPdf      = paste0(tools::file_path_sans_ext(tmpPathCommit), "-topdf.", pdfExt)
  tmpPathCommitPage     = paste0(tools::file_path_sans_ext(tmpPathCommitPdf), sprintf("-page-%02d-res-%02d.png", page, resolution))
  tmpPathCommitPageCrop = paste0(tools::file_path_sans_ext(tmpPathCommitPage), sprintf("-crop-%03d-%03d-%03d-%03d.png", xmin, xmax, ymin, ymax))

  list(
    path                  = path,
    tmpPathCommit         = tmpPathCommit,
    tmpPathCommitPdf      = tmpPathCommitPdf,
    tmpPathCommitPage     = tmpPathCommitPage,
    tmpPathCommitPageCrop = tmpPathCommitPageCrop)
}




