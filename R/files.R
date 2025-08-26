
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
#' path <- system.file("exampleData/01-Iris.pdf", package = "excelPlot")
#' epFiles(path, commit = "HEAD", page = 1)
epFiles <- function(path, commit = "HEAD", page = 1, xmin = 0, xmax = 100, ymin = 0, ymax = 100, ...) {

  path <- normalizePath(path, mustWork = TRUE, winslash = .Platform$file.sep)

  # Try to eat as much of the base folder away
  pathTrunc <- path
  pathTrunc <- gsub(paste0("^",normalizePath("~")), "", pathTrunc)
  pathTrunc <- gsub(paste0("^",dirname(normalizePath("~"))), "", pathTrunc)
  pathTrunc <- gsub(paste0("^[A-Z]:"), "", pathTrunc)

  path                  = path
  tmpPathCommit         = paste0(file.path(tempdir(), paste0(tools::file_path_sans_ext(pathTrunc), "-commit-", commit, ".", tools::file_ext(path))))
  tmpPathCommitPage     = paste0(tools::file_path_sans_ext(tmpPathCommit), sprintf("-page-%02d.png", page))
  tmpPathCommitPageCrop = paste0(tools::file_path_sans_ext(tmpPathCommitPage), sprintf("-crop-%03d-%03d-%03d-%03d.png", xmin, xmax, ymin, ymax))

  list(
    path                  = path,
    tmpPathCommit         = tmpPathCommit,
    tmpPathCommitPage     = tmpPathCommitPage,
    tmpPathCommitPageCrop = tmpPathCommitPageCrop)
}




