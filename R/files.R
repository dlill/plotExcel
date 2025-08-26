
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
  path <- normalizePath(path, mustWork = TRUE)
  regexToRemove <- paste0("^(",normalizePath("~"),"|", "[:alpha:]:", ")")

  path                  = path
  tmpPathCommit         = paste0(file.path(tempdir(), tools::file_path_sans_ext(gsub(regexToRemove, "", path))), "-commit-", commit, ".", tools::file_ext(path))
  tmpPathCommitPage     = paste0(tools::file_path_sans_ext(tmpPathCommit), sprintf("-page-%02d.png", page))
  tmpPathCommitPageCrop = paste0(tools::file_path_sans_ext(tmpPathCommitPage), sprintf("-crop-%03d-%03d-%03d-%03d.png", xmin, xmax, ymin, ymax))

  list(
    path                  = path,
    tmpPathCommit         = tmpPathCommit,
    tmpPathCommitPage     = tmpPathCommitPage,
    tmpPathCommitPageCrop = tmpPathCommitPageCrop)
}




