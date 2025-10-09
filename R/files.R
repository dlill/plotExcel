
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
epFiles <- function(path, commit = "HEAD", page = 1, resolution = 100, xmin = 0, xmax = 100, ymin = 0, ymax = 100, ...) {

  path <- normalizePath(path, mustWork = TRUE, winslash = .Platform$file.sep)
  basepath <- basename(path)
  dirpath <- dirname(path)
  dirpathDigest <- substr(digest::digest(dirpath),1,12)

  path                  = path
  tmpPathCommit         = paste0(file.path(tempdir(), dirpathDigest, paste0(tools::file_path_sans_ext(basepath), "-commit-", commit, ".", tools::file_ext(path))))
  tmpPathCommitPage     = paste0(tools::file_path_sans_ext(tmpPathCommit), sprintf("-page-%02d-res-%02d.png", page, resolution))
  tmpPathCommitPageCrop = paste0(tools::file_path_sans_ext(tmpPathCommitPage), sprintf("-crop-%03d-%03d-%03d-%03d.png", xmin, xmax, ymin, ymax))

  list(
    path                  = path,
    tmpPathCommit         = tmpPathCommit,
    tmpPathCommitPage     = tmpPathCommitPage,
    tmpPathCommitPageCrop = tmpPathCommitPageCrop)
}




