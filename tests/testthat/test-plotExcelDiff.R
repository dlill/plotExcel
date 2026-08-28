runPlotExcelDiffCase <- function(f1, f2, testName) {
  ex <- system.file("exampleData", package = "plotExcel")
  skip_if(!nzchar(ex), "example data not installed")
  p1 <- file.path(ex, f1)
  p2 <- file.path(ex, f2)
  skip_if_not(file.exists(p1) && file.exists(p2), paste("missing example:", p1, "/", p2))

  out <- tempfile(pattern = paste0("file", testName), fileext = ".xlsx")

  res <- suppressNAcoercion(
    plotExcelDiff(p1, p2, filename = out, FLAGtemp = FALSE, FLAGopenExcel = FALSE)
  )

  expect_true(file.exists(out))
  expect_gt(file.info(out)$size, 0)
  expect_equal(normalizePath(res, mustWork = FALSE), normalizePath(out, mustWork = FALSE))
}

test_that("plotExcelDiff works for two PDF files", {
  runPlotExcelDiffCase("01-Iris.pdf", "02-Iris-Brewer.pdf", "pdf")
})

test_that("plotExcelDiff works for two PNG files", {
  runPlotExcelDiffCase("01-Iris.png", "03-Iris-scale2.png", "png")
})

test_that("plotExcelDiff works for two PPTX files", {
  runPlotExcelDiffCase("11-Slides.pptx", "12-Slides.pptx", "pptx")
})

test_that("plotExcelDiff works for two DOCX files", {
  runPlotExcelDiffCase("21-Word.docx", "22-Word.docx", "docx")
})

test_that("plotExcelDiff compares one file against a commit", {
  testDir <- tempfile("plotExcelDiff-git-")
  dir.create(testDir, recursive = TRUE)
  on.exit(unlink(testDir, recursive = TRUE), add = TRUE)

  plotFile <- file.path(testDir, "plot.png")
  magick::image_write(magick::image_blank(100, 100, color = "white"), plotFile)
  git2r::init(testDir)
  git2r::add(testDir, "plot.png")
  git2r::commit(testDir, message = "Initial plot", all = TRUE)
  commit <- git2r::reflog(testDir)[[1]][[1]]
  magick::image_write(magick::image_blank(100, 100, color = "black"), plotFile)

  layout <- plotExcelDiff(
    plotFile,
    commit = commit,
    FLAGopenExcel = FALSE,
    CFLAGLayout = "return"
  )

  expect_equal(layout$File1[[2]], paste0(plotFile, "::page 1::resolution 100"))
  expect_equal(layout$File2[[2]], paste0(plotFile, "::commit ", commit, "::page 1::resolution 100"))
  expect_equal(layout$File2[[1]], paste0("* ", plotFile, " (commit ", commit, ")"))

  out <- tempfile(fileext = ".xlsx")
  result <- suppressNAcoercion(
    plotExcelDiff(plotFile, commit = commit, filename = out, FLAGtemp = FALSE, FLAGopenExcel = FALSE)
  )
  expect_true(file.exists(out))
  expect_gt(file.info(out)$size, 0)
  expect_equal(normalizePath(result), normalizePath(out))
})

test_that("plotExcelDiff requires exactly one of file2 and commit", {
  expect_error(
    plotExcelDiff("plot.png", FLAGopenExcel = FALSE, CFLAGLayout = "return"),
    "`commit` must be supplied"
  )
  expect_error(
    plotExcelDiff("plot.png", "other.png", commit = "HEAD", FLAGopenExcel = FALSE, CFLAGLayout = "return"),
    "`commit` may only be supplied"
  )
})

test_that("plotExcelDiff aligns pages via skip1 when the second file has extra slides", {
  ex <- system.file("exampleData", package = "plotExcel")
  skip_if(!nzchar(ex), "example data not installed")
  p1 <- file.path(ex, "12-Slides.pptx")
  p2 <- file.path(ex, "13-Slides-insertedDummySlidesForSkipping.pptx")
  skip_if_not(file.exists(p1) && file.exists(p2), paste("missing example:", p1, "/", p2))

  # 13-Slides has extra slides inserted at positions 1, 3, 5, 6 vs 12-Slides,
  # so skipping those output rows for File1 aligns the shared slides side-by-side.
  skip1 <- c(1, 3, 5, 6)

  d <- suppressNAcoercion(
    plotExcelDiff(p1, p2, skip1 = skip1, FLAGtemp = FALSE, FLAGopenExcel = FALSE, CFLAGLayout = "return")
  )

  # Drop the subheader row (Page == "") for the page-level checks
  dPages <- d[Page != ""]

  # File1 cells at skipped rows must be blank; other rows must reference file1
  expect_true(all(is.na(dPages$File1[skip1])))
  expect_true(all(grepl(basename(p1), dPages$File1[-skip1], fixed = TRUE)))

  # File2 cells: full coverage over the shared row range (no NAs where File1 is present)
  expect_true(all(!is.na(dPages$File2[-skip1])))

  # Diff column: aligned rows use the diff formula, skipped rows fall back to the "differ" message
  expect_true(all(dPages$Diff[-skip1] == "diff(File1, File2)"))
  expect_true(all(grepl("Page lengths differ", dPages$Diff[skip1])))

  # Also confirm the Excel export path still works with the skip argument
  out <- tempfile(pattern = "fileSkip1", fileext = ".xlsx")
  res <- suppressNAcoercion(
    plotExcelDiff(p1, p2, filename = out, skip1 = skip1, FLAGtemp = FALSE, FLAGopenExcel = FALSE)
  )
  expect_true(file.exists(out))
  expect_gt(file.info(out)$size, 0)
  expect_equal(normalizePath(res, mustWork = FALSE), normalizePath(out, mustWork = FALSE))
})

test_that("plotExcelDiff reports supported types and unequal page counts", {
  ex <- system.file("exampleData", package = "plotExcel")
  skip_if(!nzchar(ex), "example data not installed")
  p1 <- file.path(ex, "11-Slides.pptx")
  p2 <- file.path(ex, "12-Slides.pptx")
  skip_if_not(file.exists(p1) && file.exists(p2), paste("missing example:", p1, "/", p2))

  messages <- capture.output(
    layout <- plotExcelDiff(
      file1 = p1,
      file2 = p2,
      FLAGtemp = FALSE,
      FLAGopenExcel = FALSE,
      CFLAGLayout = "return"
    ),
    type = "message"
  )

  expect_equal(names(formals(plotExcelDiff))[1:2], c("file1", "file2"))
  expect_true(any(grepl("Supported.*pdf.*png.*docx.*pptx.*xlsx", messages)))
  expect_true(any(grepl("Use `skip1` or `skip2` to align corresponding pages.", messages, fixed = TRUE)))
  expect_s3_class(layout, "data.table")
})