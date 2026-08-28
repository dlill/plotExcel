test_that("plotExcelFolder can dump all files in inst/exampleData", {
  ex <- system.file("exampleData", package = "plotExcel")
  skip_if(!nzchar(ex), "example data not installed")

  out <- tempfile(fileext = ".xlsx")

  res <- suppressNAcoercion(
    plotExcelFolder(
      path = ex,
      filename = out,
      FLAGtemp = FALSE,
      FLAGopenExcel = FALSE,
      nPagesMax = 4,
      resolution = 100
    )
  )

  expect_true(file.exists(out))
  expect_gt(file.info(out)$size, 0)
  expect_equal(normalizePath(res, mustWork = FALSE), normalizePath(out, mustWork = FALSE))
})

test_that("plotExcelFolder excludes xlsx outputs from returned layouts", {
  ex <- system.file("exampleData", package = "plotExcel")
  skip_if(!nzchar(ex), "example data not installed")

  messages <- capture.output(
    layout <- plotExcelFolder(
      path = ex,
      FLAGtemp = FALSE,
      FLAGopenExcel = FALSE,
      CFLAGLayout = "return"
    ),
    type = "message"
  )

  expect_true("xlsx" %in% SUPPORTED_PLOT_EXTENSIONS)
  expect_true(all(c("html", "htm") %in% SUPPORTED_PLOT_EXTENSIONS))
  expect_true(any(grepl("Supported.*pdf.*png.*docx.*pptx", messages)))
  expect_false(any(grepl("\\.xlsx::", layout$Plot, ignore.case = TRUE)))
})
