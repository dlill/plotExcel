test_that("plotExcelFolder can dump all files in inst/exampleData", {
  ex <- system.file("exampleData", package = "excelPlot")
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
