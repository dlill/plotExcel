test_that("plotExcel announces FLAGpdf conversion before it starts", {
  plotFile <- system.file("exampleData/01-Iris.png", package = "plotExcel")
  skip_if(!nzchar(plotFile), "example data not installed")
  outputFile <- tempfile(fileext = ".xlsx")
  on.exit(unlink(outputFile), add = TRUE)

  local_mocked_bindings(
    convertOfficeToPdf = function(...) stop("conversion reached"),
    .package = "plotExcel"
  )

  expect_message(
    expect_error(
      suppressWarnings(
        plotExcel(
          data.table::data.table(Plot = plotFile),
          filename = outputFile,
          FLAGpdf = TRUE
        )
      ),
      "conversion reached",
      fixed = TRUE
    ),
    "Generating a pdf from the Excel file since FLAGpdf = TRUE.",
    fixed = TRUE
  )
})