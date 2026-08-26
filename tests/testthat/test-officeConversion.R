test_that("convertOfficeToPdf validates its inputs", {
  expect_error(
    convertOfficeToPdf("does-not-exist.xlsx", tempfile(fileext = ".pdf")),
    "Input file does not exist"
  )

  input <- tempfile(fileext = ".txt")
  writeLines("not an Office document", input)
  expect_error(
    convertOfficeToPdf(input, tempfile(fileext = ".pdf")),
    "Unsupported Office file extension"
  )
})

test_that("convertOfficeToPdf exports Excel in single and A4 modes", {
  skip_on_cran()
  skip_if(Sys.info()["sysname"] != "Windows")
  excelAvailable <- system2(
    "powershell.exe",
    args = c(
      "-NoProfile",
      "-Command",
      shQuote("if ([type]::GetTypeFromProgID('Excel.Application')) { exit 0 } else { exit 1 }")
    ),
    stdout = FALSE,
    stderr = FALSE
  ) == 0
  skip_if_not(excelAvailable, "Microsoft Excel is not installed")

  fileIn <- system.file("exampleData/plots.xlsx", package = "excelPlot")
  single <- tempfile(fileext = ".pdf")
  a4 <- tempfile(fileext = ".pdf")
  on.exit(unlink(c(single, a4)), add = TRUE)

  expect_invisible(convertOfficeToPdf(fileIn, single, pageSize = "single"))
  expect_invisible(convertOfficeToPdf(fileIn, a4, pageSize = "A4"))
  expect_equal(pdftools::pdf_length(single), 1L)
  expect_gt(pdftools::pdf_length(a4), 1L)
})

test_that("Excel inputs use a PDF intermediate in the PNG pipeline", {
  fileIn <- system.file("exampleData/plots.xlsx", package = "excelPlot")
  files <- epFiles(path = fileIn)

  expect_equal(tools::file_ext(files$tmpPathCommitPdf), "pdf")
})