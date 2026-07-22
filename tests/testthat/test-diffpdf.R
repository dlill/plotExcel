runDiffpdfCase <- function(f1, f2, testName) {
  ex <- system.file("exampleData", package = "excelPlot")
  skip_if(!nzchar(ex), "example data not installed")
  p1 <- file.path(ex, f1)
  p2 <- file.path(ex, f2)
  skip_if_not(file.exists(p1) && file.exists(p2), paste("missing example:", p1, "/", p2))

  out <- tempfile(pattern = paste0("file", testName), fileext = ".xlsx")

  res <- suppressNAcoercion(
    diffpdf(p1, p2, filename = out, FLAGtemp = FALSE, FLAGopenExcel = FALSE)
  )

  expect_true(file.exists(out))
  expect_gt(file.info(out)$size, 0)
  expect_equal(normalizePath(res, mustWork = FALSE), normalizePath(out, mustWork = FALSE))
}

test_that("diffpdf works for two PDF files", {
  runDiffpdfCase("01-Iris.pdf", "02-Iris-Brewer.pdf", "pdf")
})

test_that("diffpdf works for two PNG files", {
  runDiffpdfCase("01-Iris.png", "03-Iris-scale2.png", "png")
})

test_that("diffpdf works for two PPTX files", {
  runDiffpdfCase("11-Slides.pptx", "12-Slides.pptx", "pptx")
})

test_that("diffpdf works for two DOCX files", {
  runDiffpdfCase("21-Word.docx", "22-Word.docx", "docx")
})
