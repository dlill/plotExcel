test_that("getNPages returns expected counts across supported file types", {
  ex <- system.file("exampleData", package = "excelPlot")
  skip_if(!nzchar(ex), "example data not installed")

  expect_equal(getNPages(file.path(ex, "01-Iris.pdf")), 1)
  expect_equal(getNPages(file.path(ex, "04-IrisMulti.pdf")), 3)
  expect_equal(getNPages(file.path(ex, "01-Iris.png")), 1)
  expect_equal(getNPages(file.path(ex, "11-Slides.pptx")), 2)
  expect_equal(getNPages(file.path(ex, "12-Slides.pptx")), 3)
  expect_equal(getNPages(file.path(ex, "21-Word.docx")), 3)
  expect_equal(getNPages(file.path(ex, "22-Word.docx")), 2)
})
