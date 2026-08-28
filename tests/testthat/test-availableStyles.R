test_that("availableStyles lists styles and decorator syntax", {
  output <- capture.output(availableStyles())
  outputText <- paste(output, collapse = "\n")
  plotDecoratorKeys <- setdiff(names(formals(plotSpec)), "path")

  expect_match(outputText, "Available text styles:", fixed = TRUE)
  expect_true(all(vapply(
    names(styleList),
    function(style) grepl(paste0(": ", style), outputText, fixed = TRUE),
    logical(1)
  )))
  expect_match(outputText, "path::key value::key2 value2", fixed = TRUE)
  expect_match(outputText, paste("Valid keys:", paste(plotDecoratorKeys, collapse = ", ")), fixed = TRUE)
  expect_match(outputText, '"text::style"', fixed = TRUE)
  expect_match(outputText, "diff(col1, col2)", fixed = TRUE)
})