test_that("Various steps of the pipeline work", {
  library(ggplot2)

  testDir <- file.path(tempdir(), "excelPlot/test-pngPipeline")
  unlink(testDir, recursive = TRUE)
  dir.create(testDir, showWarnings = FALSE, recursive = TRUE)

  git2r::init(testDir)

  pl <- ggplot(iris, aes(Sepal.Length, Sepal.Width, color = Species)) +
    geom_point(shape = 1) +
    theme_bw() +
    geom_blank()
  ggsave(plot = pl, filename = file.path(testDir, "01-Iris.pdf"), width = 15.5, height = 10, scale = 1, units = "cm")

  git2r::add(testDir, "01-Iris.pdf")
  git2r::commit(testDir, message = "First plot", all = TRUE)

  ggsave(plot = pl + scale_color_brewer(), filename = file.path(testDir, "01-Iris.pdf"), width = 15.5, height = 10, scale = 1, units = "cm")
  git2r::add(testDir, "01-Iris.pdf")
  git2r::commit(testDir, message = "Color brewer", all = TRUE)

  # -------------------------------------------------------------------------#
  # Run the pipeline ----
  # -------------------------------------------------------------------------#
  # .. Checkout to temp -----
  plotSpec <- plotSpec(file.path(testDir, "01-Iris.pdf"), commit = git2r::reflog(testDir)[[1]][[1]])
  files <- do.call(epFiles, plotSpec)
  pngPipelineCheckoutToTemp(plotSpec)
  pngPipelineExtractPage(plotSpec)
  pngPipelineCrop(plotSpec)

  # .. File exists, commit is HEAD and output file is newer: No action required -----
  plotSpec <- plotSpec(file.path(testDir, "01-Iris.pdf"), commit = "HEAD")
  files <- do.call(epFiles, plotSpec)
  pngPipelineCheckoutToTemp(plotSpec)
  expect_true(idempotencyNoActionRequired(fileIn = files$path, fileOut = files$tmpPathCommit, commit = plotSpec$commit))

  # .. File exists, commit is HEAD and output file is older: Action required -----
  plotSpec <- plotSpec(file.path(testDir, "01-Iris.pdf"), commit = "HEAD")
  files <- do.call(epFiles, plotSpec)
  pngPipelineCheckoutToTemp(plotSpec)
  Sys.sleep(1)
  ggsave(plot = pl + scale_color_brewer(palette = 2), filename = file.path(testDir, "01-Iris.pdf"), width = 15.5, height = 10, scale = 1, units = "cm")
  expect_false(idempotencyNoActionRequired(fileIn = files$path, fileOut = files$tmpPathCommit, commit = plotSpec$commit))


  # -------------------------------------------------------------------------#
  # Cleanup ----
  # -------------------------------------------------------------------------#
  unlink(testDir,recursive = TRUE)
})

