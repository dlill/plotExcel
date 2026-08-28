#' Muffle expected type-coercion warnings that fire from data.table melt/dcast
#' in `parseTable()` / `plotExcelDiff()`. Specifically:
#'   * "NAs introduced by coercion"
#'   * "'measure.vars' [...] are not all of the same type"
#' Other warnings propagate normally so they still surface real issues.
suppressNAcoercion <- function(expr) {
  patterns <- c("NAs introduced by coercion",
                "are not all of the same type")
  withCallingHandlers(
    expr,
    warning = function(w) {
      msg <- conditionMessage(w)
      if (any(vapply(patterns, grepl, logical(1), x = msg, fixed = TRUE))) {
        invokeRestart("muffleWarning")
      }
    }
  )
}
