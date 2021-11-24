#' Uses VBA code to open an Excel file and to recalculate any formula in it.
#' This ensures that all Excel formulae have been executed in a file which may have been updated without being opened (e.g. using code)
#'
#' @export
#' @param path path to Excel file; can be either a relative or absolute file path
#' @name evaluate_formula
#' @title Evaluate all formulae in an existing Excel file
#'
evaluate_formula <- function(path){

  ##Normalise path to an absolute one with backslashes and surrounding quotation marks
  path <- paste0('"', normalizePath(path), '"')

  ##Execute specified VBS script to save xlsx files
  vbs_execute("recalculate.vbs", path)

}
