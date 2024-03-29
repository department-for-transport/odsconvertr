% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/evaluate_formulas.R
\name{evaluate_formula}
\alias{evaluate_formula}
\title{Evaluate all formulae in an existing Excel file}
\usage{
evaluate_formula(path)
}
\arguments{
\item{path}{path to Excel file; can be either a relative or absolute file path}
}
\description{
Uses VBA code to open an Excel file and to recalculate any formula in it.
This ensures that all Excel formulae have been executed in a file which may have been updated without being opened (e.g. using code)
}
