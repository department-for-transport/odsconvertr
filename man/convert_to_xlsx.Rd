% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/save_files.R
\name{convert_to_xlsx}
\alias{convert_to_xlsx}
\title{Save a copy of an ods file as an xlsx file}
\usage{
convert_to_xlsx(path)
}
\arguments{
\item{path}{path to ods file; can be either a relative or absolute file path}
}
\description{
Uses VBA code to save a copy of an ods file to the easy to use xlsx format, retaining all sheets and formatting.
}
\details{
Files converted can be via a relative (from working directory) or absolute (full) file path. In either case, the output ODS file will be returned in the same folder as the XLSX file.
}
