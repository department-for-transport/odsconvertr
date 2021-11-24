#' Uses VBA code to save a copy of an xlsx file to the accessible ODS format, retaining all sheets and formatting.
#'
#' Files converted can be via a relative (from working directory) or absolute (full) file path. In either case, the output ODS file will be returned in the same folder as the XLSX file.
#' @export
#' @param path path to xlsx file; can be either a relative or absolute file path
#' @name convert_to_ods
#' @title Save a copy of an xlsx file as an ods file
#'
convert_to_ods <- function(path){

  ##Normalise path to an absolute one with backslashes and surrounding quotation marks
  path <- paste('"', normalizePath(path), '"')

  ods_path <- gsub("xlsx", "ods", path)


  ##Get path of VBS script inside package
  vbs_loc <- paste0('"',
                    system.file("vbs", package = "odsconvertr"),
                    '/save.vbs"')


  #Run VBS script passing it the file paths
  system_command <- paste("WScript",
                          vbs_loc,
                          path,
                          ods_path,
                          sep = " ")
  system(command = system_command)

}

#' Uses VBA code to save a copy of an ods file to the easy to use xlsx format, retaining all sheets and formatting.
#'
#' Files converted can be via a relative (from working directory) or absolute (full) file path. In either case, the output ODS file will be returned in the same folder as the XLSX file.
#' @export
#' @param path path to ods file; can be either a relative or absolute file path
#' @name convert_to_xlsx
#' @title Save a copy of an ods file as an xlsx file
#'
convert_to_xlsx <- function(path){

  ##Normalise path to an absolute one with backslashes and surrounding quotation marks
  path <- paste('"', normalizePath(path), '"')

  xlsx_path <- gsub("ods", "xlsx", path)

  ##Get path of VBS script inside package
  vbs_loc <- paste0('"',
                    system.file("vbs", package = "odsconvertr"),
                    '/save_xlsx.vbs"')


  #Run VBS script passing it the file paths

  system_command <- paste("WScript",
                          vbs_loc,
                          path,
                          xlsx_path,
                          sep = " ")
  system(command = system_command)

}

