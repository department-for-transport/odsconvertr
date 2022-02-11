#' Uses VBA code to save a copy of an xlsx file to the accessible ODS format, retaining all sheets and formatting.
#'
#' Files converted can be via a relative (from working directory) or absolute (full) file path. In either case, the output ODS file will be returned in the same folder as the XLSX file.
#' @export
#' @param path path to xlsx file; can be either a relative or absolute file path
#' @name convert_to_ods
#' @title Save a copy of an xlsx file as an ods file
#'
convert_to_ods <- function(path){

  ##Stop if file is not found
  if(file.exists(path) == FALSE){
    stop("File not found")
  }

  #Stop if file is not an xlsx
  if(grepl(".xlsx", path, fixed = TRUE) == FALSE){
    stop("File is not an xlsx file")
  }

  ##Convert path to absolute one
  xlsx_all <-  paste0('"',normalizePath(path),  '"')

  ods_all <-  gsub(".xlsx", ".ods", xlsx_all, fixed = TRUE)


  ##Get path of VBS script inside package
  vbs_loc <- vbs_file_path('save.vbs')


  #Run VBS script passing it the file paths
  vbs_execute("save.vbs",
              xlsx_all,
              ods_all)

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

  ##Stop if file is not found
  if(file.exists(path) == FALSE){
    stop("File not found")
  }

  #Stop if file is not an xlsx
  if(grepl(".ods", path, fixed = TRUE) == FALSE){
    stop("File is not an ods file")
  }
  ##Convert path to absolute one
  ods_all <- paste0('"',normalizePath(path),  '"')

  xlsx_all <-  gsub(".ods", ".xlsx", ods_all, fixed = TRUE)


  ##Get path of VBS script inside package
  vbs_loc <- vbs_file_path('save.vbs')


  #Run VBS script passing it the file paths
  vbs_execute("save_xlsx.vbs",
              ods_all,
              xlsx_all)

}
