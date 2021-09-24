#' Uses VBA code to save a copy of an xlsx file to the accessible ODS format, retaining all sheets and formatting.
#'
#' Files converted can be via a relative (from working directory) or absolute (full) file path. In either case, the output ODS file will be returned in the same folder as the XLSX file.
#' @export
#' @param xlsx_path path to xlsx file; can be either a relative or absolute file path
#' @param relative_file_path boolean indicating if location of xlsx file is relative or absolute. Defaults to TRUE
#' @name convert_to_ods
#' @title Save a copy of an xlsx file as an ods file
#'
convert_to_ods <- function(xlsx_path, relative_file_path = TRUE){

  ##Remove single backslashes from xlsx path
  xlsx_path <- gsub("/", "\\", xlsx_path, fixed = TRUE)

  ##Select whether using a local or absolute file path
  if(relative_file_path){
  ##Join together wd and file name
    nice_wd <- paste0(gsub("/", "\\", getwd(), fixed = TRUE), "\\")

    ods_path <- gsub("xlsx", "ods", xlsx_path)

    xlsx_all <- paste0('"', nice_wd, xlsx_path, '"')

    ods_all <- paste0('"', nice_wd, ods_path, '"')

  } else{

    ods_path <- gsub("xlsx", "ods", xlsx_path)

    xlsx_all <- paste0('"', xlsx_path, '"')

    ods_all <- paste0('"', ods_path, '"')

  }

  ##Get path of VBS script inside package
  vbs_loc <- paste0('"',
                    system.file("vbs", package = "odsconvertr"),
                    '/save.vbs"')


  #Run VBS script passing it the file paths

  system_command <- paste("WScript",
                          vbs_loc,
                          xlsx_all,
                          ods_all,
                          sep = " ")
  system(command = system_command)

}

