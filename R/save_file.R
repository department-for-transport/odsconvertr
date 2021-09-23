#' Uses VBA code to save a copy of an xlsx file to the accessible ODS format, retaining all sheets and formatting.
#' @export
#' @param xlsx_path
#' @param local_file
#' @name convert_ods
#' @title Save a copy of an xlsx file as an ods file
#'
convert_ods <- function(xlsx_path, local_file = TRUE){

  ##Select whether using a local or absolute file path
  if(local_file){
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


  #Run VBS script passing it the file paths

  system_command <- paste("WScript",
                          '"save.vbs"',
                          xlsx_all,
                          ods_all,
                          sep = " ")
  system(command = system_command)

}


