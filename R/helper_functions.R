#' Create a filepath for a referenced VBS file
#'
#' Formats the name of a VBS file to be used
#' @param vbs_file Name of a vbs script file saved as part of the package
#' @name vbs_file_path
#' @title Convert an R file path to one that can be used by VBS

vbs_file_path <- function(vbs_file){
  ##Get path of VBS script inside package
  paste0('"',
         system.file("vbs", package = "odsconvertr"),
         '/', vbs_file,'"')
}


#' Create a filepath for a referenced VBS file
#'
#' Formats the name of a VBS file to be used
#' @param vbs_file Name of a vbs script file saved as part of the package
#' @param ... Arguments to be passed to the specified VBS script
#' @name vbs_execute
#' @title Execute a VBS script including arguments

vbs_execute <- function(vbs_file, ...){

  ##Convert arguments into a single string
  arguments <- paste(c(...), collapse = " ")
  #Create system command for specified vbs file
  system_command <- paste("WScript",
                          vbs_file_path(vbs_file),
                          arguments,
                          sep = " ")
  #Run specified command
  system(command = system_command)
}
