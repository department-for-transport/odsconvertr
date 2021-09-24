
# odconvertR

odsconvertR is an R package which allows converting XLSX files to ODS format within R. ODS files are an open format of data table and preferable to XLSX for publication. While some packages within R allow writing directly to an ODS file, they do not allow writing to a template or more complex formatting, limiting how much you can change the appearance of the resulting file.

This package uses a Visual Basic script (VBS) to save a copy of an existing XLSX file in an ODS format, retaining data structure and all formatting. This is particularly useful for Reproducible Analytical Pipelines for publication of data; nicely formatted and readable ODS data tables can be produced without the need for manual point-and-click steps.

## Installation

You can install odsconvertR with:

    install.packages("devtools")
    devtools::install_github("departmentfortransport/odsconvertr")

or

    install.packages("devtools")
    devtools::install_git("git://github.com/departmentfortransport/odsconvertr.git")

## Usage

The package contains a single function `convert_to_ods`.

This function takes two arguments:

`xlsx_path`: A string containing the path to the existing xlsx file. This can be a relative path (from the working directory) or an absolute path (the full path to the file). Importantly, because the code passes through VBA, you cannot use the `"../"` notation to move up in a working directory, you must provide the full file path in this instance.

`relative_file_path`: A boolean which indicates whether you have passed a relative or absolute file path to the first argument. By default this is set to TRUE, and expects a relative file path.

When called, the function will create an ODS file with the same name, in the same location as the XLSX file. The original file will not be deleted.
