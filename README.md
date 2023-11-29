
# odconvertR

odsconvertR is an R package which allows interconverting between XLSX and ODS formats within R. ODS files are an open format of data table and preferable to XLSX for publication. While some packages within R allow writing directly to an ODS file, they do not allow writing to a template or more complex formatting, limiting how much you can change the appearance of the resulting file. Reading ODS files in R is also complicated and limited, so conversion back to XLSX is also useful for this purpose.

This package uses a Visual Basic script (VBS) to save a copy of an existing XLSX file in an ODS format (and vice versa), retaining data structure and all formatting. This is particularly useful for Reproducible Analytical Pipelines for publication of data; nicely formatted and readable ODS data tables can be produced without the need for manual point-and-click steps.

## Installation

You can install odsconvertR with:

    install.packages("devtools")
    devtools::install_github("department-for-transport-public/odsconvertr")

## Usage

**`convert_to_ods`**: for conversion of XLSX files to ODS

This function takes one argument:

`xlsx_path`: A string containing the path to the existing xlsx file. This can be a relative path (from the working directory) or an absolute path (the full path to the file). Importantly, because the code passes through VBA, you cannot use the `"../"` notation to move up in a working directory, you must provide the full file path in this instance.

When called, the function will create an ODS file with the same name, in the same location as the XLSX file. The original file will not be deleted.

**`convert_to_xlsx`**: for conversion of ODS files to XLSX

This function takes one argument:

`ods_path`: A string containing the path to the existing ods file. This can be a relative path (from the working directory) or an absolute path (the full path to the file). Importantly, because the code passes through VBA, you cannot use the `"../"` notation to move up in a working directory, you must provide the full file path in this instance.

When called, the function will create an XLSX file with the same name, in the same location as the ODS file. The original file will not be deleted.

**`evaluate_formulas`**: for evaluating formulas in automatically generated xlsx files. This is useful when an Excel file is created via code with formulas in, and you want to be able to read this in without manually opening the file.

The function takes a single argument, the filepath of the file. Importantly, because the code passes through VBA, you cannot use the `"../"` notation to move up in a working directory, you must provide the full file path in this instance.


