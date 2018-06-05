# Install packages from github
#devtools::install_github("kwb-r/kwb.utils")
#devtools::install_github("kwb-r/kwb.event")
#devtools::install_github("kwb-r/kwb.fakin")

library(kwb.utils)

# Names of required packages in alphabetical order
packages <- c(
  "crayon", "data.table", "devtools", "dplyr", "fs", "readxl", "stringr",
  "tidyr"
)

# Names of installed packages
installed_packages <- rownames(installed.packages())

# Install non-installed packages from CRAN
missing_packages <- setdiff(packages, installed_packages)

if (length(missing_packages)) {
  
  install.packages(missing_packages)
}

# Define paths to scripts with functions
script_paths <- file.path("./R", c(
  "convert_xls_as_xlsx.R",
  "copy_xlsx_files.R",
  "file_database.R",
  "get_raw_text_from_xlsx.R",
  "get_tables_from_xlsx.R",
  "print_table_summary.R",
  "read_bwb_data.R",
  "utils.R"
))

# Check if all scripts exist
script_paths <- sapply(script_paths, safePath)

# Source all scripts
sourceScripts(script_paths)

# Define paths and resolve placeholders
paths <- list(
  drive_jeansen = "//medusa/projekte$/Z-Exchange/Jeansen",
  drive_stick = "F:",
  drive_hauke_home = "<downloads>/Unterstuetzung/Michael",
  downloads = "<home>/Downloads",
  input_dir = "<drive>/Daten_Labor",
  export_dir = "<drive>/ANALYSIS_R/tmp",
  export_dir_allg = "<export_dir>/K-TL_LSW-Altdaten-Werke Teil 1/Werke Teil 1/Allgemein",
  home = get_homedir()
)

paths <- kwb.utils::resolve(paths, drive = "drive_jeansen")

# Set input directory
input_dir <- kwb.utils::safePath(selectElements(paths, "input_dir"))

# Set directory in which to provide all xlsx files
export_dir <- kwb.utils::safePath(selectElements(paths, "export_dir"))

if (FALSE)
{
  # Get location of excelcnv.exe
  get_excelcnv_exe()
  
  # Convert xls to xlsx Excel files
  convert_xls_as_xlsx(input_dir, export_dir)
  
  # Copy remaining already existing .xlsx files in same directory
  copy_xlsx_files(input_dir, export_dir, overwrite = TRUE)
}

# Get all xlsx files to be imported
files <- dir(export_dir, ".xlsx", recursive = TRUE, full.names = TRUE)

file_database <- to_file_database(files)

# files:
# file_id             file_name folder_id
# file_01 haesslicher name.xlsx folder_01

# folders:
# folder_id folder_path
# - attr(*, "base_dir")

file_database$files
file_database$folders

if (FALSE)
{
  labor <- import_labor(files, export_dir = export_dir)
  
  labor_list <- import_labor(files, export_dir = export_dir, func = read_bwb_data)
  
  indices <- which(sapply(labor_list, inherits, "try-error"))
  
  # for (file in files[indices]) {
  # try(expr = read_bwb_data(file))
  # }
  files_with_problems <- files[indices]
  print(sprintf("number of xlsx files with problems: %d", 
         length(files_with_problems)))
        
  file <- files_with_problems[3]

  ### Check problems in xlsx file
  tmp_file <- read_bwb_data(file)
  
  ### Open original file and modify it
  org_files <- dir(input_dir, 
      recursive = TRUE, 
      full.names = TRUE,
      pattern = "\\.[xX][lL][sS][xX]?")
  
  kwb.utils::removeExtension(basename(file))
  
  idx <- stringr::str_detect(
    string = org_files,
    pattern = kwb.utils::removeExtension(basename(file)))
  
  org_file <- org_files[idx]
  
  kwb.utils::hsOpenWindowsExplorer(org_file)
  
  ####
  is_xls <- stringr::str_detect(string = org_file, 
                      pattern = "\\.([xX][lL][sS])$")
  
  if(is_xls) {
    convert_xls_as_xlsx(dirname(org_file),
                        export_dir)
  } else {
    
    copy_xlsx_files(from_dir = dirname(org_file), 
                    to_dir = dirname(gsub(x = org_file, 
                                  pattern = input_dir, 
                                  replacement = export_dir,
                                  fixed = TRUE)),
                    overwrite = TRUE)
  }
  
  output_files <- file.path(export_dir, c(
    "read_bwb_data_Rconsole.txt", 
    "read_bwb_header1_meta_Rconsole.txt", 
    "read_bwb_header2_Rconsole.txt"
  ))
  
  capture.output(file = output_files[1], {
    
    labor <- read_bwb_data(files)
  })
  
  capture.output(file = output_files[2], {
    
    labor_header1_meta <- import_labor(files, export_dir, read_bwb_header1_meta)
  })
  
  capture.output(file = output_files[3], {
    
    labor_header2 <- import_labor(files, export_dir, read_bwb_header2)
  })
  
  # Show file contents
  catLines(readLines(output_files[1]))
  catLines(readLines(output_files[2]))
  catLines(readLines(output_files[3]))
  
  length(which(sapply(labor_list, is.data.frame)))
  length(which(sapply(labor_header2, is.data.frame)))
  length(which(sapply(labor_header1_meta, is.data.frame)))
  
  nrow(labor)
  
  sum(unlist(sapply(labor_list, nrow)))
  sum(unlist(sapply(labor_header2, nrow)))
  sum(unlist(sapply(labor_header1_meta)))
  
  packrat::snapshot()
  
  # Tests
  read_bwb_header2(file = files[1])
}
