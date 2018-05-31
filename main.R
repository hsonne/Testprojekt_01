install.packages(c("fs", "tidyr", "dplyr", "stringr", "readxl", "devtools", 
                   "crayon", "data.table"))
devtools::install_github("kwb-r/kwb.utils")
                 

source(file = "R/convert_xls_as_xlsx.R")
source(file = "R/copy_xlsx_files.R")
source(file = "R/read_bwb_data.R")

path_list <- list(
  #drive = "//medusa/projekte$/Z-Exchange/Jeansen",
  #drive = file.path(kwb.utils::get_homedir(), "Downloads"),
  drive = "F:",
  input_dir = "<drive>/Daten_Labor",
  export_dir = "<drive>/ANALYSIS_R/tmp")

paths <- kwb.utils::resolve(x = path_list)

### Get location of excelcnv.exe 
get_excelcnv_exe()


### Convert old ".xls" to ".xlsx" Excel files
convert_xls_as_xlsx(input_dir = paths$input_dir, 
                    export_dir = paths$export_dir)


### Copy remaining already existing .xlsx files in same directory 
copy_xlsx_files(input_dir = paths$input_dir, 
                export_dir = paths$export_dir)


#### Get all xlsx files to be imported
xlsx_files <- dir(paths$export_dir,
                  pattern = ".xlsx", 
                  recursive = TRUE, 
                  full.names = TRUE)


labor <- import_labor(xlsx_files = xlsx_files,
                      export_dir = paths$export_dir)
