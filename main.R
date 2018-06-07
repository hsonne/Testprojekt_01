# Install packages from github
#devtools::install_github("kwb-r/kwb.utils")
#devtools::install_github("kwb-r/kwb.fakin")
#devtools::install_github("kwb-r/kwb.event")

library(kwb.utils)

# Names of required packages in alphabetical order
packages <- c(
  "crayon", "data.table", "devtools", "dplyr", "fs", "readxl", "stringr",
  "testthat", "tidyr"
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
  "convert_to_data_frames.R",  "copy_xlsx_files.R",
  "file_database.R",
  "get_raw_text_from_xlsx.R",
  "get_tables_from_xlsx.R",
  "guess.R",
  "metadata.R",
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
  drive_c = "C:/Jeansen",
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


files_meta <- c("Meta Info", 
                "Info-Altdaten", 
                "Brandenburg_Parameter_BWB_Stolpe", 
                "Kopie von Brandenburg_Parameter_BWB_Stolpe",
                "2005-10BeschilderungProbenahmestellenGWWIII", 
                "Bezeichnungen der Reinwasserstellen",
                "ReinwasserNomenklatur",
                "Info zu Altdaten 1970-1998"
                )

#files_no_sitecode <- c("Wuhlheide_Beelitzhof_Teildaten") 
                       

files_header_1 <- c("2018-04-11 Chlorid in Brunnen - Übersicht",
                    "2018-04-27 LIMS Reiw & Rohw Sammel ",
                    "2018-04-27 Rohwasser Bericht - Galeriefördermengen")
files_header_4 <- c("STO Rohw_1999-6_2004",
                    "Wuhlheide_1999-2003_Okt - Neu",
                    "KAU_1999-Okt2003", 
                    "KAU_Roh_Rein_1992-1998_TVO", 
                    "KAU_Roh_Rein_1994-1998_HKW",
                    "KAU_Rohw_Reinw_1992-1998")
                      

files_to_ignore <- c(files_meta, files_header_1, files_header_4)


cat(crayon::bgWhite(sprintf("
###############################################################################\n
Currently %d files are ignored for import:\n
Meta files:\n%s\n
Header1 (without metadata):\n%s\n
Header4:\n%s\n
##############################################################################", 
      length(files_to_ignore), 
      paste(files_meta, collapse = "\n"), 
      paste(files_header_1, collapse = "\n"), 
      paste(files_header_4, collapse = "\n"))))


in_files_to_ignore <- kwb.utils::removeExtension(basename(files)) %in% files_to_ignore


files_to_import <- files[!in_files_to_ignore]

cat(crayon::bold(crayon::green(sprintf("
##############################################################################
Importing %d out of %d files
##############################################################################", 
                            length(files_to_import), 
                            length(files)))))

file_database <- to_file_database(files_to_import)

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
  labor <- read_bwb_data(files = files_to_import)
  View(head(labor))
  
  ### Problems:
  labor[labor$`Datum@NA` == "-1102896000",c("file_name", "sheet_name")]
  labor[!is.na(labor$`NO3-@mg/l`), c("file_name", "sheet_name")]

  
  
  
labor_list <- import_labor(files_to_import, 
                           export_dir = export_dir, 
                           func = read_bwb_data)
  
 has_errors <- unlist(sapply(labor_list, inherits, "try-error"))
  
 has_no_data <-  unlist(sapply(labor_list, nrow))==0
 
 #### NO3-: 2 mal in 
 #### K-TL_LSW-Altdaten-Werke Teil 1\Werke Teil 1\Buch\BUC_Reinw_1992-1997.xlsx
 weired_cols <- c("frei@NA", "NO3-@mg/l")
 for (weired_col in weired_cols) {
 org_of_prob <- which(sapply(labor_list, function(x) any(names(x) %in% weired_col)))
 
 cat(crayon::bold(crayon::red(sprintf("Weired column '%s' found in:\n%s\n
Path(s):\n%s\n\n", 
               weired_col,
               paste(names(org_of_prob), collapse = "\n"),
               paste(normalizePath(files_to_import[org_of_prob]), 
                     collapse = "\n")))))
 }
 
 
 tmp_list <- labor_list[!has_errors]
 tmp_list <- tmp_list[!has_no_data]

 
 get_datatypes <- function(df, to_sort = FALSE, decreasing = TRUE) {
   tbl_datatypes <- table(unlist(sapply(df, class)))
   
   if (to_sort) tbl_datatypes <- sort(tbl_datatypes, decreasing)
   

   vec <- c(tbl_datatypes/sum(tbl_datatypes), 
            "n"=sum(tbl_datatypes))
   
   coltype_fraction <- as.data.frame(matrix(data = vec, byrow = TRUE, 
                                            ncol = length(vec)))
   names(coltype_fraction) <- names(vec)
   return(coltype_fraction)
 }
 

 check_datatypes <- function(data_list) {
   data_df <- data.table::rbindlist(l = lapply(data_list,get_datatypes), 
                                          fill = TRUE)
   
   cbind(data_df, filenames = names(data_list))
 }
 
 checked_datatypes_df <- check_datatypes(tmp_list)
 
 
 labor_df <- data.table::rbindlist(l = tmp_list, fill = TRUE)


 get_unique_rows <- function(df, col_name, 
                             export = TRUE, 
                             expDir = ".") {
  
`%>%` <-  magrittr::`%>%`  

 row_values_df <- as.data.frame(df) %>% 
   dplyr::group_by_(.dots = as.name(col_name)) %>%  
   dplyr::summarise(n = n())
 names(row_values_df) <- c(col_name, "n") 
 
 if (export) write.csv(x = row_values_df, 
                       file = sprintf("%s/%s.csv", 
                                      expDir, 
                                      col_name),
                       row.names = FALSE)
 return(row_values_df)
 }
 
#  for (colname in names(labor_df)) {
#  print(sprintf("Get unique rows for %s", colname))
#  get_unique_rows(df = labor_df,
#                  col_name = colname, 
#                  expDir = export_dir)
# }
#  
 date_name <- get_unique_rows(df = labor_df,col_name = "Datum@NA")
 variable_name <- get_unique_rows(df = labor_df,col_name = "VariableName")
 data_value <- get_unique_rows(df = labor_df,col_name = "DataValue")
 unit_name <- get_unique_rows(df = labor_df,col_name = "UnitName")
 
 col_names <- sort(names(labor_df))

 write.csv(data.frame(ColNames = col_names),
           file = "ColNames.csv", 
           row.names = FALSE)
 
 
  files_with_problems <- files_to_import[has_errors]
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
