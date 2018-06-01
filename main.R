install.packages(c("fs", "tidyr", "dplyr", "stringr", "readxl", "devtools",
                    "crayon", "data.table"))
devtools::install_github("kwb-r/kwb.utils")


source(file = "R/convert_xls_as_xlsx.R")
source(file = "R/copy_xlsx_files.R")
source(file = "R/read_bwb_data.R")

path_list <- list(
  drive = "//medusa/projekte$/Z-Exchange/Jeansen",
  #drive = file.path(kwb.utils::get_homedir(), "Downloads"),
  #drive = "F:",
  input_dir = "<drive>/Daten_Labor",
  export_dir = "<drive>/ANALYSIS_R/tmp",
  export_dir_allg = "<export_dir>/K-TL_LSW-Altdaten-Werke Teil 1/Werke Teil 1/Allgemein")

paths <- kwb.utils::resolve(x = path_list)

### Get location of excelcnv.exe 
get_excelcnv_exe()


### Convert old ".xls" to ".xlsx" Excel files
convert_xls_as_xlsx(input_dir = paths$input_dir, 
                    export_dir = paths$export_dir)



### Copy remaining already existing .xlsx files in same directory 
copy_xlsx_files(input_dir = paths$input_dir, 
                export_dir = paths$export_dir,
                overwrite = TRUE)



#### Get all xlsx files to be imported
xlsx_files <- dir(paths$export_dir,
                  pattern = ".xlsx", 
                  recursive = TRUE, 
                  full.names = TRUE)

labor_list <- import_labor(xlsx_files, 
                           export_dir =  paths$export_dir,
                           func = read_bwb_data)

txt <- capture.output(
labor <- read_bwb_data(xlsx_files = xlsx_files))
writeLines(txt, con = file.path(paths$export_dir, "read_bwb_data_Rconsole.txt"))


txt2 <- capture.output(
  labor_header1_meta <- import_labor(xlsx_files = xlsx_files,
                                export_dir = paths$export_dir, 
                                func = read_bwb_header1_meta))
writeLines(txt2, 
           con = file.path(paths$export_dir, 
                           "read_bwb_header1_meta_Rconsole.txt"))


txt3 <- capture.output(
  labor_header2 <- import_labor(xlsx_files = xlsx_files,
                                     export_dir = paths$export_dir, 
                                     func = read_bwb_header2))
writeLines(txt3, 
           con = file.path(paths$export_dir, 
                           "read_bwb_header2_Rconsole.txt"))



length(which(sapply(labor_list, function(x) {is.data.frame(x)})))
length(which(sapply(labor_header2, function(x) {is.data.frame(x)})))
length(which(sapply(labor_header1_meta, function(x) {is.data.frame(x)})))

nrow(labor)
sum(unlist(sapply(labor_list, function(x) {nrow(x)})))
sum(unlist(sapply(labor_header2, function(x) {nrow(x)})))
sum(unlist(sapply(labor_header1_meta, function(x) {nrow(x)})))


packrat::snapshot()
