# Install packages from github
#devtools::install_github("kwb-r/kwb.utils")
#devtools::install_github("kwb-r/kwb.fakin")
#devtools::install_github("kwb-r/kwb.event")

library(kwb.utils)

# Names of required packages in alphabetical order
packages <- c(
  "crayon", "data.table", "devtools", "dplyr", "fs", "readxl", "stringr",
  "testthat", "tidyr", "janitor", "ggplot2"
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
  "utils.R",
  "add_lookup_data.R"
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
  input_dir = "<drive>/02_Daten_Labor_Aufbereitung_02",
  export_dir = "<drive>/03_ANALYSIS_R/tmp",
  sel_folder = "K-TL_LSW-Altdaten-Werke Teil 1/Werke Teil 1/Buch",
  input_dir_sel = "<input_dir>/<sel_folder>",
  export_dir_sel = "<export_dir>/<sel_folder>",
  results_dir = "<drive>/03_ANALYSIS_R/results",
  meta_dir = "<export_dir>/META",
  parameters = "<meta_dir>/2018-06-01 Lab Parameter.xlsx",
  lookup_para = "<meta_dir>/lookup_para.csv",
  sites = "<meta_dir>/Info-Altdaten.xlsx",
  home = kwb.utils::get_homedir()
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
  
  if (FALSE) {
  convert_xls_as_xlsx(input_dir = paths$input_dir_sel, 
                      export_dir = paths$export_dir_sel)
  }
  
  
  # Copy remaining already existing .xlsx files in same directory
  copy_xlsx_files(input_dir, export_dir, overwrite = TRUE)
}

# Get all xlsx files to be imported
files <- dir(export_dir, ".xlsx", recursive = TRUE, full.names = TRUE)


files_meta <- c("Meta Info", 
                "Header ident",
                "Parameter ident",
                "Parameter",
                "Info-Altdaten", 
                "Brandenburg_Parameter_BWB_Stolpe", 
                "Kopie von Brandenburg_Parameter_BWB_Stolpe",
                "2005-10BeschilderungProbenahmestellenGWWIII", 
                "Bezeichnungen der Reinwasserstellen",
                "ReinwasserNomenklatur",
                "Info zu Altdaten 1970-1998",
                "2018-06-01 Lab Parameter"
                )


files_header_1_meta <- c("FRI_Br_GAL_C_Einzelparameter",
                         "FRI_Roh_Rein_NH4+NO3_2001-2003",
                         "MTBE_2003-11_2004",
                         "Reinwasser_2003_Fe_Mn", ## unclean
                         "VC_CN_in Brunnen bis Aug_2005 ",## unclean
                         "Wuhlheide_Beelitzhof_Teildaten" ## unclean
                         )

files_header_1 <- c("2018-04-11 Chlorid in Brunnen - Übersicht",
                    "2018-04-27 LIMS Reiw & Rohw Sammel ",
                    "2018-04-27 Rohwasser Bericht - Galeriefördermengen")


files_header_4 <- c("STO Rohw_1999-6_2004",
                    "Wuhlheide_1999-2003_Okt - Neu",
                    "KAU_1999-Okt2003")
                      

files_to_ignore <- c(files_meta, files_header_1, files_header_1_meta, 
                     files_header_4)


cat(crayon::bgWhite(sprintf("
###############################################################################\n
Currently %d files are ignored for import using function read_bwb_data():\n
Meta files:\n%s\n
Header1 (with manually added, but still 'unclean' metadata):\n%s\n
Header1 (without metadata):\n%s\n

###############################################################################
# Manual selection: %d files with 4 headers are imported with: read_bwb_header4() 
###############################################################################\n
Header4:\n%s\n
##############################################################################", 
      length(files_to_ignore), 
      paste(files_meta, collapse = "\n"), 
      paste(files_header_1_meta, collapse = "\n"), 
      paste(files_header_1, collapse = "\n"), 
      length(files_header_4),
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

# file_database$files
# file_database$folders

if (FALSE)
{
  files_header_4 <- files[stringr::str_detect(string = files ,
                      pattern = paste0(files_header_4,collapse = "|"))]
  
  labor_header4_list <- import_labor(files = files_header_4,
                      export_dir = export_dir, 
                      func = read_bwb_header4)
 
  has_errors <- sapply(labor_header4_list, inherits, "try-error")
  has_errors
  
    labor_header4_df <- data.table::rbindlist(l = labor_header4_list[!has_errors], 
                                            fill = TRUE)
   
  # files_header_1_meta <- files[stringr::str_detect(string = files ,
  #                                             pattern = paste0(files_header_1_meta,collapse = "|"))]
  # 
  # 
  # labor_list_1meta <- import_labor(files = files_header_1_meta,
  #                                    export_dir = export_dir,
  #                                    func = read_bwb_header1_meta)
  # 
  # has_errors <- sapply(labor_list_1meta, inherits, "try-error")
  # has_errors
  # 
  # labor_list_1meta <- data.table::rbindlist(l = labor_list_1meta[!has_errors],
  #                                           fill = TRUE)
  # 
  # labor_list_1meta %>%  filter(is.na(sheet_name)) %>%  View()
  # 
  # View(labor_list_1meta)
  
  labor <- read_bwb_data(files = files_to_import)

  #View(head(labor))
  
 labor_all <- data.table::rbindlist(l = list(x1 = labor, 
                                x2 = labor_header4_df),
                                fill = TRUE)
 
 labor_all <- labor_all %>% 
              dplyr::filter_("!is.na(DataValue)") 
 

 if(any(names(labor_all) %in% "Date")) {
   stop("Date already defined. Please rename all 'Date' column headers in 
        to the original ones (i.e. Probenahme or Datum)")
 } else {
  labor_all <- labor_all %>% 
   dplyr::mutate_("Date" = "dplyr::if_else(condition = !is.na(Datum),
        true = Datum, 
        false =  Probenahme)")  %>%  
    ### Some "Datum" rentries are missing in;
    ###K-TL_LSW-Altdaten-Werke Teil 1\Werke Teil 1\Kaulsdorf\KAU_1999-Okt2003.xlsx
    ### sheets: 66 KAU Rein 1999-2000, 65 KAU NordSüd 1999-2000
    dplyr::filter_("!is.na(Date)") 
 }
 
 labor_all <- labor_all %>% 
      dplyr::filter_("!is.na(VariableName_org)")
 
 
 
 labor_all_sel <- add_para_metadata(df = labor_all, 
                  lookup_para_path = paths$lookup_para,
                   parameters_path = paths$parameters)
 
 labor_all_sel <- add_site_metadata(df = labor_all_sel, 
                   site_path = paths$sites)
   

pdf_file <- file.path(paths$results_dir, "Zeitreihen.pdf")

pdf(file = pdf_file,width = 14, height = 9)
 for (sel_para_id in unique(labor_all_sel$para_id)) {
 
 tmp <- labor_all_sel %>%
   dplyr::filter_(.dots = sprintf("para_id == %d", sel_para_id)) %>% 
   dplyr::mutate_(year = "as.numeric(format(Date,format = '%Y'))") %>% 
   dplyr::group_by_("para_kurzname", "werk", "year") %>% 
   dplyr::summarise_(mean_DataValue = "mean(as.numeric(DataValue), na.rm=TRUE)") %>% 
   dplyr::filter_("!is.na(werk)") 
 
 
 
  g <- ggplot2::ggplot(tmp,mapping = ggplot2::aes_string(x = "year", 
                                        y = "mean_DataValue",
                                        col = "werk")) + 
     ggplot2::geom_point() +
     ggplot2::theme_bw() +
     ggplot2::ggtitle(tmp$para_kurzname[1]) +
     ggplot2::labs(x = "", y = "")
   
   print(g)
   
 }
dev.off()

 
 # ggplot2::ggplot(online, ggplot2::aes_string(x = "year",
 #                                             y = "mean_DataValue",
 #                                             col = "SiteName")) +
 #   ggforce::facet_wrap_paginate(~werk,
 #                                nrow = 1,
 #                                ncol = 1,
 #                                scales = "free_y",
 #                                page = i) +
 #   ggplot2::geom_point() +  
 #   ggplot2::theme_bw(base_size = 20) +
 #   ggplot2::theme(legend.position = "top"
 #                  , strip.text.x = element_text(face = "bold")
 #                  , legend.title = element_blank()
 #   )
 # 

 get_unique_rows <- function(df, col_name, 
                             export = TRUE, 
                             expDir = ".") {
  
`%>%` <-  magrittr::`%>%`  


 row_values_df <- df %>% 
   dplyr::group_by_(.dots = col_name) %>%  
   dplyr::summarise(n = n())
 
 
 if (export) write.csv(x = row_values_df, 
                       file = sprintf("%s/%s.csv", 
                                      expDir, 
                                      paste0(col_name, collapse = "_")),
                       row.names = FALSE)
 return(row_values_df)
 }
 

labor_all$Year_num <- as.numeric(format(labor_all$Date,format = "%Y"))
 
 
 samples_per_year <- get_unique_rows(df = labor_all,col_name = c("Year_num"))
 plot(samples_per_year$Year_num, samples_per_year$n)
 dates <- get_unique_rows(df = labor_all,col_name = c("Date"))
 plot(dates$Date, dates$n)
 site_id <- get_unique_rows(df = labor_all,col_name = c("SiteID"))
 date_name <- get_unique_rows(df = labor_all,col_name = c("Date", "file_name", "sheet_name"))
 date_name1 <- get_unique_rows(df = labor_all,col_name = c("Datum", "file_name", "sheet_name"))
 date_name2 <- get_unique_rows(df = labor_all,col_name = c("Probenahme", "file_name","sheet_name"))
 variable_name <- get_unique_rows(df = labor_all,col_name = "VariableName_org")
 data_value <- get_unique_rows(df = labor_all,col_name = "DataValue")
 unit_name <- get_unique_rows(df = labor_all,col_name = "UnitName_org")
 
 col_names <- sort(names(labor_all))

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
