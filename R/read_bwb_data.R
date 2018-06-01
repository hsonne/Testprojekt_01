# Better: call functions with <package>::

#library(readxl)
#library(stringr)
#library(dplyr)
#library(crayon)

# get_SiteID -------------------------------------------------------------------
get_SiteID <- function(string, pattern = "^[0-9][0-9]?[0-9]?[0-9]?")
{
  as.numeric(stringr::str_extract(string, pattern))
}

# gather_ignore ----------------------------------------------------------------
gather_ignore <- function()
{
  fields <- c(
    "Datum", "KN", "[iI]nterne Nr.", "Name der", "Ort", "Probe", "Pr\u00fcf", 
    "Untersuchung", "Labor", "Jahr", "Galer", "Detail", "Me\u00DF", "Zeit", 
    "Bezei"
  )
  
  kwb.utils::collapsed(fields, "|")
}

# read_bwb_header1_meta --------------------------------------------------------
read_bwb_header1_meta <- function(
  xlsx_file, 
  meta_sheet_pattern = "META", 
  pattern_gather_ignore = gather_ignore()
)
{
  result <- NULL
  
  # Get the names of the sheets in the Excel workbook
  sheets <- readxl::excel_sheets(xlsx_file)

  # Try to find the sheet containing meta data
  meta_sheet <- get_meta_sheet_or_stop(sheets, meta_sheet_pattern, xlsx_file)

  # Read the metadata sheet
  metadata <- readxl::read_excel(xlsx_file, sheet = meta_sheet)
  
  # Get the names of the sheets for which metadata are available
  described_sheets <- unique(kwb.utils::selectColumns(metadata, "Sheet"))

  # Loop through the names of the sheets for which metadata are available  
  for (sheet in described_sheets) {

    # Filter for the metadata given for the current sheet
    metadata_sheet <- metadata[metadata$Sheet == sheet, ]
    
    # Safely select the original and the clean column names, respectively
    columns_orig <- kwb.utils::selectColumns(metadata_sheet, "OriginalName")
    columns_clean <- kwb.utils::selectColumns(metadata_sheet, "Name")
    
    # Are the columns to be gathered?
    do_not_gather <- stringr::str_detect(columns_orig, pattern_gather_ignore)

    # (Clean) names of the columns not to be gathered    
    columns_not_to_gather <- columns_clean[do_not_gather]
    
    # Load the data from the current sheet
    tmp_data <- readxl::read_excel(xlsx_file, sheet = sheet)
    
    # Convert the data from wide to long format
    tmp_df <- tmp_data %>% 
      tidyr::gather_(
        key_col = "VariableName", 
        value_col = "DataValue", 
        gather_cols = setdiff(names(tmp_data), columns_not_to_gather)
      ) %>% 
      dplyr::left_join(y = metadata_sheet, by = c(VariableName = "Name"))
    
    result <- if (is.null(result)) {
      
      tmp_df
      
    } else {
      
      data.table::rbindlist(l = list(result, tmp_df), fill = TRUE)
    }
  }
  
  result
}

# get_meta_sheet_or_stop -------------------------------------------------------
get_meta_sheet_or_stop <- function(sheets, pattern, file)
{
  meta_sheets <- sheets[stringr::str_detect(sheets, pattern)]

  # Number of meta sheets  
  n <- length(meta_sheets)
  
  if (n == 0) {
    
    stop(sprintf(
      "%s does not contain a sheet matching '%s'\n", file, pattern
    ))
    
  } else if (n > 1)  {
    
    stop(sprintf(
      "%s contains %d sheets matching '%s': %s\n", 
      file, n, pattern, kwb.utils::stringList(meta_sheets)
    ))
    
  }
  
  meta_sheets
}

# read_bwb_header2 -------------------------------------------------------------
read_bwb_header2 <- function(
  xlsx_file, 
  skip = 2,
  pattern_gather_ignore = gather_ignore(),
  site_id_pattern = "^[0-9][0-9]?[0-9]?[0-9]?",
  dbg = TRUE) {
  
  result <- NULL
  
  sheets <- readxl::excel_sheets(path = xlsx_file)
  
  contains_site <- stringr::str_detect(string = sheets,
                                       pattern = site_id_pattern)
  
  if (any(contains_site)) {
    
    if (any(!contains_site)) {
      wrn_msg <- crayon::blue(
        sprintf("FROM: %s\nIgnoring the following (%d/%d) sheet(s):\n%s\n", 
                xlsx_file, 
                sum(!contains_site), 
                length(sheets), 
                paste(sheets[!contains_site],collapse = ", ")))
      warning(cat(wrn_msg))
      
    }
    
    
    for(sheet_index in which(contains_site)) {
      
      sheet_name <- sheets[sheet_index]
      
      cat(sprintf("FROM: %s\nReading sheet (%d/%d): %s\n", 
                  xlsx_file, 
                  sheet_index, 
                  length(sheets),
                  sheet_name))
      
      tmp_header <- readxl::read_excel(path = xlsx_file,
                                       sheet = sheet_name, 
                                       n_max = skip, 
                                       col_names = FALSE)
      
      tmp_header <- as.data.frame(t(tmp_header))
      names(tmp_header) <- c("VariableName", "UnitName")
      tmp_header$id <- rownames(tmp_header)
      tmp_header$key <- sprintf("%s@%s",
                                tmp_header$VariableName,
                                tmp_header$UnitName)
      tmp_header$file_name <- normalizePath(xlsx_file)
      tmp_header$sheet_name <- sheet_name
      tmp_header$SiteID <- get_SiteID(sheet_name)
      
      tmp_content <- readxl::read_excel(path = xlsx_file,
                                        sheet = sheet_name, 
                                        skip = skip, 
                                        col_names = FALSE)
      
      names(tmp_content) <- tmp_header$key[match(names(tmp_content), 
                                                 tmp_header$id)]
      
      ### Check content format:
      tbl_datatypes <- table(unlist(sapply(tmp_content, class)))
      tbl_datatypes <- sort(tbl_datatypes,decreasing = TRUE)
      
      columns_not_to_gather <- grep(pattern_gather_ignore, names(tmp_content), value = TRUE)
      
            
      if (dbg) {
        cat(crayon::green(crayon::bold(
          sprintf("The following datatypes were detected:\n%s", 
                  paste(sprintf("%d x %s", 
                                as.numeric(tbl_datatypes), 
                                names(tbl_datatypes)), 
                        collapse = ", ")))))
        
        cat(crayon::green(crayon::bold(
          stringr::str_c("\nThe following column(s) will be used as headers:\n",
                         stringr::str_c(columns_not_to_gather, collapse = ", "
                         ),"\n"))))
      }
      ### To do: check for duplicates in names
      tmp_df <- tidyr::gather_(
        data = tmp_content, key_col = "key", value_col = "DataValue", 
        gather_cols = setdiff(names(tmp_content), columns_not_to_gather)
      ) %>% 
        dplyr::left_join(y = tmp_header, by = "key")
      
      if(!is.null(result)) {
        result <- tmp_df
      } else {
        result <- data.table::rbindlist(l = list(result, 
                                              tmp_df), 
                                     fill = TRUE)
      }
      
      
    }} else {
      stop(sprintf("No data sheet in %s available!", xlsx_file))  
    }
  
  return(result)
  
}


read_bwb_data <- function(
  xlsx_files, 
  meta_sheet_pattern = "META",
  pattern_gather_ignore = gather_ignore(),
  site_id_pattern = "^[0-9][0-9]?[0-9]?[0-9]?", 
  dbg = TRUE) {


result <- NULL  


for (xlsx_file in xlsx_files) {

  
sheets <- readxl::excel_sheets(path = xlsx_file)

is_meta <- stringr::str_detect(string = sheets,
                                     pattern = "META")

contains_site <- stringr::str_detect(string = sheets,
                                     pattern = site_id_pattern)


 if (any(is_meta)) {
   tmp_header1_meta <- read_bwb_header1_meta(
     xlsx_file = xlsx_file, 
     meta_sheet_pattern = meta_sheet_pattern, 
     pattern_gather_ignore = pattern_gather_ignore
    )
   
   if(!exists("tmp_header1_meta")) break
   if(nrow(tmp_header1_meta) == 0) break
   
   if(is.null(result)) {
     result <- tmp_header1_meta
   } else {
     result <- data.table::rbindlist(l = list(result, 
                                              tmp_header1_meta), 
                                              fill = TRUE)
   }

 } else if (any(contains_site)) {
   tmp_header2 <- read_bwb_header2(
     xlsx_file, 
     pattern_gather_ignore = pattern_gather_ignore,
     site_id_pattern = site_id_pattern,
     dbg = dbg)
   
   if(!exists("tmp_header2")) break
   if(nrow(tmp_header2) == 0) break
   if(is.null(result)) {
     result <- tmp_header2
   } else {
     result <- data.table::rbindlist(l = list(result, 
                                              tmp_header2),
                                              fill = TRUE)
   }
   
   } else {
    cat(crayon::bold(crayon::red(
     sprintf("'%s' does not follow import schemes defined in functions %s\n", 
                  basename(xlsx_file), 
                  "read_bwb_header1_meta() or read_bwb_header2()"))))  
   }
}
return(result)
}

import_labor <- function(xlsx_files,
                         export_dir,
                         func = read_bwb_header2) {
  
  func_name <- as.character(substitute(func))
  
  labor <- sapply(xlsx_files, FUN = function(file){
    try(expr = func(xlsx_file = file))})
  
  names(labor) <- basename(xlsx_files)
  writeLines(text = capture.output(str(labor,
                                       nchar.max = 254,
                                       list.len = 10000)), 
             con = file.path(export_dir, 
                             sprintf("%s_structure.txt",
                                     func_name),
                             fsep = "\\"))
  
  return(labor)  
}
