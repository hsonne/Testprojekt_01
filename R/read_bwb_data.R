library(readxl)
library(stringr)
library(dplyr)
library(crayon)

get_SiteID <- function(string, 
                       pattern = "^[0-9][0-9]?[0-9]?[0-9]?") {
  
  site_id_pattern <- "^[0-9][0-9]?[0-9]?[0-9]?"
  SiteIDs <- as.numeric(stringr::str_extract(string = string,
                                             pattern = pattern))
  return(SiteIDs)
}


read_bwb_data <- function(xlsx_file, 
                          skip = 2,
                          pattern_gather_ignore = "Datum|KN|interne Nr.",
                          site_id_pattern = "^[0-9][0-9]?[0-9]?[0-9]?", 
                          dbg = TRUE) {

sheets <- readxl::excel_sheets(path = xlsx_file)

contains_site <- stringr::str_detect(string = sheets,
                                     pattern = site_id_pattern)


res <- NULL


if (any(!contains_site)) {
  wrn_msg <- crayon::blue(
    sprintf("FROM: %s\nIgnoring the following (%d/%d) sheet(s):\n%s\n", 
                     xlsx_file, 
                     sum(!contains_site), 
                     length(sheets), 
                     paste(sheets[!contains_site],collapse = ", ")))
  warning(cat(wrn_msg))
  
}


if (any(contains_site)) {
 
  for(sheet_index in which(contains_site)) {
    sheet_name <- sheets[sheet_index]
    
    cat(sprintf("FROM: %s\nReading sheet (%d/%d): %s\n", 
                xlsx_file, 
                sheet_index, 
                length(sheets),
                sheet_name))
    
    tmp_header <- readxl::read_excel(path = xlsx_file,
                                     sheet = sheet_index, 
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
  
    cols_not_to_gather <- names(tmp_content)[grepl(pattern = pattern_gather_ignore, 
                                                   x = names(tmp_content))]
    
    if (dbg) {
      cat(crayon::green(crayon::bold(
      sprintf("The following datatypes were detected:\n%s", 
              paste(sprintf("%d x %s", 
                            as.numeric(tbl_datatypes), 
                            names(tbl_datatypes)), 
                    collapse = ", ")))))
      
      cat(crayon::green(crayon::bold(
      stringr::str_c("\nThe following column(s) will be used as headers:\n",
                    stringr::str_c(cols_not_to_gather, collapse = ", "
                  ),"\n"))))
    }
    tmp_df <- tidyr::gather_(
      data = tmp_content, key_col = "key", value_col = "DataValue", 
      gather_cols = setdiff(names(tmp_content), cols_not_to_gather)
    ) %>% 
      dplyr::left_join(y = tmp_header, by = "key")
    
    if(!is.null(res)) {
      res <- tmp_df
    } else {
      res <- data.table::rbindlist(l = list(res, 
                                            tmp_df))
    }
    
    
  }} else {
    print(sprintf("No data sheet in %s available!", xlsx_file))  
  }
return(res)
}

import_labor <- function(xlsx_files,
                         export_dir) {
  
  labor <- sapply(xlsx_files, FUN = function(file){
    try(expr = read_bwb_data(xlsx_file = file))})
  
  names(labor) <- basename(xlsx_files)
  writeLines(text = capture.output(str(labor,
                                       nchar.max = 254,
                                       list.len = 10000)), 
             con = file.path(export_dir, 
                             "import_labor_structure.txt",
                             fsep = "\\"))
  
  return(labor)  
}
                      