# get_SiteID -------------------------------------------------------------------
get_SiteID <- function(string, pattern = "^[0-9]{1,4}")
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



# gather_ignore ----------------------------------------------------------------
gather_ignore_clean <- function()
{
  fields <- c(
    "LabSampleCode", "Date", "Time", "Waterbody", "ExSiteCode", 
    "Site"
  )
  
  kwb.utils::collapsed(fields, "|")
}

# read_bwb_header1_meta --------------------------------------------------------
read_bwb_header1_meta <- function(
  file, meta_pattern = "META", keep_pattern = gather_ignore_clean()
)
{
  # Get the names of the sheets in the Excel workbook
  sheets <- readxl::excel_sheets(file)
  
  # Try to find the sheet containing meta data
  meta_sheet <- get_meta_sheet_or_stop(sheets, meta_pattern, file)
  
  # Read the metadata sheet
  all_metadata <- readxl::read_excel(file, meta_sheet)
  
  # Get the names of the sheets for which metadata are available
  described_sheets <- unique(kwb.utils::selectColumns(all_metadata, "Sheet"))
  
  # Loop through the names of the sheets for which metadata are available  
  sheet_data_list <- lapply(described_sheets, function(sheet) {
    
    # Filter for the metadata given for the current sheet
    metadata <- all_metadata[all_metadata$Sheet == sheet, ]
    
    # Load the data from the current sheet
    tmp_data <- readxl::read_excel(file, sheet, guess_max = 2^20)

    # Safely select the original column names
    columns_orig <- kwb.utils::selectColumns(metadata, "OriginalName")
    
    # Safely select the clean column names
    columns_clean <- kwb.utils::selectColumns(metadata, "Name")

    # Are the columns kept as columns, i. e. excluded from gathering?
    keep <- stringr::str_detect(columns_clean, keep_pattern)
    
    # Convert the data from wide to long format
    gather_and_join_1(tmp_data, columns_clean[keep], metadata, dbg = TRUE)
  })
  
  # Merge all data frames in long format  
  data.table::rbindlist(l = sheet_data_list, fill = TRUE)
}

# get_meta_sheet_or_stop -------------------------------------------------------
get_meta_sheet_or_stop <- function(sheets, pattern, file)
{
  meta_sheets <- sheets[stringr::str_detect(sheets, pattern)]
  
  # Number of meta sheets  
  n <- length(meta_sheets)
  
  if (n == 0) {
    
    stop_formatted(
      "%s does not contain a sheet matching '%s'\n", 
      file, pattern
    )
    
  } else if (n > 1)  {
    
    stop_formatted(
      "%s contains %d sheets matching '%s': %s\n", 
      file, n, pattern, kwb.utils::stringList(meta_sheets)
    )
  }
  
  meta_sheets
}

# stop_formatted ---------------------------------------------------------------
stop_formatted <- function(fmt, ...)
{
  stop(sprintf(fmt, ...), call. = FALSE)
}

# read_bwb_header2 -------------------------------------------------------------
read_bwb_header2 <- function(
  file, skip = 2, keep_pattern = gather_ignore(), 
  site_id_pattern = "^[0-9]{1,4}", dbg = TRUE
)
{
  # Define helper functions, 2^20 = max number of rows in xlsx
  read_from_excel <- function(...) readxl::read_excel(..., col_names = FALSE, guess_max = 2^20)

  sheets <- readxl::excel_sheets(file)
  
  has_site_id <- stringr::str_detect(sheets, site_id_pattern)
  
  stop_on_missing_or_inform_on_extra_sheets(has_site_id, file, sheets)

  data_frames <- lapply(which(has_site_id), function(sheet_index) {
    
    sheet <- sheets[sheet_index]
    
    cat(sprintf(
      "FROM: %s\nReading sheet (%d/%d): %s\n", 
      file, sheet_index, length(sheets), sheet
    ))
    
    # Read the header rows      
    stopifnot(skip == 2) # Otherwise we need more column names!
    header <- read_from_excel(file, sheet, 
                              range = cellranger::cell_rows(c(1,skip)))
    header <- to_full_metadata_2(header, file, sheet)
    
    # Read the data rows      
    tmp_content <- read_from_excel(file, sheet, 
            range = cellranger::cell_limits(ul = c(skip+1,1),
                                            lr = c(NA,nrow(header))))
    
    indices <- match(names(tmp_content), header$id)
    
    names(tmp_content) <- header$key[indices]
    
    # Check content format
    tbl_datatypes <- table(unlist(sapply(tmp_content, class)))
    tbl_datatypes <- sort(tbl_datatypes, decreasing = TRUE)
    
    columns_keep <- grep(keep_pattern, names(tmp_content), value = TRUE)
    
    print_datatype_info_if(dbg, tbl_datatypes, columns_keep)
    
    # TODO: check for duplicates in names
    gather_and_join_2(tmp_content, columns_keep, header)
  })

  data.table::rbindlist(l = data_frames, fill = TRUE)
}

read_bwb_header4 <- function(
  file, skip = 4, keep_pattern = gather_ignore(), 
  site_id_pattern = "^[0-9]{1,4}", dbg = TRUE
)
{
  # Define helper functions
  read_from_excel <- function(...) {
    readxl::read_xlsx(..., col_names = FALSE, guess_max = 2^20)
  }
  
  sheets <- readxl::excel_sheets(file)
  
  has_site_id <- stringr::str_detect(sheets, site_id_pattern)
  
  stop_on_missing_or_inform_on_extra_sheets(has_site_id, file, sheets)
  
  data_frames <- lapply(which(has_site_id), function(sheet_index) {
    
    sheet <- sheets[sheet_index]
    
    cat(sprintf(
      "FROM: %s\nReading sheet (%d/%d): %s\n", 
      file, sheet_index, length(sheets), sheet
    ))
    
    # Read the header rows      
    stopifnot(skip == 4) # Otherwise we need more column names!
    header <- read_from_excel(file, sheet, 
                              range = cellranger::cell_rows(c(1,skip)))
    header <- to_full_metadata_4(header, file, sheet)
    
    # Read the data rows      
    tmp_content <- read_from_excel(file, sheet, 
          range = cellranger::cell_limits(ul = c(skip+1,1),
                                          lr = c(NA,nrow(header))))
    
    indices <- match(names(tmp_content), header$id)
    
    names(tmp_content) <- header$key[indices]
    
    # Check content format
    tbl_datatypes <- table(unlist(sapply(tmp_content, class)))
    tbl_datatypes <- sort(tbl_datatypes, decreasing = TRUE)
    
    columns_keep <- grep(keep_pattern, names(tmp_content), value = TRUE)
    
    print_datatype_info_if(dbg, tbl_datatypes, columns_keep)
    
    # TODO: check for duplicates in names
    gather_and_join_2(tmp_content, columns_keep, header)
  })
  
  data.table::rbindlist(l = data_frames, fill = TRUE)
}

# stop_on_missing_or_inform_on_extra_sheets ------------------------------------
stop_on_missing_or_inform_on_extra_sheets <- function(has_site_id, file, sheets)
{
  if (! any(has_site_id)) {
    
    stop_formatted(
      paste0(
      "No data sheet has a site code in its name!\n", 
      "Folder:\n  %s\n",
      "File:\n  '%s'\n",
      "Sheet names:\n  %s\n"
      ), 
      dirname(file), basename(file),
      kwb.utils::stringList(sheets, collapse = "\n  ")
    )
  }
  
  if (! all(has_site_id)) {
    
    crayon::blue(sprintf(
      "FROM: %s\nIgnoring the following (%d/%d) sheet(s):\n%s\n", 
      file, sum(! has_site_id), length(sheets), 
      kwb.utils::stringList(sheets[! has_site_id])
    ))
  }
}

# to_full_metadata2 -------------------------------------------------------------
to_full_metadata_2 <- function(header, file, sheet)
{
  # Start a metadata table with the Variable name and unit
  header <- as.data.frame(t(header))
  names(header) <- c("VariableName", "UnitName")
  
  # Extend the metadata      
  header$key <- kwb.utils::pasteColumns(header, sep = "@")
  header$id <- rownames(header)
  header$file_name <- normalizePath(file)
  header$sheet_name <- sheet
  header$SiteID <- get_SiteID(sheet)

  header  
}

# to_full_metadata2 -------------------------------------------------------------
to_full_metadata_4 <- function(header, file, sheet)
{
  # Start a metadata table with the Variable name and unit
  header <- as.data.frame(t(header))
  names(header) <- c("VariableName", "Method", "UnitName", "DetectionLimit")
  
  # Extend the metadata      
  header$key <- kwb.utils::pasteColumns(header, sep = "@")
  header$id <- rownames(header)
  header$file_name <- normalizePath(file)
  header$sheet_name <- sheet
  header$SiteID <- get_SiteID(sheet)
  
  header  
}

# print_datatype_info_if -------------------------------------------------------
print_datatype_info_if <- function(dbg, tbl_datatypes, columns_keep)
{
  if (dbg) {
    
    cat_green_bold_0(
      "The following datatypes were detected:\n", 
      kwb.utils::stringList(qchar = "", sprintf(
        "%d x %s", as.numeric(tbl_datatypes), names(tbl_datatypes)
      ))
    )
    
    cat_green_bold_0(stringr::str_c(
      "\nThe following column(s) will be used as headers:\n",
      stringr::str_c(columns_keep, collapse = ", "), "\n"
    ))
  }
}

# cat_green_bold_0 -------------------------------------------------------------
cat_green_bold_0 <- function(...)
{
  cat(crayon::green(crayon::bold(paste0(...))))
}

# gather_and_join_1 ------------------------------------------------------------
gather_and_join_1 <- function(tmp_data, columns_keep, metadata, dbg = FALSE)
{
  `%>%` <- magrittr::`%>%`
  
  kwb.utils::printIf(dbg, names(tmp_data))
  kwb.utils::printIf(dbg, columns_keep)

  tidyr::gather_(
    data = tmp_data, key_col = "VariableName", value_col = "DataValue", 
    gather_cols = setdiff(names(tmp_data), columns_keep)
  ) %>% 
    dplyr::left_join(y = metadata, by = c(VariableName = "Name"))
}

# gather_and_join_2 ------------------------------------------------------------
gather_and_join_2 <- function(tmp_content, columns_keep, header)
{
  `%>%` <- magrittr::`%>%`
  
  tidyr::gather_(
    data = tmp_content, key_col = "key", value_col = "DataValue", 
    gather_cols = setdiff(names(tmp_content), columns_keep)
  ) %>% 
    dplyr::left_join(y = header, by = "key")
}

# read_bwb_data ----------------------------------------------------------------
read_bwb_data <- function(
  files, meta_pattern = "META", keep_pattern = gather_ignore(), 
  keep_pattern_clean = gather_ignore_clean(),
  site_id_pattern = "^[0-9]{1,4}", dbg = TRUE
)
{
  result_list <- lapply(files, function(file) {
    
    sheets <- readxl::excel_sheets(file)
    
    is_meta <- stringr::str_detect(sheets, meta_pattern)
    
    has_site_id <- stringr::str_detect(sheets, site_id_pattern)

    if (any(is_meta)) {
      
      header <- read_bwb_header1_meta(file, meta_pattern, keep_pattern_clean)
      
      if (exists("header") && nrow(header)) {
        
        header

      } # else NULL implicitly

    } else if (any(has_site_id)) {
      
      header <- read_bwb_header2(
        file, keep_pattern = keep_pattern, site_id_pattern = site_id_pattern, 
        dbg = dbg
      )
      
      if (exists("header") && nrow(header)) {
        
        header
        
      } # else NULL implicitly

    } else {
      
      cat_red_bold_0(
        "'", basename(file), "' does not follow import schemes defined ", 
        "in functions read_bwb_header1_meta() or read_bwb_header2()\n"
      )
    }
  })
  
  data.table::rbindlist(l = kwb.utils::excludeNULL(result_list), fill = TRUE)
}

# cat_red_bold_0 ---------------------------------------------------------------
cat_red_bold_0 <- function(...)
{
  cat(crayon::bold(crayon::red(paste0(...))))
}

# import_labor -----------------------------------------------------------------
import_labor <- function(files, export_dir, func = read_bwb_header2)
{
  try_func_on_file <- function(file) try(func(file))
  
  labor <- stats::setNames(lapply(files, try_func_on_file), basename(files))
  
  file_name <- sprintf("%s_structure.txt", as.character(substitute(func)))
  
  file <- file.path(export_dir, file_name, fsep = "\\")
  
  capture.output(str(labor, nchar.max = 254, list.len = 10000), file = file)
  
  labor
}
