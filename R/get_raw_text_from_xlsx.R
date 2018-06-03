# get_raw_text_from_xlsx -------------------------------------------------------
get_raw_text_from_xlsx <- function(file, sheet = NULL, dbg = TRUE)
{
  if (is.null(sheet)) {
    
    sheets <- readxl::excel_sheets(file)
    
    kwb.utils::catIf(dbg, sprintf("\nFile: '%s'\n", basename(file)))
    
    kwb.utils::catIf(dbg, sprintf("Folder: '%s'\n", dirname(file)))
    
    # Call this function for all sheets
    result <- lapply(sheets, get_raw_text_from_xlsx, file = file)
    
    names(result) <- sprintf("sheet_%02d", seq_along(sheets))
    
    sheet_table <- to_sheet_table(sheets)
    
    structure(result, sheets = sheet_table)
    
  } else {
    
    # Explicitly select all rows starting from the first row. Otherwise empty 
    # rows at the beginning are automatically skipped. I want to keep everything
    # so that the original row numbers can be used as a reference
    range <- cellranger::cell_rows(c(1, NA))
    
    kwb.utils::catIf(dbg, sprintf("  Reading sheet '%s' ... ", sheet))
    
    result <- as.matrix(readxl::read_xlsx(
      file, sheet, range = range, col_names = FALSE, col_types = "text"
    ))
    
    kwb.utils::catIf(dbg, "ok.\n")
    
    mode(result) <- "character"
    
    result
  }
}

# to_sheet_table ---------------------------------------------------------------
to_sheet_table <- function(sheets)
{
  no_factors_data_frame(
    sheet_id = add_hex_postfix(sheets), 
    sheet_name = sheets
  )
}
