# get_raw_text_from_xlsx -------------------------------------------------------
get_raw_text_from_xlsx <- function(file, sheet = NULL, dbg = TRUE) {
  stopifnot(is.character(file), length(file) == 1)
  
  # If no sheet name is given, call this function for all sheets in the file
  
  if (is.null(sheet)) {

    # Print the file name and file path to the console
    debug_file(dbg, file)

    # Get the names of available sheets in the file
    sheets <- readxl::excel_sheets(file)

    # Call this function for all sheets
    result <- lapply(sheets, get_raw_text_from_xlsx, file = file)

    # Create sheet metadata
    sheet_table <- to_sheet_table(sheets)

    # Name the list entries according to the sheet ids
    names(result) <- kwb.utils::selectColumns(sheet_table, "sheet_id")

    # Set the sheet metadata as an attribute
    structure(result, sheet_info = sheet_table)
  } else {
    stopifnot(is.character(sheet), length(sheet) == 1)

    # Explicitly select all rows starting from the first row. Otherwise empty
    # rows at the beginning are automatically skipped. I want to keep everything
    # so that the original row numbers can be used as a reference
    range <- cellranger::cell_rows(c(1, NA))

    debug_formatted(dbg, "Reading sheet '%s' as raw text ... ", sheet)

    result <- as.matrix(readxl::read_xlsx(
      file, sheet,
      range = range, col_names = FALSE, col_types = "text"
    ))

    debug_ok(dbg)

    mode(result) <- "character"
    
    structure(result, file = file, sheet = sheet)
  }
}

# debug_file -------------------------------------------------------------------
debug_file <- function(dbg, file) {
  if (dbg) {
    cat_green_bold_0(sprintf("\n  File: '%s'\n", basename(file)))

    cat(sprintf("Folder: '%s'\n", dirname(file)))
  }
}

# debug_formatted --------------------------------------------------------------
debug_formatted <- function(dbg, fmt, ...) {
  kwb.utils::catIf(dbg, sprintf(fmt, ...))
}

# debug_ok ---------------------------------------------------------------------
debug_ok <- function(dbg) {
  kwb.utils::catIf(dbg, "ok.\n")
}

# to_sheet_table ---------------------------------------------------------------
to_sheet_table <- function(sheets) {
  no_factors_data_frame(
    sheet_id = add_hex_postfix(sheets),
    sheet_name = sheets
  )
}
