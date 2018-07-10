# get_sheet_info ---------------------------------------------------------------
get_sheet_info <- function(tables)
{
  kwb.utils::getAttribute(tables, "sheet_info")
}

# set_sheet_info ---------------------------------------------------------------
set_sheet_info <- function(tables, sheet_info)
{
  structure(tables, sheet_info = sheet_info)
}

# get_table_info ---------------------------------------------------------------
get_table_info <- function(tables)
{
  kwb.utils::getAttribute(tables, "table_info")
}

# set_table_info ---------------------------------------------------------------
set_table_info <- function(tables, table_info)
{
  structure(tables, table_info = table_info)
}

# compact_column_info ----------------------------------------------------------
compact_column_info <- function(column_info)
{
  result <- aggregate(
    table_id ~ file_id + column_names_old, # + column_type, 
    data = column_info, 
    FUN = length
  )
  
  names(result)[ncol(result)] <- "n"
  
  kwb.utils::resetRowNames(result[do.call(order, result[, 1:2]), ])
}

# export_table_metadata --------------------------------------------------------

#' Export Table Metadata
#'
#' Export Table Metadata to CSV file
#'
export_table_metadata <- function(table_info, dbg = TRUE) {
  stopifnot(is.data.frame(table_info))

  kwb.utils::checkForMissingColumns(table_info, c("table_id", "table_name"))

  file_xlsx <- kwb.utils::getAttribute(table_info, "file")

  file_csv <- paste0(kwb.utils::removeExtension(file_xlsx), "_META_tmp.csv")

  debug_formatted(dbg, "Writing table medatada to ... ")

  debug_file(dbg, file_csv)

  utils::write.csv(table_info, file = file_csv, row.names = FALSE, na = "")

  debug_ok(dbg)
}

# import_table_metadata --------------------------------------------------------
import_table_metadata <- function(file, dbg = TRUE) {
  file_csv <- paste0(kwb.utils::removeExtension(file), "_META.csv")

  if (file.exists(file_csv)) {
    debug_formatted(
      dbg, "Reading table metadata from\n    '%s' ... ", basename(file_csv)
    )

    table_info <- utils::read.csv(file_csv, stringsAsFactors = FALSE)

    debug_ok(dbg)
  } else {
    debug_formatted(dbg, "No metadata file found for this Excel file.\n")

    table_info <- NULL
  }

  table_info
}

# create_column_metadata -------------------------------------------------------
create_column_metadata <- function(
  tables, table_info = attr(tables, "table_info"), dbg = TRUE
)
{
  # kwb.utils::assignArgumentDefaults("create_column_metadata")

  get_col <- kwb.utils::selectColumns

  if (is.null(table_info)) {
    stop_formatted(
      "%s\n%s",
      "No metadata on tables given in table_info and no attribute 'table_info'",
      "available."
    )
  }

  column_infos <- lapply(names(tables), function(table_id) {

    # table_id <- names(tables)[1]

    debug_formatted(
      dbg, "Creating column metadata for table '%s'... ", table_id
    )

    table_content <- tables[[table_id]]

    selected <- get_col(table_info, "table_id") == table_id

    n_headers <- get_col(table_info, "n_headers")[selected]

    col_types <- get_col(table_info, "col_types")[selected]

    header_matrix <- table_content[seq_len(n_headers), , drop = FALSE]

    column_info <- header_matrix_to_column_info(
      header_matrix, table_id, col_types
    )

    debug_ok(dbg)

    column_info
  })

  debug_formatted(dbg, "Row-binding column info tables ... ")

  column_info <- kwb.utils::safeRowBindAll(column_infos)

  debug_ok(dbg)

  column_info
}

# header_matrix_to_column_info -------------------------------------------------
header_matrix_to_column_info <- function(header_matrix, table_id, col_types) {
  kwb.utils::stopIfNotMatrix(header_matrix)

  n_headers <- nrow(header_matrix)

  n_columns <- ncol(header_matrix)

  header_parts <- as.data.frame(t(header_matrix), stringsAsFactors = FALSE)

  header_parts[is.na(header_parts)] <- ""

  column_info <- no_factors_data_frame(
    table_id = table_id,
    column_no = seq_len(n_columns),
    column_names_old = kwb.utils::pasteColumns(header_parts, sep = "|"),
    column_name = sprintf("Column_%03d", seq_len(n_columns))
  )

  column_types <- strsplit(col_types, "\\|")[[1]]

  stopifnot(length(column_types) == n_columns)

  column_info$column_type <- column_types

  kwb.utils::resetRowNames(column_info)
}

# suggest_column_name ----------------------------------------------------------
suggest_column_name <- function(column_info)
{
  # Suggest column names based on the original column names
  raw_names <- kwb.utils::selectColumns(column_info, "column_names_old")
  
  # Keep only alphanumeric characters
  new_names <- gsub("[^A-Za-z0-9]", "", raw_names)
  
  # Replace "" with "X"
  new_names <- gsub("^$", "X", new_names)
  
  # Keep only the first eight characters
  new_names <- kwb.utils::shorten(new_names, max_chars = 11, delimiter = ".")
  
  column_info$column_name <- new_names
  
  key_columns <- c("file_id", "table_id")
  
  column_info_per_table <- split(column_info, column_info[, key_columns])
  
  column_info_per_table <- lapply(column_info_per_table, function(xx) {
    
    xx$column_name <- kwb.utils::makeUnique(xx$column_name, warn = FALSE)
    
    xx
  })
  
  kwb.utils::rbindAll(column_info_per_table)
}
