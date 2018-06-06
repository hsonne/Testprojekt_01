# export_table_metadata --------------------------------------------------------

#' Export Table Metadata
#' 
#' Export Table Metadata to CSV file 
#' 
export_table_metadata <- function(table_info, dbg = TRUE)
{
  stopifnot(is.data.frame(table_info))

  kwb.utils::checkForMissingColumns(table_info, c("table_id", "table_name"))
  
  file_xlsx <- kwb.utils::getAttribute(table_info, "file")
  
  file_csv <- paste0(kwb.utils::removeExtension(file_xlsx), "_META_tmp.csv")

  debug_formatted(dbg, "Writing table medatada to '%s'... ", file_csv)
  
  utils::write.csv(table_info, file = file_csv, row.names = FALSE, na = "")
  
  debug_ok(dbg)
}

# import_table_metadata --------------------------------------------------------
import_table_metadata <- function(file)
{
  file_csv <- paste0(kwb.utils::removeExtension(file), "_META.csv")

  if (file.exists(file_csv)) {
    
    cat(sprintf(
      "  Reading table metadata from\n    '%s'... ", basename(file_csv)
    ))
    
    table_info <- utils::read.csv(file_csv, stringsAsFactors = FALSE)
    
    cat(sprintf("ok.\n"))
    
  } else {
    
    cat(sprintf("  No metadata available for\n    '%s'.\n", basename(file)))
    
    table_info <- NULL
  }
  
  table_info
}

# create_column_metadata -------------------------------------------------------
create_column_metadata <- function(
  tables, table_info = attr(tables, "tables"), dbg = TRUE
)
{
  if (FALSE) {
    table_info = attr(tables, "tables"); dbg = TRUE
  }
  
  get_col <- kwb.utils::selectColumns
    
  if (is.null(table_info)) {
    
    stop_formatted(
      "%s\n%s", 
      "No metadata on tables given in table_info and no attribute 'tables'",
      "available."
    )
  }
  
  column_infos <- lapply(names(tables), function(table_id) {
  
    #table_id <- names(tables)[1]
    
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
    
    cat("ok.\n")
    
    column_info
  })
  
  cat("Safe row bind all... ")
  column_info <- kwb.utils::safeRowBindAll(column_infos)
  cat("ok.\n")
  
  column_info
}

# header_matrix_to_column_info -------------------------------------------------
header_matrix_to_column_info <- function(header_matrix, table_id, col_types)
{
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
