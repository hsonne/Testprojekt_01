# get_tables_from_xlsx ---------------------------------------------------------
get_tables_from_xlsx <- function(file)
{
  # Get one character matrix per sheet
  text_sheets <- get_raw_text_from_xlsx(file)

  sheet_info <- kwb.utils::getAttribute(text_sheets, "sheets")
  
  is_empty <- colSums(sapply(text_sheets, dim)) == 0
  
  if (any(is_empty)) {
    
    message(
      "Skipping empty sheet(s):\n  ", 
      kwb.utils::stringList(sheet_info$sheet_name[is_empty], collapse = "\n  ")
    )
  }
  
  all_tables <- lapply(text_sheets[! is_empty], function(text_sheet) {
    
    #text_sheet <- text_sheets[[1]]
    tables <- split_into_tables(text_sheet)
    
    if (length(tables) > 1 && length(tables[[1]]) == 1) {
      
      name_even_tables_by_odd_tables(tables)
      
    } else {
      
      tables
    }
  })

  table_infos <- lapply(all_tables, kwb.utils::getAttribute, "tables")
  
  table_info <- merge_table_infos(table_infos)
  
  all_tables <- stats::setNames(do.call(c, all_tables), table_info$table_id)
  
  structure(all_tables, tables = table_info, sheets = sheet_info, file = file)
}

# split_into_tables ------------------------------------------------------------
split_into_tables <- function(text_sheet, delete_all_empty_columns = FALSE)
{
  stopifnot(is.character(text_sheet), length(dim(text_sheet)) == 2)
  
  # Get indices of non-empty rows
  indices <- which(! kwb.utils::isNaInAllColumns(text_sheet))
  
  # Create ranges of "connected" non-empty rows
  table_info <- kwb.event::hsEvents(indices, 1, 1)[, 3:4]
  
  names(table_info) <- c("first_row", "last_row")
  
  tables <- lapply(seq_len(nrow(table_info)), function(i) {
    
    indices <- table_info$first_row[i]:table_info$last_row[i]
    
    result <- text_sheet[indices, , drop = FALSE]
    
    result <- delete_empty_columns_right(result)
    
    if (isTRUE(delete_all_empty_columns)) {
      
      df <- as.data.frame(result)
      
      result <- as.matrix(kwb.utils::removeEmptyColumns(df))
    }
    
    result
  })

  table_info <- extend_table_info(table_info, tables)
  
  names(tables) <- table_info$table_id
  
  structure(tables, tables = table_info)
}

# extend_table_info ------------------------------------------------------------
extend_table_info <- function(table_info, tables)
{
  dims <- t(sapply(tables, dim))
  
  cbind(
    no_factors_data_frame(
      table_id = add_hex_postfix(tables), 
      table_name = ""
    ),
    nrow = dims[, 1],
    ncol = dims[, 2],
    first_row = table_info$first_row, 
    last_row = table_info$last_row, 
    first_col = 1, 
    last_col = dims[, 2],
    topleft = unname(sapply(tables, "[", 1, 1)),
    stringsAsFactors = FALSE
  )
}

# name_even_tables_by_odd_tables -----------------------------------------------
name_even_tables_by_odd_tables <- function(tables)
{
  stopifnot(is.list(tables))
  
  stopifnot(all(sapply(tables, is.matrix)))
  
  table_info <- kwb.utils::getAttribute(tables, "tables")

  sizes <- sapply(tables, length)
  
  stopifnot(all(sizes == table_info$nrow * table_info$ncol))

  n_tables <- length(tables)
  
  odd_indices <- seq(1, n_tables, by = 2)
  
  even_indices <- seq(2, n_tables, by = 2)
  
  if (! all(sizes[odd_indices] == 1)) {
    
    message(
      "Not all one-value tables (expected to be captions) are on odd ",
      "positions in the list.\n",
      "Returning the original list of tables"
    )
    
    tables
    
  } else {
    
    # Get captions from one-element tables at odd indices
    table_info$table_name[even_indices] <- table_info$topleft[odd_indices]
    
    # Remove the 1x1 tables from table_info
    table_info <- table_info[even_indices, ]
    
    # Keep only the tables at even indices
    result <- tables[even_indices]

    # Renumber the tables in table_info
    table_info$table_id <- add_hex_postfix(result, "table")
    
    # Name the tables by their id and set the tables attribute
    structure(result, names = table_info$table_id, tables = table_info)
  }
}

# merge_table_infos ------------------------------------------------------------
merge_table_infos <- function(table_infos)
{
  table_info <- kwb.utils::rbindAll(table_infos, nameColumn = "sheet_id")
  
  extract_hex <- function(x) stringr::str_extract(x, "[0-9a-f]+$")
  
  table_info$table_id <- sprintf(
    "table_%s_%s", 
    extract_hex(table_info$sheet_id), 
    extract_hex(table_info$table_id)
  )
  
  kwb.utils::moveColumnsToFront(table_info, c("table_id", "sheet_id"))
}
