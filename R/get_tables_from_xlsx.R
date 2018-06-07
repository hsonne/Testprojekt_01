# get_text_tables_from_xlsx ----------------------------------------------------
get_text_tables_from_xlsx <- function(
  file, table_info = import_table_metadata(file), guess_headers = TRUE,
  dbg = TRUE
)
{
  if (FALSE) {
    table_info = import_table_metadata(file); guess_headers = TRUE
  }
  
  # Get one character matrix per sheet
  text_sheets <- get_raw_text_from_xlsx(file)

  sheet_info <- kwb.utils::getAttribute(text_sheets, "sheet_info")
  
  is_empty <- colSums(sapply(text_sheets, dim)) == 0
  
  if (any(is_empty)) {
    
    message(
      "Skipping empty sheet(s):\n  ", 
      kwb.utils::stringList(sheet_info$sheet_name[is_empty], collapse = "\n  ")
    )
  }
  
  if (is.null(table_info)) {

    all_tables <- split_into_tables_and_name(
      text_sheets, 
      sheet_names = sheet_info$sheet_name,
      indices = which(! is_empty)
    )
  
    table_infos <- lapply(all_tables, kwb.utils::getAttribute, "table_info")
    
    table_info <- merge_table_infos(table_infos)

    all_tables <- stats::setNames(do.call(c, all_tables), table_info$table_id)
    
  } else {
    
    all_tables <- extract_tables_from_ranges(text_sheets, table_info)
  }

  if (isTRUE(guess_headers)) {
  
    debug_formatted(dbg, "Guessing numbers of header lines in each table... ")

    n_headers_list <- lapply(names(all_tables), function(table_id) {
      
      #print(table_id)
      #table_id <- names(all_tables)[2]
      
      guess_number_of_headers_from_text_matrix(
        x = all_tables[[table_id]], table_id = table_id, dbg = FALSE
      )
    })
    
    debug_ok(dbg)
    
    table_info$n_headers <- unlist(n_headers_list)
    
    table_info$col_types <- sapply(n_headers_list, function(x) {
      kwb.utils::collapsed(kwb.utils::getAttribute(x, "col_types"), "|")
    })
  }
  
  structure(
    all_tables, table_info = table_info, sheet_info = sheet_info, file = file
  )
}

# split_into_tables_and_name ---------------------------------------------------
split_into_tables_and_name <- function(
  text_sheets, sheet_names, indices, dbg = TRUE
)
{
  lapply(indices, function(i) {

    text_sheet <- text_sheets[[i]]
    
    debug_formatted(
      dbg, "Extracting table areas from sheet '%s' ... ", sheet_names[i]
    )
    
    tables <- split_into_tables(text_sheet)

    debug_ok(dbg)
    
    n_tables <- length(tables)
    
    if (length(tables) > 1 && length(tables[[1]]) == 1) {
      
      debug_formatted(dbg, paste(
        "Trying to name tables at even position by tables at odd positions ... "
      ))
      
      tables <- name_even_tables_by_odd_tables(tables)
      
      debug_ok(dbg)
    }
    
    # If the number of tables did not change, they still need to be named
    if (length(tables) == n_tables) {
      
      tables <- name_tables_by_sheet_and_number(tables, sheet = sheet_names[i])
    }
    
    tables
  })
}

# split_into_tables ------------------------------------------------------------
split_into_tables <- function(
  text_sheet, delete_all_empty_columns = TRUE, dbg = TRUE
)
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
      
      result <- as.matrix(kwb.utils::removeEmptyColumns(df, dbg = FALSE))
    }
    
    result
  })

  table_info <- extend_table_info(table_info, tables)
  
  names(tables) <- table_info$table_id
  
  structure(tables, table_info = table_info, class = "xls_tables")
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
    first_row = table_info$first_row, 
    last_row = table_info$last_row, 
    first_col = 1, 
    last_col = dims[, 2],
    n_headers = as.integer(NA),
    stringsAsFactors = FALSE
  )
}

# name_even_tables_by_odd_tables -----------------------------------------------
name_even_tables_by_odd_tables <- function(tables)
{
  stopifnot(inherits(tables, "xls_tables"))

  table_info <- kwb.utils::getAttribute(tables, "table_info")

  sizes <- sapply(tables, length)
  
  stopifnot(all(sizes == table_info$nrow * table_info$ncol))

  n_tables <- length(tables)
  
  odd_indices <- seq(1, n_tables, by = 2)
  
  even_indices <- seq(2, n_tables, by = 2)
  
  if (! all(sizes[odd_indices] == 1)) {
    
    message(
      "\nNot all tables of length one (expected to be captions) are on odd\n",
      "positions in the list. Returning the original list of tables"
    )
    
    tables
    
  } else {
    
    # Get captions from one-element tables at odd indices
    table_info$table_name[even_indices] <- unname(unlist(tables[odd_indices]))
    
    # Remove the 1x1 tables from table_info
    table_info <- table_info[even_indices, ]
    
    # Keep only the tables at even indices
    result <- tables[even_indices]

    # Renumber the tables in table_info
    table_info$table_id <- add_hex_postfix(result, "table")
    
    # Name the tables by their id and set the tables attribute
    structure(result, names = table_info$table_id, table_info = table_info)
  }
}

# name_tables_by_sheet_and_number ----------------------------------------------
name_tables_by_sheet_and_number <- function(tables, sheet) 
{
  stopifnot(inherits(tables, "xls_tables"))
  
  table_info <- kwb.utils::getAttribute(tables, "table_info")
  
  table_info$table_name <- if (length(tables) == 1) {

    sheet
    
  } else {
    
    add_hex_postfix(tables, base_name = sheet)
  }
  
  structure(tables, table_info = table_info)
}

# merge_table_infos ------------------------------------------------------------
merge_table_infos <- function(table_infos)
{
  table_info <- kwb.utils::rbindAll(
    table_infos, nameColumn = "sheet_id", namesAsFactor = FALSE
  )
  
  extract_hex <- function(x) stringr::str_extract(x, "[0-9a-f]+$")
  
  table_info$table_id <- sprintf(
    "table_%s_%s", 
    extract_hex(table_info$sheet_id), 
    extract_hex(table_info$table_id)
  )
  
  kwb.utils::moveColumnsToFront(table_info, c("table_id", "sheet_id"))
}

# extract_tables_from_ranges ---------------------------------------------------
extract_tables_from_ranges <- function(text_sheets, table_info)
{
  get_col <- kwb.utils::selectColumns
  
  get_ele <- kwb.utils::selectElements
  
  selected <- kwb.utils::isNaOrEmpty(get_col(table_info, "skip"))
  
  table_info <- table_info[selected, ]
  
  tables <- lapply(seq_len(nrow(table_info)), function(i) {
    
    sheet_id <- as.character(get_col(table_info, "sheet_id")[i])
    
    sheet_rows <- get_ele(text_sheets, sheet_id)
    
    row_range <- seq(table_info$first_row[i], table_info$last_row[i])
    
    col_range <- seq(table_info$first_col[i], table_info$last_col[i])
    
    table_matrix <- sheet_rows[row_range, col_range, drop = FALSE]
  })

  stats::setNames(tables, get_col(table_info, "table_id"))
}
