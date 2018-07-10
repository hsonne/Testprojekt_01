# get_text_tables_from_xlsx ----------------------------------------------------
get_text_tables_from_xlsx <- function(file, table_info = NULL, dbg = TRUE)
{
  #kwb.utils::assignArgumentDefaults(get_text_tables_from_xlsx)

  # Get one character matrix per sheet
  all_text_sheets <- get_raw_text_from_xlsx(file)

  # Inform about and remove empty sheets
  text_sheets <- remove_empty_sheets(all_text_sheets)
  
  # If there are no table metadata, try to load them from a file
  #table_info <- kwb.utils::defaultIfNULL(table_info, import_table_metadata(file))

  # If there was no metadata file, split sheets into tables and create metadata
  if (! is.null(table_info)) {
    
    stop("Not implemented: extract tables with known ranges")
    #extract_tables_with_ranges(text_sheets, table_info)
  }
  
  #all_tables <- extract_tables_without_ranges(text_sheets)

  # Split into tables and name
  all_tables <- lapply(text_sheets, get_tables_from_sheet)
  
  table_infos <- lapply(all_tables, get_table_info)
  
  table_info <- merge_table_infos(table_infos)

  result <- do.call(c, all_tables)
  
  structure(
    result, 
    table_info = table_info, 
    sheet_info = get_sheet_info(text_sheets), 
    file = file
    #,names = table_info$table_id
  )
}

# remove_empty_sheets ----------------------------------------------------------
remove_empty_sheets <- function(text_sheets)
{
  # Identify empty sheets
  is_empty <- (sapply(text_sheets, length) == 0)
  
  if (any(is_empty)) {

    # Get meta information about the sheets
    sheet_info <- get_sheet_info(text_sheets)

    # Get the names of the empty sheets    
    sheet_names <- sheet_info$sheet_name[is_empty]
    
    # Prepare a sheet enumeration string for the output
    sheet_enumeration <- kwb.utils::stringList(sheet_names, collapse = "\n  ")
    
    # Inform the user about empty sheets
    message("Skipping empty sheet(s):\n  ", sheet_enumeration)
    
    # Remove empty sheets and update the sheet info
    set_sheet_info(text_sheets[! is_empty], sheet_info[! is_empty, ])
    
  } else {
    
    text_sheets
  }
}

# get_tables_from_sheet --------------------------------------------------------
get_tables_from_sheet <- function(text_sheet, dbg = TRUE)
{
  #text_sheet <- text_sheets[[1]]
  #kwb.utils::printIf(TRUE, head(text_sheet))
  
  tables <- split_into_tables(text_sheet)

  table_info <- provide_header_info_if_required(get_table_info(tables), tables)
  
  if (length(tables) > 1) {
    
    if (looks_like_one_table(table_info)) {
      
      tables <- merge_tables_to_one_table(tables, table_info, dbg)
      
    } else if (looks_like_tables_with_captions(tables)) {
      
      tables <- name_even_tables_by_odd_tables(tables, dbg)
    }
  }
  
  (sheet_name <- getAttribute(text_sheet, "sheet"))

  # If there are no table names yet name the tables
  set_table_info(tables, name_unnamed_tables(table_info, sheet_name))
}

# split_into_tables ------------------------------------------------------------
split_into_tables <- function(
  text_sheet, sheet_name = NULL, delete_all_empty_columns = FALSE, dbg = TRUE
)
{
  stopifnot(is.character(text_sheet), length(dim(text_sheet)) == 2)
  
  if (is.null(sheet_name)) {
    
    sheet_name <- getAttribute(text_sheet, "sheet")
  }
  
  # Get indices of non-empty rows
  indices <- which(!kwb.utils::isNaInAllColumns(text_sheet))

  # Create ranges of "connected" non-empty rows
  table_info <- kwb.event::hsEvents(indices, 1, 1)[, 3:4]

  names(table_info) <- c("first_row", "last_row")

  text <- "Extracting table areas from sheet"
  
  if (! is.null(sheet_name)) {
    
    text <- paste0(text, " '", sheet_name, "'")
  }
  
  debug_formatted(dbg, "%s ... ", text)
  
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

  debug_ok(dbg)
  
  table_info <- extend_table_info(table_info, tables)

  names(tables) <- table_info$table_id

  structure(tables, table_info = table_info, class = "xls_tables")
}

# extend_table_info ------------------------------------------------------------
extend_table_info <- function(table_info, tables) {
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

# provide_header_info_if_required ----------------------------------------------
provide_header_info_if_required <- function(table_info, tables)
{
  header_info_required <- all(is.na(get_col(table_info, "n_headers")))
  
  if (header_info_required) {

    n_headers_list <- lapply(names(tables), function(table_id) {
      
      x <- get_ele(tables, table_id)
      
      guess_number_of_headers_from_text_matrix(x, table_id, dbg = FALSE)
    })
    
    table_info <- update_table_info_with_header_info(table_info, n_headers_list)
  }
  
  table_info
}

# guess_all_headers ------------------------------------------------------------
guess_all_headers <- function(all_tables, table_info, dbg = TRUE)
{
  debug_formatted(dbg, "Guessing numbers of header lines in each table... ")
  
  n_headers_list <- lapply(names(all_tables), function(table_id) {
    
    #print(table_id)
    #table_id <- names(all_tables)[2]
    
    guess_number_of_headers_from_text_matrix(
      x = all_tables[[table_id]], table_id = table_id, dbg = FALSE
    )
  })
  
  debug_ok(dbg)
  
  if ("skip" %in% names(table_info)) {
    
    keep <- kwb.utils::isNaOrEmpty(table_info$skip)
    
    debug_formatted(dbg, "Skipping %d tables ... ", sum(! keep))
    
    table_info <- table_info[keep, ]
    
    debug_ok(dbg)
  }
  
  #headtail(table_info)
  
  update_table_info_with_header_info(table_info, n_headers_list)  
}

# update_table_info_with_header_info -------------------------------------------
update_table_info_with_header_info <- function(table_info, n_headers_list)
{
  # Define helper function
  get_collapsed_col_types <- function(x) {
    
    kwb.utils::collapsed(kwb.utils::getAttribute(x, "col_types"), "|")
  }
  
  kwb.utils::setColumns(
    table_info,
    n_headers = unlist(n_headers_list),
    col_types = sapply(n_headers_list, get_collapsed_col_types)
  )
}

# looks_like_one_table ---------------------------------------------------------
looks_like_one_table <- function(table_info)
{
  all_equal <- kwb.utils::allAreEqual
  
  if (nrow(table_info) > 1) {
    
    same_first_col <- all_equal(get_col(table_info, "first_col"))
    same_last_col <- all_equal(get_col(table_info, "last_col"))
    
    n_headers <- get_col(table_info, "n_headers")
    
    first_has_header <- (n_headers[1] > 0)
    others_dont <- all(n_headers[-1] == 0)
    
    same_first_col && same_last_col && first_has_header && others_dont
    
  } else {
    
    FALSE
  }
}

# looks_like_tables_with_captions ----------------------------------------------
looks_like_tables_with_captions <- function(tables)
{
  length(tables[[1]]) == 1
}

# merge_tables_to_one_table ----------------------------------------------------
merge_tables_to_one_table <- function(tables, table_info = NULL, dbg = TRUE)
{
  debug_formatted(dbg, "Row-binding all tables of this sheet ... ")

  table_info <- kwb.utils::defaultIfNULL(table_info, get_table_info(tables))
  
  tables <- list(do.call(rbind, tables))
  
  table_info$last_row[1] <- get_col(table_info, "last_row")[nrow(table_info)]

  table_info <- table_info[1, ]
  
  result <- set_table_info(tables, table_info)
  
  debug_ok(dbg)
  
  stats::setNames(result, get_col(table_info, "table_id"))
}

# name_even_tables_by_odd_tables -----------------------------------------------
name_even_tables_by_odd_tables <- function(tables, dbg = TRUE)
{
  kwb.utils::catIf(
    dbg, "Trying to name tables at even position by tables at odd",
    "positions ... "
  )
  
  stopifnot(inherits(tables, "xls_tables"))

  table_info <- get_table_info(tables)

  sizes <- sapply(tables, length)

  stopifnot(all(sizes == table_info$nrow * table_info$ncol))

  n_tables <- length(tables)

  odd_indices <- seq(1, n_tables, by = 2)

  even_indices <- seq(2, n_tables, by = 2)
  
  result <- if (! all(sizes[odd_indices] == 1)) {
    
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
  
  debug_ok(dbg)
  
  result
}

# name_unnamed_tables ----------------------------------------------------------
name_unnamed_tables <- function(table_info, sheet) 
{
  table_names <- get_col(table_info, "table_name")
  
  is_unnamed <- kwb.utils::isNaOrEmpty(table_names)
  
  if (any(is_unnamed)) {
    
    n_tables <- nrow(table_info)
    
    table_names <- if (n_tables == 1) {
      
      sheet
      
    } else {
      
      add_hex_postfix(seq_len(n_tables), base_name = sheet)
    }
    
    table_info$table_name[is_unnamed] <- table_names[is_unnamed]
  }
    
  table_info
}

# merge_table_infos ------------------------------------------------------------
merge_table_infos <- function(table_infos) {
  table_info <- kwb.utils::rbindAll(
    table_infos,
    nameColumn = "sheet_id", namesAsFactor = FALSE
  )

  extract_hex <- function(x) stringr::str_extract(x, "[0-9a-f]+$")

  table_info$table_id <- sprintf(
    "table_%s_%s",
    extract_hex(table_info$sheet_id),
    extract_hex(table_info$table_id)
  )

  kwb.utils::moveColumnsToFront(table_info, c("table_id", "sheet_id"))
}

# extract_tables_with_ranges ---------------------------------------------------
extract_tables_with_ranges <- function(text_sheets, table_info)
{
  selected <- kwb.utils::isNaOrEmpty(get_col(table_info, "skip"))

  table_info <- table_info[selected, ]

  tables <- lapply(seq_len(nrow(table_info)), function(i) {
    sheet_id <- as.character(get_col(table_info, "sheet_id")[i])

    sheet_rows <- get_ele(text_sheets, sheet_id)
    
    from_to <- get_col(table_info, c("first_row", "last_row"))
    
    row_indices <- seq(from = from_to[i, 1], to = from_to[i, 2])

    #col_range <- seq(table_info$first_col[i], table_info$last_col[i])
    col_range <- seq_len(ncol(sheet_rows))
    
    table_matrix <- sheet_rows[row_indices, , drop = FALSE]
    
    table_data <- as.data.frame(table_matrix)
    
    as.matrix(kwb.utils::removeEmptyColumns(table_data, dbg = FALSE))
  })

  stats::setNames(tables, get_col(table_info, "table_id"))
}
