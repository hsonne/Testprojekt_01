library(testthat)

# 1. Source main.R first!
# 2. Source this script

file_database <- to_file_database(files)
file_database$files
file_database$folders

# M A I N ----------------------------------------------------------------------
if (FALSE) {
  # Get all tables from one file
  tables <- get_text_tables_from_xlsx(file = files[1])

  # Get table metadata
  table_info <- get_table_info(tables)
  
  # Create column metadata from the table headers
  column_info <- create_column_metadata(tables)

  # Add column "skip": If the user puts an "x" into this column, the
  # corresponding table will not be imported.
  table_info$skip <- ""

  # Write table metadata to "<basename>_META.csv"
  export_table_metadata(
    structure(table_info, file = kwb.utils::getAttribute(tables, "file"))
  )
  
  # import_table_metadata returns NULL if no metadata file exists
  import_table_metadata(files[5])

  # Select all file indices
  indices <- seq_along(files)
  indices <- 1:2
  
  # Change indices to test with less files
  #indices <- 16
  
  # Clear the screen
  kwb.utils::clearConsole()

  # Get all tables from all files
  system.time(all_tables <- lapply(indices, function(index) {
    cat("File index:", index)
    get_text_tables_from_xlsx(files[index])
  }))

  #   user  system elapsed 
  # 93.604   2.936  97.305   
  
  names(all_tables) <- file_database$files$file_id[indices]
  
  # Create column metadata for all tables
  column_info_list <- lapply(all_tables, create_column_metadata)

  column_info <- rbindAll(
    column_info_list,
    nameColumn = "file_id", namesAsFactor = FALSE
  )
  
  x <- compact_column_info(column_info)

  nrow(x)
  # 6141
  
  column_info <- suggest_column_name(column_info)
  
  column_info <- merge(column_info, file_database$files)
  
  column_info <- merge(column_info, file_database$folders)
  
  base_dir <- getAttribute(file_database$folders, "base_dir")

  file_metadata <- file.path(base_dir, "METADATA_columns_tmp.csv")

  write.csv(column_info, file_metadata, row.names = FALSE)

  # TODO: Rename METADATA_columns_tmp.csv to METADATA_columns.csv, let the user
  # modify the file and read back into column_info
  #
  #column_info <- read_column_info(safePath(base_dir, "METADATA_columns.csv"))

  # Use column info to convert the text tables into data frames
  all_data <- convert_text_matrix_list_to_data_frames(all_tables, column_info)
  
  lapply(all_data, function(all_tables) lapply(all_tables, head))
    
  file_database$files$file_id
  
  # Get the path to a log file  
  logfile_summary <- tempfile("table_summary_", fileext = ".txt")
  logfile_headers <- tempfile("table_headers_", fileext = ".txt")

  # Write a summary of the read structure to the log file
  capture.output(file = logfile_summary, {
    for (tables in all_tables) print_table_summary(tables)
  })

  # Let's have a look at the tables in one Excel file only
  tables <- all_tables[[1]]

  # Get a description of the sheets in that file
  get_sheet_info(tables)
  
  # Get a description of tables in that file
  get_table_info(tables)

  # Get the name of the file that was read
  kwb.utils::getAttribute(tables, "file")

  # The tables are named by sheet number and table number within the sheet
  # The numbers are hexadecimal, i.e a = 10, f = 15, 10 = 16, ff = 255,
  names(tables)
  # "table_01_01" "table_02_01"

  # Try to guess the header rows for each table...
  n_headers <- sapply(names(tables), function(table_id) {
    guess_number_of_headers_from_text_matrix(
      tables[[table_id]],
      table_id = table_id
    )
  })

  print_logical_matrix(guess_header_matrix(x = tables$table_01_01, n_max))
  print_logical_matrix(guess_header_matrix(x = tables$table_02_01, n_max))

  print_header_guesses(tables, n_max, file = logfile_headers)

  lapply(all_tables[[3]], guess_header_matrix, 4)

  head(x <- tables$table_02_01)

  is_empty <- (is.na(x) | x == "")

  print_logical_matrix(head(is_empty))
  print_logical_matrix(head(is_empty), invert = TRUE)
}

# Text Matrices to data frames -------------------------------------------------
if (FALSE) {
  # Convert text matrices of known format
  tables <- get_text_tables_from_xlsx(file = files[1])

  selected <- grepl("^table_02_", names(tables))

  tables_year_well <- lapply(tables[selected], text_matrix_to_numeric_matrix)

  data_frames_year_well <- lapply(tables_year_well, as.data.frame)

  str(data_frames_year_well$table_02_01)
  str(data_frames_year_well$table_02_02)
}

# print_header_guesses ---------------------------------------------------------
print_header_guesses <- function(
                                 text_matrices, n_max = 5, file = NULL, dbg = TRUE) {
  if (!is.null(file)) {
    debug_formatted(dbg, "Writing output to '%s'... ", file)

    capture.output(file = file, print_header_guesses(text_matrices, n_max))

    debug_ok(dbg)
  } else {

    # matrix_name <- "table_03_01"

    for (matrix_name in names(text_matrices)) {
      header <- guess_header_matrix(x = text_matrices[[matrix_name]], n_max)

      debug_formatted(dbg, "\n%s:\n", matrix_name)

      print_logical_matrix(header)
    }
  }
}

# guess_header_matrix ----------------------------------------------------------
guess_header_matrix <- function(x, n_max = 10) {
  stopifnot(is.character(x))

  kwb.utils::stopIfNotMatrix(x)

  x_head <- as.data.frame(head(x, n_max))

  do.call(cbind, lapply(x_head, function(column_values) {
    sapply(seq_along(column_values), function(i) {
      as.integer(!(column_values[i] %in% column_values[-(1:i)]))
    })
  }))
}

# print_logical_matrix ---------------------------------------------------------
print_logical_matrix <- function(x, invert = FALSE) {
  stopifnot(is.matrix(x))

  y <- matrix(" ", nrow = nrow(x), ncol = ncol(x))

  x <- as.logical(x)

  if (invert) {
    x <- !x
  }

  y[x] <- "x"

  kwb.utils::catLines(kwb.utils::pasteColumns(as.data.frame(y), sep = "|"))
}

# text_matrix_to_numeric_matrix ------------------------------------------------
text_matrix_to_numeric_matrix <- function(x) {
  print(x)

  matrix(
    as.numeric(x[-1, -1]),
    nrow = nrow(x) - 1,
    dimnames = list(x[-1, 1], x[1, -1])
  )
}

test_that("get_raw_text_from_xlsx() works as expected", {
  text_sheets <<- get_raw_text_from_xlsx(file = files[1])

  expect_named(text_sheets)

  expect_true(all(grepl("^sheet_", names(text_sheets))))

  sheet_info <- attr(text_sheets, "sheet_info")

  expect_false(is.null(sheet_info))

  expect_true(all(c("sheet_id", "sheet_name") %in% names(sheet_info)))
})

test_that("split_into_tables() works as expected", {
  tables_list <- list(
    split_into_tables(text_sheets$sheet_01),
    split_into_tables(text_sheets$sheet_02),
    split_into_tables(text_sheets$sheet_03)
  )

  for (tables in tables_list) {
    expect_named(tables)
    expect_true(all(grepl("^table_", names(tables))))

    table_info <- attr(tables, "table_info")

    expect_false(is.null(table_info))

    columns <- c(
      "table_id", "table_name", "first_row", "last_row", "first_col",
      "last_col", "n_headers"
    )

    expect_true(all(columns %in% names(table_info)))
  }
})
