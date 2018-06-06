library(testthat)

# 1. Source main.R first!
# 2. Source this script

# M A I N ----------------------------------------------------------------------
if (FALSE)
{
  text_sheets <- get_raw_text_from_xlsx(file = files[1])
  
  expect_named(text_sheets)
  expect_true(all(grepl("^sheet_", names(text_sheets))))
  expect_true(! is.null(attr(text_sheets, "sheets")))
  
  # attr(text_sheets, "sheets")
  #   sheet_id            sheet_name
  # 1 sheet_01     Einzelwerte Liste
  # 2 sheet_02 Kreuztab Jahresmittel
  # 3 sheet_03                  Test
  
  tables_list <- list(
    split_into_tables(text_sheet <- text_sheets$sheet_01),
    split_into_tables(text_sheet <- text_sheets$sheet_02),
    split_into_tables(text_sheet <- text_sheets$sheet_03)
  )
  
  for (tables in tables_list) {
    
    expect_named(tables)
    expect_true(all(grepl("^table_", names(tables))))
    expect_true(! is.null(attr(tables, "tables")))
  }
  
  # head(attr(tables_list[[2]], "tables")[, -9])
  #   table_id table_name nrow ncol first_row last_row first_col last_col
  # 1 table_01               1    1         2        2         1        1 
  # 2 table_02              14   19         4       17         1       19
  # 3 table_03               1    1        19       19         1        1
  # 4 table_04              17   34        21       37         1       34
  # 5 table_05               1    1        39       39         1        1
  # 6 table_06              15   16        41       55         1       16
  
  # Get all tables from one file
  tables <- get_text_tables_from_xlsx(file = files[8])

  # Get table metadata
  table_info <- kwb.utils::getAttribute(tables, "tables")
  
  # Create column metadata from the table headers
  column_info <- create_column_metadata(tables)

  # Add column "skip": If the user puts an "x" into this column, the 
  # corresponding table will not be imported.
  table_info$skip <- ""

  # Write table metadata to "<basename>_META.csv"
  export_table_metadata(
    structure(table_info, file = kwb.utils::getAttribute(tables, "file"))
  )
  
  table_info <- import_table_metadata(files[5])

  # Create list representing a file database
  file_db <- to_file_database(files)
  
  indices <- seq_along(files)
  
  # Get all tables from all files
  system.time(all_tables <- lapply(files[indices], get_text_tables_from_xlsx))

  #    user  system elapsed 
  # 118.868   3.316 123.309 

  names(all_tables) <- file_db$files$file_id[indices]
  
  # Create column metadata for all tables
  column_info_list <- lapply(all_tables, create_column_metadata)
  
  column_info <- rbindAll(
    column_info_list, nameColumn = "file_id", namesAsFactor = FALSE
  )
  
  column_info <- merge(column_info, file_database$files)
  column_info <- merge(column_info, file_database$folders)
  base_dir <- getAttribute(file_database$folders, "base_dir")
  
  file_metadata <- file.path(base_dir, "METADATA_columns_tmp.csv")

  write.csv(column_info, file_metadata, row.names = FALSE)

  # TODO: Copy METADATA_columns_tmp.csv to METADATA_columns.csv, let the user
  # modify the file and read back into column_info
  #
  # column_info <- read_column_info(safePath(base_dir, "METADATA_columns.csv"))

  # Use column info to convert the text tables into data frames
  all_data <- convert_text_matrix_list_to_data_frames(all_tables, column_info)
  
  file_db$files$file_id
  
  column_info$file <- files[column_info$file_index]
  
  guess_number_of_headers_from_text_matrix(x <- all_tables[[1]]$table_03_01)

  head(x, 3)
  
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
  kwb.utils::getAttribute(tables, "sheets")
  
  # Get a description of tables in that file
  kwb.utils::getAttribute(tables, "tables")

  # Get the name of the file that was read
  kwb.utils::getAttribute(tables, "file")
  
  # The tables are named by sheet number and table number within the sheet
  # The numbers are hexadecimal, i.e a = 10, f = 15, 10 = 16, ff = 255, 
  names(tables)
  # "table_01_01" "table_02_01"

  # Try to guess the header rows for each table...
  n_headers <- sapply(tables, guess_number_of_headers_from_text_matrix)

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
if (FALSE)
{
  # Convert text matrices of known format
  tables <- get_text_tables_from_xlsx(file = files[1])
  
  selected <- grepl("^table_02_", names(tables))
  
  tables_year_well <- lapply(tables[selected], text_matrix_to_numeric_matrix)
  
  data_frames_year_well <- lapply(tables_year_well, as.data.frame)
  
  str(data_frames_year_well$table_02_01)
  str(data_frames_year_well$table_02_02)
}

# print_header_guesses ---------------------------------------------------------
print_header_guesses <- function(text_matrices, n_max = 5, file = NULL)
{
  if (! is.null(file)) {
  
    cat(sprintf("Writing output to '%s'... ", file))
    
    capture.output(file = file, print_header_guesses(text_matrices, n_max))
    
    cat("ok.\n")
    
  } else {
    
    #matrix_name <- "table_03_01"
    
    for (matrix_name in names(text_matrices)) {
      
      header <- guess_header_matrix(x = text_matrices[[matrix_name]], n_max)
      
      cat(sprintf("\n%s:\n", matrix_name))
      
      print_logical_matrix(header)
    }
  }
}

# guess_header_matrix ----------------------------------------------------------
guess_header_matrix <- function(x, n_max = 10)
{
  stopifnot(is.character(x))
  
  kwb.utils::stopIfNotMatrix(x)
  
  x_head <- as.data.frame(head(x, n_max))
  
  do.call(cbind, lapply(x_head, function(column_values) {
    sapply(seq_along(column_values), function(i) {
      as.integer(! (column_values[i] %in% column_values[-(1:i)]))
    })
  }))
}

# print_logical_matrix ---------------------------------------------------------
print_logical_matrix <- function(x, invert = FALSE)
{
  stopifnot(is.matrix(x))
  
  y <- matrix(" ", nrow = nrow(x), ncol = ncol(x))

  x <- as.logical(x)
  
  if (invert) {
    
    x <- ! x
  }
  
  y[x] <- "x"
  
  kwb.utils::catLines(kwb.utils::pasteColumns(as.data.frame(y), sep = "|"))
}

# text_matrix_to_numeric_matrix ------------------------------------------------
text_matrix_to_numeric_matrix <- function(x)
{
  print(x)
  
  matrix(
    as.numeric(x[-1, -1]), 
    nrow = nrow(x) - 1, 
    dimnames = list(x[-1, 1], x[1, -1])
  )
}
