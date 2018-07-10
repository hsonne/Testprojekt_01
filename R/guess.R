# guess_number_of_headers_from_text_matrix -------------------------------------
guess_number_of_headers_from_text_matrix <- function(
  x, table_id, guess_max = 100, dbg = TRUE
)
{
  if (FALSE) {
    method <- 2
    dbg <- TRUE
  }

  debug_formatted(
    dbg, "\nGuessing number of header rows in '%s' ... ", table_id
  )
  
  if (! is.null(guess_max)) {
    
    x <- x[seq_len(min(guess_max, nrow(x))), , drop = FALSE]
  }

  type_matrix <- get_type_matrix_from_text_matrix(x)

  col_types <- guess_column_type_from_type_matrix(type_matrix)

  indices_numeric <- which(col_types == "N")

  # If there are not columns that seem to be numeric, we give up and return 1
  n_headers <- if (length(indices_numeric) == 0) {
    debug_formatted(dbg, "No numeric columns. Returning 1.\n")

    1L
    
  } else {
    
    type_matrix_numeric <- type_matrix[, indices_numeric, drop = FALSE]

    which_first_numeric <- function(x) which(x == "N")[1]

    column_list <- kwb.utils::asColumnList(type_matrix_numeric)
    
    n_headers <- min(sapply(column_list, which_first_numeric)) - 1
  }

  debug_ok(dbg)
  
  # If n_headers is 0, return nontheless 1 in case that the upper left field is 
  # empty. This is assumed to be a cross table with numeric column headers
  if (n_headers == 0 && kwb.utils::isNaOrEmpty(x[1, 1])) {
    
    n_headers <- 1
  } 

  # Return column types in attribute
  structure(n_headers, col_types = col_types)
}

# guess_column_type_from_type_matrix -------------------------------------------
guess_column_type_from_type_matrix <- function(type_matrix) {
  sapply(kwb.utils::asColumnList(type_matrix), function(x) {
    x <- x[x != "_"]

    if (length(x)) {
      names(sort(table(x), decreasing = TRUE))[1]
    } else {
      NA
    }
  })
}

# get_type_matrix_from_text_matrix ---------------------------------------------
get_type_matrix_from_text_matrix <- function(x) {
  kwb.utils::stopIfNotMatrix(x)

  type_matrix <- matrix(" ", nrow = nrow(x), ncol = ncol(x))

  type_matrix[] <- "_"

  type_matrix[!is.na(x)] <- "T"

  type_matrix[which(seems_to_be_numeric(x))] <- "N"

  type_matrix
}

# seems_to_be_numeric ----------------------------------------------------------
# simple test for numeric
seems_to_be_numeric <- function(x) {
  grepl("^[+-]?[0-9.,]+$", kwb.utils::removeSpaces(x))
}
