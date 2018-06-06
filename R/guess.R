# guess_number_of_headers_from_text_matrix -------------------------------------
guess_number_of_headers_from_text_matrix <- function(x, method = 2, dbg = TRUE)
{
  if (FALSE) {
    method = 2; dbg = TRUE
  }
  
  debug_formatted(dbg, "\nGuessing number of header rows... ")
  
  type_matrix <- get_type_matrix_from_text_matrix(x)
  
  col_types <- guess_column_type_from_type_matrix(type_matrix)
  
  indices_numeric <- which(col_types == "N")

  # If there are not columns that seem to be numeric, we give up and return 1
  n_headers <- if (length(indices_numeric) == 0) {
    
    debug_formatted(dbg, "No numeric columns. Returning 1.\n")
    
    1L
    
  } else {
    
    type_matrix_numeric <- type_matrix[, indices_numeric, drop = FALSE]
    
    n_headers <- if (method == 1) {
      
      n_text_per_col <- colSums(type_matrix_numeric == "T")
      
      n_frequency <- sort(table(n_text_per_col), decreasing = TRUE)
      
      as.integer(names(n_frequency)[1])
      
    } else if (method == 2) {
      
      min(sapply(kwb.utils::asColumnList(type_matrix_numeric), function(types) {
        
        which(types == "N")[1] - 1
      }))
      
    } else {
      
      stop_formatted("Undefined method: %s. Possible values: 1, 2", method)
    }
  }
  
  debug_ok(dbg)
  
  # Return at least 1
  structure(max(1L, n_headers), col_types = col_types)
}

# guess_column_type_from_type_matrix -------------------------------------------
guess_column_type_from_type_matrix <- function(type_matrix)
{
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
get_type_matrix_from_text_matrix <- function(x)
{
  kwb.utils::stopIfNotMatrix(x)
  
  type_matrix <- matrix(" ", nrow = nrow(x), ncol = ncol(x))
  
  type_matrix[] <- "_"
  
  type_matrix[! is.na(x)] <- "T"
  
  type_matrix[which(seems_to_be_numeric(x))] <- "N"
  
  type_matrix  
}

# seems_to_be_numeric ----------------------------------------------------------
# simple test for numeric
seems_to_be_numeric <- function(x)
{
  grepl("^[+-]?[0-9.,]+$", kwb.utils::removeSpaces(x))
}
