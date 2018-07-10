# get_col ----------------------------------------------------------------------
get_col <- kwb.utils::selectColumns

# get_ele ----------------------------------------------------------------------
get_ele <- kwb.utils::selectElements

# delete_empty_columns_right ---------------------------------------------------
delete_empty_columns_right <- function(x)
{
  stopifnot(length(dim(x)) == 2)
  
  last_column_index <- max(which(! kwb.utils::isNaInAllRows(x)))
  
  x[, seq_len(last_column_index), drop = FALSE]
}

# paste_path_parts -------------------------------------------------------------
paste_path_parts <- function(path_parts)
{
  stopifnot(is.list(path_parts), all(sapply(path_parts, is.character)))
  
  sapply(path_parts, function(p) {
    
    if (length(p) > 0) {
      
      do.call(file.path, as.list(p))
      
    } else {
      
      ""
    }
  })
}

# add_hex_postfix --------------------------------------------------------------
add_hex_postfix <- function(x, base_name = NULL)
{
  default <- kwb.utils::defaultIfNULL
  
  base_name <- default(base_name, gsub("s$", "", deparse(substitute(x))))
  
  sprintf("%s_%02x", base_name, seq_along(x))
}

# no_factors_data_frame --------------------------------------------------------
no_factors_data_frame <- function(...) 
{
  #kwb.utils::resetRowNames(
    data.frame(..., stringsAsFactors = FALSE)
  #)
}
