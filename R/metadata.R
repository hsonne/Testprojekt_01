# export_table_metadata --------------------------------------------------------

#' Export Table Metadata
#' 
#' Export Table Metadata to CSV file 
#' 
export_table_metadata <- function(tables, dbg = TRUE)
{
  table_info <- kwb.utils::getAttribute(tables, "tables")

  # Remove column "topleft" and add columns "skip" and "n_headers". 
  # Column "skip": If the user puts an "x" into this column, the corresponding
  # table will not be imported.
  # Column "n_headers": Number of header lines of the corresponding table
  table_info <-table_info[, names(table_info) != "topleft"]

  table_info$skip <- ""
  
  table_info$n_headers <- 1
  
  file <- kwb.utils::getAttribute(tables, "file")
  
  file_csv <- paste0(kwb.utils::removeExtension(file), "_META_tmp.csv")

  kwb.utils::catIf(dbg, sprintf("Writing table medatada to '%s'... ", file_csv))
  
  utils::write.csv(table_info, file = file_csv, row.names = FALSE, na = "")
  
  kwb.utils::catIf(dbg, "ok.\n")
}

# import_table_metadata --------------------------------------------------------
import_table_metadata <- function(file)
{
  file_csv <- paste0(kwb.utils::removeExtension(file), "_META.csv")

  if (file.exists(file_csv)) {
    
    cat(sprintf(
      "  Reading table metadata from\n    '%s'... ", basename(file_csv)
    ))
    
    utils::read.csv(file_csv, stringsAsFactors = FALSE)
    
    cat(sprintf("ok.\n"))
    
  } else {
    
    cat(sprintf("  No metadata available for\n    '%s'.\n", basename(file)))
    
    NULL
  }
}
