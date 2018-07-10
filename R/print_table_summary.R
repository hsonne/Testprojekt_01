# print_table_summary ----------------------------------------------------------
print_table_summary <- function(tables)
{
  table_info <- get_table_info(tables)
  
  file <- kwb.utils::getAttribute(tables, "file")

  cat(sprintf("\nFile: '%s'\n", basename(file)))
  cat(sprintf("Folder: '%s'\n\n", dirname(file)))
  cat(sprintf("- Number of tables: %d\n", length(tables)))
  cat(sprintf("- Number of sheets: %d\n", length(unique(table_info$sheet_id))))
  cat("- Frequencies of dimensions:\n\n")
  print(get_dim_frequencies(tables))
}

# get_dim_frequencies ----------------------------------------------------------
get_dim_frequencies <- function(tables) {
  dims <- as.data.frame(t(sapply(tables, dim)))

  dim_frequency <- aggregate(
    dims$V1,
    by = list(nrow = dims$V1, ncol = dims$V2), FUN = length
  )

  names(dim_frequency)[3] <- "times"

  row_order <- do.call(order, c(dim_frequency[, 1:2], decreasing = TRUE))

  kwb.utils::resetRowNames(dim_frequency[row_order, ])
}
