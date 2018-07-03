# to_file_database -------------------------------------------------------------
to_file_database <- function(files, remove_common_base = TRUE) {
  folder_table <- to_folder_table(dirname(files), remove_common_base)

  lookup_table <- folder_table

  if (isTRUE(remove_common_base)) {
    base_dir <- kwb.utils::getAttribute(folder_table, "base_dir")

    lookup_table$folder_path <- file.path(base_dir, lookup_table$folder_path)
  }

  # Remove trailing slashes
  lookup_table$folder_path <- gsub("/$", "", lookup_table$folder_path)

  list(
    files = to_file_table(files, lookup_table),
    folders = folder_table
  )
}

# to_folder_table -------------------------------------------------------------
to_folder_table <- function(folder_paths, remove_common_base = TRUE) {
  folders <- sort(unique(folder_paths))

  folder_table <- no_factors_data_frame(
    folder_id = add_hex_postfix(folders),
    folder_path = folders
  )

  if (isTRUE(remove_common_base)) {
    subdirs <- kwb.fakin:::splitPaths(folder_table$folder_path)

    subdirs <- kwb.fakin:::removeCommonRoot(subdirs)

    folder_table$folder_path <- paste_path_parts(subdirs)

    attr(folder_table, "base_dir") <- kwb.utils::getAttribute(subdirs, "root")
  }

  folder_table
}

# to_file_table ----------------------------------------------------------------
to_file_table <- function(files, lookup_table) {
  no_factors_data_frame(
    file_id = sprintf("file_%02X", seq_along(files)),
    file_name = basename(files),
    folder_id = sapply(
      dirname(files), kwb.utils::tableLookup,
      x = lookup_table[, 2:1]
    )
  )
}
