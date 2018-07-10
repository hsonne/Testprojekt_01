copy_lookup_para_file <- function(
                                  from_dir, to_dir, overwrite = FALSE, recursive = TRUE,
                                  file_pattern = "^lookup_para\\.csv$") {
  from_dir <- normalizePath(from_dir)

  to_dir <- normalizePath(to_dir)

  from_path <- normalizePath(dir(
    from_dir, file_pattern,
    recursive = TRUE, full.names = TRUE
  ))

  if (file.exists(from_path)) {
    to_path <- gsub(from_dir, to_dir, from_path, fixed = TRUE)

    fs::dir_create(dirname(to_path), recursive = TRUE)


    cat(sprintf(
      "\nCopying 'lookup_para' file:\nFROM: %s\nTO  : %s\n",
      from_path, to_path
    ))

    fs::file_copy(from_path, to_path, overwrite = overwrite)
  } else {
    stop(sprintf(
      "No 'lookup_para' file found in the following input dir:\n%s",
      from_dir
    ))
  }
}
