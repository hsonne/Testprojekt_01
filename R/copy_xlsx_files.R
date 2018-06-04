copy_xlsx_files <- function(
  from_dir, to_dir, overwrite = FALSE, recursive = TRUE,
  file_pattern = "[xX][lL][sS][xX]"
)
{
  from_dir <- normalizePath(from_dir)
  
  to_dir <- normalizePath(to_dir)

  from_paths <- normalizePath(dir(
    from_dir, file_pattern, recursive = TRUE, full.names = TRUE
  ))

  to_paths <- gsub(from_dir, to_dir, from_paths, fixed = TRUE)

  fs::dir_create(normalizePath(dirname(from_paths)), recursive = TRUE)

  for (i in seq_along(from_paths)) {
    
    cat(sprintf(
      "\nCopying file (%d/%d):\nFROM: %s\nTO  : %s\n", 
      i, length(from_paths), from_paths[i], to_paths[i]
    ))
    
    fs::file_copy(from_paths[i], to_paths[i], overwrite = overwrite)
  }
}
