library(fs)

copy_xlsx_files <- function(input_dir,
                            export_dir, 
                            overwrite = FALSE,
                            recursive = TRUE,
                            file_pattern = "[xX][lL][sS][xX]") {
  
  
  input_dir <- normalizePath(input_dir)
  
  export_dir <- normalizePath(export_dir)
  
  
  xlsx_import_paths <- normalizePath(dir(path = input_dir, 
                                         pattern =  file_pattern, 
                                         recursive = TRUE, 
                                         full.names = TRUE))
  
  
  xlsx_export_paths <- gsub(pattern = input_dir, 
                            replacement = export_dir, 
                            x = xlsx_import_paths ,
                            fixed = TRUE)
  
  
  fs::dir_create(path = normalizePath(dirname(xlsx_import_paths)),
                 recursive = TRUE) 
  
  
  for(file_index in seq_along(xlsx_import_paths)) {
    cat(sprintf("Copying file (%d/%d):\nFROM: %s\nTO: %s\n", 
                file_index,
                length(xlsx_import_paths),
                xlsx_import_paths[file_index],
                xlsx_export_paths[file_index]))
    fs::file_copy(path = xlsx_import_paths[file_index],
                  new_path = xlsx_export_paths[file_index],
                  overwrite = overwrite) 
  }
}
