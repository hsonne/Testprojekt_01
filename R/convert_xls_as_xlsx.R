library(fs)
library(kwb.utils)

get_excelcnv_exe <- function(
  office_folder = kwb.utils::safePath("C:/Program Files (x86)/Microsoft Office")
)
{
  paths <- normalizePath(list.files(
    office_folder, "excelcnv.exe", recursive = TRUE, full.names = TRUE
  ))
  
  n_files <- length(paths)
  if (n_files == 0) {
    
    stop(
      "Excel conversion tool 'excelcnv.exe' not found under: ", office_folder
    )
  }
  
  if (n_files > 1) {
    
    # Sort (in case of multiple executables newest is at index 1)
    paths <- sort(paths, decreasing = TRUE)
    
    cat(paste0(
      "Found ", n_files, "versions of 'excelcnv.exe':\n  ",
      kwb.utils::collapsed(paths, "\n  "),
      "\nUsing the latest one:", paths[1]
    ))
  }
  
  paths[1]
}


delete_registry <- function(
  office_folder = kwb.utils::safePath("C:/Program Files (x86)/Microsoft Office"),
  dbg = TRUE) {

  office_path_split <- strsplit(dirname(get_excelcnv_exe(office_folder)),
                                split = "/")[[1]]
  
  
  office_version <- sprintf("%s.0", 
                            stringr::str_extract(office_path_split[length(office_path_split)],
                                                 pattern = "[1][0-9]"))
  
  ### Delete registry entry:
  ### http://justgeeks.blogspot.com/2014/08/free-convert-for-excel-files-xls-to-xlsx.html
  
  reg_entry <- paste0("HKEY_CURRENT_USER\\Software\\", 
                      "Microsoft\\Office\\", 
                       office_version, 
                       "\\Excel\\Resiliency\\StartupItems")
  
  
  cmd_delete_registry <- sprintf('reg delete %s /f', reg_entry)
  
  if(dbg) {
    cat(sprintf("Deleting registry entry:\n%s", 
                cmd_delete_registry))  
    
  }
  
  system(command = cmd_delete_registry)
  
  
}


### Convert xls to xlsx in temp dir

convert_xls_as_xlsx <- function(
  input_dir = "Y:/Z-Exchange/Jeansen/Daten_Labor",
  export_dir = tempdir(), 
  office_folder = kwb.utils::safePath("C:/Program Files (x86)/Microsoft Office"),
  dbg = TRUE) {


pattern <- "\\.([xX][lL][sS])$"

input_dir <- normalizePath(input_dir)

export_dir <- normalizePath(export_dir)


xls_files <-  normalizePath(dir(path = input_dir, 
                  pattern = pattern, 
                  recursive = TRUE, 
                  full.names = TRUE))


xlsx_files <- gsub(pattern = input_dir, 
                   replacement = export_dir, 
                   x = xls_files,
                   fixed = TRUE)

xlsx_files <- gsub(pattern = pattern, 
                   replacement = ".xlsx", 
                   x = xlsx_files)


fs::dir_create(path = normalizePath(dirname(xlsx_files)),
               recursive = TRUE) 


for (index in seq_along(xls_files)) {
  
  cmd <- sprintf('"%s" -oice "%s" "%s"',
                 get_excelcnv_exe(office_folder),
                 xls_files[index],
                 xlsx_files[index])
  
  if(dbg) {
  cat(sprintf("Converting xls to xlsx (%d/%d):\n%s", 
                index, 
                length(xls_files),
                cmd))
  } 
  
  system(command = cmd)
  delete_registry(office_folder = office_folder,
                  dbg = dbg)
  
}


}

