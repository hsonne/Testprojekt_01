library(fs)
library(kwb.utils)

# get_excelcnv_exe -------------------------------------------------------------
get_excelcnv_exe <- function(office_folder = safe_office_folder())
{
  x <- office_folder
  
  paths <- list.files(x, "excelcnv.exe", recursive = TRUE, full.names = TRUE)
  
  if ((n <- length(paths)) == 0) {
    
    stop("Excel conversion tool 'excelcnv.exe' not found under: ", x)
  }
  
  if (n > 1) {
    
    # Sort (in case of multiple executables newest is at index 1)
    paths <- sort(paths, decreasing = TRUE)
    
    cat(
      "Found", n, "versions of 'excelcnv.exe':\n ", 
      kwb.utils::collapsed(paths, "\n "), "\nUsing the latest one:", paths[1]
    )
  }
  
  paths[1]
}

# safe_office_folder -----------------------------------------------------------
safe_office_folder <- function()
{
  kwb.utils::safePath("C:/Program Files (x86)/Microsoft Office")
}

# delete_registry --------------------------------------------------------------
delete_registry <- function(office_folder = safe_office_folder(), dbg = TRUE) 
{
  exe_path <- get_excelcnv_exe(office_folder)
  
  parent_folder <- basename(dirname(exe_path))

  # Delete registry entry: 
  # http://justgeeks.blogspot.com/2014/08/
  # free-convert-for-excel-files-xls-to-xlsx.html

  patterns <- kwb.utils::resolve(list(
    office = "HKEY_CURRENT_USER\\Software\\Microsoft\\Office",
    reg_entry = "<office>\\<version>.0\\Excel\\Resiliency\\StartupItems",
    command = "reg delete <reg_entry> /f",
    debug = "Deleting registry entry:\n<command>",
    version = stringr::str_extract(parent_folder, pattern = "[1][0-9]")
  ))
  
  kwb.utils::catIf(dbg, patterns$debug)

  system(command = patterns$command)
}

### Convert xls to xlsx in temp dir

convert_xls_as_xlsx <- function(
  input_dir = "Y:/Z-Exchange/Jeansen/Daten_Labor", 
  export_dir = tempdir(), office_folder = safe_office_folder(), dbg = TRUE
) 
{
  input_dir <- normalizePath(input_dir)
  
  export_dir <- normalizePath(export_dir)
  
  pattern <- "\\.([xX][lL][sS])$"
  
  xls_files <- normalizePath(dir(
    input_dir, pattern, recursive = TRUE, full.names = TRUE
  ))

  xlsx_files <- gsub(input_dir, export_dir, xls_files, fixed = TRUE)
  
  xlsx_files <- gsub(pattern, ".xlsx", xlsx_files)

  fs::dir_create(path = normalizePath(dirname(xlsx_files)), recursive = TRUE) 

  exe <- normalizePath(get_excelcnv_exe(office_folder))
  
  for (i in seq_along(xls_files)) {

    command <- sprintf('"%s" -oice "%s" "%s"', exe, xls_files[i], xlsx_files[i])
    
    kwb.utils::catIf(dbg, sprintf(
      "Converting xls to xlsx (%d/%d):\n%s", i, length(xls_files), command
    ))

    system(command)
    
    delete_registry(office_folder, dbg = dbg)
  }
}
