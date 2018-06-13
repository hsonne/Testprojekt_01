add_para_metadata <- function(df, 
                              lookup_para_path,
                              parameters_path) {
  
  `%>%` <- magrittr::`%>%`
  
  if(file.exists(lookup_para_path)) {
    lookup_para <- read.csv(file = lookup_para_path) %>% 
      dplyr::filter_("!is.na(para_id)")
    
    if(length(lookup_para) > 0) {
      parameters <- readxl::read_excel(path = parameters_path, 
                                       sheet = "nur Parameterliste") %>% 
        janitor::clean_names()
      
      lookup_para <- lookup_para %>%  
        dplyr::left_join(y = parameters)
      
      analysed_paras <- unique(lookup_para$para_kurzname)
      cat(crayon::green(
        crayon::bold(
          sprintf("Successfully read parameter lookup table:\n'%s'\n and joined it with:\n
                  '%s'.\n\nIn total the following %d parameters can be analysed:\n%s\n\n", 
                  lookup_para_path,
                  parameters_path,
                  length(analysed_paras),
                  paste(analysed_paras, collapse = ", ")
          ))))
      
      labor_sel <- df %>% 
        dplyr::left_join(y = lookup_para, by = "VariableName_org") %>% 
        dplyr::filter_("!is.na(para_id)")
      
    } else {
      stop(sprintf("No parameters defined in %s", lookup_para_path)) 
    }} else {
      lookup_para_template <- data.frame(VariableName_org = sort(as.character(unique(df[, "VariableName_org"]))),
                                         para_id = NA, 
                                         stringsAsFactors = FALSE)
      
      lookup_para_export_path <- file.path(dirname(paths$lookup_para), 
                                           "lookup_para_tmp.csv")
      
      write.csv(lookup_para_template, 
                file = lookup_para_export_path, 
                row.names = FALSE)
      
      stop_msg <- cat(crayon::red(
        crayon::bold(sprintf("Template parameter lookup table created at:\n%s\n
                             Please fill with column 'para_id' with 'Para Id' from file:\n%s\n
                             Do not use EXCEL, just a text editor like Notepad++.\n
                             Afterwords rename the file to 'lookup_para.csv' and save it in the same directory!", 
                             normalizePath(lookup_para_export_path), 
                             normalizePath(parameters_path)))))
      
      stop(stop_msg)
      
    }
  return(labor_sel)
}

add_site_metadata <- function(df, 
                              site_path) {
  
  `%>%` <- magrittr::`%>%`
  
  sites <- readxl::read_excel(path = site_path) %>% 
    janitor::clean_names() %>% 
    dplyr::rename(site_id = interne_nr)  
  
  df %>%  
    dplyr::left_join(y = sites, by = "site_id")
  
}