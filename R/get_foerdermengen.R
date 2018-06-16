get_foerdermengen <- function(path) {
  
  `%>%` <- magrittr::`%>%`
  
  q_ww <- readxl::read_xlsx(path, sheet = "WW Q Rhow ", range = "A4:S127") %>% 
  tidyr::gather_(key_col = "Wasserwerk", 
                 value_col = "Foerdermenge_m3", 
                 gather_cols = setdiff(names(.), "Jahr")) %>% 
  dplyr::rename_(year = "Jahr") %>% 
  dplyr::filter_("!is.na(Foerdermenge_m3)")

lookup_werk <- data.frame(Wasserwerk = unique(q_ww$Wasserwerk), 
 werk = stringr::str_to_upper(
          stringr::str_sub(unique(q_ww$Wasserwerk), 1, 3)),
 stringsAsFactors = FALSE)

q_ww %>% 
  left_join(y = lookup_werk, by = "Wasserwerk")

}
