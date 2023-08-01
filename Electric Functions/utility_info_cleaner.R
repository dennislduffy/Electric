utility_info_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration")
  )
  
  report_year <- raw_utility |> 
    filter(row == 6, col == 3) |> 
    pull(numeric)
  
  utility_name <- raw_utility |> 
    filter(row == 9 & col == 3) |> 
    pull(character)
  
  entity_id <- raw_utility |> 
    filter(address == "C5") |> 
    pull(numeric)
  
  utility_info <- tibble(utility_name, report_year, entity_id)
  
  return(utility_info)
  
}
