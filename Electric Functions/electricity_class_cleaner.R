#required libraries - used for testing

# library(tidyverse)
# library(tidyxl)


#define electric_class_cleaner function 

electricity_class_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "ElectricityByClass")
  )
  
  report_year <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 6, col == 3) |> 
    pull(numeric)
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  elec_class <- raw_utility |> 
    filter(sheet == "ElectricityByClass", 
           row %in% c(10:16) & col %in% 1:4) |> 
    mutate(content = case_when(
      data_type == "character" ~ character, 
      data_type == "numeric" ~ as.character(numeric)
    )) |> 
    select(row, col, content) |> 
    pivot_wider(names_from = "col", values_from = "content") |> 
    select(-c(row)) |> 
    rename(`Classification of Energy Delivered to Ultimate Consumers` = `1`, 
           `Number of Customers at End of Year` = `2`, 
           `Megawatt Hours` = `3`, 
           `Revenue` = `4`) |> 
    mutate(Utility = utility_name, 
           Year = report_year, 
           `Number of Customers at End of Year` = as.numeric(`Number of Customers at End of Year`), 
           `Megawatt Hours` = as.numeric(`Megawatt Hours`), 
           `Revenue` = as.numeric(Revenue)) |> 
    relocate(Utility, .before = `Classification of Energy Delivered to Ultimate Consumers`) |> 
    relocate(Year, .before = `Classification of Energy Delivered to Ultimate Consumers`)
  
  return(elec_class)
  
}