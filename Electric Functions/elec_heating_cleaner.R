elec_heating_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "SpaceHeating")
  )
  
  report_year <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 6, col == 3) |> 
    pull(numeric)
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  utility_id <- raw_utility |> 
    filter(sheet == "Registration", 
           address == "C5") |> 
    pull(content)
  
  space_heating <- raw_utility |> 
    filter(sheet == "SpaceHeating", 
           row == 26 & col %in% c(1:3)) |> 
    select(numeric) |> 
    mutate(numeric = round(numeric)) |> 
    mutate(column_names = c("Number of Residential Electrical Space Heating Customers", 
                            "Number of Residential Units Served with Electrical Space Heating", 
                            "Total MWH")) |> 
    mutate(Utility = utility_name, 
           utility_id = utility_id) |> 
    pivot_wider(names_from = "column_names", values_from = "numeric") |> 
    mutate(Year = report_year) |> 
    relocate(Year, .after = utility_id)
  
  return(space_heating)
  
}