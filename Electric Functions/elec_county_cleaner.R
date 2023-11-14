
###
elec_county_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "ElectricityByCounty")
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
  
  county_name <- raw_utility |> 
    filter(sheet == "ElectricityByCounty", 
           row %in% c(12:56) & col == 2 | row %in% c(12:53) & col == 6) |>
    arrange(col, row) |> 
    select(character) |> 
    rename(`County Name` = "character") |> 
    mutate(`County Name` = case_when(
      `County Name` == "Olmstead" ~ "Olmsted", 
      TRUE ~ `County Name`
    ))
  
  mwh_delivered <- raw_utility |> 
    filter(sheet == "ElectricityByCounty", 
           row %in% c(12:56) & col == 3 | row %in% c(12:53) & col == 7) |>
    arrange(col, row) |> 
    select(numeric) |> 
    rename(`MWH Delivered` = "numeric")
  
  elec_by_county <- raw_utility |> 
    filter(sheet == "ElectricityByCounty",
           row %in% c(12:56) & col == 1 | row %in% c(12:53) & col == 5) |> 
    arrange(col, row) |> 
    select(numeric) |> 
    rename(`County Code` = numeric) |> 
    bind_cols(county_name) |> 
    bind_cols(mwh_delivered) |> 
    mutate(`Report Year` = report_year, 
           Utility = utility_name, 
           utility_id = utility_id) |> 
    filter(complete.cases(`MWH Delivered`))
  
  return(elec_by_county)
  
}
