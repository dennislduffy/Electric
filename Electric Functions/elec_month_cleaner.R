elec_month_cleaner <- function(directory, workbook){
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "ElectricityByMonth")
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
  
  month_table <- raw_utility |> 
    filter(sheet == "ElectricityByMonth", 
           row %% 2 == 0 & row %in% c(10:32) & col == 1) |> 
    select(row, character) |> 
    rename(Month = character)
  
  customer_types <- raw_utility |> 
    filter(sheet == "ElectricityByMonth", 
           row ==9 & col %in% c(1, 3:11)) |> 
    pull(character) |> 
    str_replace_all("\\r\\n", " ") |> 
    str_replace(" \\(Columns A through H\\)", "") 
  
  customer_types[1] <- "row" # doing this manually because one utility still has 2018 as the past year in their workbook
  
  customer_numbers <- raw_utility |> 
    filter(sheet == "ElectricityByMonth", 
           row %in% c(10:33) & col %in% c(3:11), 
           row %% 2 == 0) |> 
    select(row, col, numeric) |> 
    pivot_wider(names_from = "col", values_from = "numeric")
  
  colnames(customer_numbers) <- customer_types
  
  customer_numbers <- customer_numbers |> 
    pivot_longer(!row, names_to = "customer_type", values_to = "customer_count")
  
  mwh_numbers <- raw_utility |> 
    filter(sheet == "ElectricityByMonth", 
           row %in% c(10:33) & col %in% c(3:11), 
           row %% 2 != 0) |> 
    select(row, col, numeric) |> 
    pivot_wider(names_from = "col", values_from = "numeric")
  
  colnames(mwh_numbers) <- customer_types
  
  mwh_numbers <- mwh_numbers |> 
    pivot_longer(!row, names_to = "customer_type", values_to = "MWH") |> 
    mutate(row = row - 1)
  
  elec_month <- month_table |> 
    left_join(customer_numbers, by = "row") |> 
    left_join(mwh_numbers, by = c("row" = "row", "customer_type" = "customer_type")) |> 
    select(-c(row)) |> 
    mutate(Utility = utility_name, 
           `Report Year` = report_year, 
           utility_id = utility_id) |> 
    relocate(Utility, .before = "customer_type") |> 
    relocate(utility_id, .before = "customer_type") |> 
    relocate(`Report Year`, .before = "customer_type") |> 
    rename(`Customer Type` = "customer_type", 
           `Customer Count` = "customer_count")
  
  return(elec_month)
  
}