elec_purchase_sales_cleaner <- function(directory, workbook){
  
  
  raw_utility <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", "PurchaseSales")
  )
  
  report_year <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 6, col == 3) |> 
    pull(numeric)
  
  utility_name <- raw_utility |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  purchase_sales <- raw_utility |> 
    filter(sheet == "PurchaseSales", 
           row >= 13 & col %in% 1:4) |> 
    mutate(value = case_when(
      data_type == "character" ~ character, 
      TRUE ~ content
    )) |> 
    select(row, col, value) |> 
    pivot_wider(names_from = "col", values_from = "value") |> 
    select(-c(row)) |> 
    row_to_names(row_number = 1) |>
    rename(`Transaction Utility` = 1, 
           `Interconnected Utility` = 2, 
           `MWH Purchased` = 3, 
           `MWH Sold for Resale` = 4) |> 
    filter(complete.cases(`Transaction Utility`)) |> 
    mutate(`MWH Purchased` = as.numeric(`MWH Purchased`), 
           `MWH Sold for Resale` = as.numeric(`MWH Sold for Resale`)) |> 
    mutate(`Report Year` = report_year, 
           Utility = utility_name) |> 
    relocate(Utility, .before = `Transaction Utility`) |> 
    relocate(`Report Year`, .before = `Transaction Utility`)
  
  return(purchase_sales)
  
}