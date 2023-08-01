
single_plant_cleaner <- function(directory, workbook, plant){
  
  raw_plant <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", plant)
  )
  
  utility_name <- raw_plant |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  report_year <- raw_plant |> 
    filter(sheet == "Registration", 
           row == 6, col == 3) |> 
    pull(numeric)
  
  plant_name <- raw_plant |> 
    filter(sheet == plant, 
           address == "B12") |> 
    pull(character)
  
  plant_city <- raw_plant |> 
    filter(sheet == plant, 
           address == "B14") |> 
    pull(character)
  
  plant_zip <- raw_plant |> 
    filter(sheet == plant, 
           address == "B16") |> 
    pull(numeric)
  
  plant_county <- raw_plant |> 
    filter(sheet == plant, 
           address == "B17") |> 
    pull(character)
  
  generating_data <- raw_plant |> 
    filter(sheet == plant, 
           row %in% c(23:33) & col %in% c(2:8)) |> 
    mutate(values = case_when(
      data_type == "character"  ~ character, 
      data_type == "numeric" ~ as.character(numeric)
    )) |> 
    select(row, col, values) |> 
    pivot_wider(names_from = "col", values_from = "values") |> 
    rename(`ID` = `2`, 
           `Unit Status` = `3`, 
           `Unit Type` = `4`, 
           `Year Installed` = `5`, 
           `Energy Source` = `6`, 
           `Net Generation MWH` = `7`, 
           Comments = `8`) |> 
    filter(complete.cases(`ID`)) |> 
    select(-c(row))
  
  capability_data <- raw_plant |> 
    filter(sheet == plant, 
           row %in% c(37:47) & col %in% c(2:8)) |> 
    mutate(values = case_when(
      data_type == "character"  ~ character, 
      data_type == "numeric" ~ as.character(numeric)
    )) |> 
    select(row, col, values) |> 
    pivot_wider(names_from = "col", values_from = "values") |> 
    select(-c(row)) |> 
    rename(`ID` = `2`, 
           `Summer Capacity MW` = `3`, 
           `Winter Capacity MW` = `4`, 
           `Capacity Factor (Percent)` = `5`, 
           `Operating Factor (Percent)` = `6`, 
           `Forced Outage Rate (Percent)` = `7`,
           `Capacity Comments` = `8`) |> 
    filter(complete.cases(`ID`))
  
  fuel_data <- raw_plant |> 
    filter(sheet == plant, 
           row %in% c(51:61) & col %in% c(2:10)) |> 
    mutate(values = case_when(
      data_type == "character"  ~ character, 
      data_type == "numeric" ~ as.character(numeric)
    )) |> 
    select(row, col, values) |> 
    pivot_wider(names_from = "col", values_from = "values") |> 
    select(-c(row)) |> 
    rename(`ID` = `2`, 
           `Primary Fuel Type` = `3`, 
           `Primary Fuel Quantity` = `4`, 
           `Primary Fuel Unit` = `5`, 
           `Primary BTU Content` = `6`, 
           `Secondary Fuel Type` = `7`, 
           `Secondary Fuel Quantity` = `8`, 
           `Secondary Fuel Unit` = `9`, 
           `Secondary BTU Content` = `10`) |> 
    filter(complete.cases(`ID`))
  
  plant_data <- generating_data |> 
    left_join(capability_data, by = "ID") |> 
    left_join(fuel_data, by = "ID") |> 
    mutate(Utility = utility_name, 
           Year = report_year) |> 
    relocate(Utility, .before = "ID") |> 
    relocate(Year, .before = "ID") |> 
    mutate(`Year Installed` = as.numeric(`Year Installed`), 
           `Net Generation MWH` = as.numeric(`Net Generation MWH`), 
           `Summer Capacity MW` = as.numeric(`Summer Capacity MW`), 
           `Winter Capacity MW` = as.numeric(`Winter Capacity MW`), 
           `Capacity Factor (Percent)` = as.numeric(`Capacity Factor (Percent)`),
           `Operating Factor (Percent)` = as.numeric(`Operating Factor (Percent)`), 
           `Forced Outage Rate (Percent)` = as.numeric(`Forced Outage Rate (Percent)`), 
           `Primary Fuel Quantity` = as.numeric(`Primary Fuel Quantity`), 
           `Primary BTU Content` = as.numeric(`Primary BTU Content`), 
           `Secondary Fuel Quantity` = as.numeric(`Secondary Fuel Quantity`), 
           `Secondary BTU Content` = as.numeric(`Secondary BTU Content`))
  
  return(plant_data)
  
}

all_plants_cleaner <- function(directory, workbook){
  
  pages_toss <- c("Attachments", "BlankPlant", "ElectricityByClass", "ElectricityByCounty", "ElectricityByMonth", "PurchaseSales", 
                  "Registration", "SpaceHeating", "Substitutions", "EBM-By RevenueClass")
  
  #DD Note: one sheet has a EBM-By RevenueClass that isn't in any other sheet. Tossing it.
  
  raw_plants <- xlsx_cells(paste(directory, workbook, sep = "/")) |>
    filter(!c(sheet %in% pages_toss))
  

  plant_sheets <- raw_plants |>
    select(sheet) |>
    distinct(sheet) |>
    pull(sheet)

  if(length(plant_sheets) > 0){
    
    for(i in plant_sheets){
      
      if(i == plant_sheets[1]){
        
        plant <- single_plant_cleaner(directory, workbook, i) |>
          slice(0)
        
        x <- single_plant_cleaner(directory, workbook, i)
        
        plant <- plant |>
          bind_rows(x)
        
      }
      
      else{
        
        x <- single_plant_cleaner(directory, workbook, i)
        
        plant <- plant |>
          bind_rows(x)
        
      }
      
    }
    
    return(plant)
    
  }
  
  else{
    
    return(NULL)
    
  }

}