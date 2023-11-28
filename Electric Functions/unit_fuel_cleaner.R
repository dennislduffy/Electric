single_plant_cleaner_fuel <- function(directory, workbook, plant){
  
  raw_plant <- xlsx_cells(
    paste(directory, workbook, sep = "/"),
    sheets = c("Registration", plant)
  )
  
  utility_name <- raw_plant |> 
    filter(sheet == "Registration", 
           row == 9 & col == 3) |> 
    pull(character)
  
  utility_id <- raw_plant |> 
    filter(sheet == "Registration", 
           address == "C5") |> 
    pull(content)
  
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
  
  plant_id <- raw_plant |> 
    filter(sheet == plant, 
           address == "E12") |> 
    mutate(var = as.character(content)) |> 
    pull(var)
  
  start_points <- raw_plant |> 
    filter(sheet == plant) |> 
    filter(character == "Unit ID #")
  
  fuel_row_start <- start_points$row[3] + 1
  
  fuel_row_end <- raw_plant |> 
    filter(sheet == plant, 
           character == "ALLOWABLE CODES") |> 
    mutate(row = row - 3) |> 
    pull(row)

    
  fuel_data <- raw_plant |> 
    filter(sheet == plant, 
           row %in% c(fuel_row_start:fuel_row_end) & col %in% c(2:10)) |> 
    mutate(values = case_when(
      data_type == "character"  ~ character, 
      data_type == "numeric" ~ as.character(numeric)
    )) |> 
    select(row, col, values) |> 
    pivot_wider(names_from = "col", values_from = "values") |> 
    select(-c(row)) |> 
    rename(`Unit ID` = `2`, 
           `Primary Fuel Type` = `3`, 
           `Primary Fuel Quantity` = `4`, 
           `Primary Fuel Unit` = `5`, 
           `Primary BTU Content` = `6`, 
           `Secondary Fuel Type` = `7`, 
           `Secondary Fuel Quantity` = `8`, 
           `Secondary Fuel Unit` = `9`, 
           `Secondary BTU Content` = `10`) |> 
    filter(complete.cases(`Unit ID`)) |> 
    mutate(`Plant ID` = plant_id, 
           `Plant Name` = plant_name, 
           `Plant City` = plant_city, 
           `Plant Zip Code` = plant_zip, 
           `Plant County` = plant_county, 
           `Utility` = utility_name,
           `Primary Fuel Quantity` = as.numeric(`Primary Fuel Quantity`), 
           `Primary BTU Content` = as.numeric(`Primary BTU Content`), 
           `Secondary Fuel Quantity` = as.numeric(`Secondary Fuel Quantity`), 
           `Secondary BTU Content` = as.numeric(`Secondary BTU Content`), 
           utility_id = utility_id, 
           year = report_year) |> 
    select(`Utility`, utility_id, `Plant ID`, `Plant Name`, `Plant City`, `Plant Zip Code`, `Plant County`, `Unit ID`, `Primary Fuel Type`, 
           `Primary Fuel Quantity`, `Primary Fuel Unit`, `Primary BTU Content`, `Secondary Fuel Type`, `Secondary Fuel Quantity`, 
           `Secondary Fuel Unit`, `Secondary BTU Content`, year)
  
  return(fuel_data)
  
}

unit_fuel_cleaner <- function(directory, workbook){
  
  #Only takes Plant## sheets. Excludes the BlankPlant sheet
  sheet_names <- getSheetNames(paste(directory, workbook, sep = "/"))
  
  plant_pages <- c()
  
  for(i in sheet_names){
    
    if(grepl("Plant",i) == TRUE && grepl("BlankPlant",i) == FALSE){
      
      plant_pages <- append(plant_pages, i)
    
    }
    
  }
  
  print(plant_pages)
  
  raw_plants <- xlsx_cells(paste(directory, workbook, sep = "/")) |>
    filter(c(sheet %in% plant_pages))
  
  plant_sheets <- raw_plants |>
    select(sheet) |>
    distinct(sheet) |>
    pull(sheet)
  
  if(length(plant_sheets) > 0){
    
    for(i in plant_sheets){
      
      if(i == plant_sheets[1]){
        
        plant <- single_plant_cleaner_fuel(directory, workbook, i) |>
          slice(0)
        
        x <- single_plant_cleaner_fuel(directory, workbook, i)
        
        plant <- plant |>
          bind_rows(x)
        
      }
      
      else{
        
        x <- single_plant_cleaner_fuel(directory, workbook, i)
        
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
