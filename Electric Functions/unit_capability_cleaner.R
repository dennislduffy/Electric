single_plant_cleaner_unit <- function(directory, workbook, plant){
  
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
  
  cap_row_start <- start_points$row[2] + 1
  
  cap_row_end <- raw_plant |> 
    filter(sheet == plant) |> 
    filter(character == "D. UNIT FUEL USED") |> 
    mutate(row = row - 2) |> 
    pull(row)

  capability_data <- raw_plant |> 
    filter(sheet == plant, 
           row %in% c(cap_row_start:cap_row_end) & col %in% c(2:8)) |> 
    mutate(values = case_when(
      data_type == "character"  ~ character, 
      data_type == "numeric" ~ as.character(numeric)
    )) |> 
    select(row, col, values) |> 
    pivot_wider(names_from = "col", values_from = "values") |> 
    select(-c(row)) |> 
    rename(`Unit ID` = `2`, 
           `Summer Capacity MW` = `3`, 
           `Winter Capacity MW` = `4`, 
           `Capacity Factor (Percent)` = `5`, 
           `Operating Factor (Percent)` = `6`, 
           `Forced Outage Rate (Percent)` = `7`,
           `Capacity Comments` = `8`) |> 
    filter(complete.cases(`Unit ID`)) |> 
    mutate(Utility = utility_name, 
           `Plant ID` = plant_id, 
           `Plant Name` = plant_name, 
           `Plant City` = plant_city, 
           `Plant Zip Code` = plant_zip, 
           `Plant County` = plant_county, 
           `Summer Capacity MW` = as.numeric(`Summer Capacity MW`), 
           `Winter Capacity MW` = as.numeric(`Winter Capacity MW`), 
           `Capacity Factor (Percent)` = as.numeric(`Capacity Factor (Percent)`), 
           `Operating Factor (Percent)` = as.numeric(`Operating Factor (Percent)`), 
           `Forced Outage Rate (Percent)` = as.numeric(`Forced Outage Rate (Percent)`), 
           utility_id = utility_id, 
           year = report_year) |> 
    select(Utility, utility_id, `Plant ID`, `Plant Name`, `Plant City`, `Plant Zip Code`, `Plant County`, `Unit ID`, `Summer Capacity MW`, 
           `Winter Capacity MW`, `Capacity Factor (Percent)`, `Operating Factor (Percent)`, `Forced Outage Rate (Percent)`, 
           `Capacity Comments`, year)
 
  
  
  return(capability_data)
  
}

unit_capability_cleaner <- function(directory, workbook){
  
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
        
        plant <- single_plant_cleaner_unit(directory, workbook, i) |>
          slice(0)
        
        x <- single_plant_cleaner_unit(directory, workbook, i)
        
        plant <- plant |>
          bind_rows(x)
        
      }
      
      else{
        
        x <- single_plant_cleaner_unit(directory, workbook, i)
        
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