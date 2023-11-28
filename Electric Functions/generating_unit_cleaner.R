single_plant_cleaner_generator <- function(directory, workbook, plant){
  
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
  
  gen_row_start <- start_points$row[1] + 1
  
  gen_row_end <- raw_plant |> 
    filter(sheet == plant, 
           character == "C. UNIT CAPABILITY DATA") |> 
    mutate(row = row - 2) |> 
    pull(row)
  
  row_count <- gen_row_end - gen_row_start
  
  generating_data <- raw_plant |> 
    filter(sheet == plant, 
           row %in% c(gen_row_start:(gen_row_start + row_count)) & col %in% c(2:8)) |> 
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
    select(-c(row)) |> 
    mutate(`Plant ID` = plant_id, 
           `Plant Name` = plant_name, 
           `Plant City` = plant_city, 
           `Plant Zip Code` = plant_zip, 
           `Plant County` = plant_county, 
           `Utility` = utility_name, 
           utility_id = utility_id,
           `Net Generation MWH` = as.numeric(`Net Generation MWH`), 
           year = report_year) |> 
    rename(`Unit ID` = ID) |> 
    select(`Utility`, utility_id, `Plant ID`, `Plant Name`, `Plant City`, `Plant Zip Code`, `Plant County`, `Unit ID`, `Unit Status`, `Unit Type`, `Year Installed`, `Energy Source`, 
           `Net Generation MWH`, `Comments`, year)
  
  return(generating_data)
  
}


generating_unit_cleaner <- function(directory, workbook){
  
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
        
        plant <- single_plant_cleaner_generator(directory, workbook, i) |>
          slice(0)
        
        x <- single_plant_cleaner_generator(directory, workbook, i)
        
        plant <- plant |>
          bind_rows(x)
        
      }
      
      else{
        
        x <- single_plant_cleaner_generator(directory, workbook, i)
        
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