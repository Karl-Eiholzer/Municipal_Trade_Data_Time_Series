extract_sec_type_volume <- function(debug_it = FALSE) {
  #
  # extract_sec_type_volume.R
  #
  # This is the third in a series of scripts designed to access municipal securities market data published as a series of  
  # monthly spreadsheets by the Municipal Securities Rulemaking Board (MSRB). The first script downloaded the files, the second 
  # script extracted the new issue volume data.
  #
  # The goal of this script is to extract the cells from the "Trades by Sec Type Data" tab on each spreadsheet and write 
  # to a single continuous dataframe which can then be exported to CSV file ("sec_type_volume.csv") that can be easily loaded for future use.
  #
  # This script assumes the files are already downloaded to a subfolder previously specified as 
  # the download directory.
  #
  # This script also asumes all files are present and all files need to be extracted.
  #
  # If variable "debug_it is true then when the script gets to the loop function it 
  # will only attempt to import a sample of 10 spreadsheets.
  #
  #
  # Author: Karl Eiholzer
  # Github: https://github.com/Karl-Eiholzer/Municipal_Trade_Data_Time_Series
  # Created: 11 July 2016
  # Updated: 
  #  
  
  # Step 0: set directory and other variables as needed
  
  library("xlsx")
  library("stringr")
  require("xlsx")
  require("stringr")
  
  project_directory <- getwd()
  target_dir <- paste0(getwd(), "/", download_dir)
  setwd(target_dir)
  
  col.names <- c("Month","Year","Start.Date","Report.Type", 
                 "Security.Type", "Number.of.CUSIPs", "PCT.of.CUSIPs","Trade.Count","PCT.of.Trade.Count", 
                 "Par.Traded",  "PCT.of.Par.Traded" )
  
  # get list of files available for import and save to control file
  
  assign ("file_names",
          dir(),
          envir=.GlobalEnv )
  
  # if we are in debug mode, choose a random sample of 10% of files or 10 files to import, whichever is more
  
  if (debug_it == FALSE) { 
    loop.count <- 1:length(file_names)
  } else {
    x <- 1:length(file_names) 
    n <- max( 10,  round(length(file_names)/10, 0) )
    loop.count <- sample(x,n) 
    message(paste0(" - Operating in debug mode. Only ", length(loop.count), " files chosen at random will be imported"))
  }
  
  
  # Step 1:   Cycle through each type of trade numbers by grabbing numbers, adding factor columns, and adding to append table
  #
  #  first  XLSX data ---> "table_raw" df,    then "table_raw" df + factor variables (dates, trade types) ---> "table_append" 
  
  
  for ( i in loop.count ) {
  
    # 1a: preliminary step: get the month and year from the file - sometimes it has parenthesis characters that have to be removed
    #     month an year stiored as variable "Month.Year" 
    
    assign ("M_Y",
            read.xlsx( paste0(getwd(), "/", file_names[i]), 
                       sheetIndex = 5,
                       rowIndex = 2:2, 
                       colIndex = 1:1,  
                       header = FALSE ),
            envir=.GlobalEnv                           )
    assign("Month.Year",
           as.character(M_Y[1,1]),
           envir=.GlobalEnv                            )
    
    assign("Month.Year",
           str_trim(Month.Year, side = "both"),
           envir=.GlobalEnv                            )
    
    assign("Month.Year",
           gsub("[[:punct:]]", "", Month.Year),
           envir=.GlobalEnv                            )
    
    
    #  1b Pull trade volume data one section at a time
    
      #     1b(1)(A): raw total trading volume (rowIndex = 5:11)
    
    table_raw <-
      read.xlsx( paste0(getwd(), "/", file_names[i]), 
                 sheetIndex = 5,
                 rowIndex = 5:11,
                 colIndex = 1:7,
                 as.data.frame = TRUE,
                 header = FALSE )
    
    
      #     1b(1)(B) append month year to data frame with data
      #     NOTE: this step resets the sec_type_volume_working_table_append df, subsequent steps only append to it
    
    sec_type_volume_working_table_append <-
      cbind( str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1), 
             str_sub(Month.Year, -4, -1), 
             as.POSIXct( paste0( "01-",
                                 str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1),
                                 "-",
                                 str_sub(Month.Year, -4, -1)),
                         format = "%d-%B-%Y" ),
             "Total Customer and Interdealer",
             table_raw)
    names(sec_type_volume_working_table_append) <- col.names
    
    
      #    1b(2)(A): interdealer only (rowIndex = 16:22)
    
    table_raw <-
      read.xlsx( paste0(getwd(), "/", file_names[i]), 
                 sheetIndex = 5,
                 rowIndex = 16:22,
                 colIndex = 1:7,
                 as.data.frame = TRUE,
                 colClasses=NA,
                 header = FALSE )
    
    
    #     1b(2)(B) append month year to data frame with data
    
    table_raw <-
      cbind( str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1), 
             str_sub(Month.Year, -4, -1), 
             as.POSIXct( paste0( "01-",
                                 str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1),
                                 "-",
                                 str_sub(Month.Year, -4, -1)),
                         format = "%d-%B-%Y" ),
             "Interdealer Only",
             table_raw)
    names(table_raw) <- col.names
    
    sec_type_volume_working_table_append <- rbind(sec_type_volume_working_table_append, table_raw)
    
    #     1b(3)(A): customer only - buy and sell (rowIndex = 27:33)
    
    table_raw <-
      read.xlsx( paste0(getwd(), "/", file_names[i]), 
                 sheetIndex = 5,
                 rowIndex = 27:33,
                 colIndex = 1:7,
                 as.data.frame = TRUE,
                 header = FALSE )
    
    
    #     1b(3)(B) append month year to data frame with data
    
    table_raw <-
      cbind( str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1), 
             str_sub(Month.Year, -4, -1), 
             as.POSIXct( paste0( "01-",
                                 str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1),
                                 "-",
                                 str_sub(Month.Year, -4, -1)),
                         format = "%d-%B-%Y" ),
             "Customer Only Buy and Sell",
             table_raw)
    names(table_raw) <- col.names
    
    sec_type_volume_working_table_append <- rbind(sec_type_volume_working_table_append, table_raw)
    
    
    #    1b(4)(A): purchase by customer only (rowIndex = 38:44)
    
    table_raw <-
      read.xlsx( paste0(getwd(), "/", file_names[i]), 
                 sheetIndex = 5,
                 rowIndex = 38:44,
                 colIndex = 1:7,
                 as.data.frame = TRUE,
                 header = FALSE )
    
    
    #     1b(4)(B) append month year to data frame with data
    
    table_raw <-
      cbind( str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1), 
             str_sub(Month.Year, -4, -1), 
             as.POSIXct( paste0( "01-",
                                 str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1),
                                 "-",
                                 str_sub(Month.Year, -4, -1)),
                         format = "%d-%B-%Y" ),
             "Customer Bought Only",
             table_raw)
    names(table_raw) <- col.names
    
    sec_type_volume_working_table_append <- rbind(sec_type_volume_working_table_append, table_raw)
    
    
    #      1b(5)(A): customer only - customer sold only (rowIndex = 49:55)
    
    table_raw <-
      read.xlsx( paste0(getwd(), "/", file_names[i]), 
                 sheetIndex = 5,
                 rowIndex = 49:55,
                 colIndex = 1:7,
                 as.data.frame = TRUE,
                 header = FALSE )
    
    
    #     1b(5)(B) append month year to data frame with data
    
    table_raw <-
      cbind( str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1), 
             str_sub(Month.Year, -4, -1), 
             as.POSIXct( paste0( "01-",
                                 str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1),
                                 "-",
                                 str_sub(Month.Year, -4, -1)),
                         format = "%d-%B-%Y" ),
             "Customer Sold Only",
             table_raw)
    names(table_raw) <- col.names
    
    sec_type_volume_working_table_append <- rbind(sec_type_volume_working_table_append, table_raw)
    
    
    # 1c: attach the append table to the main table

    if(exists("sec_type_volume_working_table")){
      assign("sec_type_volume_working_table",
             rbind( sec_type_volume_working_table, 
                    sec_type_volume_working_table_append ),
             envir=.GlobalEnv                               )
    } else {
      assign("sec_type_volume_working_table",
             sec_type_volume_working_table_append,
             envir=.GlobalEnv                               )
    }
        
    message(paste0(" - added ", Month.Year, " to data"))  
  
    # END LOOP
    
  }
  
  
  # Step 2: Add average trade size column data
  #         This data is in the raw files, but sometimes has NaN values which caused complications for reading the XLSX
  
  Avg.Trade.Size <- sec_type_volume_working_table$Par.Traded / sec_type_volume_working_table$Trade.Count
  assign("sec_type_volume_working_table",
         cbind( sec_type_volume_working_table, 
                Avg.Trade.Size  ),
         envir=.GlobalEnv                       )
  
  
  # Step 3: order the data by date and other factors
  
  order_col <- c( "Start.Date","Report.Type","Security.Type" )
  setorder(sec_type_volume_working_table,
           "Start.Date","Report.Type","Security.Type"  )
  
  # Step 4: Completion. Export Results to CSV
 
  
      # 4a Export to CSV to project directory (instead of to the subfolder which has the downloaded files)
  
  setwd(project_directory)
  write.csv(sec_type_volume_working_table, file = "sec_type_volume.csv")
  
      # 4b Return
  
  if (debug_it == FALSE) { 
    return(" - extract_sec_type_volume is complete")
  } else {
    x <- 1:length(file_names) 
    loop.count <- sample(x,10) 
    return(" - Operating in debug mode. Only a random sample of files were imported")
  }
  
  
  
}