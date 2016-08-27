extract_new_issue_volume <- function() {
  #
  # extract_new_issue_volume.R
  #
  # This is the second in a series of scripts designed to access municipal securities market data published as a series of  
  # monthly spreadsheets by the Municipal Securities Rulemaking Board (MSRB). The first script downloaded the files.
  #
  # The goal of this script is to extract the cells from the "various"New Issue Data" tab on each spreadsheet and write 
  # to a single continuous dataframe which can then be exported to CSV file ("new_issue_volume.csv") that can be easily loaded for future use.
  #
  # This script assumes the files are already downloaded to a subfolder previously specified as 
  # the download directory.
  #
  # This script also asumes all files are present and all files need to be extracted.
  #
  #
  # Author: Karl Eiholzer
  # Github: https://github.com/Karl-Eiholzer/Municipal_Trade_Data_Time_Series
  # Created: 01-July 2016
  # Updated: 11 July 2016 - Minor formatting
  #  
  
  # Step 0: set directory
  
  library("xlsx")
  require("xlsx")
  
  project_directory <- getwd()
  target_dir <- paste0(getwd(), "/", download_dir)
  setwd(target_dir)
  
  # Step 1: get list of files available for import and save to control file
  #         Used to manage the loop function
  
  assign ("file_names",
          dir(),
          envir=.GlobalEnv )
  
  
  # Step 2:  Load first sheet to establish data frame
  #          After this data frame is created, all additional extractions will be appended to this one. 
  
    # 2a: get the data
  new_issue_volume_working_table <-
         read.xlsx( paste0(getwd(), "/", file_names[1]), 
                    sheetIndex = 4,
                    rowIndex = 5:7,
                    colIndex = 1:5,
                    as.data.frame = TRUE,
                    header = TRUE )
 
  
   # 2b get the month and year from the file - sometimes it has parenthesis characters that have to be removed
  assign ("M_Y",
          read.xlsx( paste0(getwd(), "/", file_names[1]), 
                     sheetIndex = 4,
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
  
  
    #2c append month year to data frame with data
  
  new_issue_volume_working_table <-
        cbind( str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1), 
                str_sub(Month.Year, -4, -1), 
                as.POSIXct( paste0( "01-",
                                    str_sub(Month.Year, 1, str_locate(Month.Year, " ")[1] - 1),
                                    "-",
                                    str_sub(Month.Year, -4, -1)),
                            format = "%d-%B-%Y" ),
                new_issue_volume_working_table)
  names(new_issue_volume_working_table)[1:3] <- c("Month","Year","Start.Date")
   
  
  # Step 3:  loop through remaining files adding the data to the data frame
  
  for ( i in 2:length(file_names) ) {
    
        # 3a: get the new issue data from Tab 4:
    
    x <-
      read.xlsx( paste0(getwd(), "/", file_names[i]), 
                 sheetIndex = 4,
                 rowIndex = 6:7,
                 colIndex = 1:5,
                 as.data.frame = TRUE,
                 header = FALSE )
    
        # 3b get the month and year from the file - sometimes it has parenthesis characters that have to be removed
    
    xM_Y <- read.xlsx( paste0(getwd(), "/", file_names[i]), 
                       sheetIndex = 4,
                       rowIndex = 2:2, 
                       colIndex = 1:1,  
                       header = FALSE )
    xMonth.Year <- as.character(xM_Y[1,1])
    xMonth.Year <- str_trim(xMonth.Year, side = "both")
    xMonth.Year <- gsub("[[:punct:]]", "", xMonth.Year)
    
        # 3c append month year to data frame with data
    
    x <-
      cbind( str_sub(xMonth.Year, 1, str_locate(xMonth.Year, " ")[1] - 1), 
             str_sub(xMonth.Year, -4, -1), 
             as.POSIXct( paste0( "01-",
                                 str_sub(xMonth.Year, 1, str_locate(xMonth.Year, " ")[1] - 1),
                                 "-",
                                 str_sub(xMonth.Year, -4, -1)),
                         format = "%d-%B-%Y" ),
             x)
    names(x) <- names(new_issue_volume_working_table)
    new_issue_volume_working_table <- rbind(new_issue_volume_working_table, x)
    message(paste0(" - added ", xMonth.Year, " to data"))
  }
  
  
  # Step 4: Completion. Create working table in global environment for DQ and Export Results to CSV
  
      # 4a create table in global environment for data quality review
  
  assign("new_issue_volume_working_table",
         new_issue_volume_working_table,
         envir=.GlobalEnv                       )
  
      # 4b Export to CSV to project directory (instead of to the subfolder which has the downloaded files)
  
  setwd(project_directory)
  write.csv(new_issue_volume_working_table, file = "new_issue_volume.csv")
  
      # 4c Return
  
  return(" - extract_new_issue_volume is complete")
  
}