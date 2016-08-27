download_msrb_spreadsheets <- function() {
  #
  # This file assumes you have run the MSRB_Data_Initialization.R
  # The script sets up certain variables and changes your working directory
  #
  
  
  # Step 0: set working firectory, check libraries 
  
  library("stringr")
  require("stringr")
  library("xlsx")
  require("xlsx")
  
  project_directory <- getwd()
  target_dir <- paste0(getwd(), "/", download_dir)
  setwd(target_dir)
  
  # Step 1:  create list of URLs for and URL ending in ".xls"
  #
  # Samples:
  #  http://msrb.org/msrb1/TRSweb/MarketStats/Monthly-Trade-Summary-April-2016.xls
  #  http://www.msrb.org/msrb1/TRSweb/MarketStats/Monthly%20Trade%20Summary%20September%202006.xls
  #  http://www.msrb.org/msrb1/TRSweb/MarketStats/Monthly%20Trade%20Summary%20December%202000.xls
  
    # Download page and shrink lines to only those referencing an ".xls" 
  download_page <- readLines(MSRB_URL)
  match_pattern <- ".xls"
  download_page <- grep(match_pattern,download_page[1:length(download_page)],value=TRUE)
  
    # Establish list to control subsequent download loop
  assign ("URL_download_control",
          URL_list <- str_sub(download_page, start = str_locate(download_page, '<a href=\"')[,2] +1, end = str_locate(download_page, ".xls")[,2]),
          envir=.GlobalEnv)
  
  assign ("URL_download_control",
          ifelse( str_sub(URL_download_control, 1, 4) == "http", 
                          URL_download_control, 
                          paste0( "http://www.msrb.org", URL_download_control)  ),
          envir=.GlobalEnv )
  
    # create vector of years - some oddball cases need handling
  file_year <- str_sub(URL_download_control, -8, -5)
  
  x <- grep("2004%20",URL_download_control)
  file_year[x] <- "2004"
  
  x <- grep("*04.xls$", URL_download_control)
  file_year[x] <- "2004"
  
  x <- grep("*03.xls$", URL_download_control)
  file_year[x] <- "2003"
  
  x <- grep("*Feb03stats_report.xls$", URL_download_control)
  file_year[x] <- "2003"
  
  x <- grep("*02.xls$", URL_download_control)
  file_year[x] <- "2002"
  
  x <- grep("*20June.xls$", URL_download_control)
  file_year[x] <- "2002"
  
  x <- grep("March%202008%20.xls",URL_download_control)
  file_year[x] <- "2008"
  
  
    # fake vectors to initially populate the data frame
  file_month <- rep("M", length(URL_download_control))
  dest_file <- rep("F", length(URL_download_control))
  trans_month <- rep(as.Date("2000-01-01"), length(URL_download_control))
  
  
    # create data frame with aditional coluimns
    assign ("URL_download_control",
          data.frame(URL_download_control, file_year, file_month, dest_file, trans_month, stringsAsFactors = FALSE),
          envir=.GlobalEnv )
  
  names(URL_download_control) <- c("URL", "file_year", "file_Month", "dest_file", "trans_month")
  
  
  # populate file_months column
  
  target_months <- as.character(target_months <- seq(as.Date("2000-01-01"), as.Date("2000-12-01"), by = "month"),"%b")
  
  for ( i in 1:length(target_months) ) {
    URL_download_control[which(grepl(target_months[i], URL_download_control[,1] )), 3] <- target_months[i]
  }
  
  target_months <- as.character(target_months <- seq(as.Date("2000-01-01"), as.Date("2000-12-01"), by = "month"),"%B")
  
  for ( i in 1:length(target_months) ) {
    URL_download_control[which(grepl(target_months[i], URL_download_control[,1] )), 3] <- target_months[i]
  }
  
  
  # populate dest_file column
  
  URL_download_control[,4] <- paste0(URL_download_control[,2], "-", URL_download_control[,3], ".xls")
  
  
  # populate trans_month column
  
  URL_download_control[,5] <- as.Date(paste0(URL_download_control[,2], "-", URL_download_control[,3], "-01"), "%Y-%b-%d")
  
  
  #add factors
  
  URL_download_control[,2] <- as.factor(URL_download_control[,2])
  URL_download_control[,3] <- as.factor(URL_download_control[,3])
  
  
    # error checking
  x <- length(URL_download_control[,1])
  y <- length(grep("http://www.msrb.org/msrb1", URL_download_control[,1]))
  if ( x != y ) {
    return(" - Failure! not all rows in URL_download_control begin with 'http://www.msrb.org/msrb1'")
  } else { message(" - Success! all rows in URL_download_control begin with 'http://www.msrb.org/msrb1'")
         }
  y <- length(grep(".xls", str_sub(URL_download_control[,1], start = -4, end =  -1)))
    if ( x != y ) {
    return(" - Failure! not all rows in URL_download_control end with '.xls'")
    } else { message(" - Success! all rows in URL_download_control end with '.xls'")
           }
  
  # There are errors on the web page for July 2000 and September 2010 - this section checks that script uses the correct URLs
  
  
  if( length( grep("trs/TRSweb/MarketStats/Monthly%20Trade%20Summary%20July%202000.xls", URL_download_control[,1] ) ) > 0 ) {
    
      x <- grep("trs/TRSweb/MarketStats/Monthly%20Trade%20Summary%20July%202000.xls", URL_download_control[,1] )
      new_url <- "http://www.msrb.org/msrb1/TRSweb/MarketStats/Monthly%20Trade%20Summary%20July%202000.xls"
      URL_download_control[x,1] <- new_url
      message(" - URL for July 2000 updated to reflect error on page")
    
  } else {
    message(paste0(" - No action. July 2000 URL was already corrected"))
  }
  
  
  if( length( grep("20Summary%20September%202010", URL_download_control[,1] ) ) == 0 ) {
      if( length( grep("Monthly-Trade-Summary-October-2010", URL_download_control[,1]) ) > 1 ) {
        x <- tail(grep("Monthly-Trade-Summary-October-2010", URL_download_control[,1]), n = 1)
        new_url <- "http://www.msrb.org/msrb1/TRSweb/MarketStats/Monthly%20Trade%20Summary%20September%202010.xls"
        URL_download_control[x,1] <- new_url
        URL_download_control[x,c(3, 4)] <- c( "September", "2010-September.xls"  )
        message(" - URL for September 2010 updated to reflect error on page")
      }
  } else {
    message(paste0(" - No action. File 2010_September.xls has been fixed or already downloaded"))
  }
  
  
  
  # Step 2: Loop through files that are not already downloaded and download each one 
  
   
    # loop through contol list downloading the files
  x <- 0
  
  
  # for ( i in c(4, 38, 53, 54, 55, 56, 57, 90, 149, 161, 189) ) {   # use this loop while testing
    
  # for ( i in 1:nrow(URL_download_control) ) {  
  for (i in 184:187) {
    download.file(URL_download_control[i,1], destfile = URL_download_control[i,4], method = "curl")
    message(paste0(" - downloaded file at ", URL_download_control[i,1]))
    Sys.sleep(max(round(rnorm(1, mean = 8, sd = 4),1), 1))  # makes sure our download process does not cause a denial-of-service to others
    
    # to make sure the file name matches the data in the file, I am getting the month and year from inside the file itself
    xM_Y <- read.xlsx( paste0(getwd(), "/", URL_download_control[i,4]), 
                       sheetIndex = 1,
                       rowIndex = 2:2, 
                       colIndex = 1:1,  
                       header = FALSE )
    xMonth.Year <- as.character(xM_Y[1,1])
    xMonth.Year <- str_trim(xMonth.Year, side = "both")
    xMonth.Year <- gsub("[[:punct:]]", "", xMonth.Year)
    
    # rename the file to match extracted month and year
    
    from_name <- paste0( getwd(), "/", URL_download_control[i,4] )
    to_name   <- paste0( str_sub(xMonth.Year, -4, -1), "_", str_sub(xMonth.Year, 1, str_locate(xMonth.Year, " ")[1] - 1), ".xls" )
    to_name_path   <- paste0( getwd(), "/", to_name )
    file.rename( from = from_name, to = to_name_path )
    message(paste0(" - file renamed from ", URL_download_control[i,4], " to ", to_name))
                         
    x <- x + 1
  }
  message(paste0(" - ", x, " files downloaded"))
  
 
  assign ("URL_download_control",
          URL_download_control,
          envir=.GlobalEnv        )
  
  setwd(project_directory)
  
  write.csv(URL_download_control, file = paste0(project_name, "_file_list_", as.character(Sys.Date()), ".csv"))
  
  return(" - script download_msrb_spreadsheets.R has completed")
}