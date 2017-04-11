GetProdCountData <- function(path="D:/DATA/PUR/PLANT",back.one.month=FALSE) {
  # This function descends a directory tree and processes the production report (excel) file 
  # for one month in one year. The current date determines the year/ month. 
  # on a month boundary, the previous month is processed for 7 days to capture changes
  # path.p1 = The top level directory is "D:/DATA/PUR/PLANT", a general store for all types of production files
  # path.p2 = The next level down is "./Production Reports yyyy", one folder per year. is.numeric(year)
  # path.p3 = The next level down is "./mmmm yyyy Production Reports", one folder per month. 
  #   Each folder has one file. is.character(month)
  # path.p4 = The actual files e.g "mmmm yyyy Production Reports.xlsx", 
  #   an excel file for the month with one sheet per day
  # a fully qualified file name example (p1/p2/p3/p4) is :- 
  # D:/DATA/PUR/PLANT/Production Reports 2016/January 2016 Production Reports/January 2016 Production Reports.xlsx
  # 
  #
  if (!require(rJava)) install.packages('rJava')
  library(rJava)
  
  # The XLConnect library is required for the scrip to work
  # options(java.parameters = "-Xmx6g" )
  if (!require(XLConnect)) install.packages('XLConnect')
  library(XLConnect)
  
  # The plyr library is required for the scrip to work
  if (!require(plyr)) install.packages('plyr')
  library(plyr)
  
  # set the working directory
  setwd(path)
  p1 <- getwd()
  # get list of sub-directories
  dirnames <- list.dirs(recursive = FALSE, full.names = FALSE)
  
  # deduce the year and month to process,
  # on a month boundary process both months for 7 calendar days
  # new logic, always process 2 months of data
  mydate = Sys.Date()
  mydate.year <- character(0)
  mydate.cMonth <- character(0)
  mydate.year[1] = format(mydate,"%Y")
  mydate.cMonth[1] = format(mydate,"%B")
  mydate.day = as.numeric(format(mydate,"%d")) 
#  if (back.one.month) { mydate <-mydate - as.numeric(format(mydate,"%d")) 
#  if (mydate.day < 8) { 
    mydate <-mydate - as.numeric(format(mydate,"%d")) 
    mydate.year[2] = format(mydate,"%Y")
    mydate.cMonth[2] = format(mydate,"%B") 
#  }
  
  # filter for the production report sub-directory, for the required year,
  # use the second year if on a year boundary
  # also assume the current directory if no sub dir if found
  pattern <- paste0("^production reports ",mydate.year[1],"$")
  dirnames.years <- dirnames[grepl(pattern,dirnames,ignore.case = TRUE)]
  if (length(dirnames.years) == 0) { dirnames.years[1] <- "." }
  if (length(mydate.year) == 2) {
    pattern <- paste0("^production reports ",mydate.year[2],"$")
    dirnames.years <- rbind(dirnames.years, dirnames[grepl(pattern,dirnames,ignore.case = TRUE)])
    if (length(dirnames.years) == 1) { dirnames.years[2] <- "." }
  }
  rm(dirnames)
  
  # a java garbage collection function
  jgc <- function()
  {
    .jcall("java/lang/System", method = "gc")
  } 
  
  # create a day substring function to process names(list.df)
  d.day <- function(x) {
    pattern <- "([[:alpha:]]+)([[:digit:]]+)([[:alpha:]]+)"
    # remove white space
    s <- gsub(" ","",x)
    # break the string into 3 parts alpha, numeric, alpha
    lst <- lapply(regmatches(s, gregexpr(pattern, s)),
                  function(e) regmatches(e, regexec(pattern, e)))
    # return the numeric part
    sapply(lst, function(x) x[[1]][[3]]) 
  }
  
  # set up the function that retrieves and processes the file for one month  
  get.month <- function(mthdir,p2,nyear) {
    # some needed month lists
    cmList <- c("01","02","03","04","05","06","07","08","09","10","11","12")
    names(cmList) <- c("January","February","March","April","May","June","July","August","September","October","November","December")
    
    #p3 <- paste0(p2,"\\",dirname.month)
    if(regexpr("/$",p2) == -1) {
      p3 <- paste0(p2,"/", mthdir) }
    else {
      p3 <- paste0(p2, mthdir)
    }
    
    setwd(p3)
    # look for a production report file, with or without an "s" at the end of the file name
    if(regexpr("[Ss]$",mthdir) == -1) {
      fn <- mthdir }
    else {
      fn <- substr(mthdir,1, nchar(mthdir)-1)
      
    }
    
    fnames <- dir(pattern=".xlsx")
    
    # remove any files that don't meet the file name pattern
    pattern = paste0("^",fn,"([s ])?",".xlsx","$")
    fnames <- fnames[grepl(pattern,fnames,ignore.case = TRUE)]
    
    #warn if we can't find fn in the file list
    if (!length(grep(fn,fnames)) > 0) {warning(paste("can't find",fn,"in directory. Check spelling, spaces, DIR"))}
    else {
      # load the workbook into memory, only take the first file
      wb <- loadWorkbook(fnames[1], create = FALSE)
      
      # get all the sheet names
      wb.names <- getSheets(wb)
      
      # subset on names that have the month fully spelt out, ignore case
      cMonth <- substring(mthdir,1,regexpr(" ",mthdir)[1]-1)
      wb.names <- wb.names[sapply(wb.names,function(x) grepl(cMonth, x, ignore.case = TRUE))]
      wb.names = wb.names[!grepl("-",wb.names)]
      #create list.df, a list of dataframes. each dataframe will represent a sheet
      #list.df <- readWorksheet(wb,sheet=wb.names,startRow = 7, endRow = -11, startCol = 1,endCol = 9,
      list.df <- readWorksheet(wb,sheet=wb.names,startRow = 7, startCol = 1,endCol = 9,
                               header = FALSE,useCachedValues = TRUE, keep = c(1,6,5,8,9))
      # names(list.df) will default to the names of each sheet, 
      # where each name is equivalent to a date e.g. "December 3rd" is a valid sheet name
      #each list.df dataframe will have the following structure
      # Col1 = Inventory_ID         = column A in each worksheet of the workbook = 1st column
      # Col2 = Targer               = column F in each worksheet of the workbook = 6th column
      # Col3 = Inventory            = column E in each worksheet of the workbook = 5th column
      # Col4 = WrapCount            = column E in each worksheet of the workbook = 8th column
      # Col5 = SHORTS(-) / OVERS(+) = column I in each worksheet of the workbook = 9th column
      
      # remove rows with N/A in Column 1, i.e lines with no Inventory_ID value
      #list.df <- lapply(list.df,subset,! is.na(Col1))
      list.df <- lapply(list.df,subset,
                        !( is.na(Col1) | is.na(Col6) | is.na(Col5) | is.na(Col8) | is.na(Col9)  )   )
      # remove additional rows with N/A in any column
      #list.df <- list.df[complete.cases(list.df),]
      
      # construct the actual date as a character string
      d.year <- paste0("/",nyear)
      d.month = paste0(cmList[cMonth],"/")
      # now replace the orginal "sheet name" style names with date formatted names
      names(list.df) <- paste0(names(list.df),"th") # fix a small error in the data
      names(list.df) <- paste0(d.month,d.day(names(list.df)),d.year)
      
      # convert list of dataframes into one dataframe using ldply with no extra arguments. 
      #The default behaviour for ldply is to insert a new column1 using the names(list.df) (which we just made a set of dates.as.characters) 
      # this effectively add the Date column 
      df.month <- ldply(list.df)
      rm(wb,wb.names,list.df)
      gc()
      jgc()
      # rename the columns
      names(df.month) <- c("Date", "InventoryID", "Target", "Inventory", "WrapCount", "ShortsOvers")
      df.month
    }
  }
  
  # set up the function to retrieve the directory files for each year  
  get.year <- function(nyear,cMonth, ndx) {
    # select the directory for the year
    dn <- dirnames.years[grepl(nyear,dirnames.years,ignore.case = TRUE)]
    # pad the dir list to a length of 2. needed for the period boundary loop to work correctly
    for (j in 1:2) {
	if(length(dn) < 2) {dn[length(dn)+1] <- "." }
    }
    if(regexpr("/$",path) == -1) {
      p2 <- paste0(path,"/",dn[ndx])} 
    else {
      p2 <- paste0(path,dn[ndx])}
    setwd(p2)
    # get the list of month sub-directores
    dirnames.months <- list.dirs(recursive = FALSE, full.names = FALSE)
    # restrict list to one month only
    dirnames.months <- dirnames.months[grepl(cMonth,dirnames.months,ignore.case = TRUE)]
    df.year <- data.frame("Date"=character(0), "InventoryID"=character(0), 
                          "Target"=character(0), "Inventory"=character(0), 
                          "WrapCount"=character(0), "ShortsOvers"=numeric(0))
    #df.year <- mapply(get.month,dirnames.months, MoreArgs = list(p2 = p2, nyear = nyear))
    for (mdir in dirnames.months) df.year <- rbind(df.year, get.month(mthdir=mdir,p2=p2,nyear=nyear))
    df.year
  }
  
  # df.all <- mapply(get.year,years.to.get)
  df.all <- data.frame("Date"=character(0), "InventoryID"=character(0), 
                       "Target"=character(0), "Inventory"=character(0), 
                       "WrapCount"=character(0), "ShortsOvers"=numeric(0))
  #for (yr in mydate.year) df.all <- rbind(df.all, get.year(nyear=yr))
  for (i in 1:length(mydate.year)) df.all <- rbind(df.all, get.year(nyear=mydate.year[i],
                                                                  cMonth=mydate.cMonth[i],ndx=i))
  row.names(df.all) <- NULL  
  
  setwd(path)
  write.csv(df.all, file = "ProdCounts.csv", row.names = FALSE)
 # writeWorksheetToFile("ProdCounts.xlsx", df.all, sheet = "ProdCounts", startRow = 1, startCol = 1)
  #if (back.one.month == TRUE) 
  #  {write.csv(df.all, file = "ProdCountsback.csv", row.names = FALSE)}
  #else 
  #  {write.csv(df.all, file = "ProdCounts.csv", row.names = FALSE)}
}