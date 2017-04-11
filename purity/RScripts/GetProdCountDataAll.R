GetProdCountDataAll <- function(path="D:/DATA/PUR/PLANT",years.to.get = c("2014","2015","2016","2017")){
  # This function descends a directory tree and processes all the production report (excel) files 
  # found at the leaves of the tree
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
  if (!require(rJava)) install.packages('rjava')
  library(rJava)
  
  # The XLConnect library is required for the scrip to work
  #options(java.parameters = "-Xmx6g" )
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
  # filter for the production report sub-directories only, for all possible years
  dirnames.years <- dirnames[grepl("^production reports [0-9]{4}$",dirnames,ignore.case = TRUE)]
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
  get.year <- function(nyear) {
    # select the directory for the year
    dn <- dirnames.years[grepl(nyear,dirnames.years,ignore.case = TRUE)]
    if(regexpr("/$",path) == -1) {
      p2 <- paste0(path,"/",dn)} 
    else {
      p2 <- paste0(path,dn)}
    setwd(p2)
    # get the list of month sub-directores
    dirnames.months <- list.dirs(recursive = FALSE, full.names = FALSE)
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
  for (yr in years.to.get) df.all <- rbind(df.all, get.year(nyear=yr))
  row.names(df.all) <- NULL  

  setwd(path)
  #write.csv(df.all, file = "ProdCountsAll.csv", row.names = FALSE)
  writeWorksheetToFile("ProdCounts.xlsx",df.all,sheet = "ProdCounts", startRow = 1, startCol = 1)
  
}