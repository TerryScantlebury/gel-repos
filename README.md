# gel-repos  
GEL repos  

# Purity Repos  

## Purity Rscripts  

Location in Repo:  ./purity/Rscripts
A directory of R scripts that are currently implemented at Purity  
Server Name & IP: PUR-DELLSVR01: 205.214.111.5  
Location on Server: C:\Users\bis-sol\Documents\RScripts  

R script files:  
GetProdCountData.r:  
This function descends a directory tree and processes the production report (excel) files for the two (2) most recent months. 
The assumption is that at the end of two (2) months, the prior month file will be stable and all edits would have been done. 
The current date determines the year/ month.  
* path.p1 = The top level directory is "D:/DATA/PUR/PLANT", a general store for all types of production files
* path.p2 = The next level down is "./Production Reports yyyy", one folder per year. is.numeric(year)
* path.p3 = The next level down is "./mmmm yyyy Production Reports", one folder per month. 
* Each folder has one file. is.character(month)
* path.p4 = The actual files e.g "mmmm yyyy Production Reports.xlsx", 
* an excel file for the month with one sheet per day
* a fully qualified file name example (p1/p2/p3/p4) is :- 
* D:/DATA/PUR/PLANT/Production Reports 2016/January 2016 Production Reports/January 2016 Production Reports.xlsx
* N.B. because of inconsistency in saving the excel workbooks, files are sometimes found as level path.p2 or path.p3.
* The script has enough logic to check for this inconsistency. 

GetProdCountDataAll.r:  
This function descends a directory tree and processes **ALL** the production report (excel) files. 
This function was used to do the initial population of the data files.  
* path.p1 = The top level directory is "D:/DATA/PUR/PLANT", a general store for all types of production files
* path.p2 = The next level down is "./Production Reports yyyy", one folder per year. is.numeric(year)
* path.p3 = The next level down is "./mmmm yyyy Production Reports", one folder per month. 
* Each folder has one file. is.character(month)
* path.p4 = The actual files e.g "mmmm yyyy Production Reports.xlsx", 
* an excel file for the month with one sheet per day
* a fully qualified file name example (p1/p2/p3/p4) is :- 
* D:/DATA/PUR/PLANT/Production Reports 2016/January 2016 Production Reports/January 2016 Production Reports.xlsx
* N.B. because of inconsistency in saving the excel workbooks, files are sometimes found as level path.p2 or path.p3.
* The script has enough logic to check for this inconsistency. 

LoadGetProdData.r:  
This is the r script that is called by the windows scheduler - to process GetProdCountData.r.  
The current schedule runs once daily at 3:30 A.M.  
The current schedule runs on the server PUR-DELLSVR01  

## Purity Windows Schedule  

Scheduled task name: _GetProdCountdata  
Scheduled task command: "CMD "/c  ""D:\Program Files\R\R-3.3.2\bin\Rscript.exe" "C:/Users/Administrator.PURITY/Documents/RScripts/LoadGetProdData.R"""  
Scheduled task start in folder: "C:\Users\Administrator.PURITY\Documents\RScripts"

## Purity SQL Jobs
Location in Repo:  ./purity/SQLScripts  
A directory of SQL scripts that are currently implemented at Purity  
Server Name & IP: PUR-SOLOMON: 205.214.111.8  
SQL Server: PUR-SOLOMON  
Job name : Purity_PowerBI_updated  
Location on SQL Server: (SQL jobs object)  

Purity_PowerBI_updated.sql : Scripted steps to update Power_BI_data tables.  
* Step 1: CopyProdCounts : Runs the following DTS package to import the count data "C:\Users\administrator.PURITY\Documents\Integration Services Script Task\GetProdCounts.dtsx"  
* Step 2: All steps : (all other steps to create the power bi data files from the live applicatikon data)  

Purity_PowerBI_Create_New.sql : Scripted steps to recreate an empty Power_BI_data tables database.  
