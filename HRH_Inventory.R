rm(list = ls())

library(purrr)
library(readxl)
library(openxlsx)


#################################################################################
#Importing a list of excel files in R by sheet names containing a specific string
#################################################################################

#List all the excel files
file_path <- list.files(path = 'C:/Users/cnhantumbo/Documents/HRH Inventory Handbook/HRH Inventory Data', pattern = '\\.xlsx$', full.names = TRUE)

#Read each excel file and combine them in one dataframe
map_df(file_path, ~{
  #get all the names of the sheet
  sheets <- excel_sheets(.x)
  #Select the one which has 'StaffList' in them
  sheet_name <- grep('StaffList', sheets, value = TRUE)
  #Read the excel with the sheet name
  read_excel(.x, sheet_name)
}) -> data

#To extract rows where all cells are empty
data<-data[-which(apply(data,1,function(x)all(is.na(x)))),]

data

HRH_ALL <- "'C:/Users/cnhantumbo/Documents/HRH Inventory Handbook//HRH_ALL.xlsx"
write.xlsx(data,HRH_ALL)



