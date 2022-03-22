
# Loading packages --------------------------------------------------------

library(tidyverse)
library(googlesheets4)
library(openxlsx)



# Importing the Google Sheet template -------------------------------------



### This is the sheet we will use 
gsheet <- "https://docs.google.com/spreadsheets/d/INSERT YOUR SPREAD SHEET ID HERE/edit#gid=0"



# Demo 1 ---------------------------------------

## Adding sheets to the spreadsheet

### creating a vector of the months we will add
sheet_months <- list("JAN2017",
                  "FEB2017",
                  "MAR2017",
                  "APR2017",
                  "MAY2017",
                  "JUN2017",
                  "JUL2017",
                  "AUG2017",
                  "SEP2017",
                  "OCT2017",
                  "NOV2017",
                  "DEC2017")

### Creating the tabs

for (gtab in sheet_months) {
  
  sheet_add(gsheet,
            sheet = gtab)
  
}



### Adding data to sheets

for(gtab in sheet_months){
tmonth<- read.csv(file = paste0("Transaction Data/",gtab,".csv",
                       sep = ""))

sheet_write(tmonth,
            ss = gsheet,
            sheet = gtab)
}


# Demo 2 ------------------------------------------------------------------


## Adding all months to the main data tab

for (gtab in sheet_months) {
  
  tmonth<- read.csv(file = paste0("Transaction Data/",gtab,".csv",
                                  sep = ""))
  
  sheet_append(ss = gsheet,
               data = tmonth,
               sheet = "Data")
 
  
}




# Demo 3 ---------------------------------


## Sending filter data to other department

### Lets grab the data tab
main_data<- read_sheet(ss = gsheet,
                       sheet = "Data")

glimpse(main_data)

### Cleaning for just units sold by SKU for Manufacturing department

ManDep <- 
  main_data%>%
  select(SKU = sku,
         Units = quantity)%>%
  group_by(SKU)%>%
  summarise(Units = sum(Units))


View(ManDep)

### Add data to spreadsheet

sheet_add(ss = gsheet,
          sheet = "ManDep Data")

write_sheet(ManDep,
            ss = gsheet,
            sheet = "ManDep Data")




