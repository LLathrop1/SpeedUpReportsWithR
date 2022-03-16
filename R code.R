
# Loading packages --------------------------------------------------------

library(tidyverse)
library(googlesheets4)
library(openxlsx)



# Importing the Google Sheet template -------------------------------------



### This is the sheet we will use
gsheet <- "https://docs.google.com/spreadsheets/d/1XZrDf0eYgSfDiJNj74lBNK3T6AhACeXBn9wsMRlkPoM/edit#gid=0"



# Demo 1 ---------------------------------------

## Adding sheets tot he spradsheet

### creating a vector of the months we will add
sheet_months <- c("JAN2017",
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


for (gtab in sheet_months) {
  
  sheet_add(gsheet,
            sheet = gtab)
  
}



### Adding data to sheets

tmonth<- read.csv(file = paste0("/",gtab,".csv",
                       sep = ""))

sheet_write(tmonth,
            ss = gsheet,
            sheet = gtab)



# Demo 2 ------------------------------------------------------------------


## Adding all months to the main data tab


tmonth<- read.csv(file = paste0("/",gtab,".csv",
                                sep = ""))

sheet_append(ss = gsheet,
             data = tmonth,
             sheet = gtab)




# Demo 3 ---------------------------------


## Sending filter data to other department

### Lets grab the data tab
main_data<- read_sheet(ss = gsheet,
                       sheet = Data)


### Cleaning for just units sold by Sku for Manufacturing department

ManDep <- 
  main_data%>%
  select(SKU = sku,
         Units = quanity)%>%
  group_by(SKU)%>%
  summarise(Units = sum(Units))

### Add data to spreadsheet

sheet_add(ss = gsheet,
          sheet = "ManDep Data")

write_sheet(ManDep,
            ss = gsheet,
            sheet = "ManDep Data")







read.csv(file= "/SpeedUpReportsWithR/Transaction Data/JAN2017.csv")

