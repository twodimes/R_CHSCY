########---########---########---########---########---########---########---########---########---########---########---########---
# #1 OPEN FILE AND LOAD DATA, LIBRARIES
########---########---########---########---########---########---########---########---########---########---########---########---
Sys.time()

library(tidyverse) # Data management and reading excel files
library(survey) # For generating survey object and cv/ci calculations
library(memisc) # for cases
library(openxlsx) # Should be best option?
library(foreign) # for reading stata 12 and earlier


#FUNCTION TO COMBINE CHSCY BOOTSTRAPS
merge_bootstraps <- function() {
  CHSCY19_ON <- foreign::read.dta(path_CHSCY19_ON)
  CHSCY19_ON_BOOTWT <- foreign::read.dta(path_CHSCY19_ON_BOOTWT)
  CHSCY19_ON$WTS_S <- NULL #This removes the duplicate weights field so the merged file uses the one from the bootstrap, which is more precise.
  output1 <- merge(CHSCY19_ON,CHSCY19_ON_BOOTWT,by.x = 'ONT_ID', by.y = 'ONT_ID')
  rm(CHSCY19_ON)
  rm(CHSCY19_ON_BOOTWT)
  return(output1)
}

#FUNCTION TO REWRITE COLUMN TITLES FOR OUTPUT
myresult_formatting <- function(input_table) {
  colnames(input_table)<-gsub(newmetricname,"",colnames(input_table))
  colnames(input_table)<-gsub("ci_l.","Lower CI for: ",colnames(input_table))
  colnames(input_table)<-gsub("ci_u.","Upper CI for: ",colnames(input_table))
  colnames(input_table)<-gsub("cv%.","CV for: ",colnames(input_table))
  
  output_table <- input_table
  return(output_table)
}



### README
### SET VALUES BELOW
path_CHSCY19_ON <- "C:/CHSCY/STATA Files/CHSCY19_ON_DISTR.dta"
path_CHSCY19_ON_BOOTWT <- "C:/CHSCY/STATA Files/CHSCY19_ON_BOOTWT.dta"

chscy_descriptions <- read.xlsx("C:/CHSCY/Doc/CHSCY_MetricNames.xlsx",sheet="LabelDescriptions") # Not part of CHSCY - custom made by me to help name the metrics in the excel file.

outputfile <- "C:/CHSCY/Output/FullSummaryOutput.xlsx"


### README
### Only needs to be run once. 
chscy <- merge_bootstraps() ### Comment this out after creating the chscy object in your workspace


### README
### SET VALUES BELOW

chscy$YEAR <- 2019 #Hard Coded because there's only one year now and I'm lazy.
metriclist <- c("YAL_005","YAL_010","YAL_015","YAL_020","YALDVTOD")

### README
### CUSTOM LIST ##---########---########---########---########---########---########---########---########---########---
jv_yearlist <- list(c(2019))
jv_yearlist_labels <- list("2019") ####MAKE LIST HERE


jv_geolist <- list(unique(chscy[chscy$GEODVP16 == 'Health Region Peer Group D',]$GEODVCD),unique(chscy[chscy$GEODVHR4 == 3566,]$GEODVCD),unique(chscy$GEODVCD)) ####MAKE LIST HERE
jv_geolist_labels <- list("Peer Group D","WDGPH","Ontario") ####MAKE LIST HERE

jv_sexlist <- list(c("Male","Female"))
jv_sexlist_labels <- list("Both") ####MAKE LIST HERE

jv_age <- c("custom_age_group")

chscy$custom_age_group <- cases(
  "01-06"=(chscy$DHH_AGE<=6),
  "07-11"=(chscy$DHH_AGE>=7 & chscy$DHH_AGE<=11),
  "12-17"=(chscy$DHH_AGE>=12)
)


chscy$geo_prv <- "ON" ### Grouping variable for all Ontario


########---########---########---########---########---########---########---########---########---########---########---########---
# #2 LOOP TO RECODE AND CREATE VARIABLE GROUPS
########---########---########---########---########---########---########---########---########---########---########---########---


#RE-CODING THE VARIABLE TO REMOVE N/A & NOT STATED RESPONSES
for (i in metriclist) { ### LOOP 1
  metric<-tools::file_path_sans_ext(i)
  metric <- paste0(metric) #FIX/REMOVE
  newmetricname<-paste0("jvm_",metric,"_lab") #FIX/REMOVE
  print(i)
  print(metric)
  
    
  ### CLEAN DATA - REMOVE NON ANSWERS, CREATE NEW VARIABLE
  chscy[,newmetricname] <- dplyr::recode(chscy[,metric],
                                        "Not stated"=NA_character_,
                                        "Not Stated"=NA_character_,
                                        "Don't know"=NA_character_,
                                        "Don't Know"=NA_character_,
                                        "Valid skip"=NA_character_,
                                        "Valid Skip"=NA_character_,
                                        "Not tested"=NA_character_,
                                        "Not Tested"=NA_character_,
                                        "Refusal"=NA_character_
                                        )
  
  
  levels(chscy[,newmetricname])[levels(chscy[,newmetricname])==""] <- NA_character_ ### THIS CLEANS BLANK FIELDS THAT CAN"T BE RE-CODED AS THEY ARE ZERO-LENGTH
  

}# END LOOP 1



Sys.time()
### CREATE SURVEY OBJECT BETWEEN LOOPS 1 AND 2
csurvey <- svrepdesign(data=chscy,
                       type="bootstrap",
                       weights=chscy$WTS_S,
                       repweights="BSW[0-9]+",
                       combined.weights=TRUE)


########---########---########---########---########---########---########---########---########---########---########---########---
# #4 LOOP TO CHECK CVS
########---########---########---########---########---########---########---########---########---########---########---########---



# # CREATE INITIAL WORKBOOK ##---########---########---########---########---########---########---########---########---########---
wb <- createWorkbook()


# # CREATE YEAR TAB   ########---########---########---########---########---########---########---########---########---########---
addWorksheet(wb, sheetName = "Summary")
mytab_row <- 1

y<-list("year","Geography","Metric","MetricDescription","Sex","answer","AgeRange","Percentage","Lower CI","Upper CI","CV","RawCounts")
writeData(wb, sheet = "Summary", x = y, startCol = 1, startRow = mytab_row)
mytab_row <- mytab_row+1


for (i in metriclist) { ### LOOP 2, FOR EACH METRIC
  metric<-tools::file_path_sans_ext(i)
  metric <- paste0(metric)
  newmetricname<-paste0("jvm_",metric,"_lab")
  print(metric)


  
  
  for (y in 1:length(jv_yearlist)) { ### LOOP 2A for year/geography tab
    jv_year<-jv_yearlist[[y]]
    jv_year_label<-jv_yearlist_labels[[y]]    
    #print(jv_year)
    #print(jv_year_label)
    
    
    
    for (g in 1:length(jv_geolist)) { ### LOOP 2A for year/geography tab
      jv_geo<-jv_geolist[[g]]
      jv_geo_label<-jv_geolist_labels[[g]]
      #print(jv_geo)
      #print(jv_geo_label)
      
      
      for (s in 1:length(jv_sexlist)) { ### LOOP 2A for year/geography tab
        jv_sex<-jv_sexlist[[s]]
        jv_sex_label<-jv_sexlist_labels[[s]]
        #print(s)
        #print(jv_sex)
        #print(jv_sex_label)
    
        
        print(paste0(jv_year_label," - ",jv_geo_label," - ",jv_sex_label))
        
        
        ### Junk for counts starts here
        
        #Counts by Age
        myresult_counts_t <- table(
          subset(chscy[,newmetricname], chscy$GEODVCD %in% jv_geo & chscy$YEAR %in% jv_year & chscy$DHH_SEX %in% jv_sex),
          subset(chscy$custom_age_group, chscy$GEODVCD %in% jv_geo & chscy$YEAR %in% jv_year & chscy$DHH_SEX %in% jv_sex)
        )
        myresult_counts <- as.data.frame.matrix(myresult_counts_t)
        
        
        #Counts total, all ages
        myresult_totalcounts_t <- table(subset(chscy[,newmetricname], chscy$GEODVCD %in% jv_geo & chscy$YEAR %in% jv_year & chscy$DHH_SEX %in% jv_sex),
                                        subset(chscy$geo_prv, chscy$GEODVCD %in% jv_geo & chscy$YEAR %in% jv_year & chscy$DHH_SEX %in% jv_sex)
        )
        myresult_totalcounts <- as.data.frame.matrix(myresult_totalcounts_t)
        
        ### Junk for counts ends here
        
        
        
        if (sum(myresult_counts_t)>0) { ### TEST FOR DATA, ELSE DROP FOR UNASKED YEARS
          
          
          
          
          
          myresult_age <- svyby(make.formula(newmetricname), # variable to pass to function
                                by = make.formula(jv_age),  # grouping by age range - which age grouping is defined above
                                design = subset(csurvey, GEODVCD %in% jv_geo & YEAR %in% jv_year & DHH_SEX %in% jv_sex), # design object with subset definition
                                vartype = c("ci","cvpct"), # report variation as ci, and cv percentage
                                na.rm=TRUE,
                                na.rm.all=TRUE,
                                FUN = svymean # specify function from survey package, mean here
          )
          
          
          myresult_total <- svyby(make.formula(newmetricname), # variable to pass to function
                                  by = ~geo_prv,  # grouping by age range - which age grouping is defined above
                                  design = subset(csurvey, GEODVCD %in% jv_geo & YEAR %in% jv_year & DHH_SEX %in% jv_sex), # design object with subset definition
                                  vartype = c("ci","cvpct"), # report variation as ci, and cv percentage
                                  na.rm=TRUE,
                                  na.rm.all=TRUE,
                                  FUN = svymean # specify function from survey package, mean here
          )
          
          
          
          
          
          
          
          answerlist<-levels(chscy[,newmetricname])
          for (a in answerlist) {
            
            
            writeData(wb, sheet = "Summary", x = jv_year_label, startCol = 1, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = jv_geo_label, startCol = 2, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = metric, startCol = 3, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = chscy_descriptions[chscy_descriptions$variable_lc==metric,]$Description, startCol = 4, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = jv_sex_label, startCol = 5, startRow = mytab_row)
            
            writeData(wb, sheet = "Summary", x = a, startCol = 6, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = "TOTAL (1-17)", startCol = 7, startRow = mytab_row)
            
            writeData(wb, sheet = "Summary", x = myresult_total[,paste0(newmetricname,a)], startCol = 8, startRow = mytab_row)
            
            writeData(wb, sheet = "Summary", x = myresult_total[,paste0("ci_l.",newmetricname,a)], startCol = 9, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = myresult_total[,paste0("ci_u.",newmetricname,a)], startCol = 10, startRow = mytab_row)
            writeData(wb, sheet = "Summary", x = myresult_total[,paste0("cv%.",newmetricname,a)], startCol = 11, startRow = mytab_row)
            
            
            ### Junk here for counts starts
            writeData(wb, sheet = "Summary", x = myresult_totalcounts[a,], startCol = 12, startRow = mytab_row)
            ### Junk here for counts ends 
            
            
            mytab_row <- mytab_row+1
            
            
            
            
            
            age_rangeprint<-levels(factor(chscy[,jv_age]))
            for (r in age_rangeprint) {
              
              writeData(wb, sheet = "Summary", x = jv_year_label, startCol = 1, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = jv_geo_label, startCol = 2, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = metric, startCol = 3, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = chscy_descriptions[chscy_descriptions$variable_lc==metric,]$Description, startCol = 4, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = jv_sex_label, startCol = 5, startRow = mytab_row)
              
              writeData(wb, sheet = "Summary", x = a, startCol = 6, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = r, startCol = 7, startRow = mytab_row)
              
              writeData(wb, sheet = "Summary", x = myresult_age[r,paste0(newmetricname,a)], startCol = 8, startRow = mytab_row)
              
              writeData(wb, sheet = "Summary", x = myresult_age[r,paste0("ci_l.",newmetricname,a)], startCol = 9, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = myresult_age[r,paste0("ci_u.",newmetricname,a)], startCol = 10, startRow = mytab_row)
              writeData(wb, sheet = "Summary", x = myresult_age[r,paste0("cv%.",newmetricname,a)], startCol = 11, startRow = mytab_row)
              
              
              ### Junk here for counts starts
              writeData(wb, sheet = "Summary", x = myresult_counts[a,r], startCol = 12, startRow = mytab_row)
              ### Junk here for counts ends 
              
              
              
              
              mytab_row <- mytab_row+1
            } # End age range print R
          } # End AnswerList A
        
        
        } ### END 'IF' STATEMENT TO CHECK IF THERE IS VALID DATA - SKIP ABOVE IF NOT ASKED IN CERTAIN YEAR

      } ### END LOOP 2C for Sex tab
  
    } ### END LOOP 2B for Geo tab
    
  } ### END LOOP 2A for Year tab

} ### END LOOP 2 for Metric tab
  
  
  
  

  
  
saveWorkbook(wb, file = outputfile, overwrite = TRUE)




########---########---########---########---########---########---########---########---########---########---########---########---
# #5 VALIDATE VALUES, CV AND CI
########---########---########---########---########---########---########---########---########---########---########---########---


#CALCULATIONS ON JDV_METRIC, BY JDV_AGE AND REGIONS (JDV_REGIONS?)
#CALCULATE BY YEAR, 2 YEAR, OR COMBINED


#AND FORMATTING?
#  FLAG HIGH CV/INVALID CV/QUALITY INDICATOR ABCDE

#  FLAG LOW COUNTS


########---########---########---########---########---########---########---########---########---########---########---########---
# #6 END SCRIPT, CLOSE LOG
########---########---########---########---########---########---########---########---########---########---########---########---

#CLOSE AND SAVE, DONE LOOPING
Sys.time()




# ### EXAMPLE OUTPUT TO EXCEL FOR SECONDARY VALIDATION
# library(writexl) # No formatting
# exd <- select_at(chscy, vars(ont_id, year, DHH_SEX, DHH_AGE, starts_with("jv_age"), GEODVCD, starts_with("jvm"), starts_with("phc"), starts_with("mex"), starts_with("mxa"), starts_with("mxs"), wts_s))
# exd <- select_at(chscy, vars(ont_id, year, DHH_SEX, DHH_AGE, starts_with("jv_age"), starts_with("jvm"), starts_with("geo"), starts_with("sui"), GEODVCD, wts_s, wts_shh))
# write_xlsx(exd, paste0("C:/Users/jamesm.WDGHU/OneDrive - Wellington Dufferin Guelph Public Health/Projects/ExcelExport_suicide.xlsx"))
# 
# 
