library("readxl")
library(ggplot2)
library(tidyverse)

FileName <- 'Salary and Count.xlsx'
SalaryData <- read_excel(FileName, sheet = 'Salary')
CountData <- read_excel(FileName, sheet = 'Count')

YOSStart <- 1
YOSEnd <- 39
YOSIncrement <- 5
YOSDivision <- 5

AgeStart <- 20
AgeEnd <- 73

#We separate Increment and division because of the chane in increment from 5 to 9 towards the end for 65-73
AgeIncrement <- 5
AgeDivision <- 5

CountMatrix <- as.data.frame(matrix(0,AgeEnd-AgeStart+1,YOSEnd-YOSStart+1))
for(i in 1:nrow(CountMatrix)){
  LastYOSIndexCheckCol <- 0
  for(j in 1:ncol(CountMatrix)){
    #Reset these values every time
    AgeDivision <- 5
    YOSDivision <- 5
    LastYOSIndexCheckRow <- 0
    
    #Recalc the Age to reference from the datasheet
    NewAge <- AgeStart + i - 1
    #Normally the increment is 5 but since the last bracket is 73, 
    #The increment is 9. Its 9 because 73-65 + 1
    #Add the +1 to include 65
    if(NewAge >= 65){
      NewAge <- 65
      AgeDivision <- (AgeEnd - 65) + 1
    }
    
    #SearchValueRow is the value based on the range
    #So if it's 57, it will search for 55
    SearchValueRow <- NewAge - (NewAge %% AgeIncrement)
    SearchIndexRow <- which(CountData[,1] == SearchValueRow)
    
    #Based on the Specific row, find the last column for YOS which is not blank
    LastYOSIndex <- which(is.na(CountData[SearchIndexRow,]))[1] - 1
    if(is.na(LastYOSIndex)){
      #IF NA, take the last column
      LastYOSIndex <- ncol(CountData)
      
      #Only do the last YOS Column the first time. Next time around do the
      #regular 35-40 incremenets
      #This means that for Age 55, the last YOS is 35 and we divide by 1
      #But going forward for age 60 where the last YOS is 35, we divide by 5
      #So to do this, we check the previous row and whether that was NA
      if((!is.na(CountData[SearchIndexRow-1,LastYOSIndex])) & (j >= 5)){
        LastYOSIndexCheckRow <- 1
      }
    }
    
    #The ColCheck value refers to the value used to reference the column data
    #Since j is not a multiple of 5.
    #Before 5, take j as is
    if(j > 5){
      ColCheckValue <- j - (j %% YOSIncrement)
    } else {
      ColCheckValue <- j
    }
    
    #If j is before 5 or if it's the last column AND the previous row was NA
    #i.e 55 and 35 YOS vs 60 35 YOS
    #then and only then do you divide by 1
    if(((j < 5) | (ColCheckValue == colnames(CountData)[LastYOSIndex])) & (LastYOSIndexCheckRow == 0)){
      SearchIndexCol <- which(colnames(CountData) == j)
      YOSDivision <- 1
      
      #if it's the last column before null, i.e. age 20 5 YOS, then only do it for YOS 5
      #LastYOSIndexCheckCol is used to ensure YOS 6, etc. are not processed
      if((j > 5) & ((j %% 5) != 0)){
        LastYOSIndexCheckCol <- 1
        SearchIndexCol <- which(colnames(CountData) == ColCheckValue)
      }
    } else {
      SearchValueCol <-  j - (j %% YOSIncrement)
      SearchIndexCol <- which(colnames(CountData) == SearchValueCol)
      YOSDivision <- YOSIncrement
    }
    
    #If no na in search indecies and its not a non-factor of 5 in the last column
    #i.e.LastYOSIndexCheckCol == 1
    if(!is.na(CountData[SearchIndexRow,SearchIndexCol]) & (LastYOSIndexCheckCol == 0)){
      CountMatrix[i,j] <- as.matrix(CountData[SearchIndexRow,SearchIndexCol] / AgeDivision / YOSDivision)
    }
  }
}

colnames(CountMatrix) <- seq(YOSStart:YOSEnd)
AgeLabels <- AgeStart + seq(0,(AgeEnd-AgeStart))
CountMatrix <- cbind(AgeLabels,CountMatrix)