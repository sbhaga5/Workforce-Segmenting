library("readxl")
library(ggplot2)
library(tidyverse)

FileName <- 'Salary and Count.xlsx'
CountData <- read_excel(FileName, sheet = 'Count')
SalaryData <- read_excel(FileName, sheet = 'Salary')
SalaryIncreaseData <- read_excel(FileName, sheet = 'Salary Increase')
#Add 1 to make this calculation easier for later
SalaryIncreaseData$Rate <- SalaryIncreaseData$Rate + 1

YOSStart <- 1
YOSEnd <- 39
YOSIncrement <- 5
YOSDivision <- 5

AgeStart <- 20
AgeEnd <- 73

#We separate Increment and division because of the chane in increment from 5 to 9 towards the end for 65-73
AgeIncrement <- 5
AgeDivision <- 5
#
#####################################################################################################################
#
#Functions

#The compound salary rates are to make the multiplication later easier
#Here we want the salaries to increase for age and YOS, following a zig-zag pattern
#If the rates are 4% across the board, the rates would be 4%, 9%, etc. relative to 
#the starting salary
CompoundSalaryIncrease <- function(Data){
  for(i in 2:nrow(Data)){
    for(j in 2:ncol(Data)){
      Data[i,j] <- Data[i,j]*Data[i-1,j-1]
    }
  }
  
  return(Data)
}

#After getting the compound salary data, calc the starting salary for that group
GetStartingSalary <- function(Average,SalaryIncreaseCompound,AgeDivision,YOSDivision){
  #Since we assume the same starting salary for all ages in that group, we can calculate
  # a "base average" which will be the sum of the compound values divide by 25 (age and yos div)
  #This makes calculating the starting salary much easier
  #This is only possible because the starting salaries for all ages in that group are the same
  #if not, new method required
  BaseAverage <- sum(SalaryIncreaseCompound) / AgeDivision / YOSDivision
  StartingSalary <- Average / BaseAverage
  return(StartingSalary)
}
#
#####################################################################################################################
#
#Salary increase data for later
#Take the long form data from the excel sheet and covert to matrix form
SalaryIncreaseMatrix <- expand_grid(20:(20+AgeEnd-AgeStart+1),1:(YOSEnd-YOSStart+1))
colnames(SalaryIncreaseMatrix) <- c('Age','YOS')
SalaryIncreaseMatrix <- left_join(SalaryIncreaseMatrix,SalaryIncreaseData) %>%
  pivot_wider(names_from = YOS,values_from = Rate)

#Create salary matrix
SalaryMatrix <- as.data.frame(matrix(0,AgeEnd-AgeStart+1,YOSEnd-YOSStart+1))
colnames(SalaryMatrix) <- seq(YOSStart:YOSEnd)
AgeLabels <- AgeStart + seq(0,(AgeEnd-AgeStart))
#Salary and headcount matrix have the same structure so initialize both at the same time
SalaryMatrix <- CountMatrix <- cbind(AgeLabels,SalaryMatrix)

for(j in 2:ncol(SalaryData)){
  for(i in 1:nrow(SalaryData)){
    #Reset these values
    AgeDivision <- 5
    YOSDivision <- 1
    
    #Up to 65, standard division by 5
    #After that, its 73-65+1
    if(i == nrow(SalaryData)){
      AgeRows <- as.double(SalaryData[i,1]) + seq(0,8)
      AgeDivision <- AgeEnd - 65 + 1
    } else {
      AgeRows <- as.double(SalaryData[i,1]) + seq(0,4)
    }
    #Row Index for all matching matrix values from the data sheet
    RowIndex <- which(SalaryMatrix[,1] == AgeRows)
    
    #Condition 1, all YOS less than 5
    if(as.double(colnames(SalaryData)[j]) < 5){
      ColIndex <- which(colnames(SalaryMatrix) == colnames(SalaryData)[j])
      
      #Condition 2, YOS >5, less than the final value
    } else if(j < ncol(SalaryData)){
      #Condition 2.1, if the value after this column is NA, its the last relavent column
      #and thus divide by 1 for YOS
      if(is.na(SalaryData[i,j+1])){
        ColIndex <- which(colnames(SalaryMatrix) == colnames(SalaryData)[j])
        
        #Condition 2.2 if not, then it's the regular range and thus divide by 5
        #for YOS. in addition, the ColIndex is for a range of YOS
      } else {
        YOSCols <- as.double(colnames(SalaryData)[j]) + seq(0,4)
        YOSDivision <- 5
        ColIndex <- which(colnames(SalaryMatrix) == YOSCols)
      }
      
      #Condition 3, at the last column of the input sheet
    } else if(j == ncol(SalaryData)){
      if(i > 1){
        
        #Condition 3.1 If its the last column and the previous row for that column
        #was NA, then divide by 1. For example, age 50 goes until YOS 35. Divide by 1 for that year
        #But Age 55 also goes to YOS 35, dont divide by 1 then, divide by 5
        if(is.na(SalaryData[i-1,j])){
          ColIndex <- which(colnames(SalaryMatrix) == colnames(SalaryData)[j])
        } else {
          
          #Condition 3.2 regular division by for both age and YOS
          YOSCols <- as.double(colnames(SalaryData)[j]) + seq(0,4)
          YOSDivision <- 5
          ColIndex <- which(colnames(SalaryMatrix) == YOSCols)
        }
      }
    }
    
    #Take the Row and Col Indecies and divide headcounts
    CountMatrix[RowIndex,ColIndex] <- CountData[i,j] / AgeDivision / YOSDivision
    
    #For Payroll, if the YOS division, is 1, take the same salary across all ages
    if(YOSDivision == 1){
      SalaryMatrix[RowIndex,ColIndex] <- SalaryData[i,j]
    } else {
      #The method here has 3 steps:
      #1. Get the compunded salary rates for that subset
      #2. Calulate the starting salary for that subset to get the required avg
      #3. multipl 1&2 to get the salaries for that group
      SalaryIncreaseCompound <- CompoundSalaryIncrease(SalaryIncreaseMatrix[RowIndex,ColIndex])
      SalaryMatrix[RowIndex,ColIndex] <- as.double(GetStartingSalary(SalaryData[i,j],SalaryIncreaseCompound,AgeDivision,YOSDivision)) * SalaryIncreaseCompound
    }
    
    #Remove NAs
    CountMatrix[is.na(CountMatrix)] <- 0
    SalaryMatrix[is.na(SalaryMatrix)] <- 0
  }
}
