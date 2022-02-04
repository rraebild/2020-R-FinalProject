
# Preparation of the input data
#

# Install and load libraries ---------------------------------------------------------------


  # Libraries needed to run the code in this file
  requiredLibraries <- c( "tidyverse",
                          "magrittr",
                          "readxl")
  
  
  # install and load the libraries if not done so far
  for(pkg in requiredLibraries){
    
    if(!require(pkg, character.only = TRUE)) {
      install.packages(pkg)
      library(pkg, character.only = TRUE) }
    
  }
  
  
  # remove the package installation vectors
  rm( requiredLibraries, pkg )

  
  path <- "D:/SyncPrivat/Programmieren/R/2020-R-Learning/edLuxFinalCapstone/2020-R-FinalProject/DataPreperation/Inputs"
  fileHandle <- paste(path, "P018 - P01 Airport Carpark.xlsx", sep="/" )
  excel_sheets( fileHandle ) 
