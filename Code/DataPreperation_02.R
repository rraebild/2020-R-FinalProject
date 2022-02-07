
# Preparation of the input data
#

# Install and load libraries ===================================================
# that are required to process the files

  # Vector with the library names that areneeded to run the code in this file
  requiredLibraries <- c( "tidyverse",
                          "magrittr",
                          "readxl",
                          "tidyxl",
                          "janitor",
                          "stringr")
  
  
  # install and load the libraries if not done so far
  for(pkg in requiredLibraries){
    
    if(!require(pkg, character.only = TRUE)) install.packages(pkg)
    
    library(pkg, character.only = TRUE) 
  }
  
  
  # remove the vectors used during installation
  rm( requiredLibraries, pkg )

  
# Create a Project list  =======================================================

  
  # Path to the project files
  path <- "D:/SyncPrivat/Programmieren/R/2020-R-Learning/edLuxFinalCapstone/2020-R-FinalProject/DataPreperation/Inputs"

  
  # RegEx Pattern for the filename with the following groups
  # \\1  ^(P[:digit:]+)    Starts with P and several numbers = Unknown Projectnumbering 
  # \\2  (\\s-\\sP)        whitespace minus whitespace P = seperator
  # \\3  ([:digit:]+)      P several digits whitespace = Regular Project Number
  # \\4  (\\s)             whitespace = separator
  # \\5  (.*)              multiple characters = Project name
  # \\6  (\\.)             separator
  # \\7  (xl.*)$           xl followed by further characters = filetype
  filenameRegExPattern = "^(P[:digit:]+)(\\s-\\sP)([:digit:]+)(\\s)(.*)(\\.)(xl.*)$"
  
  # double the [ for standard R regEx
  filenameRegExPattern_forStandardR = str_replace_all( filenameRegExPattern, 
                                                       c( "\\[" = "\\[\\[" ,
                                                          "\\]" = "\\]\\]" ))


  # Create a Project List
  projects <- 
    tibble( fileName = list.files( path = path, 
                                   pattern = filenameRegExPattern_forStandardR ),
            
            fileType = str_replace( fileName,
                                    pattern = filenameRegExPattern,
                                    replacement = "\\7"),
            
            
            projectNumber = str_replace( fileName,
                                         pattern = filenameRegExPattern,
                                         replacement = "\\3") %>% 
                              as.integer(),
    
            projectName = str_replace( fileName,
                                       pattern = filenameRegExPattern,
                                       replacement = "\\5")) %>% 
      
      select( projectNumber, projectName, everything())


  
# Load the estimation sheets ===================================================
  
  
  # function to read the project estimates
  read_estimateSheet <- function ( filename = NULL, path = "." ) {

    
    # file to analyse        
    fileHandle = str_c( path, filename, sep = "/")
    
    
    # Load the Excel cell format information and find the formats which hold a "grey" (FFE6E6E6) cell background
    formatsInExcelFile <- xlsx_formats( fileHandle )  
    formatsThatHoldGreyBG <- which( formatsInExcelFile$local$fill$patternFill$fgColor$rgb == "FFE6E6E6")
    
    
    # get the row numbers where the Resource Name (col 3) has a grey background
    rowsWithGreyGolour <- xlsx_cells( fileHandle,
                                      sheets = "Estimate" ) %>% 
      subset( col == 3 & local_format_id %in% formatsThatHoldGreyBG ) %>% 
      pull( row )
    
    
    # Information about columnname and -type
    columnInfo <- tribble( ~name ,             ~type ,
                           "originalWbs" ,     "numeric" ,
                           "lineNo" ,          "text" ,
                           "resourceName" ,    "text" ,
                           "resourceUnit",     "text" ,
                           "resourceType",     "text" ,
                           "resourceNo",       "numeric" ,
                           "resourceProdn",    "numeric" ,
                           "resourceQuantity", "numeric" ,
                           "resourceRate",     "numeric" ,
                           "cost",             "numeric" ,
                           "usage",            "numeric" ,
                           "duration",         "numeric" ,
                           "labour",           "numeric" ,
                           "material",         "numeric" ,
                           "plant",            "numeric" ,
                           "subcontract" ,     "numeric" ,
                           "total" ,           "numeric" ,
                           "portfolioWbs" ,    "text" )
    
    
    # read Excel information
    read_excel(fileHandle,
               sheet = "Estimate",
               range = cell_cols("A:R"),
               col_names = columnInfo$name,
               col_types = columnInfo$type ) %>% 
      
      # add information whether this is a summary item (WBS) or not (PortfolioWBS)
      mutate( isSummary = row_number() %in% rowsWithGreyGolour) %>% 

      # drop lines with no information
      subset( !is_empty(resourceName) & 
              !is.na(resourceName) & 
              !str_detect( str_to_upper(resourceName),
                           pattern = "RESO.*NAM.*|COMMENT") ) %>% 
            
      # create a clean WBS field and 
      mutate( wbs = ifelse (isSummary == TRUE ,
                            originalWbs, NA )) %>% 
      
      # give summary items without original WBS a negative row_number as WBS
      mutate( wbs = ifelse (isSummary == TRUE & is.na(wbs),
                            -1 * row_number(), wbs )) %>% 
      
      # fill the empty wbs (of the portfolio items) with the ones from above
      fill( wbs, .direction ="down") %>% 

      # add variable "isOutsideWbsStructure" for negative wbs
      mutate( isOutsideWbsStructure = ifelse( wbs < 0,
                                              TRUE, FALSE)) %>% 

      # add information whether this is a "subtotal" below the summary item (e,g, streets in project 18) 
      # this is only relevant for P18
      mutate( isSubTotal = ifelse( filename == "P039 - P18 Rural Road Repairs.xlsx" & 
                                     isSummary == TRUE & 
                                     !str_detect( str_to_upper(lineNo),
                                                  pattern = "LINE" ),
                                   TRUE, FALSE )) %>% 
      
      # change the order of the columns
      select( wbs, portfolioWbs, isSummary, isSubTotal, isOutsideWbsStructure, originalWbs, everything()) %>% 
      
      # return the tibble
      return() 
  }

  projects %<>% 
    add_column( estimate = map( projects$fileName, read_estimateSheet, path=path) )
  
# Load the Portfolio WBS sheets ================================================
  
  
  # function to read the portfolio wbs sheet
  read_portfolioWbsSheet <- function ( filename = NULL, path = "." ) {

#filename <- projects[3,]$fileName
        
    # file to analyse        
    fileHandle = str_c( path, filename, sep = "/")
    
    
    # read Excel information
    sheet <- read_excel(fileHandle,
                        #range = cell_cols(1:1000),   #to fix an error that sometimes don't catch columns
                        sheet = "Portfolio WBS")

    
    # remove empty columns
    sheet %<>% 
      remove_empty(which = c("cols"))

    
    # Calculate how many periods there are
    # 9 is the number of leading columns and each period has 5 columns
    numberOfPeriods <- ( length(sheet) - 9) %>% 
                         divide_by_int(5)

        
    # cut off excess columns (if any) 
    sheet %<>% 
      select( 1 : (9 + numberOfPeriods * 5) )

    
    # Information about columnname and -type
    columnInfo <- tribble( ~name ,
                           "portfolioWbs" ,
                           "description" ,
                           "BACQuantity", 
                           "Unit",
                           "BAC",
                           "BACRate",
                           "AACQuantity",
                           "AAC",
                           "AACRate" )
    
    
    # Generate ColumnNames for the Periods
    perioColumnNames <- paste(c("PQ","PV","AQ","AC","EV"), 
                              rep( seq(numberOfPeriods*5 /5), each = 5, len = numberOfPeriods*5), 
                              sep = ".")

    
    # period columns that shall be renamend 
    periodColumnNumbers <- 10:(9+numberOfPeriods*5)
    
    
    #renaming of columns
    sheet %<>% 
      rename_with( ~ columnInfo$name, 1:9) %>% 
      rename_with( ~ perioColumnNames, .cols= periodColumnNumbers)
    
    
    # filter the rows for portfolioWbs entries
    sheet %<>%
      filter( !is_empty(portfolioWbs) & !is.na(portfolioWbs) & portfolioWbs !="" ) %>% 
      filter( !str_detect( str_to_upper(description),
                           pattern = "DESCRIP.*|TOTAL.*" ))
    
    # change the column types
    textColumns = c("portfolioWbs","description","Unit")
    
    sheet %<>% 
      mutate( across( textColumns, as.character )) %>% 
      mutate( across( -textColumns, as.numeric ))

    return (sheet)
  }
      
  # add estimate sheets to projects
  projects %<>% 
    add_column( actual = map( projects$fileName, read_portfolioWbsSheet, path=path) )

 
# Other stuff --------------------------
    
  projects %>% 
    unnest( cols=estimate) %>% 
    filter( projectNumber == 2) %>% 
    filter( isSubTotal == FALSE , isOutsideWbsStructure == FALSE) %>% 
    group_by( projectNumber, projectName, isSummary) %>% 
    #View()
    #summarise( total = sum(total, na.rm = TRUE))
    filter( isSummary == TRUE) %>% 
    left_join( t, by ="wbs") %>% 
    mutate( check = round(total, 2) == round(Budget,2)) %>% View()
    
  

# ---------------
  read.excel <- function(header=TRUE,...) {
    read.table("clipboard",sep="\t",header=header,...)
    
  }
  
  t <- read.table("clipboard",sep="\t",header=TRUE) %>% 
         mutate( Budget = as.numeric( scan(text=Budget, dec=","))) 
    
  
  t %>% 
    filter(isSummary == FALSE, isSubTotal == FALSE, isOutsideWbsStructure == FALSE) %>% 
    summarise( total = sum(total, na.rm = TRUE))
  
  
  