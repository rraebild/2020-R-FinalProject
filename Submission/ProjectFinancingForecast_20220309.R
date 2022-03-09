
# Install and load libraries and global setting ================================
# that are required to process the files

  # Vector with the library names that areneeded to run the code in this file
  requiredLibraries <- c( "tidyverse",
                          "magrittr",
                          "httr",
                          "utils",
                          "readxl",
                          "tidyxl",
                          "janitor",
                          "stringr",
                          "caret",
                          "knitr",
                          "kableExtra")
  
  
  # install and load the libraries if not done so far
  for(pkg in requiredLibraries){
    
    if(!require(pkg, character.only = TRUE)) install.packages(pkg)
    
    library(pkg, character.only = TRUE) 
  }
  
  
  # remove the vectors used during installation
  rm( requiredLibraries, pkg )
  
  # remove the annoying summarise message about groupings :-)
  options(dplyr.summarise.inform = FALSE)

  
# Download files and prepare Project list  =====================================

  ## Create a data directory if it does not exist
  {
    
    # path of data directory in the current working directory
    dataDirectoryPath <- file.path( getwd(), "DataPreperation" )
    
    # create if it does not exist
    if ( ! dir.exists( dataDirectoryPath )) dir.create(dataDirectoryPath)
    
  }

  
  ## download the input data if it does not already exists
  ## then unzip 
  {
    if ( ! file.exists( file.path( dataDirectoryPath, "ProjectData.zip") )) {
      
      url <- "https://github.com/rraebild/2020-R-FinalProject/raw/main/DataPreperation/Inputs/ProjectData.zip"
      GET(url, write_disk( file.path( dataDirectoryPath, "ProjectData.zip"), overwrite=TRUE))
      
    }
    
    
    unzip( file.path( dataDirectoryPath, "ProjectData.zip"),
           overwrite = TRUE,
           junkpaths = TRUE,
           exdir = dataDirectoryPath )
    
  }

  
  ## Prepare an overview  of all projects that have been downloaded
  {
    
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
      tibble( fileName = list.files( path = dataDirectoryPath, 
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
    
    
    # remove unnecessary variables
    rm( filenameRegExPattern, filenameRegExPattern_forStandardR)
  }

  
# Import the project data ======================================================
  
  ## Load the estimation sheet
  {
    # function to read the project estimates
    read_estimateSheet <- function ( filename = NULL, path = "." ) {
  
      
      # file to analyse        
      fileHandle = str_c( dataDirectoryPath, filename, sep = "/")
      
      
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
      add_column( estimate = map( projects$fileName, 
                                  read_estimateSheet, 
                                  path=path) )
    
    # remove the function, because it is no longer needed
    rm ( read_estimateSheet)
  }
  

  ## Load the porfolio sheet 
  {
    # function to read the portfolio wbs sheet
    read_portfolioWbsSheet <- function ( filename = NULL, path = "." ) {
  
    #used only for manual debugging
    #filename <- projects[3,]$fileName
            
      # file to analyse        
      fileHandle = str_c( dataDirectoryPath, filename, sep = "/")
      
          
      # read Excel information
      sheet <- read_excel(fileHandle,
                          #range = cell_cols(1:1000),   #to fix an error that sometimes don't catch columns
                          sheet = "Portfolio WBS")
  
      
      # remove empty columns
      sheet %<>% 
        remove_empty(which = c("cols"))
  
      
      # Information about columnnames ( is different for some sheet)
      columnInfo <- list( name = c( "portfolioWbs", 
                                    "description",
                                    "bacQuantity",
                                    "unit",
                                    "bac",
                                    "bacRate",
                                    "aacQuantity",
                                    "aac",
                                    "aacRate" ),
                          
                          name2 = c( "portfolioWbs" ,
                                     "description" ,
                                     "bacQuantity",
                                     "unit",
                                     "bacRate",
                                     "bac",
                                     "aacQuantity" ,
                                     "aac",
                                     "aacRate" ),
                          
                          name3 = c( "portfolioWbs" ,
                                     "description" ,
                                     "bacQuantity" ,    
                                     "unit" ,
                                     "originalBACRate" ,
                                     "originalBAC",
                                     "bac" ,
                                     "bacRate" ,
                                     "aacQuantity" ,
                                     "aac" ,
                                     "aacRate" ) )
  
      
      # General columnnames that shall be used, depending on the project
      generalColumnNames <- 
        if (filename == "P021 - P04 Regional Arterial Pavement Repairs.xlsx") 
           columnInfo$name2 else 
        if (filename == "P038 - P17 Marina Sub-division.xlsx")
           columnInfo$name3 else 
        columnInfo$name
  
          
      # number of General columns
      numberOfGeneralColumns <- length( generalColumnNames )
  
      
      # Calculate how many periods there are
      # each period has 5 columns
      numberOfPeriods <- ( length(sheet) - numberOfGeneralColumns) %>% 
        divide_by_int(5)
      
      
      # cut off excess columns (if any) 
      sheet %<>% 
        select( 1 : (numberOfGeneralColumns + numberOfPeriods * 5) )
      
      
      # Generate ColumnNames for the Periods
      perioColumnNames <- paste(c("pq","pv","aq","ac","ev"), 
                                rep( seq(numberOfPeriods*5 /5), each = 5, len = numberOfPeriods*5), 
                                sep = ".")
  
      
      # period columns that shall be renamend 
      periodColumnNumbers <- (numberOfGeneralColumns +1) :( numberOfGeneralColumns +numberOfPeriods*5)
      
      
      #renaming of columns
      sheet %<>% 
        rename_with( ~ generalColumnNames, .cols= 1:numberOfGeneralColumns) %>% 
        rename_with( ~ perioColumnNames, .cols= periodColumnNumbers)
      
      
      # filter the rows for portfolioWbs entries
      sheet %<>%
        filter( !is_empty(portfolioWbs) & !is.na(portfolioWbs) & portfolioWbs !="" ) %>% 
        filter( !str_detect( str_to_upper(description),
                             pattern = "DESCRIP.*|TOTAL.*" ))
      
      # change the column types
      textColumns = c("portfolioWbs","description","unit")
      
      sheet %<>% 
        mutate( across( textColumns, as.character )) %>% 
        mutate( across( -textColumns, as.numeric ))
  
      return (sheet)
    }
        
    # add estimate sheets to projects
    projects %<>% 
      add_column( progress = map( projects$fileName, read_portfolioWbsSheet, path=path) )
    
    # remove the funcion, because it is no longer needed
    rm ( read_portfolioWbsSheet )
  }
  
 
  ## transform progress data of the periods into long format and add to project data set
  {
    
    projects %<>% 
      
      # start with the existing data
      select ( projectName, projectNumber, progress) %>% 
      unnest( progress ) %>% 
      
      # transform period values from wide to long format
      select( projectNumber, portfolioWbs, description, contains(".")) %>% 
      pivot_longer( cols = contains(".") ,
                    values_to = "periodValue") %>% 
      
      separate( col = name,
                into = c("valueType", "period"),
                convert = TRUE) %>% 
      
      # remove rows for empty periods (e.g. if some project have a shorter duration than others)
      drop_na( periodValue ) %>% 
      
      # nest the data into a list colum
      nest( periodProgress = -projectNumber) %>% 
      
      # add the list column to the original table
      right_join( projects, by = "projectNumber" ) %>% 
      
      # reorder the columns
      select( projectNumber, projectName, everything())
    
  } 
  
  
  ## remove unnecessary variables
  rm( dataDirectoryPath )
  
  
  ## save the loaded data into an RData file within the current working directory
  ## so that the "R markdown report" can make use of this data
  
  save.image( file.path( getwd() , 
                         "Report",
                         "projectDataLoaded.RData" ) )
  
  
# DataSet Wrangling ============================================================

  ## function to define the base data set
  baseSet <- function( .removeNAs = TRUE , .baseIndex = FALSE , 
                       .includeValueTypes = c( "ac", "pv" ) ) {
    
    #only for debugging
    
    # start with the full project data set
    projects %>% 
    
      # omit project data that is not relevant for this analysis
      select( projectNumber, projectName, periodProgress ) %>% 
      
      # unnest the progress data to have a tidy, long table with all observations
      unnest( periodProgress ) %>% 
        
      # filter for the periodValues that are part of this analysis
      filter( valueType %in% .includeValueTypes ) %>% 
      
      # if an index shall be added, e.g. in order to restore to original values
      { if (.baseIndex) rownames_to_column( . , var = "baseIndex" ) else . } %>% 
      
      # remove NA entries if necessary
      { if (.removeNAs) na.omit(.) else . } 

  }
    

  ## create test and training data sets from the same base data
  {
    # set the seed to create reproducible partitions
    set.seed(20)
    
    # central definition of the base data set
    baseSet() %T>%
      
      # create training index (within the global environment)
      { trainIndex <<- createDataPartition( y = .$periodValue, 
                                           times = 1, 
                                           p = 0.8, 
                                           list = FALSE)} %>% 
        
      # create a list with training and testing data sets (in global environment)
      { dataSet <<- list ( training = .[ trainIndex , ],
                           testing = .[ -trainIndex , ] )}
    
    
    # nest the progress data to project level
    for ( i in c(1,2) ) {
      dataSet[[i]] %<>% 
        nest( progress = -c(projectNumber, projectName))
    }
    
    
    # remove variables  no longer needed
    rm ( trainIndex, i)
  }
  
  ## function to explore the full dataSet
  fullDataSet <- function ( .dataSet = dataSet ) {
    
    a <- .dataSet[["training"]] %>% 
      add_column( ID = "training")
    
    b <- dataSet[["testing"]] %>% 
      add_column( ID = "testing")
    
    bind_rows( a, b ) %>% 
      return()
  }


  ## Create a mapping table between periods and standard periods
  {
    
    # maximum number of periods that exist in the dataset
    maxPeriods <- fullDataSet() %>% 
      unnest( progress ) %$%
      max( period )
    
    
    # Number of standardized periods (shall always be bigger than normal periods, usually 100)
    stdPeriodNumber = max( 100,  maxPeriods) 
    
    
    # there are always more stdPeriod das map to a period, therefore each
    # stdperiod covers up to a certain fraction of a period
    # the conversion break is the lenght of one such fraction
    conversionBreak = maxPeriods / stdPeriodNumber
    
    # Mapping table between periods and stdPeriods
    periodMappings <-
      
      # create a tibble with a row for each period
      tibble ( period = 1:maxPeriods) %>% 
      
      # find the fraction until which the stdPeriod covers
      rowwise () %>% 
      mutate( toBreak = max( seq( from = conversionBreak, 
                                  to = period, 
                                  by = conversionBreak ))) %>% 
      
      # find the fraction from which the stdPeriod covers
      ungroup () %>% 
      mutate( fromBreak = lag(toBreak, default = 0) + conversionBreak ) %>% 
      
      # nest the list of stdPeriod that match do a specific period into a new column
      rowwise() %>% 
      mutate( stdPeriod = list( seq( from = fromBreak,
                                     to = toBreak,
                                     by = conversionBreak )
                                / conversionBreak )) %>%
      
      # remove the temporary columns 
      select( -fromBreak, -toBreak) 
    
    # add the mapping table to the dataSet
    dataSet %<>% 
      c( periodMapping = list( periodMappings ) )
    
    # remove the temporary table
    rm( periodMappings )
  
  }
  
  
  ## transform project progress data to standardized periods 
  for (i in 1:2 ) {

      # must incluce change %<>%
    dataSet[[i]] %<>% 
      unnest( progress ) %>% 
      
      # add the list of stdPeriods to progress data
      left_join( dataSet$periodMapping, by = "period") %>% 
      
      # expand the new rows for stdPeriods (by definition: stdPeriods are always more than normal Periods) 
      unnest(stdPeriod) %>% 
      
      # nest back the progress data
      nest( progress = -any_of( c( "projectNumber", "projectName", "total" )))
      
  }
  

  ## define a function to pull the portfolioWbs elements
  portfolioWbs <- function ( .projectNumber = NULL, .portfolioWbs = NULL , .onlyUngroupedResources = FALSE ) {
    
    # only if debugging the function
    # .projectNumber <- 1:2
    # .portfolioWbs <- 11:904
    # .onlyUngroupedResources <- TRUE
    
    projects %>%
      # Gather the data
      select ( projectNumber, estimate) %>% 
      unnest( estimate ) %>%
      select( -c( isSummary, isSubTotal, originalWbs, lineNo)) %>% 
      filter( !is.na(portfolioWbs) ) %>% 
      
      # filter for project if required
      { if ( !is.null(.projectNumber)  ) filter( . , projectNumber %in% .projectNumber ) else . } %>% 
      
      # filter for portfoliowbs  if required
      { if ( !is.null(.portfolioWbs)  )  filter( . , portfolioWbs %in% as.character( .portfolioWbs ) ) else . } %>% 
      
      # add an index "resourceWeight" for the resources in the WBS,
      # this tells how much "weight" a particular resource has within the overall cost of the portfolioWBS
      # e.g. a resource may have 10% of the toal cost of a portfolioWbs
      group_by(projectNumber, portfolioWbs ) %>% 
      mutate( resourceWeight = total / sum( total, na.rm = TRUE)) %>% 
      ungroup() %>% 
      
      # summarize to resourceType level
      group_by( projectNumber, portfolioWbs, resourceType ) %>% 
      summarise( resourceWeight = sum( resourceWeight, na.rm = TRUE)) %>% 
      ungroup() %>% 
      
      # check in which form the data shall be return
      { if ( .onlyUngroupedResources == TRUE) {
        
        # only return the resource information g
        select( . ,  -c( projectNumber, portfolioWbs )) } 
        
        else {
        
        # Nest the data into a tibble with project and portfolioWbs info
        nest( .,  resources = -c( projectNumber, portfolioWbs))
        }
        
      } %>% 
      
      
      # if nothing can be found (e,g, because the project or portfolioWbs does not exist) then return an NA
      { if ( nrow(.) == 0 ) { message("No portfolioWbs available for this selection")
                              NA } 
        else .                                               #the previous data
        
      } %>% 
      
      return()
  }
  

  ## add resource information to the dataset
  for (i in 1:2 ) {

    # change the  progress data in the data set
    dataSet[[i]] %<>% 
      unnest( progress ) %>% 
      
      # add the resource information 
      left_join( portfolioWbs(), by = c("projectNumber", "portfolioWbs") ) %>%
      
      # robust way of applying a group on all existing columns except the resource column (i.e. one observation) 
      # then unnest the resource column into these groups 
      group_by(  select(., -resources) ) %>%
      unnest( resources, keep_empty = TRUE ) %>% 
    
      # add grouping level 2 which are the different resource types 
      # and calculate the weight of each resource type within one observation
      group_by( resourceType, .add = TRUE ) %>% 
        summarise( weight = sum( resourceWeight, na.rm = TRUE) ) %>% 
      
      # back to grouping level 1 (the observation)
      # create separate variables for each resource type
      pivot_wider( names_from = resourceType,
                   names_prefix = "weight",
                   values_from = weight,
                   values_fill = 0) %>% 
      
      # add another variable for the case that no resources were estimated for this observation
      select( -any_of("weightNA")) %>% 
      mutate( weightNoResource = ifelse( round( sum( weightM,
                                                     weightS,
                                                     weightL,
                                                     weightP), 0) == 1,
                                         0,1))  %>%
      
      # result as clean, nested tibble
      ungroup() %>% 
      
      # nest the table back to its original format
      nest( progress = -any_of( c( "projectNumber", "projectName", "total" )))
  }

  
# Data Exploration =============================================================


  # company level expenses by period
  baseSet() %>% 
    group_by( period, valueType ) %>%
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>% 
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) 
  
  
  # project level expenses by period
  baseSet() %>% 
    group_by( projectName, period, valueType ) %>%
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>% 
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) +
      facet_wrap ( projectName ~ ., scales = "free" )
  
  
  # portfolioWbs level expanses for all work packages of project 1
  # (Note: not used in project report)
  baseSet() %>% 
    filter( projectNumber == 1) %>% 
    group_by( portfolioWbs, period, valueType ) %>%
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%   
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) +
      facet_wrap ( portfolioWbs ~ ., scales = "free")
  
  
  # portfolioWbs level expanses for an sample illustrative sample of work packages of project 1
  baseSet() %>% 
    filter( projectNumber == 1) %>% 
    group_by( portfolioWbs, period, valueType ) %>%
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%   
    filter( portfolioWbs %in% c(11, 13, 62, 95, 131, 141,73, 902, "CSA", "SWP", "QMR", "WW")) %>% 
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) +
      facet_wrap ( portfolioWbs ~ ., scales = "free" )
  
  
  # total number of work packages across all projects
  baseSet() %>% 
    distinct( projectNumber, portfolioWbs ) %>% 
    nrow()
  
  
  # influence of resources on delta between plan and actual
  # (Note: not used in project report)
  baseSet() %>% 
    group_by( projectNumber ) %>% 
    filter( period == max(period)) %>% 
    ungroup() %>% 
    group_by( projectNumber, portfolioWbs, valueType ) %>%
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%
    ungroup() %>% 
    pivot_wider( names_from = valueType,
                 values_from = periodValue) %>% 
    mutate( delta = abs(pv - ac),
            delta_percent = delta / ac ) %>% 
    rowwise() %>% 
    mutate( resources = list( portfolioWbs( projectNumber, portfolioWbs, TRUE))) %>% 
    unnest( resources) %>% 
    pivot_wider( names_from = resourceType,
                 values_from = resourceWeight) %>% 
    ungroup() %>% 
    select( -'NA') %>% 
    mutate( across( c(M,S,L,P), ~ replace_na(.x, 0)) ) %$%  
    plot( M, delta_percent)
  
  
  
  ## see projects in std Periods
  
  fullDataSet()%>% 
    unnest( progress ) %>% 
    pivot_wider( names_from = valueType, 
                 names_prefix = "value.",
                 values_from = periodValue) %>% 
    group_by( projectNumber ) %>%
    mutate ( totalAc = sum( value.ac, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by( projectNumber, stdPeriod) %>% 
    summarise( value.ac     = sum(value.ac, na.rm = TRUE),
               value.ac.rel = sum(value.ac, na.rm = TRUE) / totalAc,
               
               value.pv     = sum(value.pv, na.rm = TRUE),
               value.pv.rel = sum(value.pv, na.rm = TRUE) / totalAc ) %>% 
    
    pivot_longer( cols = starts_with("value."),
                  names_to = "valueType",
                  values_to = "periodValue") %>% 
    
    filter( valueType %in% c( "value.ac", "value.pv" )) %>% 
    
    ggplot( aes(x=stdPeriod, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f", .x)) +
      facet_wrap ( projectNumber ~ ., scales = "free" )
  
  

  
  ## see projects in std Periods and relative values
  
  fullDataSet()%>% 
    unnest( progress ) %>% 
    pivot_wider( names_from = valueType, 
                 names_prefix = "value.",
                 values_from = periodValue) %>% 
    group_by( projectNumber ) %>%
    mutate ( totalAc = sum( value.ac, na.rm = TRUE)) %>% 
    ungroup() %>% 
    group_by( projectNumber, stdPeriod) %>% 
    summarise( value.ac     = sum(value.ac, na.rm = TRUE),
               value.ac.rel = sum(value.ac, na.rm = TRUE) / totalAc,
               
               value.pv     = sum(value.pv, na.rm = TRUE),
               value.pv.rel = sum(value.pv, na.rm = TRUE) / totalAc ) %>% 
    
    pivot_longer( cols = starts_with("value."),
                  names_to = "valueType",
                  values_to = "periodValue") %>% 
    
    filter( valueType %in% c( "value.ac.rel", "value.pv.rel" )) %>% 
    
    ggplot( aes(x=stdPeriod, y=periodValue, color=valueType)) +
    geom_line() +
    scale_y_continuous( labels= ~ sprintf("%.2f", .x)) +
    facet_wrap ( projectNumber ~ ., scales = "free" )
  
    
  
  
  
  fullDataSet() -> t
  

  
  
  
  

  
  
  dataSet$training %>% 
    select ( projectName, projectNumber, progress) %>% 
    unnest( progress ) %>% 
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
    geom_line() +
    scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) 
    
  
  
  
  # Plot Cash need across all projects 
  projects %>% 
    select ( projectName, projectNumber, periodProgress) %>% 
    unnest( periodProgress ) %>% 
    filter( valueType %in% c("ac", "pv")) %>% 
    group_by( period, valueType ) %>% 
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%
    rowwise() %>% 
    # mutate( periodShare = periodValue /
    #                       switch( valueType,
    #                               "ac" = sum( unnest( projects, total)$ac ),
    #                               "pv" = sum( unnest( projects, total)$pv ),
    #                               NA )) %>% 
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) 

    
  # Plot cash needs by each project
  projects %>% 
    select ( projectName, projectNumber, periodProgress) %>% 
    unnest( periodProgress ) %>% 
    filter( valueType %in% c("ac", "pv")) %>% 
    group_by( projectName, period, valueType ) %>% 
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%
    rowwise() %>% 
    ggplot( aes(x=period, y=periodValue, color=valueType)) +
      geom_line() +
      scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) +
      facet_wrap ( projectName ~ ., scales = "free" )
  
  
  
  # Plot cash need across standardized projects
  
  stdPeriodProgress() %>% 
    #filter( projectNumber == 1) %>% 
    filter( valueType %in% c("ac", "pv")) %>% 
    group_by( stdPeriod, valueType ) %>% 
    summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%
    rowwise() %>% 
    mutate( periodShare = periodValue /
                          switch( valueType,
                                  "ac" = sum( unnest( projects, total)$ac ),
                                  "pv" = sum( unnest( projects, total)$pv ),
                                  NA )) %>% 
    ggplot( aes(x=stdPeriod, y=periodShare, color=valueType)) +
    geom_line() +
    scale_y_continuous( labels= ~ sprintf("%.2f", .x)) 
  

  # projects %>% 
  #   #filter( projectNumber == 1) %>% 
  #   select ( projectName, projectNumber, progress) %>% 
  #   unnest( progress ) %>% 
  #   select( projectNumber, projectName, portfolioWbs, description, contains(".")) %>% 
  #   pivot_longer( cols = contains(".") ,
  #                 values_to = "periodValue") %>% 
  #   separate( col = name,
  #             into = c("valueType", "period"),
  #             convert = TRUE) %>% 
  #   filter( valueType %in% c("ac", "pv")) %>% 
  #   group_by( period, valueType ) %>% 
  #   summarise( periodValue = sum( periodValue, na.rm = TRUE)) %>%
  #   rowwise() %>% 
  #   mutate( periodShare = periodValue /
  #             switch( valueType,
  #                     "ac" = sum( projectTotalValue$aac ),
  #                     "pv" = sum( projectTotalValue$bac ),
  #                     NA )) %>% 
  #   ggplot( aes(x=period, y=periodShare, color=valueType)) +
  #   geom_line( ) 
    
  

  

  # plot pv Vs ac across wbs in one project
  periodProgress %>%
    filter( projectNumber == 1) %>% 
    filter( valueType %in% c("ac", "pv")) %>% 
    group_by( portfolioWbs, period, valueType ) %>% 
    summarise( projectNumber, periodValue = sum( periodValue, na.rm = TRUE)) %>%
    rowwise() %>% 
    mutate( periodShare = periodValue /
              switch( valueType,
                      "ac" = projectTotalValue[projectTotalValue$projectNumber == projectNumber,]$aac,
                      "pv" = projectTotalValue[projectTotalValue$projectNumber == projectNumber,]$bac,
                      NA )) %>% 
    ggplot( aes(x=period, y=periodShare, color=valueType)) +
    geom_line( ) +
    facet_wrap ( portfolioWbs ~ ., scales = "free" )
  
  
  # join the resources with portfolioWbs 
  # and count how many unique resources there are in percent of the data set
  # Insight: There are typically (more than) 50% unique values in a project, therefore the information is limited 
  projects %>% 
    #filter( projectNumber == 1) %>% 
    select ( projectName, projectNumber, progress) %>%  
    unnest( progress ) %>% 
    select( projectNumber, projectName, portfolioWbs) %>% 
    left_join( select(portfolioWbs, -"projectName"), 
               by = c("projectNumber", "portfolioWbs") ) %>% 
    group_by( projectName ) %>% 
    summarise( uniques = n_distinct( resourceName) / n() ,
               duplicates = ( n() - n_distinct(resourceName)) / n() ) %>% 
    pivot_longer( cols = c("uniques", "duplicates"),
                  names_to = "resourceNames",
                  values_to = "share") %>% 
    ggplot( aes( x=projectName , 
                 y=share, 
                 fill=resourceNames)) +
      geom_col() +
      coord_flip()
  
  
  # join the resourcesTypes with portfolioWbs 
  # and count how many unique resources there are in percent of the data set
  # Insight: There are very few unique values in a project, 
  projects %>% 
    #filter( projectNumber == 1) %>% 
    select ( projectName, projectNumber, progress) %>%  
    unnest( progress ) %>% 
    select( projectNumber, projectName, portfolioWbs) %>% 
    left_join( select(portfolioWbs, -"projectName"), 
               by = c("projectNumber", "portfolioWbs") ) %>% 
    group_by( projectName ) %>% 
    summarise( uniques = n_distinct( resourceType) / n() ,
               duplicates = ( n() - n_distinct(resourceType)) / n() ) %>% 
    pivot_longer( cols = c("uniques", "duplicates"),
                  names_to = "resourceTypes",
                  values_to = "share") %>% 
    ggplot( aes( x=projectName , 
                 y=share, 
                 fill=resourceTypes)) +
    geom_col() +
    coord_flip()
    # might want to add a horizontal average across all projects
  

  
  
# Modeling ====================================================================
  
  ## Set up the models
  {
    modelComparison <- 
      
      # define the parameters of each model
      tribble( 
        
        ~modelName,            ~batch,  ~modelmethod,  ~formula,
        
        "just use plan value", "A",     "",            "",
        
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightM",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightS",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightL",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightP",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + stdPeriod",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightNoResource",
        
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightM + weightS + weightL + weightP",
        "autogenerate",        "B",     "glm",         "value.ac ~ value.pv + weightM + weightS + weightL + weightP + weightNoResource"
        
      ) %>% 
      
      # Autogenerate names for some models by combining the method and the factors
      mutate( across( modelName, 
                      ~ ifelse( modelName == "autogenerate", 
                                paste0( modelmethod, " with ", formula ),
                                modelName ))) %>% 
            
      # Add a model ID 
      rowid_to_column( var = "modelId" )
  }
  

  ## function to train a model
  training <- function ( .dataset = dataSet$training,
                         .modelNo = 1, 
                         .trainControl = trainControl( method="none" )) {
  
    
    # only for debugging
    # .dataset <- dataSet$training
    # .modelNo <- 2
    # .trainControl <- trainControl( method="none" )

    
    # pull the batch that belongs to the current model    
    batch <- modelComparison[ .modelNo , ]$batch 
    
    
    # Chose which batch shall be processed
    switch( batch,
            
            
            # model based on the planned value
            "A" = {
              
              # end the function
              return( list( model = "Just use the plan values as forecast" ))

            },
            
            
            # models that are using a simple train function (multiple formulas can use this code)
            "B" = {
              
              
              # getting the method that shall be used
              method <- modelComparison[ .modelNo , ]$modelmethod
              
              
              # getting the formula that shall be used
              formula <- as.formula( modelComparison[ .modelNo , ]$formula ) 
              
              
              # turning the data set into wide format and add a row index
              .dataset %<>%
                select( projectNumber, projectName, progress) %>%
                unnest( progress ) %>%
                select( - any_of( "relPeriodValue" )) %>%                 # remove to omit sideeffects (was added to the code)
                pivot_wider( names_from = valueType,
                             names_prefix = "value.",
                             values_from = periodValue) %>%
                rowid_to_column( var = "rowindex" )
              
              
              # create a forecast set by removing NAs from the input dataset
              forecastset <- .dataset %>% 
                na.omit()
              
              
              # train the model
              model <- train( form = formula,
                              data = forecastset,
                              method= method,
                              trControl = .trainControl )
              
              
              # end the function
              return( list( model = model))

            }
            
    )
  }
  

  ## Function to predict values from a dataset and a model
  prediction <- function( .dataSet = NULL, .model = NULL ) {
    
    # only for debugging
    # .model <- training( .modelNo = 9 )$model
    # .dataSet <- dataSet$training
    
    
    # pull the supplied dataset
    stdPeriodPrediction <- .dataSet %>% 
      select( projectNumber, projectName, progress) %>% 
      unnest( progress ) 
    
    
    # forecasting values -------------------------------------------------------
    
    
    # For the simple plan value model no 1
    if (.model[1] == "Just use the plan values as forecast") {

      
      # use the supplied data and copy plan to forecast
      stdPeriodPrediction %<>% 
        pivot_wider( names_from = valueType,
                     names_prefix = "value.",
                     values_from = periodValue) %>% 
        mutate( value.fc = value.pv ) 
      
    # for all other models  
    } else {
      
      # preserve the original data Set and turn it into wide format and add a row index
      originalDataSet <-
        
        # start with the supplied original data set and unnest the progress column
        .dataSet %>%
        select( projectNumber, projectName, progress) %>% 
        unnest( progress ) %>%
        select( - any_of( "relPeriodValue" )) %>%                   # omit to prevent side-effects (this item was added to the code later)
        
        # turn into wide format
        pivot_wider( names_from = valueType,
                     names_prefix = "value.",
                     values_from = periodValue) %>%
        
        # add a row index
        rowid_to_column( var = "rowindex" )
      
      
      # create a forecastset by removing NAs from the original dataset
      forecastset <- originalDataSet %>% 
        na.omit()
      
      
      # create a tibble with the predicted forcast values and the original rowids  
      forecast <- tibble ( rowindex = forecastset$rowindex,
                           value.fc = unlist( predict( .model,
                                                       forecastset )))
      
      # create stdPeriodPredictions 
      stdPeriodPrediction <- 
        
        # start with the supplied dataset
        originalDataSet %>% 
        
        # add the available predictions (note: some entries don't have a forecast, because e.g. the pv is NA)
        left_join( forecast, by = "rowindex" ) %>% 
        
        # remove rowindex
        select( -rowindex ) 
        
    }
    
    return( stdPeriodPrediction = stdPeriodPrediction) 
  }
  
  
  ## Function to evaluate a model with a dataSet
  evaluation <- function( .prediction = NULL,
                          .overrunPenatly = 0.05, .underrunePenalty = 0.05,
                          .plotSubtitleText = NULL,
                          .noLinePlot = FALSE) {
    
    
    # only for debugging (to generate test values)
    # .prediction <-   prediction( .model = training( .model=2,
    #                                                 .dataset = dataSet$training)$model,
    #                              .dataSet = dataSet$training)
    # .overrunPenatly <- 0.04
    # .underrunePenalty <- 0.04
    # .plotSubtitleText <- NULL
    

    # transform standardPeriods to normal Periods ------------------------------
    predictionPeriods <- 
      
      # start from  the prediction per stdPeriod that were supplied as input 
      .prediction %>%   
      
      # group the data into (normal) periods  (i.e. one group is one normal period ) 
      group_by( select( . , - starts_with("value."),
                        - stdPeriod) ) %>% 
      
      # generate the normal period values (since all data is cummultive, it is just the maximum value from the standard period)
      summarise( stdPeriod = list( stdPeriod ),
                 across( .cols =  starts_with("value."),
                         .fns = ~ max( .x ) )) %>% 
      ungroup()
    
    
    # caluclate the prediction on company level incl. penalties ----------------
    predictionCompany <-
      
      # start with predictions per period
      predictionPeriods %>%
      
      # aggregate to periods on company level
      group_by( period, stdPeriod ) %>%
      summarize( across( .cols =  starts_with("value."),
                         .fns = ~ sum( .x , na.rm = TRUE ))) %>% 
      ungroup() %>% 
      
      # Calculate the delta between forecast and actual and then calculate the corresponding penalty
      mutate(
        # the delta between forecast and actual in a period
        value.delta = value.fc - value.ac,
        
        # the penalty that applies in this specific period
        value.periodSpecificPenalty = 
          case_when( value.delta > 0      ~ abs(value.delta) * .underrunePenalty,
                     value.delta < 0      ~ abs(value.delta) * .overrunPenatly,
                     TRUE                 ~ 0  ),   # all other cases, e.g. if NA or 0
        
        # the cumulative penalties up to this period (remember: all values in the data set are cummulative)
        value.penalty = cumsum( value.periodSpecificPenalty )) %>%
      
      # bring the data back to a tidy format
      pivot_longer( cols = starts_with("value."),
                    names_to = "valueType",
                    names_transform = ~ str_sub( string = .x,
                                                 start = 7),
                    values_to = "periodValue")
    
    
    # calculate the LOSS of the model ------------------------------------------
    modelLoss <- 
      
      # start with company predictions
      predictionCompany %>%
      
      # filter for the last period
      filter( period == max( period, na.rm = TRUE)) %>% 
      
      # grab the penalties
      filter( valueType == "penalty") %$%
      periodValue
    
    
      # end the function and return only LOSS if no lineplot is required
    
    # create a lineplot on company level ---------------------------------------
    # if it has not been deselected during the function call
    if ( .noLinePlot == FALSE ) { 
      
      companyPlot <- 
        
        # start with the prediction on company level
        predictionCompany %>% 
        
        # select the valueTypes that shall be plotted %>% 
        filter( valueType %in% c("ac","fc","penalty")) %>% 
        #filter( valueType %in% c("ac","fc")) %>% 
        
        # draw the plot
        ggplot( aes( x=period, y=periodValue, color= valueType) ) +
        geom_line() +
        scale_y_continuous( labels= ~ sprintf("%.2f Mio", .x/10^6)) +
        labs( 
          title = paste0( "Company Expenses with a LOSS of: ", format( round( modelLoss, 2),
                                                                       big.mark = ",",
                                                                       nsmall = 2,
                                                                       scientific = FALSE )),
          subtitle = if (is.null(.plotSubtitleText)) { waiver() } else .plotSubtitleText, 
          caption = "(c) edxLux 2022" )
      
      # if no company lineplot is required
    } else { companyPlot <- NULL }
      
    
    # return values ------------------------------------------------------------ 
    return( list( predictions = predictionPeriods,
                  companyPlot = companyPlot,
                  loss = modelLoss ))
    
  }
  

  ## First Run of Modeling (Linear Regression) ---------------------------------
  
    # Function that goes through all the modeling steps for a specific model and returns the loss
    modelingLoss <- function ( currentModel ) {
      
      model <- training( .dataset = dataSet$training,
                         .modelNo = currentModel) 
      
      forecast <- prediction( .model = model,
                              .dataSet = dataSet$testing )
      
      evaluation( forecast ) %>% 
        pluck(3) %>% 
        return()
    }
  
  
    # add a column with the first run results to the modelComparison
    modelComparison %<>% 
      mutate( firstRunLOSS = map_dbl( modelComparison$modelId,
                                      ~ modelingLoss( .x )))

    
  
  ## Second Run of Modeling (Regression Tree )  --------------------------------
  ## splitting the data into two ranges and then finding the optimal splitting factor
  
    # define id for head and tail as global factors
    head <- factor(1, levels = 1:2, ordered = TRUE, labels = c("head", "tail"))
    tail <- factor(2, levels = 1:2, ordered = TRUE, labels = c("head", "tail"))
    
    
    # function to split a dataSet into head and tail section at a standard period
    splitAtStdPeriod <- function( .dataSet = NULL,
                                  .breakStdPeriod = NULL) {
  
  
      # only for debugging
      # .dataSet <- dataSet$training
      # .breakStdPeriod <- 25
    
      # split the supplied data
      .dataSet %>%
        
        unnest( progress ) %>% 
        
        # add a factor indicator whether it is head or tail 
        mutate( branch = ifelse( stdPeriod < .breakStdPeriod, head, tail ) %>% 
                           factor( levels = 1:2, ordered = TRUE, labels = c("head", "tail") )) %>% 
        
        # nest the data back but keep the branches separate
        nest( progress = -any_of( c( "projectNumber", "projectName", "total", "branch" ))) %>% 
  
        # nest according to the two branches
        nest( subset = -branch) %>% 
        
        return()
      
    }
    
    
    # function to train a data model separately on both head and tail data
    trainingOnSplitData <- function( .splitDataSet = NULL ,
                                     .modelNo ) {
      
      # only for debugging
       # .splitDataSet <- splitAtStdPeriod( dataSet$training,  25)
       # .modelNo <- 1
      
      # train a first model on the head data
      headModel <- 
        .splitDataSet %>% 
        filter( branch == head) %>%
        select( -branch ) %>% 
        unnest("subset") %>% 
        training( .modelNo )  
      
      
      # train a second model on the tail data
      tailModel <- 
        .splitDataSet %>% 
        filter( branch == tail) %>%
        select( -branch ) %>% 
        unnest("subset") %>% 
        training( .modelNo )  
      
      
      # return both models 
      return( list( headModel = headModel,
                    tailModel = tailModel ) )
      
    }
  
  
    # function to forecast a dataset with a split datamodel for head and tail
    preditionOnSplitData <- function ( .dataSet = NULL,
                                       .model = NULL) {
      
      # only for debugging
       # .dataSet <- splitAtStdPeriod( dataSet$training, 25 )
       # .model <- trainingOnSplitData( splitAtStdPeriod( dataSet$training, 25 ), 1)
  
      headPrediction <-
        .dataSet %>% 
        filter( branch == head ) %>% 
        select( -branch ) %>% 
        unnest( subset ) %>% 
        prediction( .model = .model$headModel$model )
        
      
      tailPrediction <-
        .dataSet %>% 
        filter( branch == tail) %>% 
        select( -branch ) %>% 
        unnest( subset ) %>% 
        prediction( .model = .model$tailModel$model )
  
      
      # return the combined values from head and tail
      bind_rows( tailPrediction,
                 headPrediction ) %>% 
        
        return()
      
    }

  
  
    # Find optimal break point ------------------------------------------------
      
      
      evaluationOfStdBreakPeriod <- function( .dataSet = NULL, 
                                              .breakStdPeriodRange = NULL,
                                              .modelNo = NULL) {
        
        # only for debugging 
        # .dataSet <- dataSet$training
        # .modelNo <- 1
        # .breakStdPeriodRange <- 42
        
        
        # break the data into head and tail set
        splitSet <- splitAtStdPeriod( .dataSet = .dataSet, 
                                      .breakStdPeriod = .breakStdPeriodRange )
        
        
        # train a model on both sections
        splitModel <-  trainingOnSplitData ( .splitDataSet = splitSet ,
                                             .modelNo = .modelNo )
        
        
        # predict the foretasted values
        splitPredicton <- preditionOnSplitData( .dataSet = splitSet,
                                                .model = splitModel )
        
        # evaluate the results
        evaluation( splitPredicton , .noLinePlot = TRUE) %$% 
          
          # return the restul
          return( loss )
        
      }
    
      
      evaluationOfStdBreakPeriod( dataSet$testing,
                                  71,
                                  7)
      
      
      
      breakEvaluation <-
        
        tibble( breakPeriod = c(10:90),
                
                loss = map_dbl ( breakPeriod, 
                                 ~ evaluationOfStdBreakPeriod( dataSet$testing,
                                                               .x,
                                                               7)))
                                 
      
      breakEvaluation %$% 
        plot( breakPeriod, loss)
      
      breakEvaluation %>% 
        slice_min( loss )
      
    # --------------------------------------------------------------------------
      
    # use the best model 
    modelNo <- 7
      
    
    # Run a trial an error over the breakpoints 10:90 
    breakEvaluation <-
      
      tibble( 
              breakPeriod = c(10:90),
              
              loss = map_dbl ( breakPeriod, function( breakPeriod ){
                
                splitTrainingData <- splitAtStdPeriod( .dataSet = dataSet$training,
                                                       .breakStdPeriod = breakPeriod )
                
                 
                trainedModel <- trainingOnSplitData( .splitDataSet = splitTrainingData,
                                                     .modelNo = modelNo)
              
                
                splitTestData <- splitAtStdPeriod( .dataSet = dataSet$testing,
                                                   .breakStdPeriod = breakPeriod )
            
                
                predictionOnTestData <- preditionOnSplitData ( .dataSet = splitTestData,
                                                               .model = trainedModel)
                
                
                evaluation( .prediction = predictionOnTestData ) %>% 
                  pluck("loss")
                
              }))
    
    
    # Plot the results from the trial an error run
    breakEvaluation %$% 
      plot( breakPeriod, loss)
    
    
    # Find the best soultion (i.e. the smallest loss)
    
    bestBreakPoint <- breakEvaluation %>% 
                        slice_min( loss ) 
    
    bestBreakPoint$breakPeriod
    
    
    # Analysing runs
    # -------------------------------------------------------------------------- 
    
    # Defining the best model from the second run
    bestModelSecondRun <- 
      
      # split the Trainiung Data
      splitAtStdPeriod( .dataSet = dataSet$training,
                        .breakStdPeriod = bestBreakPoint$breakPeriod ) %>% 
      
      # train the model
      trainingOnSplitData( .modelNo = modelNo)
      
    
    # Predicting the second run model on company level
    predictionSecondRun <-
      
      # splitt the test data
      splitAtStdPeriod( .dataSet = dataSet$testing,
                      .breakStdPeriod = bestBreakPoint$breakPeriod ) %>% 
      
      # predict the values
      preditionOnSplitData( .model = bestModelSecondRun)
    
    
    # Plotting the second run model on company level
    predictionSecondRun %>% 
      evaluation() %>% 
      pluck( 2 ) 
    
    
    # Showing NA values
    predictionSecondRun %>% 
      select( -starts_with("weight")) %>% 
      slice_head( n=100)
    
      
    predictionSecondRun %>% 
      slice_head( n=10) %>% 
      select( -starts_with("weight")) %>% 
      knitr::kable( tableFormat , caption = "First 10 observations from Second run", 
                    digits = 2, format.args = list( big.mark = ","))

 