

readPdvExcel <- function(path){
  es <- excel_sheets(path)
  allSheets <- vector(mode = "list",length = length(es))
  for(i in 1:length(es)){
    if(startsWith(es[i],"Meta")){
      allSheets[[i]] <- read_excel(path,sheet = i,col_names = c("Key","Value"))
    }else{
      allSheets[[i]] <- read_excel(path,sheet = i)
    }
  }
  esg <- stringr::str_remove_all(es," ")
  names(allSheets) <- esg

  return(allSheets)
}


pdv2lipd <- function(pdvExcel){
  #metadata
  meta <- pdvExcel$MetaData

  #initialize LiPD file
  L <- list()
  L$dataSetName <- meta$Value[meta$Key == "Core"]
  L$investigators <- meta$Value[grepl(meta$Key,pattern = "importer",ignore.case = TRUE)]
  L$notes <- meta$Value[grepl(meta$Key,pattern = "comment",ignore.case = TRUE)]
  if(is.na(L$notes)){L$notes <- NULL}

  L$archiveType <- "MarineSediment" #THIS IS HARDCODED IN!

  #make up a datasetId
  L$datasetId <- paste0("pdv2l",str_remove_all(L$dataSetName,pattern = "[^A-Za-z0-9]"))

  #make it random?
  # L$datasetId <- paste0(  L$datasetId ,paste(sample(size = 16,c(letters,LETTERS,replace = TRUE)),collapse = ""))

  L$lipdVersion <- 1.3
  L$createdBy <- "https://github.com/nickmckay/paleoDataView2Lipd"




  #initialize geo
  geo <- list()
  geo$latitude <- meta$Value[grepl(meta$Key,pattern = "latitude",ignore.case = TRUE)] %>%
    as.numeric()
  geo$longitude <- meta$Value[grepl(meta$Key,pattern = "longitude",ignore.case = TRUE)] %>%
    as.numeric()
  geo$elevation <- meta$Value[grepl(meta$Key,pattern = "depth",ignore.case = TRUE)] %>%
    as.numeric() * -1

  L$geo <- geo

  #pub
  L$pub <- list(list(citation = meta$Value[grepl(meta$Key,pattern = "paper",ignore.case = TRUE)]))


  #paleoData
  paleo <- pdvExcel$Proxy

  mt <- vector(mode = "list",length = length(paleo))

  for(m in 1:length(mt)){
    mt[[m]] <- makeColumn(paleo[m])
    mt[[m]]$number <- m
    if(startsWith(mt[[m]]$variableName,"d1")){
      mt[[m]]$sensorSpecies <- meta$Value[grepl(meta$Key,pattern = "species",ignore.case = TRUE)]
    }
  }
  an <- map_chr(mt,"variableName")

  names(mt) <- an

  L$paleoData <- vector(mode = "list",length = 1)
  L$paleoData[[1]]$measurementTable <- list()
  L$paleoData[[1]]$measurementTable[[1]] <- mt

  #chronData
  chron <- pdvExcel$Age

  cmt <- vector(mode = "list",length = length(chron))

  for(m in 1:length(cmt)){
    cmt[[m]] <- makeColumn(chron[m])
    cmt[[m]]$number <- m
  }
  an <- map_chr(cmt,"variableName")

  names(cmt) <- an

  L$chronData <- vector(mode = "list",length = 1)
  L$chronData[[1]]$measurementTable <- list()
  L$chronData[[1]]$measurementTable[[1]] <- cmt


return(L)

}

makeColumn <- function(col){
  longName <- names(col)

  if(str_detect(longName,"\\[")){
  variableName <- str_extract(longName,pattern = ".*\\[") %>%
    str_remove_all(pattern = "\\[") %>% str_trim()

  units <- str_extract(longName,pattern = "\\[.*\\]") %>%
    str_remove_all(pattern = "\\[") %>%
    str_remove_all(pattern = "\\]") %>%
    str_trim()
  }else{
    variableName <- longName
    units <- "unitless"
  }

  values = col[[1]]

  out <- list(variableName = variableName,
              units = units,
              values = values,
              longName = longName,
              TSid = createTSid("pdv2l"))

  return(out)

}
