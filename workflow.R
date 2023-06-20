library(lipdR)
library(tidyverse)
library(readxl)
source("convert.R")
pdvFilePath <- "~/Downloads/PDV/"

ax <- list.files(pdvFilePath,pattern = "xls",full.names = TRUE)

#load in all the sheets
allExcel <- map(ax,readPdvExcel)

#create LiPD
D <- map(allExcel,pdv2lipd)

#get all the dataset names
dsns <- map_chr(D,"dataSetName")

if(length(unique(dsns)) > 1){
  stop("There should only be one dataset and multiple are showing up")
}

#add ages to paleoData measurement tables

#loop though and combine tables
if(length(D) >= 1){
  for(n in 2:length(D)){
    D[[1]]$chronData[[1]]$measurementTable[[n]] <- D[[n]]$chronData[[1]]$measurementTable[[1]]
    D[[1]]$paleoData[[1]]$measurementTable[[n]] <- D[[n]]$paleoData[[1]]$measurementTable[[1]]
  }
}

L <- D[[1]]
#make a LiPD object
L <- lipdR:::new_lipd(L)

#write
writeLipd(L)
