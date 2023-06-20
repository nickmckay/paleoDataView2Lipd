library(lipdR)
library(tidyverse)
library(readxl)
source("convert.R")
pdvFilePath <- "~/Downloads/PDV/"

ax <- list.files(pdvFilePath,pattern = "xls",full.names = TRUE)

#load in all the sheets
allExcel <- map(ax,readPdvExcel)

#create LiPD
L <- pdv2lipd(allExcel[[1]])
