#test apartment analyzer

library(tidyverse)
library(readxl)
library(purrr)
library(lubridate)
library(xts)
library(timetk)
library(asrsMethods)
library(leaflet)
library(plotly)

source("apt_functions.r")
apt_config=read_config("modera_decatur.xlsx")
ans=apt_analyzer(apt_config)


#str(ans)
analist=ans$analist
bslist=ans$bslist



feelist=ans$feelist
eq=ans$equitystructure
hurdle=eq$hurdlbal

bslist=ans$bslist
teq=bslist$Sponsor_Equity+bslist$Investor_Equity
ploan=bslist$Perm_loan
tidx=index(ploan[ploan>0])
tidx=tidx[-(1:18)]
refi=(2*teq[tidx])-ploan[tidx]
refidf=xtodf(refi,name="Refi_potential")
plot=ggplot(refidf,aes(x=Date,y=Refi_potential))+
  geom_line()+
  ggtitle("Potential proceeds from refinance")
ggplotly(plot)