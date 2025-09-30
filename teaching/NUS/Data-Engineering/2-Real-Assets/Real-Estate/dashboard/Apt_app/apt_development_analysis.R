#
#Apartment analyzer app by Karl Polen
#
# This app uses functions in the asrsMethods package
# which can be installed by uncommenting the following code
#devtools::install_github("karlp-asrs/asrsMethods/asrsMethods")
#



library(shiny)
library(shinydashboard)
library(tidyverse)
library(readxl)
library(purrr)
library(lubridate)
library(xts)
library(timetk)
library(asrsMethods)
library(knitr)
library(DT)
library(leaflet)
library(plotly)
source("apt_functions.R")


ui <- dashboardPage(
    
    dashboardHeader(title = textOutput("projname")),
    dashboardSidebar(
        selectInput(inputId="file",
                    label="Select file",
                    choices=c("project_bulldog.xlsx"),
                    selected="project_bulldog.xlsx"),
        sliderInput(inputId="rent_sens",
                    label="Rent Sensitivity %",
                    min=-.1,
                    max=.1,
                    value=0,
                    step=.01),
        sliderInput(inputId="exp_sens",
                    label="Expense Sensitivity %",
                    min=-.1,
                    max=.1,
                    value=0,
                    step=.01),
        sliderInput(inputId="overrun",
                    label="Construction cost overrun %",
                    min=-.1,
                    max=.1,
                    value=0,
                    step=.01),
        sliderInput(inputId="predev_delay",
                    label="Predevelopment delay months",
                    min=0,
                    max=12,
                    value=0,
                    step=1),
        sliderInput(inputId="constr_delay",
                    label="Construction delay months",
                    min=0,
                    max=12,
                    value=0,
                    step=1),
        sliderInput(inputId="leaseup_delay",
                    label="Leaseup delay months",
                    min=0,
                    max=12,
                    value=0,
                    step=1)
    ),
    dashboardBody(
        tabBox(
            width=12,
            tabPanel("Map",leafletOutput("map")), 
            tabPanel("Timeline",plotlyOutput("plot_timeline")),
            tabPanel("IRR Analysis",
                     plotlyOutput("plot_irrs"),
                     plotlyOutput("plot_3yr")
            ),
            tabPanel("Cash Flow/NAV",
                     plotlyOutput("plot_cumcf"),
                     plotlyOutput("plot_unlcumcf")
                     ),
            tabPanel("Promote",
                     plotlyOutput("plot_seq"),
                     #plotlyOutput("plot_seqlog"),
                     plotlyOutput("plot_seqpct")),
            tabPanel("Fee Analysis",
                     plotlyOutput("plot_fee"),
                     plotlyOutput("plot_feepct")),
            tabPanel("Invstor Sensitivity",
                     tabBox(width=12,
                        tabPanel("ConstrDelay",
                            plotlyOutput("plot_invsens_constr_delay")),
                        tabPanel("LeaseupDelay",
                            plotlyOutput("plot_invsens_leaseup_delay")),
                        tabPanel("PredevDelay",
                            plotlyOutput("plot_invsens_predev_delay")),
                        tabPanel("ConstrOverrun",
                                 plotlyOutput("plot_invsens_constr_overrun")),
                        tabPanel("RentSens",
                                 plotlyOutput("plot_invsens_rentsens"))
                     )
            ),
            tabPanel("Sponsor Sensitivity",
                     tabBox(width=12,
                            tabPanel("ConstrDelay",
                                     plotlyOutput("plot_sponsens_constr_delay")),
                            tabPanel("LeaseupDelay",
                                     plotlyOutput("plot_sponsens_leaseup_delay")),
                            tabPanel("PredevDelay",
                                     plotlyOutput("plot_sponsens_predev_delay")),
                            tabPanel("ConstrOverrun",
                                     plotlyOutput("plot_sponsens_constr_overrun")),
                            tabPanel("RentSens",
                                     plotlyOutput("plot_sponsens_rentsens"))
                     )
            ),
            tabPanel("Financials",
                     tabBox(width=12,
                        tabPanel("Income Statement",DT::DTOutput("is")),
                        tabPanel("Balance Sheet",
                            DT::DTOutput("bs"),
                            plotlyOutput("plot_refi")),
                        tabPanel("Cash Flow",DT::DTOutput("cf"))
                    )
            )
            )
    )
)


server <- function(input, output,session) {
    #  apt_config <- reactiveFileReader(
    #  intervalMillis = 1000,
    #      session = session,
    #      filePath = input$file,
    #      readFunc = read_config
    # )
    ans=reactive({
        apt_config=read_config(input$file)
        apt_analyzer(apt_config,
                     rent_sensitivity=input$rent_sens,
                     expense_sensitivity=input$exp_sens,
                     constr_overrun=input$overrun,
                     predev_delay=input$predev_delay,
                     constr_delay=input$constr_delay,
                     leaseup_delay=input$leaseup_delay)
    })
    output$projname=renderText({
        apt_config=ans()$apt_config
        specsid=apt_config[["Identification"]]
        filter(specsid,name=="ProjectName")$descr
    })
    output$is=DT::renderDT({
        is_list=ans()$islist
        ltodfyear(is_list,tname="Net_Income",roundthou=TRUE)
    })
    output$bs=DT::renderDT({
        bs_list=ans()$bslist
        ltodfyear(bs_list,total=FALSE,roundthou=TRUE,isbs=TRUE)
    })
    output$cf=DT::renderDT({
        cf_list=ans()$sumcflist
        ltodfyear(cf_list,total=TRUE,tname="Net_CF",roundthou=TRUE)
    })
    output$plot_irrs=renderPlotly({
        analist=ans()$analist
        x=evolution_irr(analist$Investor_CF,analist$Inv_FV)
        y=evolution_irr(analist$Unlev_CF,analist$Unl_FV)
        z=evolution_irr(analist$Lev_CF,analist$Lev_FV)
        irrlist=list(z,x,y)
        names(irrlist)=c("Project_IRR","Investor_IRR","Unlevered_IRR")
        irrdf=ltodf(irrlist)
        irrdf=gather(irrdf,key="type",value="IRR",-Date)
        plot=ggplot(irrdf,aes(x=Date,y=IRR,col=type))+
            geom_line()+
            ggtitle("Inception IRRs")
        ggplotly(plot)
    })
    output$plot_3yr=renderPlotly({
        analist=ans()$analist
        x=rolling_irr(analist$Investor_CF,analist$Inv_FV,36)
        y=rolling_irr(analist$Unlev_CF,analist$Unl_FV,36)
        z=rolling_irr(analist$Lev_CF,analist$Lev_FV,36)
        irrlist=list(z,x,y)
        names(irrlist)=c("Project_IRR","Investor_IRR","Unlevered_IRR")
        irrdf=ltodf(irrlist)
        irrdf=gather(irrdf,key="type",value="IRR",-Date)
        plot=ggplot(irrdf,aes(x=Date,y=IRR,col=type))+
            geom_line()+
            ggtitle("Rolling 3 Yr IRRs")
        ggplotly(plot)
    })
    output$plot_cumcf=renderPlotly({
        analist=ans()$analist
        lcumcf=list(-1*cumsum(analist$Investor_CF),analist$Inv_FV)
        names(lcumcf)=c("Invstr_Cum_CF","Invstr_NAV")
        cumcfdf=ltodf(lcumcf,wtotal=FALSE,donona=FALSE)
        cumcfdfg=gather(cumcfdf,key="type",value="Dollars",-Date)
        cumcfdfg=filter(cumcfdfg,!is.na(Dollars))
        plot=ggplot(cumcfdfg,aes(x=Date,y=Dollars,col=type))+
            geom_line()+
            ggtitle("Investor Cumulative CF and NAV")
        ggplotly(plot)
    })
    output$plot_unlcumcf=renderPlotly({
        analist=ans()$analist
        lcumcf=list(-1*cumsum(analist$Unlev_CF),analist$Unl_FV)
        names(lcumcf)=c("Cum_CF","Gross_NAV")
        cumcfdf=ltodf(lcumcf,wtotal=FALSE,donona=FALSE)
        cumcfdfg=gather(cumcfdf,key="type",value="Dollars",-Date)
        cumcfdfg=filter(cumcfdfg,!is.na(Dollars))
        plot=ggplot(cumcfdfg,aes(x=Date,y=Dollars,col=type))+
            geom_line()+
            ggtitle("Unlevered Cumulative CF and NAV")
        ggplotly(plot)
    })
    
    output$plot_fee=renderPlotly({
        feelist=ans()$feelist
        feedf=makefeedf(feelist)
        plotdat=feedf %>%
            select(Date,Total_Fee,Total_Excess) %>%
            gather(key="type",value="Dollars",-Date) 
        plot=ggplot(plotdat,aes(x=Date,y=Dollars,col=type))+
            geom_line()+
            ggtitle("Fees and cumulative profit above 1st tier hurdle")
        ggplotly(plot)
        
    })
    output$plot_feepct=renderPlotly({
        feelist=ans()$feelist
        feedf=makefeedf(feelist)
        plotdat=feedf %>%
            select(Date,Fee_pct_of_Excess)
        plot=ggplot(plotdat,aes(x=Date,y=Fee_pct_of_Excess)) +
            geom_line()+
            ggtitle("Fees as % of profit above 1st tier hurdle") 
        ggplotly(plot)
    })
    output$plot_seq=renderPlotly({
        bslist=ans()$bslist
        sponsor_eq=bslist$Sponsor_Equity
        seqdf=xtodf(sponsor_eq,name="Sponsor_Equity")
        plot=ggplot(seqdf,aes(x=Date,y=Sponsor_Equity))+
            geom_line()+
            ggtitle("Sponsor Promote")
        ggplotly(plot)
    })
    output$plot_seqlog=renderPlotly({
        bslist=ans()$bslist
        sponsor_eq=bslist$Sponsor_Equity
        seqdf=xtodf(sponsor_eq,name="Sponsor_Equity")
        plot=ggplot(seqdf,aes(x=Date,y=Sponsor_Equity))+
            geom_line()+
            ggtitle("Sponsor Promote Log Scale")+
            scale_y_log10()
        ggplotly(plot)
    })
    output$plot_seqpct=renderPlotly({
        bslist=ans()$bslist
        sponsor_eq=bslist$Sponsor_Equity
        s_eq_roll12=rollapply(sponsor_eq,width=12,FUN=mean)
        s_eq_r12_change=diff(s_eq_roll12,lag=12)/s_eq_roll12
        seqdf=xtodf(s_eq_r12_change,name="Sponsor_Eq_pct_change")
        plot=ggplot(seqdf,aes(x=Date,y=Sponsor_Eq_pct_change))+
            geom_line()+
            ggtitle("Sponsor Promote Growth Rate")
            #coord_cartesian(ylim=c(-.05,1))
        ggplotly(plot)
    })
    output$plot_refi=renderPlotly({
        bslist=ans()$bslist
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
    })
    output$plot_timeline=renderPlotly({
        timeline=ans()$timeline
        idx=which(timeline$Risk_Ctgry=="Stable")[12]
        plot=ggplot(timeline[1:idx,],aes(x=Date,y=Risk_Ctgry))+
            geom_line(aes(col=Risk_Ctgry),size=15)+
            theme(legend.position="none")+
            ggtitle("Timeline")+
            ylab("")+
            xlab("")
        ggplotly(plot)
    })
 
    ## investor sensitiviy plots
       output$plot_invsens_constr_delay=renderPlotly({
        sensvarm=0:12
        irrvec=vector()
        for (i in 1:length(sensvarm)) {
        ans2=apt_analyzer(ans()$apt_config,
                     rent_sensitivity=input$rent_sens,
                     expense_sensitivity=input$exp_sens,
                     constr_overrun=input$overrun,
                     predev_delay=input$predev_delay,
                     constr_delay=sensvarm[i],
                     leaseup_delay=input$leaseup_delay)
        irrvec[i]=irr10yr(ans2$analist$Investor_CF,ans2$analist$Inv_FV)
        }
        irrdf=data.frame(Delay_mths=sensvarm,Invstr_10_yr_irr=irrvec)
        plot=ggplot(irrdf,aes(x=Delay_mths,y=Invstr_10_yr_irr))+
            geom_line()+
            ggtitle("Impact of construction delay on investor returns")+
            coord_cartesian(ylim=c(0,max(irrvec)))
        ggplotly(plot)
        
    })
       
           
    output$plot_invsens_leaseup_delay=renderPlotly({
        sensvarm=0:12
        irrvec=vector()
        for (i in 1:length(sensvarm)) {
            ans2=apt_analyzer(ans()$apt_config,
                              rent_sensitivity=input$rent_sens,
                              expense_sensitivity=input$exp_sens,
                              constr_overrun=input$overrun,
                              predev_delay=input$predev_delay,
                              constr_delay=input$constr_delay,
                              leaseup_delay=sensvarm[i])
            irrvec[i]=irr10yr(ans2$analist$Investor_CF,ans2$analist$Inv_FV)
        }
        irrdf=data.frame(Delay_mths=sensvarm,Invstr_10_yr_irr=irrvec)
        plot=ggplot(irrdf,aes(x=Delay_mths,y=Invstr_10_yr_irr))+
            geom_line()+
            ggtitle("Impact of leaseup delay on investor returns")+
            coord_cartesian(ylim=c(0,max(irrvec)))
        ggplotly(plot)
        
    })
    output$plot_invsens_predev_delay=renderPlotly({
        sensvarm=0:12
        irrvec=vector()
        for (i in 1:length(sensvarm)) {
            ans2=apt_analyzer(ans()$apt_config,
                              rent_sensitivity=input$rent_sens,
                              expense_sensitivity=input$exp_sens,
                              constr_overrun=input$overrun,
                              predev_delay=sensvarm[i],
                              constr_delay=input$constr_delay,
                              leaseup_delay=input$leaseup_delay)
            irrvec[i]=irr10yr(ans2$analist$Investor_CF,ans2$analist$Inv_FV)
        }
        irrdf=data.frame(Delay_mths=sensvarm,Invstr_10_yr_irr=irrvec)
        plot=ggplot(irrdf,aes(x=Delay_mths,y=Invstr_10_yr_irr))+
            geom_line()+
            ggtitle("Impact of Predevelopment delay on investor returns")+
            coord_cartesian(ylim=c(0,max(irrvec)))
        ggplotly(plot)
        
    })
    output$plot_invsens_rentsens=renderPlotly({
        sensvarp=seq(-.1,.1,.02)
        irrvec=vector()
        for (i in 1:length(sensvarp)) {
            ans2=apt_analyzer(ans()$apt_config,
                              rent_sensitivity=sensvarp[i],
                              expense_sensitivity=input$exp_sens,
                              constr_overrun=input$overrun,
                              predev_delay=sensvarm[i],
                              constr_delay=input$constr_delay,
                              leaseup_delay=input$leaseup_delay)
            irrvec[i]=irr10yr(ans2$analist$Investor_CF,ans2$analist$Inv_FV)
        }
        irrdf=data.frame(Pct_change=sensvarp,Invstr_10_yr_irr=irrvec)
        plot=ggplot(irrdf,aes(x=Pct_change,y=Invstr_10_yr_irr))+
            geom_line()+
            ggtitle("Impact of rent changes on investor returns")+
            coord_cartesian(ylim=c(0,max(irrvec)))
        ggplotly(plot)
        
    })
    output$plot_invsens_constr_overrun=renderPlotly({
        sensvarp=seq(0, .2, .02)
        irrvec=vector()
        for (i in 1:length(sensvarp)) {
            ans2=apt_analyzer(ans()$apt_config,
                              rent_sensitivity=input$rent_sens,
                              expense_sensitivity=input$exp_sens,
                              constr_overrun=sensvarp[i],
                              predev_delay=sensvarm[i],
                              constr_delay=input$constr_delay,
                              leaseup_delay=input$leaseup_delay)
            irrvec[i]=irr10yr(ans2$analist$Investor_CF,ans2$analist$Inv_FV)
        }
        irrdf=data.frame(Pct_change=sensvarp,Invstr_10_yr_irr=irrvec)
        plot=ggplot(irrdf,aes(x=Pct_change,y=Invstr_10_yr_irr))+
            geom_line()+
            ggtitle("Impact of construction overrun on investor returns")+
            coord_cartesian(ylim=c(0,max(irrvec)))
        ggplotly(plot)
        })
        
    ## sponsor sensitivity plots
        output$plot_sponsens_constr_delay=renderPlotly({
            sensvarm=0:12
            provec=vector()
            for (i in 1:length(sensvarm)) {
                ans2=apt_analyzer(ans()$apt_config,
                                  rent_sensitivity=input$rent_sens,
                                  expense_sensitivity=input$exp_sens,
                                  constr_overrun=input$overrun,
                                  predev_delay=input$predev_delay,
                                  constr_delay=sensvarm[i],
                                  leaseup_delay=input$leaseup_delay)
                provec[i]=ans2$bslist$Sponsor_Equity[120]
            }
            prodf=data.frame(Delay_mths=sensvarm,Promote_yr_10=provec)
            plot=ggplot(prodf,aes(x=Delay_mths,y=Promote_yr_10))+
                geom_line()+
                ggtitle("Impact of construction delay on sponsor promote at year 10")+
                coord_cartesian(ylim=c(0,max(provec)))
            ggplotly(plot)
            
        })
        output$plot_sponsens_leaseup_delay=renderPlotly({
            sensvarm=0:12
            provec=vector()
            for (i in 1:length(sensvarm)) {
                ans2=apt_analyzer(ans()$apt_config,
                                  rent_sensitivity=input$rent_sens,
                                  expense_sensitivity=input$exp_sens,
                                  constr_overrun=input$overrun,
                                  predev_delay=input$predev_delay,
                                  constr_delay=input$constr_delay,
                                  leaseup_delay=sensvarm[i])
                provec[i]=ans2$bslist$Sponsor_Equity[120]
            }
            prodf=data.frame(Delay_mths=sensvarm,Promote_yr_10=provec)
            plot=ggplot(prodf,aes(x=Delay_mths,y=Promote_yr_10))+
                geom_line()+
                ggtitle("Impact of leaseup delay on sponsor promote at year 10")+
                coord_cartesian(ylim=c(0,max(provec)))
            ggplotly(plot)
            
        })
        output$plot_sponsens_predev_delay=renderPlotly({
            sensvarm=0:12
            provec=vector()
            for (i in 1:length(sensvarm)) {
                ans2=apt_analyzer(ans()$apt_config,
                                  rent_sensitivity=input$rent_sens,
                                  expense_sensitivity=input$exp_sens,
                                  constr_overrun=input$overrun,
                                  predev_delay=sensvarm[i],
                                  constr_delay=input$constr_delay,
                                  leaseup_delay=input$leaseup_delay)
                provec[i]=ans2$bslist$Sponsor_Equity[120]
            }
            prodf=data.frame(Delay_mths=sensvarm,Promote_yr_10=provec)
            plot=ggplot(prodf,aes(x=Delay_mths,y=Promote_yr_10))+
                geom_line()+
                ggtitle("Impact of Predevelopment delay on sponsor promote at year 10")+
                coord_cartesian(ylim=c(0,max(provec)))
            ggplotly(plot)
            
        })
        output$plot_sponsens_rentsens=renderPlotly({
            sensvarp=seq(-.1,.1,.02)
            provec=vector()
            for (i in 1:length(sensvarp)) {
                ans2=apt_analyzer(ans()$apt_config,
                                  rent_sensitivity=sensvarp[i],
                                  expense_sensitivity=input$exp_sens,
                                  constr_overrun=input$overrun,
                                  predev_delay=sensvarm[i],
                                  constr_delay=input$constr_delay,
                                  leaseup_delay=input$leaseup_delay)
                provec[i]=ans2$bslist$Sponsor_Equity[120]
            }
            prodf=data.frame(Pct_change=sensvarp,Promote_yr_10=provec)
            plot=ggplot(prodf,aes(x=Pct_change,y=Promote_yr_10))+
                geom_line()+
                ggtitle("Impact of rent changes on sponsor promote at year 10")+
                coord_cartesian(ylim=c(0,max(provec)))
            ggplotly(plot)
            
        })
        
        output$plot_sponsens_constr_overrun=renderPlotly({
        sensvarp=seq(0, .2, .02)
        provec=vector()
        for (i in 1:length(sensvarp)) {
            ans2=apt_analyzer(ans()$apt_config,
                              rent_sensitivity=input$rent_sens,
                              expense_sensitivity=input$exp_sens,
                              constr_overrun=sensvarp[i],
                              predev_delay=sensvarm[i],
                              constr_delay=input$constr_delay,
                              leaseup_delay=input$leaseup_delay)
            provec[i]=ans2$bslist$Sponsor_Equity[120]
        }
        prodf=data.frame(Pct_change=sensvarp,Promote_yr_10=provec)
        plot=ggplot(prodf,aes(x=Pct_change,y=Promote_yr_10))+
            geom_line()+
            ggtitle("Impact of construction overrun on sponsor promote at year 10")+
            coord_cartesian(ylim=c(0,max(provec)))
        ggplotly(plot)
        })    
        
        
        
        
        
    output$map=renderLeaflet({
        apt_config=ans()$apt_config
        specsid=apt_config$Identification
        lati=filter(specsid,name=="Latitude")$num
        long=filter(specsid,name=="Longitude")$num
        leaflet() %>%
            addTiles() %>%
            addMarkers(
                lat=lati,
                lng=long
            )
    })
}

# Run the application 
shinyApp(ui = ui, server = server)
