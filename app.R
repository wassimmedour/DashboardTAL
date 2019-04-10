library(tidyverse)
library(readxl)
library(data.table)
library(shiny)
library(rpivotTable)
library(shinyjs)

jscode<-"shinyjs.closeWindow=function() {window.close(); } "
ui<-fluidPage(
  
  
 headerPanel(title ="Tableau comparatif d'une cotation d'achat de TASSILI Airlines"),
  sidebarLayout(
    sidebarPanel(
      
      # Use different language and different first day of week
      dateInput("date5", "Date:",
                language = "e",
                weekstart = 1),
      textInput("Numéro de cotation", "Numéro de cotation", "Ex: 066 127"),
      
      fileInput('file1', 'Chosissez une cotation' ,
                accept = c(".xlsx"), width = NULL,
                multiple = T, buttonLabel = "cotation" ),
      useShinyjs(),
      extendShinyjs(text = jscode, functions = c("closeWindow")),
      actionButton("close", "Fermer la fenêtre")
      
      
      
    ), 
    mainPanel(

      
      rpivotTableOutput('contents'))
  )
)
server<-function(input, output ){

  
  verbatimTextOutput("value")
  output$contents <- renderRpivotTable({
    
    if(is.null(input$file1))
      return(NULL)
    
    df<-data.frame(fournissseurs=character(),c=double(),prix=double(),
                   délai=integer(),
                   
                   stringsAsFactors=FALSE)
    
    #y<-read_excel("F1.xlsx")
    #x<-read_excel("F2.xlsx")
    
    #r<-rbind(x,y)
    
    she_list<-lapply(excel_sheets(input$file1$datapath),read_excel,path=input$file1$datapath,na = "NQ")
    r<-rbindlist(she_list)%>%as.data.frame()
    r<-filter(r,!is.na(prix))
    r<-filter(r,!is.na(délai))
  # View(r)
    #str(r)
   r$prix<-as.numeric(r$prix)
    #str(r$prix)
    r$délai<-as.numeric(r$délai)
    r$Quantité<-as.numeric(r$Quantité)
    r$etat<-as.numeric(r$etat)
    r$alternative<-as.numeric(r$alternative)
    #sr<-arrange(r,PN)
  # View(sr)
     #print(sr)
    #o<-r[,-(1:2)]
    upn<-unique(r$PN)
    #View(upn)
  print(r)
    for(n in 1:length(upn)){
      
      #a<-o[r$PN==r$PN[i],]
      
      
      b<-r[r$PN==r$PN[n],]
      a<-b[,-(1:2)]
      a<-as.data.frame(a)
      #View(a) 
    #  print(a)
      #matrice  des preferences normaliser 
      e<-matrix (rep(0, nrow(a)*ncol(a)), nrow(a), ncol(a)) 
      # print(e)
      #matrice  des preference normaliser  et pondérer
      p<-matrix(rep(0,nrow(a)*ncol(a)),nrow(a),ncol(a))  
      #print(p)
      # vecteur point ideal
      v<-matrix(rep(0,1*ncol(a)),1,ncol(a))
      # print(v)
      # vecteur point anti-ideal
      vm<-matrix(rep(0,1*ncol(a)),1,ncol(a))
      # vecteur des distance ideal par rapport aux profils
      D<-matrix(rep(0,1*nrow(a)),1,nrow(a))
      # vecteur des distance anti-ideal par rapport aux profils
      Dm<-matrix(rep(0,1*nrow(a)),1,nrow(a))
      # vecteur des coeffecient de mesure du rapprochement du profile ideal
      c<-matrix(rep(0,1*nrow(a)),1,nrow(a))
      #poid de chaque critere 
      w<-c(0.582,
           0.185,
           0.114,
           0.072,
           0.046
      )
      #sens de la fonction de chaque critere 
      si<-c("min","min","min","max","min")
      
      #Normalisation de la matrice de perference 
      
      for (j in 1:ncol(a)) 
      {
        
        for (i in 1:nrow(a))
        {
          s<-0
          for (k in 1:nrow(a))
          {
            s<-s+a[k,j]*a[k,j]
          }
          
          e[i,j]<-a[i,j]/sqrt(s)
        }
      }
      #Ponderation de la matrice des preference normaliser  selon les criteres 
      for (i in 1:nrow(a))
      {
        
        p[i,]<-e[i,]*w
        
      }
      #determination du point ideal 
      
      # print(p)
      for (i in 1:ncol(a)) 
      { 
        if(si[i]=="max") # a revoir 
        {
          v[i]<-max(p[,i])
        }
        else
        {
          v[i]<-min(p[,i])
        }
      }
      #print(v)
      #determination du point anti-ideal 
      for (i in 1:ncol(a) ) 
      { 
        if(si[i]=="max") # a revoir 
        {
          vm[i]<-min(p[,i])
        }
        else
        {
          vm[i]<-max(p[,i])
        }
      }
      #print(vm)
      #calculer la distance 
      
      
      
      for (i in 1:nrow(a))
      {
        s<-0
        for (j in 1:ncol(a))
        {
          s<-s+((p[i,j]-v[j])*(p[i,j]-v[j]))
        }
        D[i]<-sqrt(s)
      }
      #print(D)
      #calculer la distance des points anti ideal 
      
      
      
      for (i in 1:nrow(a))
      {
        s<-0
        for (j in 1:ncol(a))
        {
          s<-s+((p[i,j]-vm[j])*(p[i,j]-vm[j]))
        }
        Dm[i]<-sqrt(s)
      }
      #print(Dm)
      for(i in 1:nrow(a))
      {
        c[i] <-Dm[i]/(D[i]+Dm[i])
      }
      #print(c)
      c<-t(c)
      #print(c)
      q<-cbind(b[,2],c,a)
      # q<-c(r[,2]%>%unique(),c,a)
    #  View(q)
      #q<-str_replace(q,"1000000000000.00","0")
      #q[,2]<-as.numeric( q[,2])
      q<-arrange(q,desc(c))
      q<-cbind(b[,1],q)
      names(q)<-c("PN","Fournisseur","score","Prix","Délai","Quantité","etat","alternative")
      #View(q)
      df<-rbind(df,q)
      print(n)
      
    } 
  View(df)
    #df<-cbind(sr[,1],df)
   print(df)
    names(df)<-c("PN","Fournisseur","score","Prix","Délai","Quantité","etat","alternative")
   # print(df)
    #df<-as.data.frame(df)
    #df<-filter(df,prix =1000000000000.00 & délai=120000000)
    #df<-unique(df)
    write.csv(df, file="df.csv")
    #group_by(df,PN)%>%ggplot(aes(Fournisseur,score))+geom_col()+facet_grid(~PN)
    rpivotTable(df,rows = c("PN","Fournisseur","score","Prix","Délai","Quantité","etat","alternative"), vals = "score",aggregatorName = "Sample Variance")
    
  })
  observeEvent(input$close, {js$closeWindow() 
    stopApp()})
}
shinyApp(ui,server)

 
 
 