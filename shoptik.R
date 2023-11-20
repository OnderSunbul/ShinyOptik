
require(stringr)
require(openxlsx)
require(psych)
require(itemanalysis)
require(crayon)
require(ggplot2)




# Define UI for data upload app ----
ui <- fluidPage(
  
  # App title ----
  titlePanel("ÖLÇME DEĞERLENDİRME MERKEZİ OPTİK PUANLAMA"),
  
  # Sidebar layout with input and output definitions ----
  sidebarLayout(
    
    # Sidebar panel for inputs ----
    sidebarPanel(
      
      # Input: Select a file ----
      fileInput("file1", "Puanlanacak Olan Dosyayı Seçiniz",
                multiple = TRUE,
                accept = c("dat",
                           ".dat")),
      
      # Horizontal line ----
      tags$hr(),
      
      # Input: Checkbox if file has header ----
      checkboxInput("header", "Header", TRUE),
      
      # Input: Select separator ----
      radioButtons("sep", "Separator",
                   choices = c(Comma = ",",
                               Semicolon = ";",
                               Tab = "\t"),
                   selected = ","),
      
      # Input: Select quotes ----
      radioButtons("quote", "Quote",
                   choices = c(None = "",
                               "Double Quote" = '"',
                               "Single Quote" = "'"),
                   selected = '"'),
      
      # Horizontal line ----
      tags$hr(),
      
      # Input: Select number of rows to display ----
      radioButtons("disp", "Display",
                   choices = c(Head = "head",
                               All = "all"),
                   selected = "head")
      
    ),
    
    # Main panel for displaying outputs ----
    mainPanel(
      
      # Output: Data file ----
      tableOutput("contents")
      
    )
    
  )
)

# Define server logic to read selected file ----
server <- function(input, output) {
  
  output$contents <- renderTable({
    
    # input$file1 will be NULL initially. After the user selects
    # and uploads a file, head of that data file by default,
    # or all rows if selected, will be shown.
    
    req(input$file1)
    
    okuma120  <- read.fwf(
      input$file1$datapath,
      header = F,
      stringsAsFactors = F,
      widths = c(11, 12, 12,11,3,1,1,1,1,1,1,1,120),
      fileEncoding = "latin5"
    )
    
    tam_okuma<-okuma120
    
    madde_sayisi <-
      str_count(sub('    *\\    ', '', okuma120[1, 13]))# cevap anahtarında boşluk olan durum için kontrol edilmeli
  
    
    
    tam_okuma <-
      read.fwf(
        input$file1$datapath,
        header = F,
        widths = c(11, 12, 12,11,3,1,1,1,1,1,1,1, rep(1, madde_sayisi)),
        fileEncoding = "latin5",
        stringsAsFactors = F,
      )
    
    
    names(tam_okuma)[1] ="BinNo"
    names(tam_okuma)[2] = "Ad"
    names(tam_okuma)[3] = "Soyad"
    names(tam_okuma)[4] = "ÖgrenciNumarası"
    names(tam_okuma)[5] = "GrupNumarası"
    names(tam_okuma)[6] = "ÖğretimTürü"
    names(tam_okuma)[7] = "Sınıf"
    names(tam_okuma)[8] = "SınavTürü"
    names(tam_okuma)[9] = "SınavDönemi"
    names(tam_okuma)[10] = "Sınavİptali"
    names(tam_okuma)[11] = "SınavaKatılmama"
    names(tam_okuma)[12] = "KitapçıkTürü"
    
    tam_okuma$ÖgrenciNumarası<-as.numeric(tam_okuma$ÖgrenciNumarası)
    
    names(tam_okuma)[13:(dim(tam_okuma)[2])] <-
      paste0("M", 1:madde_sayisi)
    
 df<-tam_okuma     
      
    
    if(input$disp == "head") {
      return(head(df))
    }
    else {
      return(df)
    }
    
  })
  
}

# Create Shiny app ----
shinyApp(ui, server)