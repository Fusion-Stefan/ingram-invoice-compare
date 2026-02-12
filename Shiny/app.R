
library(shiny)
library(bslib)

# Scripts
setwd("..")
source("./Scripts/write_excel.R")
# import settings
source("./Scripts/settings.R")

ui <- fluidPage(
  titlePanel("Compare Ingram Subscriptions"),
  fileInput("filePrevious",
            label = "Previous"),
  fileInput("fileCurrent",
            label = "Current"),
  fileInput("fileSubscriptions",
            label = "Subscription Info"),
  layout_columns(
    col_widths = c(6, 6),
    actionButton("btnGenerate",
                 label = "Generate"),
    downloadButton("dlFile", 
                   label = "Download")
  )
  
)

server <- function(input, output, session) {
  setwd("..")
  
  
  observeEvent(input$btnGenerate, {
    
    out_path <- paste0("Ingram_invoice_", format.POSIXct(Sys.time(), format = "%d.%m.%Y_%H:%M:%S"), ".xlsx")
    
    excel_file <- generate_xlsx(
      old_path = input$filePrevious$datapath,
      new_path = input$fileCurrent$datapath,
      sub_path = input$fileSubscriptions$datapath
    )
    
    showNotification("File successfully created. Click 'Download' to download.", type = "message")
    
    output$dlFile <- downloadHandler(
      filename = function() {
        out_path
      },
      content =  function(file) {
        wb_save(excel_file, file)
      }
    )
  })
  
  
}

shinyApp(ui, server)