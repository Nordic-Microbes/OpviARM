library(shiny)
library(bslib)
library(readxl)
library(DT)

ui <- page_fluid(
  title = "Excel Upload Tool",
  
  card(
    card_header("Upload ARM export"),
    fileInput(
      "excel_file", 
      accept = c(".xlsx"),
      buttonLabel = "Browse...",
      placeholder = "No file selected"
    ) |> 
      tooltip("Upload an Excel file (.xlsx) containing your data")
  ),
  
  conditionalPanel(
    condition = "output.file_uploaded",
    card(
      card_header("Data Preview"),
      tabsetPanel(
        id = "preview_tabs",
        tabPanel("Table 1", DTOutput("table1") |> tooltip("Preview of first sheet data")),
        tabPanel("Table 2", DTOutput("table2") |> tooltip("Preview of second sheet data")),
        tabPanel("Table 3", DTOutput("table3") |> tooltip("Preview of third sheet data")),
        tabPanel("Table 4", DTOutput("table4") |> tooltip("Preview of fourth sheet data"))
      )
    ),
    
    card(
      actionButton(
        "send_to_opvia", 
        "Send to Opvia", 
        class = "btn-primary"
      ) |> 
        tooltip("Upload the formatted data to Opvia")
    )
  )
)

server <- function(input, output, session) {
  # Reactive value to store the excel data
  excel_data <- reactiveVal(list())
  
  # Flag to check if file is uploaded
  output$file_uploaded <- reactive({
    return(!is.null(input$excel_file))
  })
  outputOptions(output, "file_uploaded", suspendWhenHidden = FALSE)
  
  # Read Excel file when uploaded
  observeEvent(input$excel_file, {
    req(input$excel_file)
    
    # Get the file path
    file_path <- input$excel_file$datapath
    
    # Get sheet names
    sheet_names <- readxl::excel_sheets(file_path)
    
    # Read up to 4 sheets (or fewer if less are available)
    data_list <- list()
    for (i in 1:min(4, length(sheet_names))) {
      data_list[[i]] <- readxl::read_excel(file_path, sheet = i)
    }
    
    # Store the data
    excel_data(data_list)
  })
  
  # Render the tables
  output$table1 <- renderDT({
    req(excel_data())
    if (length(excel_data()) >= 1) {
      datatable(excel_data()[[1]], options = list(pageLength = 5))
    }
  })
  
  output$table2 <- renderDT({
    req(excel_data())
    if (length(excel_data()) >= 2) {
      datatable(excel_data()[[2]], options = list(pageLength = 5))
    }
  })
  
  output$table3 <- renderDT({
    req(excel_data())
    if (length(excel_data()) >= 3) {
      datatable(excel_data()[[3]], options = list(pageLength = 5))
    }
  })
  
  output$table4 <- renderDT({
    req(excel_data())
    if (length(excel_data()) >= 4) {
      datatable(excel_data()[[4]], options = list(pageLength = 5))
    }
  })
  
  # Handle Send to Opvia button
  observeEvent(input$send_to_opvia, {
    # Example API calls - replace with actual implementation
    send_data_to_opvia(excel_data())
    
    # Show success notification
    showNotification("Data successfully sent to Opvia", type = "message")
  })
}

# Function to handle API calls to Opvia
send_data_to_opvia <- function(data) {
  # Placeholder for API call implementation
  # Example of what might go here:
  # 
  # for (table_data in data) {
  #   response <- httr::POST(
  #     url = "https://api.opvia.com/data",
  #     body = jsonlite::toJSON(table_data),
  #     httr::content_type_json()
  #   )
  # }
  
  # For now, we'll just print a message
  message("API call to Opvia would happen here")
}

shinyApp(ui, server)
