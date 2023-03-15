library(shiny)
library(tidyverse)
library(officer)
library(glue)
library(waiter)
library(shinyWidgets)
library(sortable)
library(flextable)
library(shinyjs)

source("assets.R")

ui <- fluidPage(
  useWaiter(),
  useShinyjs(),
  tags$head(
    tags$style(HTML(custom_css))
  ),
  fluidRow(
    column(6,
      offset = 3,
      div(
        align = "center",
        tags$h1("Document Parser"),
        tags$hr()
      ),
      br(),
      br()
    )
  ),
  fluidRow(
    column(6,
      offset = 3,
      div(
        align = "center",
        class = "fancy",
          fileInputArea(
            "upload",
            label = "Drop Protocol File Here",
            buttonLabel = ".docx format only!",
            multiple = TRUE,
            accept = "text/plain"
          )
      ),
      hr()
    )
  ),
  br(),
  br(),
  fluidRow(
    column(6,
      offset = 3,
      div(
        align = "center",
        uiOutput("ui.action")
      )
    )
  ),
  br(),
  br(),
  hidden(tabPanel(
    id = "xxx", title = "Hello world", value = "HB",
    tabsetPanel(
      id = "subtabz", type = "pills",
      tabPanel(
        title = "Text", value = "ILPF",
        fluidRow(
          column(6,
            offset = 3,
            div(
              align = "center",
              uiOutput("controls"),
              uiOutput("controls2"),
              br(),
              br(),
              uiOutput("ui.dl")
            )
          )
        ),
        br(),
        br(),
        fluidRow(
          column(6,
            offset = 3,
            div(
              align = "center",
              uiOutput("tables")
            )
          )
        )
      ),
      tabPanel(
        title = "Tables", value = "FS",
        br(),
        fluidRow(
          column(6,
            offset = 3,
            div(
              align = "center",
              uiOutput("table_list")
            )
          )
        ),
        fluidRow(
          column(6,
            offset = 3,
            div(
              align = "center",
              uiOutput("table_preview")
            )
          )
        )
      ),
      tabPanel(
        title = "Misc", value = "huh",
        br(),
        fluidRow(
          column(6,
            offset = 3,
            div(
              align = "center",
              br(),
              p("For protocols with many sub-sections, the drag-and-drop method might not be the best. Below is an alternative."),
              br(),
              uiOutput("alt")
            )
          )
        )
      )
    )
  ))
)


server <- function(input, output) {
  # raw .docx
  filedata <- reactive({
    infile <- input$upload
    if (is.null(infile)) {
      # User has not uploaded a file yet
      return(NULL)
    }
    officer::read_docx(infile$datapath)
  })

  filedata2 <- reactive({
    infile <- input$upload
    if (is.null(infile)) {
      # User has not uploaded a file yet
      return(NULL)
    }
    docxtractr::read_docx(infile$datapath)
  })

  output$ui.action <- renderUI({
    if (is.null(input$upload)) {
      return()
    }
    actionButton("parse", "Parse Protocol File", class = "button-53")
  })


  output$ui.dl <- renderUI({
    if (is.null(input$sections)) {
      return()
    }
    downloadButton("go", "Extract", class = "button-53")
  })


  # because life is easier with reactiveValues()
  values <- reactiveValues()

  # on Parsing
  observeEvent(input$parse, {
    # Actually Parse
    waiter_show(html = div(spin_pixel(), br(), "Parsing Document. This can take a while!"))
    values$summarized <- officer::docx_summary(filedata())
    waiter_hide()
    shinyjs::show("xxx")


    # Create TOC, indicies
    values$toc <- values$summarized %>%
      dplyr::filter(content_type == "paragraph" & tolower(style_name) %in% c("heading 1", "heading 2")) %>%
      mutate(index = 1:n()) %>%
      mutate(subtext = ifelse(tolower(style_name) == "heading 1", text, NA)) %>%
      tidyr::fill(subtext) %>%
      dplyr::select(text, index, style_name, subtext)

    # Create Indexed file
    values$sumdoc_indexed <- values$summarized %>%
      left_join(values$toc) %>%
      tidyr::fill(index)
  })


  output$controls2 <- renderUI({
    req(values$toc)
    bucket_list(
      class = c("default-sortable", "custom-sortable"),
      header = "Protocol Contents",
      group_name = "bucket_list_group",
      orientation = "horizontal",
      add_rank_list(
        text = "Drag sections from here to the right",
        labels = values$toc$text,
        input_id = "rank_list_1"
      ),
      add_rank_list(
        text = "You can reorder sections here too",
        labels = NULL,
        input_id = "sections"
      )
    )
  })

  # Download
  output$go <- downloadHandler(
    # file name
    filename = function() {
      paste("derp-", Sys.Date(), ".docx", sep = "")
    },
    # content
    content = function(con) {
      temp <- values$sumdoc_indexed %>%
        slice(match(input$sections, text)) %>%
        mutate(key = 1:n()) %>%
        select(index, key) %>%
        inner_join(values$sumdoc_indexed)

      temp2 <- temp %>%
        mutate(text = ifelse(style_name == "List Paragraph", paste("*", text), text)) %>%
        filter(!is.na(text) | text != "")

      out <- paste(temp2$text, collapse = "\n\n")

      text <- out %>%
        strsplit("\n\n") %>%
        unlist()

      doc <- officer::read_docx("template.docx")

      for (t in length(text):1) {
        body_add_par(doc, text[[t]])
      }

      officer:::print.rdocx(doc, target = con)
    }
  )

  # table stuff
  output$table_list <- renderUI({
    n <- docxtractr::docx_tbl_count(filedata2())

    selectInput("table_choice",
      "Choose a Table to Preview",
      choices = 1:n
    )
  })

  output$table_preview <- renderUI({
    print(input$table_choice)

    selection <- as.numeric(input$table_choice)

    # read
    temp <- docx_extract_tbl(filedata2(), selection, preserve = TRUE, header = FALSE)

    temp %>%
      flextable::flextable() %>%
      flextable::autofit() %>%
      flextable::fontsize(part = "body", size = 9) %>%
      flextable::font(part = "all", fontname = "Times") %>%
      flextable::delete_part(., part = "header") %>%
      flextable::theme_box() %>%
      flextable::htmltools_value()
  })

  output$alt <- renderUI({
    # Basic validation
    req(values$toc)


    pickerInput(
      inputId = "sections",
      label = "Choose Protocol Sections",
      multiple = TRUE,
      choices = split(values$toc$text, values$toc$subtext)
    )
  })
}



# Run the application
shinyApp(ui = ui, server = server)
