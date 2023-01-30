library(shiny)
# library(cli)

ui <- fluidPage(
  titlePanel("Miles Per Gallon"),
  sidebarLayout(
    sidebarPanel(
      selectInput("variable", "Variable:",
                  c("Cylinders" = "cyl",
                    "Transmission" = "am",
                    "Gears" = "gear")),
      checkboxInput("outliers", "Show outliers", TRUE)
                ),
    mainPanel(
      h3(textOutput("caption")),
      plotOutput("mpgPlot")
    )
  )
)

# 数据预处理 ----
# 将"am"变量转换成拥有更好标签的因子变量 -- 由于这个变量不依赖于任何输入，
# 我们可以在一开始就对它进行处理，这样就可以在整个app中使用处理后的变量了
mpgData <- mtcars[1:4,]
mpgData$am <- factor(mpgData$am, labels = c("Automatic", "Manual"))
 
# 定义用来展示与mpg相关联的多个变量的server逻辑 ----
server <- function(input, output){
    formulaText <- reactive({
      paste("mpg ~", input$variable)
  })
    output$caption <- renderText({
      formulaText()
  })
    output$mpgPlot <- renderPlot({
      boxplot(as.formula(formulaText()),
            data = mpgData,
            outline = input$outliers,
            col = "#75AADB", pch = 19)
  })
}

shinyApp(ui, server)



