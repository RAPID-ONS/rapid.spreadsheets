#' @title Create styles
#'
#' @description A function to create common openxlsx styles.
#'
#' @return List containing styles to use within openxlsx::addStyle() function.
#'
#' @import openxlsx
#' @export

create_styles <- function() {

  s <- list()

  s$text <- openxlsx::createStyle(
    "Arial", 12,
    halign = "left"
  )

  s$bold_text <- openxlsx::createStyle(
    "Arial", 12,
    halign = "left",
    textDecoration = "bold"
  )

  s$heading <- openxlsx::createStyle(
    "Arial", 16,
    halign = "left",
    textDecoration = "bold"
  )

  s$subheadings <- openxlsx::createStyle(
    "Arial", 14,
    halign = "left",
    textDecoration = "bold"
  )

  s$table_header <- openxlsx::createStyle(
    "Arial", 12,
    halign = "left",
    valign = "top",
    textDecoration = "bold",
    wrapText = TRUE
  )

  s$wrap_text <- openxlsx::createStyle(
    "Arial", 12,
    halign = "left",
    wrapText = TRUE
  )

  s$centre <- openxlsx::createStyle(
    "Arial", 12,
    halign = "left",
    valign = "center",
    wrapText = TRUE
  )

  s$border <- openxlsx::createStyle(
    border = c("top", "bottom", "left", "right"),
    borderStyle = "thin"
  )

  s$no_decimal <- openxlsx::createStyle(numFmt = "COMMA")
  s$one_decimal <- openxlsx::createStyle(numFmt = "#,##0.0")
  s$two_decimal <- openxlsx::createStyle(numFmt = "#,##0.00")

  s
}
