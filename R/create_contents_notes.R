#' @title Adds contents or notes worksheet to workbook.
#'
#' @description This function adds a contents or notes worksheet to an openxlsx
#' workbook.
#'
#' @param wb An openxlsx workbook object.
#' @param df Data frame containing two or three columns e.g.: worksheet names
#' (containing tab names), worksheet descriptions and an optional third column
#' containing hyperlink text.
#' @param tab_name Worksheet name as string, default is "Contents".
#' @param heading Heading or title name as string, default is
#' "Table of contents".
#' @param contents_links If TRUE then internal hyperlinks are created for
#' worksheet names specified in the first column of the df. Default is TRUE.
#' @param additional_text Vector or data frame containing additional rows of
#' text to place above the table, default is empty.
#' @param hyperlinks Data frame with two columns:
#' \itemize{
#'   \item first column as a numeric vector listing rows in df with hyperlinks,
#'   \item second column as character vector with url or email addresses.
#' }
#' The df has to have three columns as hyperlinks will be added to the third
#' column. Optional argument.
#' @param column_width Width of Excel columns, defaults are: 20, 80 and 15.
#' @param num_tables Number of tables in a worksheet, default is 1.
#'
#' @return Adds a new contents/notes worksheet to existing openxlsx workbook.
#'
#' @import openxlsx
#'
#' @export

create_contents_notes <- function(wb,
                                  df,
                                  tab_name = "Contents",
                                  heading = "Table of contents",
                                  contents_links = TRUE,
                                  additional_text = c(),
                                  hyperlinks = NA,
                                  column_width = c(20, 80, 15),
                                  num_tables = 1) {

  add_format_worksheet(
    wb, ncol(df), nrow(df),
    tab_name,
    heading,
    length(additional_text),
    num_tables
  )
  heading_length <- length(heading)
  table_start_row <- heading_length + length(additional_text) + 2
  table_end_row <- table_start_row + nrow(df)

  openxlsx::writeData(
    wb, tab_name, additional_text,
    startCol = 1,
    startRow = heading_length + 2
  )
  openxlsx::writeDataTable(
    wb, tab_name, df,
    startCol = 1,
    startRow = table_start_row,
    tableName = tab_name,
    tableStyle = "None",
    withFilter = FALSE
  )

  if (ncol(df) == 1) {
    openxlsx::setColWidths(wb, tab_name, 1:2, column_width[1])
  } else if (ncol(df) == 2) {
    openxlsx::setColWidths(wb, tab_name, 1:2, column_width[1:2])
  } else {
    openxlsx::setColWidths(wb, tab_name, 1:3, column_width)
  }

  s <- create_styles()
  openxlsx::addStyle(
    wb, tab_name, s$centre,
    rows = table_start_row:table_end_row,
    cols = 1:ncol(df),
    gridExpand = TRUE,
    stack = TRUE
  )

  if (contents_links == TRUE) {
    for (y in 1:nrow(df)) {
      openxlsx::writeFormula(
        wb, tab_name,
        startRow = table_start_row + y,
        x = openxlsx::makeHyperlinkString(
          sheet = df[y, 1],
          row = 1,
          col = 1,
          text = df[y, 1]
          )
        )
    }
  }

  if (is.data.frame(hyperlinks)) {
    for (x in seq_len(nrow(hyperlinks))) {
      hyperlink_name <- hyperlinks[x, 2]
      names(hyperlink_name) <- df[hyperlinks[x, 1], 3]
      class(hyperlink_name) <- "hyperlink"

      openxlsx::writeData(
        wb, tab_name,
        startCol = 3,
        startRow = table_start_row + hyperlinks[x, 1],
        colNames = FALSE,
        x = hyperlink_name
      )
    }
  }
}
