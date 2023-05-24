#' @title Adds cover sheet worksheet to workbook.
#'
#' @description This function adds a Cover sheet worksheet to an openxlsx
#' workbook.
#'
#' @param wb An openxlsx workbook object.
#' @param text_df A data frame containing the cover sheet text in a single
#' column.
#' @param tab_name Worksheet name as character string, default is
#' "Cover_sheet".
#' @param subheadings Numeric vector specifying row numbers for subheadings.
#' These will be in bold and have increased row height. Optional argument.
#' @param subheadings_row_height Subheadings row height as a numeric value.
#' Default is 30.
#' @param hyperlinks Data frame with two columns:
#' \itemize{
#'   \item first as numeric vector listing rows which should have links,
#'   \item second as character vector with url or email addresses.
#'   }
#' Optional argument.
#' @param rows_to_bold A numeric vector listing rows to bold. Optional
#' argument.
#' @param replacements Data frame with two columns:
#' \itemize{
#'   \item "replace" specifying strings to replace
#'   \item "with" specifying what to replace the strings with
#' }
#' Optional argument.
#' @param column_width Column width of the first column as a numeric value.
#' Default is 130.
#'
#' @examples
#' \dontrun{
#' wb <- openxlsx::createWorkbook()
#'
#' replacements <- data.frame(
#'  replace = c('!release_date!', '!copyright!'),
#'  with = c('23 August 2022', '2022')
#'  )
#'
#' cover_sheet_text <- read.csv(
#'   "data/Cover_sheet.txt",
#'   header = FALSE,
#'   sep = "\n", quote = "",
#'   fileEncoding = "UTF-8-BOM")
#'
#' hyperlinks <- data.frame(
#'   rows = c(19, 31),
#'   links = c('https://www.ons.gov.uk', 'mailto:Health.Data@ons.gov.uk'))
#'
#' create_cover_sheet(wb,
#'   text_df = cover_sheet_text,
#'   replacements = replacements,
#'   subheadings = c(5, 9, 16),
#'   rows_to_bold = c(2:4, 45),
#'   hyperlinks = hyperlinks)
#' }
#'
#'
#' @return Adds a cover sheet worksheet to existing openxlsx workbook.
#'
#' @import openxlsx
#'
#' @export

create_cover_sheet <- function(wb,
                               text_df,
                               tab_name = "Cover_sheet",
                               subheadings = NA,
                               hyperlinks = NA,
                               rows_to_bold = NA,
                               replacements = NA,
                               subheadings_row_height = 30,
                               column_width = 130) {

  if (is.data.frame(replacements)) {
    text_df <- replace_text_for_spreadsheets(replacements, text_df)
  }
  nrows <- nrow(text_df)

  openxlsx::addWorksheet(wb, tab_name, gridLines = FALSE)
  openxlsx::modifyBaseFont(wb, fontSize = 12, fontName = "Arial")

  openxlsx::writeData(
    wb, tab_name,
    startCol = 1,
    startRow = 1,
    colNames = FALSE,
    x = text_df
  )

  openxlsx::setColWidths(wb, tab_name, 1, column_width)
  openxlsx::setRowHeights(wb, tab_name, subheadings, subheadings_row_height)

  s <- create_styles()
  openxlsx::addStyle(
    wb, tab_name, s$wrap_text,
    rows = 1:nrows,
    cols = 1,
    gridExpand = TRUE,
    stack = TRUE
  )
  openxlsx::addStyle(
    wb, tab_name, s$heading,
    rows = 1,
    cols = 1,
    gridExpand = TRUE,
    stack = TRUE
  )
  openxlsx::addStyle(
    wb, tab_name, s$subheadings,
    rows = subheadings,
    cols = 1,
    stack = TRUE
  )
  openxlsx::addStyle(
    wb, tab_name, s$bold_text,
    rows = rows_to_bold,
    cols = 1,
    stack = TRUE
  )

  if (is.data.frame(hyperlinks)) {

    for (x in 1:nrow(hyperlinks)) {

      hyperlink_name <- hyperlinks[x, 2]
      names(hyperlink_name) <- text_df[hyperlinks[x, 1], 1]
      class(hyperlink_name) <- "hyperlink"

      if (hyperlinks[x, 1] <= nrows) {
        openxlsx::writeData(
          wb, tab_name,
          startCol = 1,
          startRow = hyperlinks[x, 1],
          colNames = FALSE,
          x = hyperlink_name
        )
      } else {
        stop("Row number for a hyperlink is higher than number of rows with text.")
      }
    }
  }
}
