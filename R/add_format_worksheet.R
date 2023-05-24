#' @title Add a new worksheet with basic formatting.
#'
#' @description This function will add a new worksheet with formatting that
#' follows accessibility guidelines to an openxlsx workbook.
#'
#' @param wb An openxlsx workbook object.
#' @param ncol_df Number of columns in table.
#' @param nrow_df Number of rows in table.
#' @param tab_name Worksheet name as a string.
#' @param heading Worksheet heading as a character vector, usually length one
#' unless the title is very long. To split heading into 2 rows save it as two
#' characters in a vector.
#' @param subtitle_length Number of rows for subtitles and extra notes that
#' will be under the heading and before the table, default is 0.
#' @param num_tables Number of tables in worksheet as integer, default is 1.
#'
#' @return Adds a worksheet to existing openxlsx workbook. The formatting will
#' follow accessibility guidelines.
#'
#' @import openxlsx
#'
#' @export

add_format_worksheet <- function(wb,
                                 ncol_df,
                                 nrow_df,
                                 tab_name,
                                 heading,
                                 subtitle_length = 0,
                                 num_tables = 1) {

  if (num_tables > 1) {
    stop("Avoid worksheets with multiple tables. This function cannot format
         multiple tables correctly. If you do need to have multiple tables
         please consult accessability guidance and maunually format.")
  }

  openxlsx::addWorksheet(wb, tab_name, gridLines = FALSE)
  openxlsx::modifyBaseFont(wb, fontSize = 12, fontName = "Arial")

  heading_length <- length(heading)
  table_start_row <- heading_length + subtitle_length + 2
  table_end_row <- table_start_row + nrow_df

  openxlsx::writeData(wb, tab_name, heading, startCol = 1, startRow = 1)
  openxlsx::writeData(wb, tab_name, "This worksheet contains 1 table",
                      startCol = 1,
                      startRow = heading_length + 1
                      )

  s <- create_styles()
  openxlsx::addStyle(wb, tab_name, s$heading,
                     rows = 1:heading_length,
                     cols = 1,
                     gridExpand = TRUE,
                     stack = TRUE
                     )
  openxlsx::addStyle(wb, tab_name, s$table_header,
                     rows = table_start_row,
                     cols = 1:ncol_df,
                     gridExpand = TRUE,
                     stack = TRUE
                     )
  openxlsx::addStyle(wb, tab_name, s$border,
                     rows = table_start_row:table_end_row,
                     cols = 1:ncol_df,
                     gridExpand = TRUE,
                     stack = TRUE
                     )
}
