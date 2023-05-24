#' @title Adds data table worksheet to workbook.
#'
#' @description This function adds a data table worksheet to an openxlsx
#' workbook.
#'
#' @param wb An openxlsx workbook object.
#' @param df Data frame containing data to put in data table.
#' @param tab_name Worksheet name as string, default is "Table_1".
#' @param heading Table title/ heading as character vector. Vector length
#' should be one unless table title is very long. If length is more than one
#' text will be split in multiple rows. Default is "Heading".
#' @param additional_text Character vector containing additional rows of text
#' to be placed above the data table. Default is empty.
#' @param column_width Width of table columns, default is 10.
#' @param num_tables Number of tables in a worksheet. Default is 1.
#' @param num_char_cols Use if you get Excel warning "numbers stored as text".
#' Vector of column numbers that contain both numbers and text values.
#' Optional argument.
#' @param no_decimal Vector containing column numbers where values should be
#' shown without any decimal points. Will also add thousand separator "1,000"
#' where needed. Optional argument.
#' @param one_decimal Vector containing column numbers where values should be
#' shown with one decimal point. Will also add thousand separator "1,000"
#' where needed. Optional argument.
#' @param two_decimal Vector containing column numbers where values should
#' be shown with two decimal points. Will also add thousand separator "1,000"
#' where needed. Optional argument.
#'
#' @return Adds a worksheet with data table to existing openxlsx workbook.
#'
#' @import openxlsx
#'
#' @export

create_data_table_tab <- function(wb,
                                  df,
                                  tab_name = "Table_1",
                                  heading = "Heading",
                                  additional_text = c(),
                                  column_width = rep(10, ncol(df)),
                                  num_tables = 1,
                                  num_char_cols = NA,
                                  no_decimal = NA,
                                  one_decimal = NA,
                                  two_decimal = NA) {

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

  openxlsx::setColWidths(wb, tab_name, 1:ncol(df), column_width)

  openxlsx::writeData(
    wb, tab_name,
    additional_text,
    startCol = 1,
    startRow = heading_length + 2
  )
  openxlsx::writeDataTable(
    wb, tab_name,
    df,
    startCol = 1,
    startRow = table_start_row,
    tableName = tab_name,
    tableStyle = "None",
    withFilter = FALSE
  )

  if (is.numeric(num_char_cols)) {
    overwrite_df(wb, df, tab_name, table_start_row, num_char_cols)
  }

  decimals <- list(no_decimal, one_decimal, two_decimal)
  s <- create_styles()
  style <- c(s$no_decimal, s$one_decimal, s$two_decimal)

  for (i in which(sapply(decimals, is.numeric) == TRUE)) {

    openxlsx::addStyle(
      wb, tab_name,
      style = style[[i]],
      rows = (table_start_row + 1):table_end_row, # first table row is colname
      cols = decimals[[i]],
      gridExpand = TRUE,
      stack = TRUE
    )
  }
}

#' @title Fixes column where numbers are stored as text.
#'
#' @description Overwrites selected columns that contain both numeric and
#' character elements.
#'
#' @param wb Openxlsx workbook name.
#' @param df Data frame containing columns with worksheet names and
#' descriptions.
#' @param tab_name Worksheet name as string.
#' @param table_start_row Row number where table starts (as integer).
#' @param num_char_cols Use if you get Excel warning "numbers stored as text".
#' Vector of column numbers to be overwritten.
#'
#' @return Updated workbook with modified columns.
#'
#' @import openxlsx

overwrite_df <- function(wb, df, tab_name, table_start_row, num_char_cols) {

  df_list <- as.list(df)
  style <- openxlsx::createStyle(halign = "right")

  for (i in seq_along(num_char_cols)) {

    icol <- df_list[[num_char_cols[i]]]

    for (j in seq_along(icol)) {

      if (is.na(suppressWarnings(as.numeric(icol[[j]])))) {

        x <- icol[[j]]
        openxlsx::addStyle(
          wb, tab_name,
          style,
          rows = j + table_start_row,
          cols = num_char_cols[i],
          stack = TRUE
        )

      } else {
        x <- as.numeric(icol[[j]])
      }

      openxlsx::writeData(
        wb, tab_name,
        x,
        startCol = num_char_cols[i],
        startRow = j + table_start_row
      )
    }
  }
}
