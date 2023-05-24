#' @title Replace text for spreadsheets
#'
#' @description A function to replaces set strings with specified replacements
#'
#' @param replacements A data frame containing a "replace" column with strings
#' to be replaced by strings contained in the "with" column
#' @param text_file A data frame with a single column of characters
#'
#' @importFrom dplyr mutate_all
#' @importFrom stringr str_replace_all
#'
#' @return Modified data frame where certain strings were replaced

replace_text_for_spreadsheets <- function(replacements, text_file) {

  for (x in 1:nrow(replacements)) {
    text_file <-dplyr::mutate_all(
      text_file,
      stringr::str_replace_all,
      replacements$replace[x],
      replacements$with[x]
      )
  }
  text_file
}
