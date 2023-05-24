test_that("Function works for without optional args", {
  wb <- openxlsx::createWorkbook()
  text_df <- data.frame(X1 = c("a", "b", "c", "d", "e"))
  halefunctionlib::create_cover_sheet(wb = wb, text_df = text_df)

  expect_equal(wb$sheet_names, "Cover_sheet")
  expect_equal(read.xlsx(wb, 1, colNames = FALSE), text_df)
  expect_equivalent(wb$rowHeights[[1]], "30")
  expect_equivalent(wb$colWidths[[1]], "130")
  expect_equal(
    getBaseFont(wb),
    list(
      size = list(val = "12"),
      colour = list(rgb = "FF000000"),
      name = list(val = "Arial")
    )
  )
})

test_that("Function runs with optional arguments", {

  wb <- openxlsx::createWorkbook()
  text_df <- data.frame(
    X1 = c("a", "b1", "c1", "d", "e", "f", "ONS", "email")
  )
  replacements <- data.frame(
    replace = c("b1", "c1"),
    with = c("b replaced", "c replaced")
  )
  links <- data.frame(
    rows = c(7, 8),
    links = c("https://www.ons.gov.uk", "mailto:Health.Data@ons.gov.uk")
  )
  replaced_df <- data.frame(
    X1 = c("a", "b replaced", "c replaced", "d", "e", "f", "ONS", "email")
  )
  halefunctionlib::create_cover_sheet(
    wb = wb,
    tab_name = "Cool_sheet",
    text_df = text_df,
    subheadings = c(3, 6),
    hyperlinks = links,
    rows_to_bold = 9,
    replacements = replacements,
    subheadings_row_height = 40,
    column_width = 100
  )

  expect_equal(wb$sheet_names, "Cool_sheet")
  expect_equal(read.xlsx(wb, 1, colNames = FALSE), replaced_df)
  expect_equivalent(wb$rowHeights[[1]], c("40", "40"))
  expect_equivalent(wb$colWidths[[1]], "100")
  expect_equal(
    getBaseFont(wb),
    list(
      size = list(val = "12"),
      colour = list(rgb = "FF000000"),
      name = list(val = "Arial")
    )
  )
})

test_that("Function formatting style works as expected", {
  wb <- openxlsx::createWorkbook()
  text_df <- data.frame(
    X1 = c("a", "b1", "c1", "d", "e", "f", "ONS", "email", "bold")
  )
  links <- data.frame(
    rows = c(7, 8),
    links = c("https://www.ons.gov.uk", "mailto:Health.Data@ons.gov.uk")
  )
  halefunctionlib::create_cover_sheet(
    wb = wb,
    tab_name = "Cool_sheet",
    text_df = text_df,
    subheadings = c(3, 6),
    hyperlinks = links,
    rows_to_bold = 9,
    subheadings_row_height = 40,
    column_width = 100
  )
  fl <- tempfile(fileext = ".xlsx")
  openxlsx::saveWorkbook(wb, file = fl, overwrite = TRUE)
  formats <- tidyxl::xlsx_formats(fl)
  x <- tidyxl::xlsx_cells(fl)

  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$bold), "address"],
    data.frame(address = c("A1", "A3", "A6", "A9"))
    )
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$underline == "single"),
      "address"],
    data.frame(address = c("A7", "A8"))
    )
  expect_equivalent(
    x[x$local_format_id %in%
        which(formats$local$font$color$theme == "hyperlink"),
      "address"],
    data.frame(address = c("A7", "A8"))
    )
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$color$rgb == "FF0000FF"),
      "address"],
    data.frame(address = c("A7", "A8"))
    )
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$color$rgb == "FF000000"),
      "address"],
    data.frame(address = c("A1", "A2", "A3", "A4", "A5", "A6", "A9"))
    )
  expect_equal(x$height, c(15, 15, 40, 15, 15, 40, 15, 15, 15))
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$size == 12), "address"],
    data.frame(address = c("A2", "A4", "A5", "A7",  "A8", "A9"))
    )
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$size == 14), "address"],
    data.frame(address = c("A3", "A6"))
    )
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$size == 16), "address"],
    data.frame(address = c("A1"))
    )
  expect_equivalent(
    x[x$local_format_id %in% which(formats$local$font$name == "Arial"),
      "address"],
    data.frame(address = paste0("A", 1:9))
    )
})
