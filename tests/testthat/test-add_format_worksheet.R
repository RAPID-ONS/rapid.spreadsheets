test_that("Adds worksheet to a workbook", {
  wb <- openxlsx::createWorkbook()
  add_format_worksheet(
    wb,
    ncol_df = 2,
    nrow_df = 2,
    tab_name = "One",
    heading = "Title"
  )
  expect_equal(wb$sheet_names, "One")
  expect_equal(
    read.xlsx(wb, 1, colNames = FALSE),
    data.frame(X1 = c("Title", "This worksheet contains 1 table"))
  )
})

test_that("Error is thrown if more than 1 table used", {
  wb <- openxlsx::createWorkbook()

  expect_error(
    add_format_worksheet(
      wb,
      ncol_df = 2,
      nrow_df = 2,
      tab_name = "One",
      heading = "Title",
      num_tables = 2
    )
  )
})

test_that("Worksheet formatting style is set as expected", {
  wb <- openxlsx::createWorkbook()
  add_format_worksheet(
    wb,
    ncol_df = 2,
    nrow_df = 2,
    tab_name = "One",
    heading = "Title"
  )
  cells_with_borders <- c("A3", "B3", "A4", "B4", "A5", "B5")
  fl <- tempfile(fileext = ".xlsx")
  openxlsx::saveWorkbook(wb, file = fl, overwrite = TRUE)
  formats <- tidyxl::xlsx_formats(fl)
  x <- tidyxl::xlsx_cells(fl)

  expect_equal(
    getBaseFont(wb),
    list(
      size = list(val = "12"),
      colour = list(rgb = "FF000000"),
      name = list(val = "Arial")
    )
  )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$bold), "address"],
    data.frame(address = c("A1", "A3", "B3")),
    ignore_attr = TRUE
  )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$size == 12), "address"],
    data.frame(address = c("A2", "A3", "B3", "A4", "B4", "A5", "B5")),
    ignore_attr = TRUE
  )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$size == 16), "address"],
    data.frame(address = c("A1")),
    ignore_attr = TRUE
  )

  expect_equal(
    x[x$local_format_id %in% which(formats$local$border$right$style == "thin"),
      "address"],
    data.frame(address = cells_with_borders),
    ignore_attr = TRUE
  )

  expect_equal(
    x[x$local_format_id %in% which(formats$local$border$left$style == "thin"),
      "address"],
    data.frame(address = cells_with_borders),
    ignore_attr = TRUE
  )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$border$top$style == "thin"),
      "address"],
    data.frame(address = cells_with_borders),
    ignore_attr = TRUE
  )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$border$bottom$style == "thin"),
      "address"],
    data.frame(address = cells_with_borders),
    ignore_attr = TRUE
  )
})
