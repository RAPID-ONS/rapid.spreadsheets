test_that("Basic table is created with default values", {
  wb <- openxlsx::createWorkbook()
  df <- data.frame(a = c("a","b","c"), d = c(1, 2, 3))
  create_data_table_tab(wb, df)

  expect_equal(wb$sheet_names, "Table_1")
  expect_equal(
    read.xlsx(wb, 1, colNames = FALSE),
    data.frame(
      X1 = c("Heading", "This worksheet contains 1 table", "a", df$a),
      X2 = c(NA, NA, "d", df$d)
    )
  )
  expect_equivalent(wb$colWidths[[1]], c("10", "10"))
  expect_equal(
    getBaseFont(wb),
    list(
      size = list(val = "12"),
      colour = list(rgb = "FF000000"),
      name = list(val = "Arial")
    )
  )
})

test_that("Numeric values stored as numeric even if mixed with text rows", {
  wb <- openxlsx::createWorkbook()
  df <- data.frame(a = c("a","b","c"), d = c(1, 2, 3), r = c(2.3, 4.1, "x"))
  create_data_table_tab(wb, df, num_char_cols = 3)
  fl <- tempfile(fileext = ".xlsx")
  openxlsx::saveWorkbook(wb, file = fl, overwrite = TRUE)
  x <- tidyxl::xlsx_cells(fl)

  expect_equal(
    x$data_type[x$address %in% c("C4", "C5")],
    c("numeric", "numeric")
  )
  expect_equal(x$data_type[x$address %in% c("C6")], c("character"))
})

test_that("Number formatting works as expected", {
  wb <- openxlsx::createWorkbook()
  v_numFmt <- c(1.123, 2.345, 1003.456)
  df <- data.frame(a = v_numFmt, b = v_numFmt, c = v_numFmt, d = v_numFmt)
  create_data_table_tab(
    wb, df,
    no_decimal = 2,
    one_decimal = 3,
    two_decimal = 4
  )
  fl <- tempfile(fileext = ".xlsx")
  openxlsx::saveWorkbook(wb, file = fl, overwrite = TRUE)
  x <- tidyxl::xlsx_cells(fl)
  formats <- tidyxl::xlsx_formats(fl)

  cells_gen <- x$local_format_id %in% which(formats$local$numFmt == "General")
  expect_identical(
    sort(x$address[cells_gen]),
    c(paste0("A", 1:6), paste0(c("B", "C", "D"), 3))
  )
  expect_identical(
    x$address[x$local_format_id %in% which(formats$local$numFmt == "#;##0")],
    paste0("B", 4:6)
  )
  expect_identical(
    x$address[x$local_format_id %in% which(formats$local$numFmt == "#,##0.0")],
    paste0("C", 4:6)
  )
  expect_identical(
    x$address[x$local_format_id %in% which(formats$local$numFmt == "#,##0.00")],
    paste0("D", 4:6)
  )
})
