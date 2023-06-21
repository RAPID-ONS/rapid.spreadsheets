test_that("Creating contents sheet with default options and a single column", {

  wb <- openxlsx::createWorkbook()
  df <- data.frame(Name = c("Cover_sheet", "Contents", "Table_1", "Table_2"))
  create_contents_notes(wb = wb, df = df)

  expect_equal(wb$sheet_names, "Contents")
  expect_equal(
    read.xlsx(wb, 1, colNames = FALSE),
    data.frame(X1 = c("Table of contents", "This worksheet contains 1 table", "Name"))
    )
  expect_identical(is.list(wb$rowHeights[[1]]), TRUE)
  expect_equal(wb$colWidths[[1]], c("20", "20"), ignore_attr = TRUE)
  expect_equal(
    getBaseFont(wb),
    list(size = list(val = "12"),
         colour = list(rgb = "FF000000"),
         name = list(val = "Arial")
         )
    )
})

test_that("Create contents formats excel cells as expected", {

  wb <- openxlsx::createWorkbook()
  df <- data.frame(Name = c("Cover_sheet", "Contents", "Table_1", "Table_2"))
  create_contents_notes(wb = wb, df = df)
  fl <- tempfile(fileext = ".xlsx")
  openxlsx::saveWorkbook(wb, file = fl, overwrite = TRUE)
  formats <- tidyxl::xlsx_formats(fl)
  x <- tidyxl::xlsx_cells(fl)

  expect_identical(x$formula[4], "=HYPERLINK(\"#'Cover_sheet'!A1\", \"Cover_sheet\")")
  expect_identical(x$formula[5], "=HYPERLINK(\"#'Contents'!A1\", \"Contents\")")
  expect_identical(x$formula[6], "=HYPERLINK(\"#'Table_1'!A1\", \"Table_1\")")
  expect_identical(x$formula[7], "=HYPERLINK(\"#'Table_2'!A1\", \"Table_2\")")
  expect_identical(x$height, rep(15, 7))

  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$bold), "address"],
    data.frame(address = c("A1", "A3")),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in%
        which(formats$local$font$color$theme == "hyperlink"),
      "address"],
    data.frame(address = paste0("A", 4:7)),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$underline == "single"),
      "address"],
    data.frame(address = paste0("A", 4:7)),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in%
        which(formats$local$font$color$rgb == "FF0000FF"),
      "address"],
    data.frame(address = paste0("A", 4:7)),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$color$rgb == "FF000000"),
      "address"],
    data.frame(address = c(paste0("A", 1:3))),
    ignore_attr = TRUE
    )

  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$size == 12), "address"],
    data.frame(address = paste0("A", 2:7)),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$size == 16), "address"],
    data.frame(address = c("A1")),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$name == "Arial"),
      "address"],
    data.frame(address = paste0("A", 1:7)),
    ignore_attr = TRUE
    )
})

test_that("Creating contents sheet with two column has set column widths", {
  wb <- openxlsx::createWorkbook()
  df <- data.frame(
    Name = c("Cover", "Contents", "Table_1"),
    Description = c("Cover sheet", "Contents", "Some data")
  )
  create_contents_notes(wb = wb, df = df)

  expect_equal(wb$colWidths[[1]], c("20", "80"), ignore_attr = TRUE)
})

test_that("Creating contents sheet with three column has set column widths", {
  wb <- openxlsx::createWorkbook()
  df <- data.frame(
    Name = c("Cover", "Contents", "Table_1"),
    Description = c("Cover sheet", "Contents", "Some data"),
    Links = c(NA, NA, "link")
  )
  create_contents_notes(wb = wb, df = df)

  expect_equal(wb$colWidths[[1]], c("20", "80", "15"), ignore_attr = TRUE)
})

test_that("Creating contents with double (long) heading, additional text and no interlinks", {
  wb <- openxlsx::createWorkbook()
  df <- data.frame(
    Name = c("Cover sheet", "Contents", "Table_1"),
    Description = c("Cover sheet", "Contents", "Table"),
    Links = c(NA, NA, "link")
  )
  create_contents_notes(
    wb, df,
    heading = c("Long", "title"),
    contents_links = FALSE,
    additional_text = c("A", "B")
  )

  expect_equal(
    read.xlsx(wb, 1, colNames = FALSE),
    data.frame(
      X1 = c("Long", "title", "This worksheet contains 1 table", "A", "B",
             "Name", df$Name),
      X2 = c(rep(NA, 5), "Description", df$Description),
      X3 = c(rep(NA, 5), "Links", df$Links)
      )
    )
})

test_that("Function runs as expected with third column with hyperlinks", {
  wb <- openxlsx::createWorkbook()
  df <- data.frame(
    Name = c("Cover", "Contents", "T_1", "T_2"),
    Description = c("Cover sheet", "Contents", "Some data", "Other data"),
    Links = c(NA, NA, "link", "email")
  )
  hyperlinks <- data.frame(
    rows = 3:4,
    links = c("https://www.ons.gov.uk/", "mailto:Health.Data@ons.gov.uk")
  )
  create_contents_notes(wb, df,
                                         hyperlinks = hyperlinks
                                         )
  fl <- tempfile(fileext = ".xlsx")
  openxlsx::saveWorkbook(wb, file = fl, overwrite = TRUE)
  formats <- tidyxl::xlsx_formats(fl)
  x <- tidyxl::xlsx_cells(fl)

  expect_equal(
    read.xlsx(wb, 1, colNames = FALSE),
    data.frame(
      X1 = c("Table of contents", "This worksheet contains 1 table", "Name",
             rep(NA, 4)),
      X2 = c(NA, NA, "Description", df$Description),
      X3 = c(NA, NA, "Links", df$Links)))

  expect_equal(
    x[x$local_format_id %in%
        which(formats$local$font$color$theme == "hyperlink"),
      "address"],
    data.frame(address = c("A4", "A5", "A6", "C6", "A7", "C7")),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$underline == "single"),
      "address"],
    data.frame(address = c("A4", "A5", "A6", "C6", "A7", "C7")),
    ignore_attr = TRUE
    )
  expect_equal(
    x[x$local_format_id %in% which(formats$local$font$color$rgb == "FF0000FF"),
      "address"],
    data.frame(address = c("A4", "A5", "A6", "C6", "A7", "C7")),
    ignore_attr = TRUE
    )
})
