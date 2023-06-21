s <- create_styles()

test_that("List has been created", {
  expect_true(is.list(s))
})

test_that("List contains 11 styles used by other speadsheet funtions", {
  s_names <-  c("text", "bold_text", "heading", "subheadings", "table_header",
                "wrap_text", "centre", "border", "no_decimal", "one_decimal",
                "two_decimal")
  expect_true(all(s_names %in% names(s)))
  for (i in seq(s_names)){
    expect_equal(class(s[[s_names[i]]]), "Style", ignore_attr = TRUE)
  }
})

test_that("Font in text styles is set to Arial", {
  s_names <-  c("text", "bold_text", "heading", "subheadings", "table_header",
                "wrap_text", "centre")
  for (i in seq(s_names)){
    expect_equal(s[[s_names[i]]]$fontName$val, "Arial", ignore_attr = TRUE)
  }
})

test_that("Fonts in text styles have expected sizes", {
  s_names <-  c("text", "bold_text", "heading", "subheadings", "table_header",
                "wrap_text", "centre")
  sizes <- c(12, 12, 16, 14, 12, 12, 12)
  for (i in seq(s_names)){
    expect_equal(s[[s_names[i]]]$fontSize$val, sizes[i], ignore_attr = TRUE)
  }
})

test_that("Bold styles are bold", {
  s_names <-  c("bold_text", "heading", "subheadings", "table_header")
  for (i in seq(s_names)){
    expect_equal(s[[s_names[i]]]$fontDecoration, "BOLD", ignore_attr = TRUE)
  }
})

test_that("Border style has thin borders on all sides", {
  borders <- c("borderTop", "borderBottom", "borderLeft", "borderRight")
  for (i in seq(borders)){
    expect_equal(s$border[[borders[i]]], "thin", ignore_attr = TRUE)
  }
})

test_that("Styles formatting numbers have numFmt list", {
  num_s <- c("no_decimal", "one_decimal", "two_decimal")
  for (i in seq(num_s)){
    expect_true(is.list(s[[num_s[i]]]$numFmt))
  }
})
