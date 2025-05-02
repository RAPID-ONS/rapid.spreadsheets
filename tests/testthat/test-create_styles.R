s <- create_styles()

test_that("List has been created", {
  expect_true(is.list(s))
})

test_that("List contains 16 styles used by other speadsheet funtions", {
  s_names <-  c("text", "bold_text", "heading", "subheadings", "table_header",
                "wrap_text", "centre", "all_borders", "top_bottom_borders",
                "vertical_borders_left", "vertical_borders_right",
                "bottom_borders", "left_align", "no_decimal", "one_decimal",
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

test_that("All borders style has thin borders on all sides", {
  borders <- c("borderTop", "borderBottom", "borderLeft", "borderRight")
  for (i in seq(borders)){
    expect_equal(s$all_borders[[borders[i]]], "thin")
  }
})

test_that("Top and bottom borders style has thin borders on top and bottom", {
  borders <- c("borderTop", "borderBottom")
  for (i in seq(borders)){
    expect_equal(s$top_bottom_borders[[borders[i]]], "thin")
  }
  no_borders <- c("borderLeft", "borderRight")
  for (i in seq(no_borders)){
    expect_equal(s$top_bottom_borders[[no_borders[i]]], NULL)
  }
})

test_that("Vertical borders left style has thin borders on the left", {
  borders <- c("borderTop", "borderBottom", "borderRight")
  for (i in seq(borders)){
    expect_equal(s$vertical_borders_left[[borders[i]]], NULL)
  }
  expect_equal(s$vertical_borders_left$borderLeft, "thin")
})

test_that("Vertical borders right style has thin borders on the right", {
  borders <- c("borderTop", "borderBottom", "borderLeft")
  for (i in seq(borders)){
    expect_equal(s$vertical_borders_right[[borders[i]]], NULL)
  }
  expect_equal(s$vertical_borders_right$borderRight, "thin")
})

test_that("Bottom borders style has thin borders on the bottom", {
  borders <- c("borderTop", "borderLeft", "borderRight")
  for (i in seq(borders)){
    expect_equal(s$bottom_borders[[borders[i]]], NULL)
  }
  expect_equal(s$bottom_borders$borderBottom, "thin")
})

test_that("Left aligned style has alignment set to left", {
  expect_equal(s$left_align$halign, "left")
})

test_that("Styles formatting numbers have numFmt list", {
  num_s <- c("no_decimal", "one_decimal", "two_decimal")
  for (i in seq(num_s)){
    expect_true(is.list(s[[num_s[i]]]$numFmt))
  }
})
