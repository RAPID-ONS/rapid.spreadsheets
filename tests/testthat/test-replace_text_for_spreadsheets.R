test_that("Phrases to replace are changed in the output data frame", {
  text_file <- data.frame(a = c("Heading", "Release date: !date!", "smth"))
  replacements <- data.frame(
    replace = c("!date!", "smth"),
    with = c("8 Nov", "hi!")
  )
  output <- replace_text_for_spreadsheets(replacements, text_file)
  expected <- data.frame(a = c("Heading", "Release date: 8 Nov", "hi!"))
  expect_identical(output, expected)
})
