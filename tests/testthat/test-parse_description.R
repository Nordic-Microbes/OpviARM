test_that("read_descr_sheet requires Description sheet", {
  tmp_file <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(
    list(
      Description = data.frame(x = 1),
      Another = data.frame(y = 2)
    ),
    tmp_file
  )
  expect_silent(read_descr_sheet(tmp_file))

  tmp_file_no <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(
    list(
      SomethingElse = data.frame()
    ),
    tmp_file_no
  )
  expect_error(read_descr_sheet(tmp_file_no), "Missing required sheet")
})

test_that("parse_description returns list for format_trial and format_soil", {
  # Fake description tibble
  df <- data.frame(V1 = c("Site", "Plots", "Sowing"), V2 = c("A", "B", "C"))
  md <- parse_description(df)
  expect_type(md, "list")
  # next functions should accept the output without error
})
