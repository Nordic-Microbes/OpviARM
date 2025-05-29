test_that("read_assessment_sheet accepts only .xlsx with required sheet", {
  # create a temp .xlsx with the correct and incorrect sheets
  tmp_file <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(
    list(
      Assessments = data.frame(x = 1),
      Other = data.frame(y = 2)
    ),
    tmp_file
  )
  expect_silent(read_assessment_sheet(tmp_file))
  # remove the Assessments sheet
  tmp_file_no <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(
    list(
      NotAssessments = data.frame(z = 3)
    ),
    tmp_file_no
  )
  expect_error(read_assessment_sheet(tmp_file_no), "Missing required sheet")

  # try wrong extension
  tmp_file_txt <- tempfile(fileext = ".txt")
  writeLines("fake content", tmp_file_txt)
  expect_error(
    read_assessment_sheet(tmp_file_txt),
    "File must have .xlsx extension"
  )

  # try non-existing file
  expect_error(read_assessment_sheet("not_a_file.xlsx"), "does not exist")
})
