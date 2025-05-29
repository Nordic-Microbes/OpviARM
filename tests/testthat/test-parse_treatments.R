test_that("read_treatment_sheet requires Treatments sheet", {
  tmp_file <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(
    list(
      Treatments = data.frame(x = 1)
    ),
    tmp_file
  )
  expect_silent(read_treatment_sheet(tmp_file))

  tmp_file_no <- tempfile(fileext = ".xlsx")
  openxlsx::write.xlsx(
    list(
      NotTreatments = data.frame(x = 1)
    ),
    tmp_file_no
  )
  expect_error(read_treatment_sheet(tmp_file_no), "Missing required sheet")
})
