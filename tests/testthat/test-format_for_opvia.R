test_that("extract_assessment_germ returns correct columns for format_germ", {
  # Mock data with necessary columns
  df <- data.frame(
    Name = "T1",
    stage_majority = 1,
    Plot = 1,
    `Rating Date` = "2024-01-01",
    value = 10,
    assessment = "a"
  )
  out <- format_germ(df)
  expect_s3_class(out, "data.frame")
})
