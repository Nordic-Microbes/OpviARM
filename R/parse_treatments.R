#' Import a treatment sheet from ARM
#'
#' Reads the Treatments worksheet from an ARM Excel file and returns a tibble for parsing.
#'
#' @param file Path to the ARM export containing a Treatments sheet.
#'
#' @return Tibble with all treatment definitions and properties.
#' @importFrom readxl read_excel
#' @export
read_treatment_sheet <- function(excel_path) {
  .check_xlsx_file_and_sheet(
    excel_path,
    required_sheets = c("Trial Treatments")
  )
  sheets <- excel_sheets(excel_path)

  treatments <- read_xlsx(
    excel_path,
    col_names = c(
      # subject to change **id_code**
      "no",
      "arm_type",
      "name",
      "description",
      "rate",
      "rate_unit",
      "appl_code",
      "appl_timing"
    ),
    trim_ws = F,
    sheet = which(str_detect(sheets, "Trial Treatments")),
    skip = 10 # DANGER: Expects fixed rownumber
  ) |>
    filter(str_detect(no, "^[:digit:]*$"))

  return(treatments)
}
