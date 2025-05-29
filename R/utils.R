#' Sanitize input xlsx file
#'
#' Ensure that file format and sheets match expectations
#'
#' @param file A character string with a path pointing to an ARM xlsx export file
#' @param required_sheets A vector of character strings that must be available in the xlsx export file
#'
#' @importFrom tools file_ext
#' @importFrom readxl excel_sheets
#' @importFrom stringr str_detect
#'
.check_xlsx_file_and_sheet <- function(file, required_sheets) {
  if (!is.character(file) || length(file) != 1)
    stop("File path must be a character string.")
  if (!file.exists(file)) stop("File does not exist: ", file)
  if (tolower(tools::file_ext(file)) != "xlsx")
    stop("File extension must be .xlsx")
  sheets <- tryCatch(excel_sheets(file), error = function(e) character())
  if (str_detect(sheets, paste0(required_sheets, collapse = "|"), negate = T))
    stop(
      "Missing required sheet(s): ",
      paste(setdiff(required_sheets, sheets), collapse = ", ")
    )
}
