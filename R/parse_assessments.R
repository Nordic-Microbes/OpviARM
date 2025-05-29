#' Import an assessment sheet from ARM
#'
#' Loads a specified assessment worksheet from an ARM Excel file, preserving raw values for further parsing.
#'
#' @param file Path to the ARM-generated Excel file; should point directly to an assessment sheet.
#'
#' @return Tibble with assessment data as present in the source file.
#' @importFrom readxl read_excel
#' @export
read_assessment_sheet <- function(excel_path) {
  .check_xlsx_file_and_sheet(excel_path, required_sheets = c("Assessment"))
  sheets <- excel_sheets(excel_path)

  assessment_sheet <- read_xlsx(
    excel_path,
    col_names = F,
    trim_ws = F,
    sheet = which(str_detect(sheets, "Assessment"))
  )

  return(assessment_sheet)
}

#' Parse raw assessment data
#'
#' Reshapes and cleans an assessment data frame, extracting core measurement and metadata columns for downstream analyses.
#'
#' @param assessment_df Tibble with columns as produced by \code{read_assessment_sheet}.
#'
#' @return Cleaned and organized tibble suitable for extracting trait-specific data.
#' @importFrom dplyr mutate select
#' @importFrom tidyr pivot_longer
#' @export
parse_assessment_data <- function(assessment_sheet) {
  assessments <- assessment_sheet |>
    slice_tail(n = -8) |> # DANGER: Expects fixed rownumber - Implement check
    slice_head(n = 19) |> # DANGER: Expects fixed rownumber - Implement check
    select_if(~ any(!is.na(.))) |>
    pivot_longer(cols = !...1, names_to = "row_id", values_to = "value") |>
    pivot_wider(names_from = ...1, values_from = value) |>
    select(-row_id) |>
    mutate(
      assessment = row_number(),
      `Rating Date` = dmy(`Rating Date`)
    ) |>
    separate(
      col = `Crop Stage Majority/Min/Max`,
      into = c("stage_majority", "stage_min", "stage_max"),
      sep = ";",
      convert = T
    )

  assess_cols <- assessment_sheet |>
    slice_tail(n = -28) |> # DANGER: Expects fixed rownumber - Implement check
    slice(1) |>
    unlist() |>
    as.character()

  assess_data <- assessment_sheet |>
    slice_tail(n = -29) |> # DANGER: Expects fixed rownumber - Implement check
    set_names(assess_cols) |>
    filter(!is.na(Plot)) |>
    fill(No.:Timing) |>
    pivot_longer(
      cols = matches("^\\d+$"),
      names_to = "assessment",
      values_to = "value"
    ) |>
    mutate(
      Rate = as.numeric(Rate),
      Timing = as.numeric(Timing),
      assessment = as.integer(assessment),
      value = as.numeric(value)
    ) |>
    left_join(assessments, by = join_by("assessment"))

  return(assess_data)
}

#' Filter assessment data for germination traits
#'
#' Subsets and renames columns from parsed assessment data to include only germination-related rows.
#'
#' @param assessment_df Tibble output from \code{parse_assessment_data}. Should contain trait-identifying columns.
#'
#' @return Tibble with germination records and Opvia-compliant fields.
#' @importFrom dplyr filter select rename
#' @export
extract_assessment_germ <- function(assessment_data) {
  germ <- assessment_data |>
    filter(str_detect(`SE Name`, "COT")) |>
    mutate(
      Replicate = row_number()
    )

  return(germ)
}


#' Filter assessment data for yield traits
#'
#' Extracts yield-related rows and reformats columns to meet Opvia import requirements.
#'
#' @param assessment_df Tibble from \code{parse_assessment_data}, containing all assessments.
#'
#' @return Tibble with yield records, harmonized for Opvia upload.
#' @importFrom dplyr filter select rename
#' @export
extract_assessment_yield <- function(assessment_data) {
  yield <- assessment_data |>
    filter(str_detect(`SE Name`, "Y")) |>
    mutate(
      `SE Name` = case_when(
        `SE Name` == "Y085" ~ "yield",
        `SE Name` == "Y006" ~ "density",
        `SE Name` == "Y005" ~ "prot_perc",
        `SE Name` == "Y086" ~ "moist_perc",
        `SE Name` == "Y004" ~ "tgw"
      )
    ) |>
    rename(report = `Reporting Basis`) |>
    pivot_wider(
      id_cols = c(No.:Plot, `Assessed By`:`Rating Date`),
      names_from = `SE Name`,
      values_from = c(value, report)
    ) |>
    mutate(
      value_yield = case_when(
        # If results are reported per plot, normalize to area
        str_detect(report_yield, regex("plot", ignore_case = T)) ~
          value_yield / md$plots$plot_area, # Notify the user
        # If results are reported per x rowm, normalize with row_width
        str_detect(report_yield, regex("rowm", ignore_case = T)) ~
          value_yield /
            as.numeric(str_extract(report_yield, "^[:digit:]*")) *
            md$plots$row_width /
            100, # Notify the user
        .default = value_yield
      )
    )
  return(yield)
}
