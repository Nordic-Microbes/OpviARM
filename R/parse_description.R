#' Import a description sheet from ARM
#'
#' Loads the description sheet from an ARM Excel file, providing all contextual metadata for the trial.
#'
#' @param file Path to the ARM export containing a Description sheet.
#'
#' @return Tibble containing the description fields for parsing.
#' @importFrom readxl read_excel
#' @export
read_descr_sheet <- function(excel_path) {
  .check_xlsx_file_and_sheet(
    excel_path,
    required_sheets = c("Site Description")
  )
  sheets <- excel_sheets(excel_path)

  descr_sheet <- read_xlsx(
    excel_path,
    col_names = F,
    trim_ws = F,
    sheet = which(str_detect(sheets, "Site Description"))
  )
  return(descr_sheet)
}

#' Extract and organize trial metadata
#'
#' Parses high-level study and management information from a raw description tibble, returning a structured list for later formatting.
#'
#' @param descr_df Tibble as output by \code{read_descr_sheet}; must contain all standard ARM description fields.
#'
#' @return Named list with elements for sowing, site, plots, etc.
#' @importFrom dplyr filter pull
#' @export
parse_description <- function(descr_sheet) {
  # Grab metadata
  # TODO: Convert to function that sanitizes input
  # DANGER: Expects fixed columns
  md <- list(
    study = descr_sheet |>
      filter(str_detect(...1, "Trial ID")) |>
      pull(2),

    year = descr_sheet |>
      filter(str_detect(...5, "Trial Year")) |>
      pull(6),

    status = descr_sheet |>
      filter(str_detect(...1, "Status:")) |>
      pull(3),

    site = list(
      address = descr_sheet |>
        filter(str_detect(...1, "Address \\(Location\\)")) |>
        pull(2),

      city = descr_sheet |>
        filter(str_detect(...1, "City")) |>
        slice_head(n = 1) |>
        pull(2),

      postal = descr_sheet |>
        filter(str_detect(...1, "Postal")) |>
        pull(2),

      latitude = descr_sheet |>
        filter(str_detect(...1, "Latitude")) |>
        pull(2),

      longitude = descr_sheet |>
        filter(str_detect(...1, "Longitude")) |>
        pull(2),

      randomization = descr_sheet |>
        filter(str_detect(...8, "Study Design")) |>
        pull(11),

      prev_crop = descr_sheet |>
        slice(which(descr_sheet[1] == "Location Quality:") + 4) |>
        pull(2)
    ),

    map = loadWorkbook(here(arm_path))@.xData$media[1] |>
      readPNG(),

    crop = list(
      crop = descr_sheet |>
        filter(str_detect(...1, "Crop  1:")) |>
        pull(8),

      crop_eppo = descr_sheet |>
        filter(str_detect(...1, "Crop  1:")) |>
        pull(3),

      tsw = descr_sheet |>
        filter(str_detect(...8, "1000 Seed Weight")) |>
        pull(9) |>
        as.numeric(),

      tsw_unit = descr_sheet |>
        filter(str_detect(...8, "1000 Seed Weight")) |>
        pull(10)
    ),

    sowing = list(
      sowing_depth = descr_sheet |>
        filter(str_detect(...1, "Depth:")) |>
        pull(4) |>
        as.numeric(),

      sowing_date = descr_sheet |>
        filter(str_detect(...1, "Planting Date")) |>
        pull(4) |>
        dmy(),

      sowing_rate = descr_sheet |>
        filter(str_detect(...8, "Planting Rate")) |>
        pull(9) |>
        as.numeric(),

      sowing_rate_unit = descr_sheet |>
        filter(str_detect(...8, "Planting Rate")) |>
        pull(11),

      sowing_density = descr_sheet |>
        filter(str_detect(...8, "Planting Density")) |>
        pull(9) |>
        as.numeric(),

      row_spacing = descr_sheet |>
        filter(str_detect(...1, "Row Spacing")) |>
        pull(4) |>
        as.numeric(),

      spacing_within_row = descr_sheet |>
        filter(str_detect(...1, "Spacing within Row")) |>
        pull(4) |>
        as.numeric(),

      tillage = descr_sheet |>
        filter(str_detect(...8, "Tillage Type")) |>
        pull(11)
    ),

    plots = list(
      plot_width = descr_sheet |>
        filter(str_detect(...1, "Plot Width")) |>
        pull(2) |>
        as.numeric(),

      plot_length = descr_sheet |>
        filter(str_detect(...1, "Plot Length")) |>
        pull(2) |>
        as.numeric(),

      plot_area = descr_sheet |>
        filter(str_detect(...1, "Plot Width")) |>
        pull(2) |>
        as.numeric(),

      row_width = descr_sheet |>
        filter(str_detect(...1, "Row Spacing")) |>
        pull(4) |>
        as.numeric(),

      n_reps = descr_sheet |>
        filter(str_detect(...1, "Replications")) |>
        pull(2) |>
        as.numeric(),

      n_treatments = descr_sheet |>
        filter(str_detect(...3, "Treatments")) |>
        pull(4) |>
        as.numeric()
    ),

    soil = list(
      type = descr_sheet |>
        filter(str_detect(...5, "Soil Name")) |>
        pull(6),

      sand = descr_sheet |>
        filter(str_detect(...1, "Sand")) |>
        pull(2) |>
        as.numeric(),

      silt = descr_sheet |>
        filter(str_detect(...1, "Silt")) |>
        pull(2) |>
        as.numeric(),

      clay = descr_sheet |>
        filter(str_detect(...1, "Clay")) |>
        pull(2) |>
        as.numeric(),

      carbon = descr_sheet |>
        filter(str_detect(...3, "Carbon")) |>
        pull(4) |>
        as.numeric(),

      lime = descr_sheet |>
        filter(str_detect(...3, "Lime")) |>
        pull(4) |>
        as.numeric(),

      pH = descr_sheet |>
        filter(str_detect(...3, "pH")) |>
        pull(4) |>
        as.numeric()
    )
  )
  writePNG(md$map, "trial_map.png") # Overwriting?
  return(md)
}

#' Parse trial maintenance activities
#'
#' Extracts information on maintenance events and management practices from the description data. Intended for generating treatment/operation tables.
#'
#' @param descr_df Tibble with trial metadata, as from \code{read_descr_sheet}.
#'
#' @return Data frame or list of parsed maintenance events, with dates and categories.
#' @export
parse_maintenance <- function(descr_sheet) {
  maint_start = descr_sheet |>
    slice_tail(n = -(which(descr_sheet[1] == "Maintenance") + 2))

  maintenance = maint_start |>
    slice_head(n = which(is.na(maint_start[[1]]))[1] - 1) |>
    select(1:13)

  names(maintenance) = c(
    "no",
    "date",
    "type",
    "product_name",
    "conc",
    "unit",
    "form_type",
    "reg_number",
    "descr",
    "rate",
    "rate_unit",
    "tank_mix",
    "comment"
  )

  maintenance <- maintenance |>
    mutate(
      no = as.numeric(no),
      date = dmy(date),
      conc = as.numeric(conc),
      rate = as.numeric(rate),
      tank_mix = as.logical(case_when(
        str_detect(tank_mix, regex("yes", ignore_case = T)) ~ T,
        str_detect(tank_mix, regex("no", ignore_case = T)) ~ F
      ))
    )
}
