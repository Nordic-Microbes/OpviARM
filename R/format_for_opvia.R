#' Create Opvia-compatible trial record
#'
#' Converts ARM trial metadata and additional information into a format that matches Opvia's required schema for trials. Handles mapping, sowing, and plot structure fields.
#'
#' @param trial_df Data frame with trial-level information, typically parsed directly from ARM exports.
#' @param md Named list with sublists for sowing, plots, and site; must contain all metadata required for Opvia.
#' @param map_id Opvia-generated identifier for the field map, as returned by the map upload function.
#'
#' @return A tibble with standardized trial information, ready for Opvia import.
#' @importFrom dplyr mutate select case_when
#' @importFrom glue glue
#' @importFrom httr2 resp_body_json
#' @export
format_trial <- function(trial_df, md, map_id) {
  formatted_trial <- trial_df |>
    mutate(
      `Sowing depth (cm)` = md$sowing$sowing_depth,
      `Actual sowing quantity (kg/ha)` = md$sowing$sowing_rate,
      Tillage = md$sowing$tillage,
      `Sowing date` = md$sowing$sowing_date,
      `Replicate plots` = md$plots$n_reps,
      Randomisation = case_when(
        is.na(md$site$randomization) ~ F,
        .default = T
      ),
      `Field map` = resp_body_json(map_id),
      `Field address` = glue(
        "{md$site$address}, {md$site$postal} {md$site$city}"
      ),
      `Field Location` = glue(
        "Lat: {md$site$latitude}; Long: {md$site$longitude}"
      )
    ) |>
    select(
      `Sowing depth (cm)`,
      `Actual sowing quantity (kg/ha)`,
      Tillage,
      `Sowing date`,
      `Sowing bed quality`,
      `Replicate plots`,
      Randomisation,
      `Field map`
    )

  return(formatted_trial)
}

#' Prepare soil data for Opvia ingestion
#'
#' Formats soil and site information, fetching additional metadata from Opvia and integrating previous crop EPPO codes. Useful for ensuring harmonized soil records.
#'
#' @param md List of metadata fields; must include soil (type, pH, sand, silt, clay, carbon) and site information (trial/study name, previous crop code).
#'
#' @return Tibble with one row per trial, in the structure required by Opvia soil tables.
#' @importFrom dplyr filter mutate
#' @importFrom ropvia download_csv
#' @export
format_soil <- function(md) {
  opv_soil <- download_csv("ad049118-4bfd-41ab-840c-2122b5960053") |>
    filter(Trial == md$study)

  prev_crop <- read_eppo(
    md$site$prev_crop,
    tableID = "19b9de58-1da8-4e90-aa43-4244a1a39ee7"
  )

  formatted_soil <- opv_soil |>
    mutate(
      `Soil index` = md$soil$type, # Doesn't match with Opvia select column - Remove "[:digit:]: " in Opvia
      `Reaction number (Rt)` = md$soil$pH + 0.5,
      pH = md$soil$pH,
      `Sand (%)` = md$soil$sand,
      `Silt (%)` = md$soil$silt,
      `Clay (%)` = md$soil$clay,
      `Organic matter (%)` = md$soil$carbon,
      `Previous crop` = prev_crop # Do we want multiple previous crops? Currently just the one from the year before
    )

  return(formatted_soil)
}

#' Tidy and standardize germination assessment data
#'
#' Transforms ARM germination assessment data into the format required by Opvia, performing date formatting and column renaming.
#'
#' @param germ_df Tibble with germination data, with columns expected to include `Name`, `stage_majority`, `Plot`, `Rating Date`, `value`, and `assessment`.
#'
#' @return Tibble with Opvia-compatible germination columns.
#' @importFrom dplyr mutate select rename
#' @export
format_germ <- function(germ_df) {
  formatted_germ <- germ_df |>
    mutate(
      `Rating Date` = format(`Rating Date`, "%Y-%m-%dT%H:%M:%S%z")
    ) |>
    select(Name, stage_majority, Plot, `Rating Date`, value, assessment)
  rename(
    `Treatment code` = Name, # Not correct yet! I need to find a way to couple the treatments - Identification code : Treatment code
    `Median growth stage` = stage_majority,
    `Date of assessment` = `Rating Date`,
    # If reporting basis != m2, warn
    `Emergence (plants/m2)` = value, # plants / m2  -- allegedly (what does the reporting mean?)
    `Assessment timing` = assessment
  )
}

#' Prepare yield assessment data for Opvia
#'
#' Takes yield assessment data from ARM, cleans and reshapes it to align with Opvia yield record specifications. Handles date formatting and unit conversion if necessary.
#'
#' @param yield_df Tibble with yield assessment fields: `Name`, `Plot`, `Rating Date`, `value_yield`, `value_density`, `value_prot_perc`, `value_moist_perc`, `value_tgw`.
#'
#' @return Tibble ready for upload to Opvia's yield table.
#' @importFrom dplyr mutate select rename
#' @export
format_yield <- function(yield_df) {
  formatted_yield <- yield_df |>
    mutate(
      `Rating Date` = format(`Rating Date`, "%Y-%m-%dT%H:%M:%S%z")
    ) |>
    select(
      Name,
      Plot,
      `Rating Date`,
      value_yield,
      value_density,
      value_prot_perc,
      value_moist_perc,
      value_tgw
    ) |>
    rename(
      `Treatment code` = Name, # Not correct yet! I need to find a way to couple the treatments - Identification code : Treatment code
      `Date of assessment` = `Rating Date`,
      # If reporting basis != m2, warn
      `Fresh yield (kg/m2)` = value_yield, # I need to convert this to tonne/ha (multiply by 10) - but what if they report in tonne/ha?
      `Grain density (kg/hL)` = value_density,
      `Protein content (%)` = value_prot_perc,
      `Moisture content (%)` = value_moist_perc,
      `TGW (g)` = value_tgw
    )
}
