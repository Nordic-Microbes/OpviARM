arm_path <- here("data", "arm_export_artyield.xlsx")
sheets <- readxl::excel_sheets(arm_path)

descr <- readxl::read_xlsx(
  arm_path,
  col_names = F,
  trim_ws = F,
  sheet = which(str_detect(sheets, "Site Description"))
)

# Grab metadata
# TODO: Convert to function that sanitizes input
# DANGER: Expects fixed columns
md <- list(
  study = descr |>
    filter(str_detect(...1, "Trial ID")) |>
    pull(2),

  year = descr |>
    filter(str_detect(...5, "Trial Year")) |>
    pull(6),

  status = descr |>
    filter(str_detect(...1, "Status:")) |>
    pull(3),

  site = list(
    address = descr |>
      filter(str_detect(...1, "Address \\(Location\\)")) |>
      pull(2),

    city = descr |>
      filter(str_detect(...1, "City")) |>
      slice_head(n = 1) |>
      pull(2),

    postal = descr |>
      filter(str_detect(...1, "Postal")) |>
      pull(2),

    latitude = descr |>
      filter(str_detect(...1, "Latitude")) |>
      pull(2),

    longitude = descr |>
      filter(str_detect(...1, "Longitude")) |>
      pull(2),

    randomization = descr |>
      filter(str_detect(...8, "Study Design")) |>
      pull(11),

    prev_crop = descr |>
      slice(which(descr[1] == "Location Quality:") + 4) |>
      pull(2)
  ),

  map = openxlsx::loadWorkbook(here(arm_path))@.xData$media[1] |>
    png::readPNG(),

  crop = list(
    crop = descr |>
      filter(str_detect(...1, "Crop  1:")) |>
      pull(8),

    crop_eppo = descr |>
      filter(str_detect(...1, "Crop  1:")) |>
      pull(3),

    tsw = descr |>
      filter(str_detect(...8, "1000 Seed Weight")) |>
      pull(9) |>
      as.numeric(),

    tsw_unit = descr |>
      filter(str_detect(...8, "1000 Seed Weight")) |>
      pull(10)
  ),

  sowing = list(
    sowing_depth = descr |>
      filter(str_detect(...1, "Depth:")) |>
      pull(4) |>
      as.numeric(),

    sowing_date = descr |>
      filter(str_detect(...1, "Planting Date")) |>
      pull(4) |>
      lubridate::dmy(),

    sowing_rate = descr |>
      filter(str_detect(...8, "Planting Rate")) |>
      pull(9) |>
      as.numeric(),

    sowing_rate_unit = descr |>
      filter(str_detect(...8, "Planting Rate")) |>
      pull(11),

    sowing_density = descr |>
      filter(str_detect(...8, "Planting Density")) |>
      pull(9) |>
      as.numeric(),

    row_spacing = descr |>
      filter(str_detect(...1, "Row Spacing")) |>
      pull(4) |>
      as.numeric(),

    spacing_within_row = descr |>
      filter(str_detect(...1, "Spacing within Row")) |>
      pull(4) |>
      as.numeric(),

    tillage = descr |>
      filter(str_detect(...8, "Tillage Type")) |>
      pull(11)
  ),

  plots = list(
    plot_width = descr |>
      filter(str_detect(...1, "Plot Width")) |>
      pull(2) |>
      as.numeric(),

    plot_length = descr |>
      filter(str_detect(...1, "Plot Length")) |>
      pull(2) |>
      as.numeric(),

    plot_area = descr |>
      filter(str_detect(...1, "Plot Width")) |>
      pull(2) |>
      as.numeric(),

    row_width = descr |>
      filter(str_detect(...1, "Row Spacing")) |>
      pull(4) |>
      as.numeric(),

    n_reps = descr |>
      filter(str_detect(...1, "Replications")) |>
      pull(2) |>
      as.numeric(),

    n_treatments = descr |>
      filter(str_detect(...3, "Treatments")) |>
      pull(4) |>
      as.numeric()
  ),

  soil = list(
    type = descr |>
      filter(str_detect(...5, "Soil Name")) |>
      pull(6),

    sand = descr |>
      filter(str_detect(...1, "Sand")) |>
      pull(2) |>
      as.numeric(),

    silt = descr |>
      filter(str_detect(...1, "Silt")) |>
      pull(2) |>
      as.numeric(),

    clay = descr |>
      filter(str_detect(...1, "Clay")) |>
      pull(2) |>
      as.numeric(),

    carbon = descr |>
      filter(str_detect(...3, "Carbon")) |>
      pull(4) |>
      as.numeric(),

    lime = descr |>
      filter(str_detect(...3, "Lime")) |>
      pull(4) |>
      as.numeric(),

    pH = descr |>
      filter(str_detect(...3, "pH")) |>
      pull(4) |>
      as.numeric()
  )
)

# Grab treatments
treatments <- readxl::read_xlsx(
  arm_path,
  col_names = c(
    # Will change when Sofie updates export with id_code
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

treatments


#png::writePNG(md$map, here("trial_map.png"))

# Grab assessment data
assessmentxl <- readxl::read_xlsx(
  arm_path,
  col_names = F,
  trim_ws = F,
  sheet = which(str_detect(sheets, "Assessment"))
)

assessments <- assessmentxl |>
  slice_tail(n = -8) |> # DANGER: Expects fixed rownumber - Implement check
  slice_head(n = 19) |> # DANGER: Expects fixed rownumber - Implement check
  select_if(~ any(!is.na(.))) |>
  pivot_longer(cols = !...1, names_to = "row_id", values_to = "value") |>
  pivot_wider(names_from = ...1, values_from = value) |>
  select(-row_id) |>
  mutate(
    assessment = row_number(),
    `Rating Date` = lubridate::dmy(`Rating Date`)
  ) |>
  separate(
    col = `Crop Stage Majority/Min/Max`,
    into = c("stage_majority", "stage_min", "stage_max"),
    sep = ";",
    convert = T
  )


assess_cols <- assessmentxl |>
  slice_tail(n = -28) |> # DANGER: Expects fixed rownumber - Implement check
  slice(1) |>
  unlist() |>
  as.character()

assess_data <- assessmentxl |>
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


germ <- assess_data |>
  filter(str_detect(`SE Name`, "COT")) |>
  mutate(
    Replicate = row_number()
  )

yield <- assess_data |>
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

# Grab maintenance data
maint_start = descr |>
  slice_tail(n = -(which(descr[1] == "Maintenance") + 2))

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

maintenance |>
  mutate(
    no = as.numeric(no),
    date = lubridate::dmy(date),
    conc = as.numeric(conc),
    rate = as.numeric(rate),
    tank_mix = as.logical(case_when(
      str_detect(tank_mix, stringr::regex("yes", ignore_case = T)) ~ T,
      str_detect(tank_mix, stringr::regex("no", ignore_case = T)) ~ F
    ))
  )


# Update trial info
md$study = "NMFT25-RS-WW-01"
opv_trials <- ropvia::download_csv("2a0e36f4-3f0c-4162-b804-4ff8a5b2b506")

opv_study <- opv_trials |> filter(ID == md$study)

map_id <- httr2::request("https://api.opvia.io/v0/files/upload") |>
  httr2::req_auth_bearer_token(ropvia::get_bearer_token()) |>
  httr2::req_url_query(filename = "trial_map.png") |>
  httr2::req_body_file(here::here("trial_map.png"), type = "image/png") |> # Don't use here::here in package
  httr2::req_perform()


opv_study_upd <- opv_study |>
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
    `Field map` = httr2::resp_body_json(map_id),
    `Field address` = glue::glue(
      "{md$site$address}, {md$site$postal} {md$site$city}"
    ),
    `Field Location` = glue::glue(
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

# TODO: Implement checks
opv_study_upd |>
  ropvia::df_to_payload() |>
  ropvia::update_records(tableID = "2a0e36f4-3f0c-4162-b804-4ff8a5b2b506")


# Update soil info:
md$study = "NMFT24-RS-SB-01" # Dev only
opv_soil <- ropvia::download_csv("ad049118-4bfd-41ab-840c-2122b5960053") |>
  filter(Trial == md$study)

# Safeguarding against moving columns around
# Soooo much unnecessary downloading from the API...
crops_csv = ropvia::download_csv("19b9de58-1da8-4e90-aa43-4244a1a39ee7")
eppo_col_idx = which(str_detect(names(crops_csv), "EPPO")) # Dangerous, if they think of adding another EPPO column
crops = ropvia::get_recordlist("19b9de58-1da8-4e90-aa43-4244a1a39ee7") |>
  ropvia::recordlist_to_df() |>
  mutate(
    eppo = map_chr(pluck(crops, "cells"), \(x) {
      pluck(x, eppo_col_idx, "value", .default = NA_character_)
    })
  )

prev_crop <- crops |>
  filter(eppo == md$site$prev_crop) |>
  pull(id) # If EPPO doesn't exist, this will fail

opv_soil_upd <- opv_soil |>
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

opv_soil_upd |>
  ropvia::df_to_payload() |>
  ropvia::update_records(tableID = "19b9de58-1da8-4e90-aa43-4244a1a39ee7")


# Add assessments
md$study = "NMFT25-CF-WW-01"
opv_study <- opv_trials |> filter(ID == md$study) # Duplicate of L358 - For dev only
opv_treatments <- ropvia::download_csv("e97bf1ee-4bab-47a2-bf94-43e6546e3750")
study_treatments <- opv_study |> pull(Treatments) |> jsonlite::fromJSON()
study_treatments <- opv_treatments |>
  filter(`ID` %in% study_treatments) |>
  select(ID, `Treatment code`, UUID)

# If any assessments are available, get them ready for Opvia and send it
if (length(germ[1]) > 0) {
  opv_study_germ <- ropvia::download_csv(
    "b9175b48-5e56-463f-8661-d65d3f57b3ef"
  ) |>
    filter(`Treatment UUID` %in% study_treatments$UUID)

  # This is how I remove previously inserted data
  germ_new <- anti_join(
    germ,
    opv_study_germ,
    by = join_by(
      Name == `Treatment code`,
      Plot,
      assessment == `Assessment timing`,
      `Rating Date` == `Date of assessment`
    ) # "Name" is wrong
  ) |>
    mutate(
      `Rating Date` = format(`Rating Date`, "%Y-%m-%dT%H:%M:%S%z")
    ) |>
    select(Name, stage_majority, Plot, `Rating Date`, value, assessment)
  rename(
    `Treatment code` = Name, # Not correct yet! I need to find a way to couple the treatments - Identification code : Treatment code
    `Median growth stage` = stage_majority,
    `Date of assessment` = `Rating Date`,
    # If reporting basis != m2, warn
    `Emergence (plants/m2)` = value,
    `Assessment timing` = assessment
  )
}

# If any assessments are available, get them ready for Opvia and send it
if (length(yield[1]) > 0) {
  opv_study_yield <- ropvia::download_csv(
    "852ee4d8-e987-4624-98d2-4c38af33c5c8"
  ) |>
    filter(`Treatment UUID` %in% study_treatments$UUID)

  yield_new <- anti_join(
    yield,
    opv_study_yield,
    by = join_by(
      Name == `Treatment code`,
      Plot,
      `Rating Date` == `Date of assessment`
    ) # "Name" is wrong
  ) |>
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
      `Fresh yield (kg/m2)` = value_yield, # plants / m2  -- allegedly (what does the reporting mean?)
      `Grain density (kg/hL)` = value_density,
      `Protein content (%)` = value_prot_perc,
      `Moisture content (%)` = value_moist_perc,
      `TGW (g)` = value_tgw
    )
}
