#' Fetch EPPO codes from Opvia
#'
#' Queries Opvia for crop or pest EPPO code metadata, used to standardize crop rotation or treatment records.
#'
#' @param code Single EPPO code as a character string. Must match a code in the referenced Opvia table.
#' @param tableID UUID string referencing the correct Opvia lookup table.
#'
#' @return UUID string that refers to the record that matches the EPPO code.
#' @importFrom ropvia download_csv
#' @export
read_eppo <- function(eppo_code, tableID) {
  csv <- download_csv(tableID = tableID)

  eppo_col_idx <- which(str_detect(
    names(crops_csv),
    regex("EPPO", ignore_case = T)
  )) # Dangerous, if they think of adding another EPPO column

  rl <- get_recordlist(tableID = tableID) |>
    recordlist_to_df() |>
    mutate(
      eppo = map_chr(pluck(rl, "cells"), \(x) {
        pluck(x, eppo_col_idx, "value", .default = NA_character_)
      })
    )

  if (!eppo_code %in% rl$eppo) {
    stop(sprintf(
      "EPPO code '%s' not found in tableID '%s'",
      eppo_code,
      tableID
    ))
  }

  eppo <- rl |>
    filter(eppo == eppo_code) |>
    pull(id) # If EPPO doesn't exist, this will fail
}
