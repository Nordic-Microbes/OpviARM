#' Upload a field map file to Opvia
#'
#' Submits a field map (image or PDF) to Opvia and returns a unique identifier for later linking.
#'
#' @param file Path to the field map file. Must be a valid file format accepted by Opvia (e.g., PNG, PDF).
#'
#' @return Character string; the Opvia-generated map ID.
#' @importFrom ropvia upload_file
#' @export
upload_map <- function(map_file) {
  id <- request("https://api.opvia.io/v0/files/upload") |>
    req_auth_bearer_token(ropvia::get_bearer_token()) |>
    req_url_query(filename = map_file) |>
    req_body_file(map_file, type = "image/png") |>
    req_perform()

  return(id)
}

#' Add or update records in an Opvia table
#'
#' Pushes new records or updates existing entries in an Opvia table using data from an R data frame.
#'
#' @param tableID UUID string for the target Opvia table.
#' @param data Data frame with rows and columns matching the Opvia table schema.
#'
#' @return Logical value indicating whether the upload/update succeeded.
#' @importFrom ropvia update_records
#' @export
update_table <- function(formatted_df, tableID) {
  req_list <- formatted_df |>
    df_to_payload() |>
    update_records(tableID = tableID)

  response <- map(req_list, req_perform, .progress = T) # TODO: Test map call

  return(response)
}
