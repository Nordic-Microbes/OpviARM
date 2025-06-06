% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/parse_treatments.R
\name{read_treatment_sheet}
\alias{read_treatment_sheet}
\title{Import a treatment sheet from ARM}
\usage{
read_treatment_sheet(excel_path)
}
\arguments{
\item{file}{Path to the ARM export containing a Treatments sheet.}
}
\value{
Tibble with all treatment definitions and properties.
}
\description{
Reads the Treatments worksheet from an ARM Excel file and returns a tibble for parsing.
}
