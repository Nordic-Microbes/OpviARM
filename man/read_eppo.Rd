% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/read_from_opvia.R
\name{read_eppo}
\alias{read_eppo}
\title{Fetch EPPO codes from Opvia}
\usage{
read_eppo(eppo_code, tableID)
}
\arguments{
\item{tableID}{UUID string referencing the correct Opvia lookup table.}

\item{code}{Single EPPO code as a character string. Must match a code in the referenced Opvia table.}
}
\value{
UUID string that refers to the record that matches the EPPO code.
}
\description{
Queries Opvia for crop or pest EPPO code metadata, used to standardize crop rotation or treatment records.
}
