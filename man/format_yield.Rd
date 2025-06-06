% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/format_for_opvia.R
\name{format_yield}
\alias{format_yield}
\title{Prepare yield assessment data for Opvia}
\usage{
format_yield(yield_df)
}
\arguments{
\item{yield_df}{Tibble with yield assessment fields: \code{Name}, \code{Plot}, \verb{Rating Date}, \code{value_yield}, \code{value_density}, \code{value_prot_perc}, \code{value_moist_perc}, \code{value_tgw}.}
}
\value{
Tibble ready for upload to Opvia's yield table.
}
\description{
Takes yield assessment data from ARM, cleans and reshapes it to align with Opvia yield record specifications. Handles date formatting and unit conversion if necessary.
}
