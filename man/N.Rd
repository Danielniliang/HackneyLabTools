% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/utility.R
\name{N}
\alias{N}
\title{Count Non-Missing and Missing by Group}
\usage{
N(vi, group)
}
\arguments{
\item{vi}{A data frame or matrix with variables to assess.}

\item{group}{A vector indicating the group for each observation.}
}
\value{
A data frame with counts and missing values by group.
}
\description{
Returns a table of counts and missing values for each variable, grouped by a given factor.
}
