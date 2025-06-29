% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/regression.R
\name{multivariate}
\alias{multivariate}
\title{Multivariate Regression Table}
\usage{
multivariate(
  dt,
  group,
  vi_name,
  con_name,
  cat_name,
  ref = c(),
  method = "Linear",
  CI = FALSE,
  log.transform = FALSE,
  outfile = "Multivariate Analysis Results.xlsx"
)
}
\arguments{
\item{dt}{A data frame including all variables.}

\item{group}{A vector or a variable name in \code{dt} specifying treatment groups.}

\item{vi_name}{A character vector of outcome variable names.}

\item{con_name}{A character vector of continuous covariate names.}

\item{cat_name}{A character vector of categorical covariate names.}

\item{ref}{Reference group level name (optional).}

\item{method}{Regression method: "Linear" or "logistic".}

\item{CI}{Logical; whether to include confidence intervals in results.}

\item{log.transform}{Logical; whether to log-transform the outcome.}

\item{outfile}{Output Excel file path.}
}
\value{
A list with formatted result tables and regression summaries.
}
\description{
Generate a table of multivariate analysis results using linear or logistic regression.
}
