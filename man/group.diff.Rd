% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/summary_stats.R
\name{group.diff}
\alias{group.diff}
\title{Perform Univariate Group Comparison}
\usage{
group.diff(
  dt,
  group,
  median = FALSE,
  paired = FALSE,
  cohend.correct = TRUE,
  outfile = "Comparison of difference between group result.xlsx"
)
}
\arguments{
\item{dt}{A data frame or matrix of numeric variables to compare.}

\item{group}{A vector or factor indicating the grouping variable.}

\item{median}{Logical. If TRUE, uses non-parametric tests and reports medians.}

\item{paired}{Logical. If TRUE, uses paired tests.}

\item{cohend.correct}{Logical. If TRUE, applies Hedges' correction to Cohen's d.}

\item{outfile}{Output Excel file name. Default is 'Comparison of difference between group result.xlsx'.}
}
\value{
A data frame of group comparison results, and exports to Excel.
}
\description{
Conducts statistical tests to compare continuous outcomes between two groups,
providing p-values, effect sizes (Cohen's d), and summary statistics (mean/SD or median/IQR).
}
