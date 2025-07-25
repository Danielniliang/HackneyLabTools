% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/plotting.R
\name{dist.plots}
\alias{dist.plots}
\title{Normality Test and Distribution Plots}
\usage{
dist.plots(
  dt,
  group,
  ks = FALSE,
  plot = TRUE,
  table = TRUE,
  plot.filename = "Distribution Plots for Variables",
  Table.filename = "Normality Test Results"
)
}
\arguments{
\item{dt}{A data frame or matrix containing continuous variables of interest.}

\item{group}{A grouping vector used for stratified plots and summary statistics.}

\item{ks}{Logical. If TRUE, performs Kolmogorov-Smirnov test; if FALSE, uses Shapiro-Wilk test.}

\item{plot}{Logical. If TRUE, generates PDF plots of distributions.}

\item{table}{Logical. If TRUE, generates a CSV file with normality test statistics and descriptive statistics.}

\item{plot.filename}{A character string for the output PDF file of distribution plots.}

\item{Table.filename}{A character string for the output CSV file of normality test results.}
}
\value{
A data frame containing normality test results and summary statistics (if \code{table = TRUE}).
}
\description{
Generates normality test results and visualizations for variables, including histograms, density plots, Q-Q plots, and boxplots.
}
