# HackneyLabTools

A collection of R functions for statistical analysis, reporting, scoring, and visualization — developed for use in Dr. Hackney's Lab.

## Installation

You can install the package directly from GitHub:

```r
# install.packages("devtools")
devtools::install_github("Danielniliang/HackneyLabTools")
```

## Features

- Data cleaning utilities (`toNA`, `To_6Cs_ID`, etc.)
- SF-12 scoring tool
- Table generators for regression and multivariate analysis
- Correlation matrix with significance levels
- Normality test + distribution plots
- Exportable Excel tables with styling

## Usage Example

```r
library(HackneyLabTools)

# Replace '.', 'MISSING', and empty strings with NA
clean_data <- toNA(df)

# Run SF-12 scoring
scores <- SF12(my_sf12_data)

# Generate multivariate regression table
multivariate(dt = df, group = "treatment", vi_name = "outcome", con_name = c("age", "bmi"), ...)
```

## License

MIT © Liang Ni
