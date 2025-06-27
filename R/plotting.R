#' Normality Test and Distribution Plots
#'
#' Generates normality test results and visualizations for variables, including histograms, density plots, Q-Q plots, and boxplots.
#'
#' @param dt A data frame or matrix containing continuous variables of interest.
#' @param group A grouping vector used for stratified plots and summary statistics.
#' @param ks Logical. If TRUE, performs Kolmogorov-Smirnov test; if FALSE, uses Shapiro-Wilk test.
#' @param plot Logical. If TRUE, generates PDF plots of distributions.
#' @param table Logical. If TRUE, generates a CSV file with normality test statistics and descriptive statistics.
#' @param plot.filename A character string for the output PDF file of distribution plots.
#' @param Table.filename A character string for the output CSV file of normality test results.
#'
#' @return A data frame containing normality test results and summary statistics (if \code{table = TRUE}).
#' @export
dist.plots <- function(dt, group,
                       ks = FALSE,
                       plot = TRUE,
                       table = TRUE,
                       plot.filename = 'Distribution Plots for Variables',
                       Table.filename = 'Normality Test Results') {

  dt <- apply(dt, 2, as.numeric)
  if (table == TRUE) {
    result1 <- c()
    for (i in 1:ncol(dt)) {
      if (sum(is.na(dt[, i]) == FALSE) < 3) next
      re <- if (ks == FALSE) {
        shapiro.test(as.numeric(dt[, i]))
      } else {
        ks.test(dt[, i], "pnorm", mean = mean(dt[, i], na.rm = TRUE), sd = sd(dt[, i], na.rm = TRUE))
      }

      kt <- e1071::kurtosis(as.numeric(dt[, i]), na.rm = TRUE)
      skew <- e1071::skewness(as.numeric(dt[, i]), na.rm = TRUE)
      result1 <- rbind(result1, round(c(re$statistic, re$p.value, kt, skew), 4))
    }

    gt <- dplyr::group_by(cbind.data.frame(dt, group), group)
    sum_result <- c()
    for (i in 1:(ncol(gt) - 1)) {
      sum <- dplyr::summarize_at(gt[, 1:(ncol(gt) - 1)],
                                 dplyr::vars(names(gt)[i]),
                                 dplyr::funs(mean(., na.rm = TRUE),
                                             sd(., na.rm = TRUE),
                                             median(., na.rm = TRUE),
                                             IQR(., na.rm = TRUE),
                                             min(., na.rm = TRUE),
                                             max(., na.rm = TRUE)))
      N <- sum(is.na(gt[, i]) == FALSE)
      N_missing <- sum(is.na(gt[, i]))
      sum_result <- rbind(sum_result, cbind(N, N_missing, sum))
    }

    result <- cbind(colnames(dt), sum_result, result1)
    colnames(result) <- c('Variables', colnames(sum_result),
                          ifelse(ks == TRUE,
                                 'Kolmogorov-Smirnov test statistics',
                                 'Shapiro test statistics'),
                          'P value', 'Kurtosis', 'Skewness')
    write.csv(result, paste0(Table.filename, '.csv'))
  }

  if (plot == TRUE) {
    dt <- as.data.frame(dt)
    pdf(paste0(plot.filename, ".pdf"))
    for (i in 1:ncol(dt)) {
      p1 <- ggplot2::ggplot(dt, ggplot2::aes(x = as.numeric(dt[, i]))) +
        ggplot2::geom_histogram(ggplot2::aes(y = ..density..),
                                color = 'Black', fill = "#56B4E9") +
        ggplot2::geom_density(alpha = .2) +
        ggplot2::labs(title = paste("Histogram and Density Curve of", names(dt)[i]),
                      x = names(dt)[i], y = "Density")

      p2 <- ggplot2::ggplot(dt, ggplot2::aes(x = as.numeric(dt[, i]), fill = group)) +
        ggplot2::geom_density(alpha = .3) +
        ggplot2::labs(title = paste("Density Curve of", names(dt)[i]),
                      x = names(dt)[i], y = "Density", fill = group)

      p3 <- ggplot2::ggplot(dt, ggplot2::aes(sample = as.numeric(dt[, i]))) +
        ggplot2::stat_qq() + ggplot2::stat_qq_line() +
        ggplot2::labs(title = paste("Q-Q plot of", names(dt)[i]),
                      x = group, y = names(dt)[i])

      p4 <- ggplot2::ggplot(dt, ggplot2::aes(sample = as.numeric(dt[, i]), colour = factor(group))) +
        ggplot2::stat_qq() + ggplot2::stat_qq_line() +
        ggplot2::labs(title = paste("Q-Q plot of", names(dt)[i]),
                      x = group, y = names(dt)[i])

      p5 <- ggplot2::ggplot(dt, ggplot2::aes(y = as.numeric(dt[, i]), x = group, color = group)) +
        ggplot2::geom_boxplot() +
        ggplot2::labs(title = paste("Box plot of", names(dt)[i]),
                      x = group, y = names(dt)[i])

      print(p1)
      print(p2)
      print(p3)
      print(p4)
      print(p5)
    }
    dev.off()
  }

  return(result)
}
