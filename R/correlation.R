#' Correlation Matrix Table
#'
#' Computes and exports a correlation matrix between two sets of variables, including correlation coefficients and p-values.
#'
#' @param col A data frame or matrix containing the first set of variables (columns).
#' @param row A data frame or matrix containing the second set of variables (rows).
#' @param method A character string indicating which correlation coefficient is to be used: "pearson", "spearman", or "kendall".
#' @param outfile File path to the output Excel file.
#'
#' @return A list containing:
#' \itemize{
#'   \item A formatted table with correlation coefficients and p-values
#'   \item A compact matrix of R and P values
#' }
#' @export
corr.table <- function(col, row, method="pearson", outfile="Correlation Tables.xlsx") {
  col <- apply(apply(col, 2, as.character), 2, as.numeric)
  row <- apply(apply(row, 2, as.character), 2, as.numeric)
  col <- as.data.frame(col)
  row <- as.data.frame(row)

  library("Hmisc")
  cor <- rcorr(x = as.matrix(col), y = as.matrix(row), type = method)

  R <- apply(cor$r[(ncol(col)+1):nrow(cor$r), 1:ncol(col)], 2, round, digit = 2)
  P <- apply(cor$P[(ncol(col)+1):nrow(cor$r), 1:ncol(col)], 2, round, digit = 3)

  CM <- rbind(R, c(""), P)

  RP <- c()
  rowname <- c()
  for (i in 1:ncol(row)) {
    r <- R[i,]
    p <- P[i,]
    r[which(p < 0.05)] <- paste(r[which(p < 0.05)], "*", sep = "")
    rp <- cbind(c("Correlation Coefficient", "P value"), rbind(r, p))
    RP <- rbind(RP, "", rp)
    rowname <- c(rowname, c(colnames(row)[i], "", ""))
  }
  rownames(RP) <- rowname
  l <- list(RP, CM)

  wb <- ExcelTable(
    table = RP,
    title = "Correlation Matrix",
    caption = paste("Correlation Table. Values were ", method, " correlation coefficients. '*' P< 0.05", sep = ""),
    outfile = outfile,
    openfile = F
  )

  openXL(outfile)
  return(l)
}
