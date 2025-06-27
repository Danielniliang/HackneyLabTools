#' Convert Specific Text Values to NA
#'
#' Replaces specified string values in a matrix or data frame with NA.
#'
#' @param vi A matrix or data frame to process.
#' @param list A character vector of values to convert to NA. Defaults to common missing indicators.
#'
#' @return A data frame with specified values replaced by NA.
#' @export
toNA <- function(vi, list = c(".", "MISSING", "na", " ", "")) {
  vi <- as.matrix(vi)
  for (i in 1:ncol(vi)) {
    col <- vi[, i]
    col[which(col %in% list)] <- NA
    vi[, i] <- col
  }
  vi <- as.data.frame(vi)
  return(vi)
}

#' Count Items in a Delimited Text Field
#'
#' Counts the number of items in each element of a text column, separated by a specified delimiter.
#'
#' @param df A data frame containing the text column.
#' @param var_name The name of the text column.
#' @param seperate_by A string used as the delimiter for splitting text. Default is space (" ").
#'
#' @return A numeric vector with the count of items in each row.
#' @export
count_items <- function(df, var_name, seperate_by = " ") {
  var <- as.character(df[, c(var_name)])
  count <- rep(0, length(var))
  for (i in 1:length(var)) {
    tmp <- str_split(var[i], seperate_by)[[1]]
    count[i] <- length(tmp[which(nchar(tmp) > 0)])
  }
  return(count)
}

#' Count Non-Missing and Missing by Group
#'
#' Returns a table of counts and missing values for each variable, grouped by a given factor.
#'
#' @param vi A data frame or matrix with variables to assess.
#' @param group A vector indicating the group for each observation.
#'
#' @return A data frame with counts and missing values by group.
#' @export
N <- function(vi, group) {
  nmiss <- c()
  for (i in 1:ncol(vi)) {
    tmp <- cbind.data.frame(is.na(vi[, i]), group)
    tmp1 <- rbind("", table(tmp$`is.na(vi[, i])`, tmp$group))
    rownames(tmp1) <- c(colnames(vi)[i], "N", "Missing")
    nmiss <- rbind(nmiss, "", tmp1)
  }
  return(nmiss)
}

#' Convert Participant ID to Six Characters
#'
#' Ensures all participant IDs are six characters long by padding with zeros.
#'
#' @param idname A character string of the ID variable name in the data frame.
#' @param dt A data frame containing the ID variable.
#'
#' @return A data frame with modified ID values.
#' @export
To_6Cs_ID <- function(idname, dt) {
  participant <- dt[, c(idname)]
  for (i in 1:length(participant)) {
    participant <- as.character(participant)
    id <- participant[i]
    if (nchar(id) == 4) {
      id <- paste(substr(id, 1, 3), "00", substr(id, 4, nchar(id)), sep = "")
    }
    if (nchar(id) == 5) {
      id <- paste(substr(id, 1, 3), "0", substr(id, 4, nchar(id)), sep = "")
    }
    participant[i] <- id
  }
  dt[, c(idname)] <- participant
  return(dt)
}

#' Export a Formatted Excel Table
#'
#' Exports a matrix or data frame to an Excel file with formatted headers and footnotes.
#'
#' @param table A table or data frame to write to Excel.
#' @param title A character title for the Excel column header.
#' @param caption A caption to add as a footer.
#' @param outfile The full path for the output Excel file.
#' @param openfile Logical indicating whether to open the file after saving. Default is TRUE.
#'
#' @return The workbook object (invisible).
#' @export
ExcelTable <- function(table,
                       title = "Table.N",
                       caption = "haha",
                       outfile = outfile,
                       openfile = TRUE) {
  table <- rbind(as.matrix(cbind(rownames(table), table)),
                 c(caption, rep("", ncol(table))))
  x <- rbind(colnames(table), table)
  colnames(x) <- c(title, rep("", ncol(x) - 1))

  wb <- write.xlsx(x, outfile)

  First2 <- createStyle(border = "Bottom", borderColour = "black", borderStyle = "thin", textDecoration = "bold")
  Border <- createStyle(border = "Bottom", borderColour = "black", borderStyle = "thin")
  wraptext <- createStyle(wrapText = TRUE, valign = 'top', halign = 'left')

  addStyle(wb, sheet = 1, First2, rows = 1, cols = 1:ncol(x), gridExpand = TRUE)
  addStyle(wb, sheet = 1, First2, rows = 2, cols = 2:ncol(x), gridExpand = TRUE)
  addStyle(wb, sheet = 1, Border, rows = nrow(x), cols = 1:ncol(x), gridExpand = TRUE)

  mergeCells(wb, 1, cols = 1:ncol(x), rows = 1)
  mergeCells(wb, 1, cols = 1:ncol(x), rows = (nrow(x) + 1):(nrow(x) + 4))
  addStyle(wb, 1, wraptext, cols = 1:ncol(x), rows = nrow(x) + 1)

  setColWidths(wb, 1, cols = 1:ncol(x), widths = "auto", ignoreMergedCells = TRUE)
  modifyBaseFont(wb, fontSize = 12, fontColour = "black", fontName = "Times New Roman")
  saveWorkbook(wb, outfile, overwrite = TRUE)

  if (openfile == TRUE) openXL(outfile)
  return(invisible(wb))
}
