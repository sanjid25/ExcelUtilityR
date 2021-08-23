#' Round up to nearest n decimal places
#'
#'
#' R does not always round up from xxx.xx5 like the way you may have learnt in high school.
#' There is a good reason for it.
#' https://stackoverflow.com/questions/56487484/what-explains-the-1-decimal-place-rounding-of-x-x5-in-r
#' In case, you don't like that reason, use this function
#'
#' @param x `numeric` that need rounding
#' @param n decimal places to round by, default is 0, i.e. nearest whole number
#'
#' @return
#' @export
#'
#' @examples
#'
#' round(c(2.5, 3.5))
#' roundUp(c(2.5, 3.5))
roundUp <- function(x, n = 0){
  posneg = sign(x); z = abs(x)*10^n; z = z + 0.5; z = trunc(z); z = z/10^n; z*posneg
}

#' Anticipate and add space between words in column names
#'
#' @param dta `data.frame` or `data.table`
#'
#' @return the same data
#' @export
#'
#' @examples
spaceBetweenNames <- function(dta){
  names(dta) <- gsub("([a-z])([A-Z])", "\\1 \\2", names(dta))
  names(dta) <- gsub("_", " ", names(dta))
  names(dta) <- gsub("[.]", " ", names(dta))
  return(dta)
}

#' Add number format to a column
#'
#' @param workbook create with createWorkbook() or load with loadWorkbook("path/to/workbook.xlsx")
#' @param sheet
#' @param dta `data.frame` or `data.table`
#' @param column_pattern provide a pattern that captures the column names(s). Exact names for specific column.
#' @param format_name use `PERCENTAGE`, `CURRENCY`, `ACCOUNTING`
#' @param ... `optional` add more options to the `grep` call within this function.
#' This might include `ignore.case = TRUE` or `invert = TRUE`
#' @return the same workbook
#' @export
#'
#' @examples
assignNumberFormat <- function(workbook, sheet, dta, column_pattern, format_name = "PERCENTAGE", ...){
  column_numbers <- grep(column_pattern, names(dta), ...)
  for(i in column_numbers){
    openxlsx::addStyle(
      workbook, sheet, openxlsx::createStyle(numFmt = format_name), rows = 2:(dim(dta)[1] + 1), cols = i)
  }
  return(workbook)
}

#' Add cell style to a column
#'
#' @param workbook create with createWorkbook() or load with loadWorkbook("path/to/workbook.xlsx")
#' @param sheet
#' @param dta `data.frame` or `data.table`
#' @param column_name column that needs formatting
#' @param rule_value detault is `!=\"\"` which indicates all the cell in that column.
#' @param style_name use `input`, `negative`, `positive`. Corresponds to MS Excel. Default is `input`
#'
#' @return the same workbook
#' @export
#'
#' @examples
assignCellStyle <- function(workbook, sheet, dta, column_name, rule_value = "!=\"\"", style_name = "input"){
  if(style_name == "input"){
    openxlsx::conditionalFormatting(workbook, sheet, cols=grep(pattern = column_name, names(dta)), rows=2:(dim(dta)[1] + 1),
      rule = rule_value, style = openxlsx::createStyle(bgFill = "#FFCC99", borderColour = "#7F7F7F", border = c("top", "right", "bottom", "left"), borderStyle = "thin"))
  }
  if(style_name == "negative"){
    openxlsx::conditionalFormatting(workbook, sheet, cols=grep(pattern = column_name, names(dta)), rows=2:(dim(dta)[1] + 1),
      rule = rule_value, style = openxlsx::createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE"))
  }
  if(style_name == "positive"){
    openxlsx::conditionalFormatting(workbook, sheet, cols=grep(pattern = column_name, names(dta)), rows=2:(dim(dta)[1] + 1),
      rule = rule_value, style = openxlsx::createStyle(fontColour = "#006100", bgFill = "#C6EFCE"))
  }
  return(workbook)
}

#' Make the spreadsheet look nicer with minimal effort
#'
#' This function will bolden, wrap text and freeze the column name by default, with options customize them as parameters below if needed
#'
#' @param workbook create with createWorkbook() or load with loadWorkbook("path/to/workbook.xlsx")
#' @param sheet
#' @param dta `data.frame` or `data.table`
#' @param freeze_row_at first row to be affected by scroll
#' @param freeze_col_at first column to be affected by scroll
#' @param set_width_at numeric vector, length is the same as the number of columns in the dta
#' @param hide_columns character vector of the names of columns that need hiding
#'
#' @return the same workbook
#' @export
#'
#' @examples
prettifyColumns <- function(workbook, sheet, dta, freeze_row_at = 2, freeze_col_at = 1, set_width_at = "auto", hide_columns = "HideThisColumn"){
  hidden_columns_boolean <- grepl(paste(hide_columns, collapse = "|"), names(dta))
  input_column_numbers <- grep("Merch|Corrected|Brand|Comment", names(dta))
  for(i in input_column_numbers){
    openxlsx::addStyle(workbook, sheet, rows = 1, cols = i, style = openxlsx::createStyle(wrapText = TRUE, textDecoration = "bold",
      fgFill = "#FFCC99", borderColour = "#7F7F7F", border = c("top", "right", "bottom", "left"), borderStyle = "thin"))
  }
  #openxlsx::addStyle(workbook, sheet, rows = 1, cols = 1, style = openxlsx::createStyle(wrapText = TRUE, textDecoration = "bold", textRotation = 90, valign = "top"))
  openxlsx::freezePane(workbook, sheet, firstActiveRow = freeze_row_at, firstActiveCol = freeze_col_at)
  openxlsx::setColWidths(workbook, cols = 1:dim(dta)[2], sheet, widths = set_width_at, hidden = hidden_columns_boolean) #
  openxlsx::addStyle(workbook, sheet, rows = 1, cols = 1:dim(dta)[2], style = openxlsx::createStyle(wrapText = TRUE, textDecoration = "bold"))
  openxlsx::addFilter(workbook, sheet, rows = 1, cols = 1:dim(dta)[2])
  return(workbook)
}
