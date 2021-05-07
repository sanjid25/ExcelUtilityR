round2 <- function(x, n){ 
  posneg = sign(x); z = abs(x)*10^n; z = z + 0.5; z = trunc(z); z = z/10^n; z*posneg 
} 
 
spaceBetweenNames <- function(dta){ 
  setnames(dta, names(dta), gsub("([a-z])([A-Z])", "\\1 \\2", names(dta))) 
  setnames(dta, names(dta), gsub("_", " ", names(dta))) 
  setnames(dta, names(dta), gsub("[.]", " ", names(dta))) 
  return(dta) 
} 
 
assignNumberFormat <- function(workbook, sheet, dta, column_name, format_name = "PERCENTAGE"){ 
  addStyle(workbook, sheet, createStyle(numFmt = format_name), rows = 2:(dim(dta)[1] + 1), cols = grep(column_name, names(dta))) 
  return(workbook) 
} 
 
assignCellStyle <- function(workbook, sheet, dta, column_name, rule_value = "!=\"\"", style_name = "input"){ 
  if(style_name == "input"){ 
    conditionalFormatting(workbook, sheet, cols=grep(pattern = column_name, names(dta)), rows=2:(dim(dta)[1] + 1), 
      rule = rule_value, style = createStyle(bgFill = "#FFCC99", borderColour = "#7F7F7F", border = c("top", "right", "bottom", "left"), borderStyle = "thin")) 
  } 
  if(style_name == "negative"){ 
    conditionalFormatting(workbook, sheet, cols=grep(pattern = column_name, names(dta)), rows=2:(dim(dta)[1] + 1), 
      rule = rule_value, style = createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")) 
  } 
  if(style_name == "positive"){ 
    conditionalFormatting(workbook, sheet, cols=grep(pattern = column_name, names(dta)), rows=2:(dim(dta)[1] + 1), 
      rule = rule_value, style = createStyle(fontColour = "#006100", bgFill = "#C6EFCE")) 
  } 
  return(workbook) 
} 
 
prettifyColumns <- function(workbook, sheet, dta, freeze_row_at = 2, freeze_col_at = 5, set_width_at = "auto", hide_columns = "HideThisColumn"){ 
  hidden_columns_boolean <- grepl(paste(hide_columns, collapse = "|"), names(dta)) 
  input_column_numbers <- grep("Merch|Corrected|Brand|Comment", names(dta)) 
  for(i in input_column_numbers){ 
    addStyle(workbook, sheet, rows = 1, cols = i, style = createStyle(wrapText = TRUE, textDecoration = "bold", 
      fgFill = "#FFCC99", borderColour = "#7F7F7F", border = c("top", "right", "bottom", "left"), borderStyle = "thin")) 
  } 
  #addStyle(workbook, sheet, rows = 1, cols = 1, style = createStyle(wrapText = TRUE, textDecoration = "bold", textRotation = 90, valign = "top")) 
  freezePane(workbook, sheet, firstActiveRow = freeze_row_at, firstActiveCol = freeze_col_at) 
  setColWidths(workbook, cols = 1:dim(dta)[2], sheet, widths = set_width_at, hidden = hidden_columns_boolean) # 
  addStyle(workbook, sheet, rows = 1, cols = 1:dim(dta)[2], style = createStyle(wrapText = TRUE, textDecoration = "bold")) 
  addFilter(workbook, sheet, rows = 1, cols = 1:dim(dta)[2]) 
  return(workbook) 
} 
 