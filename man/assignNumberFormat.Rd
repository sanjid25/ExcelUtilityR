% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/ExcelUtility.R
\name{assignNumberFormat}
\alias{assignNumberFormat}
\title{Add number format to a column}
\usage{
assignNumberFormat(
  workbook,
  sheet,
  dta,
  column_pattern,
  format_name = "PERCENTAGE",
  ...
)
}
\arguments{
\item{workbook}{create with createWorkbook() or load with loadWorkbook("path/to/workbook.xlsx")}

\item{dta}{`data.frame` or `data.table`}

\item{column_pattern}{provide a pattern that captures the column names(s). Exact names for specific column.}

\item{format_name}{use `PERCENTAGE`, `CURRENCY`, `ACCOUNTING`}

\item{...}{`optional` add more options to the `grep` call within this function.
This might include `ignore.case = TRUE` or `invert = TRUE`}
}
\value{
the same workbook
}
\description{
Add number format to a column
}
