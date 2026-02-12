# Libraries
library(openxlsx2)

# Other scripts
source("./Scripts/compare_invoices.R")

# Functions

# Create new excel workbook
create_workbook <- function(sheets, ...) {
  wb <- wb_workbook(...)
  for (sheet in sheets) {
    wb$add_worksheet(sheet = sheet)
  }
  
  return(wb)
}

# Adjust all columns
adjust_cols <- function(wb) {
  sheets <- wb$sheet_names
  for (sheet in sheets) {
    wb$set_col_widths(sheet = sheet,
                      cols = 1:100,
                      widths = "auto")
  }
  
  return(wb)
}

# Colour rows by column value
colour_rows_by_type <- function(wb, sheet, data, column_name) {
  if (nrow(data) == 0) return(wb)
  
  # Find the column index
  col_idx <- which(names(data) == column_name)
  
  # Get unique values in the column
  unique_types <- unique(data[[column_name]])
  
  # Define colors for each type (customize these as needed)
  color_map <- setNames(
    c("#c3e3f7", "#e0e0e0dd"),
    unique_types[1:2]
  )
  
  # Apply colors row by row (starting from row 2 since row 1 is headers)
  for (i in seq_len(nrow(data))) {
    row_num <- i + 1  # +1 because of header row
    type_value <- data[[column_name]][i]
    fill_color <- color_map[type_value]
    
    # Apply fill to entire row
    wb$add_fill(
      sheet = sheet,
      dims = paste0("A", row_num, ":", int2col(ncol(data)), row_num),
      color = wb_color(hex = fill_color)
    )
  }
  
  # Add borders to make columns visible
  # Add borders to each cell individually
  for (row in 1:(nrow(data) + 1)) {  # Include header row
    for (col in 1:ncol(data)) {
      cell <- paste0(int2col(col), row)
      wb$add_border(
        sheet = sheet,
        dims = cell,
        left_border = "thin",
        right_border = "thin",
        top_border = "thin",
        bottom_border = "thin",
        left_color = wb_color(hex = "#888888"),
        right_color = wb_color(hex = "#888888"),
        top_color = wb_color(hex = "#aaaaaa"),
        bottom_color = wb_color(hex = "#aaaaaa")
      )
    }
  }
  
  return(wb)
}

# Add formatting to headers
format_headers <- function(wb) {
  sheets <- wb$sheet_names
  
  for (sheet in sheets) {
    # Get the number of columns in the sheet
    # We'll assume headers are in row 1
    wb$add_fill(
      sheet = sheet,
      dims = "A1:ZZ1",  # Covers up to column ZZ
      color = wb_color(hex = "#4472C4")  # Dark blue
    )
    
    wb$add_font(
      sheet = sheet,
      dims = "A1:ZZ1",
      color = wb_color(hex = "FFFFFF"),  # White text
      bold = TRUE,
      size = 11
    )
    
    # Optional: Add borders
    wb$add_border(
      sheet = sheet,
      dims = "A1:ZZ1",
      bottom_color = wb_color(hex = "000000"),
      bottom_border = "thin"
    )
  }
  
  return(wb)
}


# make Excel file

generate_xlsx <- function(old_path, new_path, sub_path, out_path) {
  # Load data
  old <- read.csv(old_path)
  new <- read.csv(new_path)
  sub <- read.csv(sub_path)
  
  compare_data <- easy_compare_data(old, new, sub)
  
  
  
  # write to excel
  wb <- create_workbook(names(compare_data)) |> 
    wb_add_data(sheet = "Gone", x = compare_data$Gone, row_names = FALSE) |> 
    wb_add_data(sheet = "New Additions", x = compare_data$`New Additions`, row_names = FALSE) |> 
    wb_add_data(sheet = "Changed", x = compare_data$Changed, row_names = FALSE) |> 
    colour_rows_by_type(sheet = "Changed", data = compare_data$Changed, column_name = "Compare.Type") |> 
    wb_add_data(sheet = "Expires soon", x = compare_data$`Expires soon`, row_names = FALSE) |> 
    wb_add_data(sheet = "View all", x = compare_data$`View all`, row_names = FALSE) |> 
    format_headers() |> 
    adjust_cols()
  
  if (!missing(out_path)) wb_save(wb, out_path)
  else return(wb)

}

# Example
#generate_xlsx(
#  old_path = "./Data/Invoice_Compare_Data Dec 2025.csv",
#  new_path = "./Data/Invoice_Compare_Data Jan 2026.csv",
#  out_path = "./Output/test.xlsx"
#)

