# Load necessary libraries
library(readxl)
library(openxlsx)
library(stringr)

# Set the path to your folder
folder_path <- "/Users/turkialmalki/Desktop/GCC weekly data"

# List of your exact Excel files
files <- c(
  "Data tool collection - Bahrain.xlsm",
  "Data tool collection - Oman.xlsm",
  "Data tool collection - Qatar.xlsm",
  "Data tool collection - Saudi Arabia.xlsm",
  "Data tool collection - UAE.xlsm",
  "Data tool collection - Kuwait.xlsm"
)

# Full paths to the files
file_paths <- file.path(folder_path, files)

# Initialize an empty list to store data for each disease
disease_data <- list()

# Read the names of the sheets (diseases) from the first file
first_file <- file_paths[1]
diseases <- excel_sheets(first_file)

# Function to ensure correct date format
convert_to_date <- function(date_column) {
  # Try different date formats and convert
  dates <- suppressWarnings(as.Date(date_column, format = "%d/%m/%Y"))
  if (any(is.na(dates))) {
    dates <- suppressWarnings(as.Date(date_column, format = "%Y-%m-%d"))
  }
  if (any(is.na(dates))) {
    dates <- suppressWarnings(as.Date(date_column, format = "%m/%d/%Y"))
  }
  return(dates)
}

# Loop through each file (each country)
for (file_path in file_paths) {
  # Loop through each disease (sheet)
  for (disease in diseases) {
    # Try to read the specific sheet (disease) from the current country file
    sheet_data <- tryCatch({
      read_excel(file_path, sheet = disease)
    }, error = function(e) NULL)  # Skip if sheet not found
    
    # Check if the sheet has data (non-empty)
    if (!is.null(sheet_data) && nrow(sheet_data) > 0) {
      # Clean up the country name
      country_name <- str_replace(tools::file_path_sans_ext(basename(file_path)), "Data tool collection - ", "")
      
      # Add the cleaned country name as the first column
      sheet_data <- cbind(Country = country_name, sheet_data)
      
      # Convert "Starting date of Epidemiological week" to a proper date format
      if ("Starting date of Epidemiological week" %in% names(sheet_data)) {
        sheet_data$`Starting date of Epidemiological week` <- convert_to_date(sheet_data$`Starting date of Epidemiological week`)
        
        # Identify rows with invalid dates
        invalid_dates <- is.na(sheet_data$`Starting date of Epidemiological week`)
        if (any(invalid_dates)) {
          warning(sprintf("Invalid dates found in '%s' for disease '%s' and removed.", country_name, disease))
          sheet_data <- sheet_data[!invalid_dates, ]  # Remove rows with invalid dates
        }
      }
      
      # Standardize columns if needed
      if (!is.null(disease_data[[disease]])) {
        ref_cols <- names(disease_data[[disease]])  # Use existing columns as reference
        missing_cols <- setdiff(ref_cols, names(sheet_data))
        sheet_data[missing_cols] <- NA
        sheet_data <- sheet_data[, ref_cols]
      }
      
      # Append the data to the respective disease in the list
      if (is.null(disease_data[[disease]])) {
        disease_data[[disease]] <- sheet_data
      } else {
        disease_data[[disease]] <- rbind(disease_data[[disease]], sheet_data)
      }
    }
  }
}

# Create a new Excel workbook to save the combined data
output_file <- file.path(folder_path, "merged_data_corrected.xlsx")
wb <- createWorkbook()

# Add a sheet for each disease and write the combined data
for (disease in names(disease_data)) {
  addWorksheet(wb, sheetName = disease)
  writeData(wb, sheet = disease, disease_data[[disease]])
  
  # Adjust column widths to avoid ##### in Excel
  setColWidths(wb, sheet = disease, cols = 1:ncol(disease_data[[disease]]), widths = "auto")
}

# Save the workbook in the same folder
saveWorkbook(wb, output_file, overwrite = TRUE)

cat("The combined Excel file has been created at:", output_file)