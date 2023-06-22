# Installing the packages required for the data
install.packages("readxl")
install.packages("dplyr")
install.packages("ggplot2")
install.packages("openxlsx")


# Loading the packages which are installed.
library(readxl)
library(dplyr)
library(ggplot2)
library(openxlsx)


## Setting the file path
file_path <- "E:/Data Analytics - Spring 2023/ANA 515 - Fundamentals of Data Storage/Practicum/GRAIN.xlsx"


# Reading the data from the Excel file
sheet1_data <- read_excel(file_path, sheet = "Sheet1")
sheet2_data <- read_excel(file_path, sheet = "Sheet2")


# Here  One sheet seems to have the "Year" column as a double (numeric) data type, while the other sheet has it as a character data type. To resolve this, I can convert the "Year" column to a consistent data type before combining the sheets.
# Convert "Year" column in sheet2_data to numeric
sheet2_data$Year <- as.numeric(sheet2_data$Year)


# Convert specific rows (246, 247, 248 and 249) in the "Year" column to numeric. This updated code will convert rows 246, 247, 248, and 249 in the "Year" column to numeric by specifying the range [246:249]. Since, those cells have character.
# Converting the datatype in specific rows with out changing any other because everything is important, so without changing any format just changed specific cells.
sheet2_data$Year[246:249] <- as.numeric(sheet2_data$Year[246:249])


# Combine the two data frames into one. We have two separate sheets so combined both the sheets into single sheet.
combined_data <- bind_rows(sheet1_data, sheet2_data)


# Specifying the columns where I want to add NA values to the blank cells
columns_to_fill <- c(4, 5, 6, 7, 8)


# Loop through the specified columns and replace blank cells with NA
# In this column, I adjusted the columns_to_fill vector to include the specific columns 4, 5, 6, 7, and 8, where I want to add NA values.
# Here, the logical condition combined_data[, col] == "" | is.na(combined_data[, col]) is used. It checks both for blank cells ("") and missing values (NA) in the specified columns. The is.na() function is used to identify the missing values.
for (col in columns_to_fill) {
  combined_data[combined_data[, col] == "" | is.na(combined_data[, col]), col] <- NA
}


# Create a new Excel workbook
wb <- createWorkbook()

# Add the combined_data to a new worksheet
addWorksheet(wb, "Combined Data")
writeData(wb, sheet = "Combined Data", x = combined_data)

# Adjust column widths
col_widths <- c(20, 20, 20, 20, 20, 30, 20, 20, 20, 50)  
for (col in 1:length(col_widths)) 
{setColWidths(wb, sheet = "Combined Data", cols = col, widths = col_widths[col])}

# Adjust row heights
row_heights <- rep(60, nrow(combined_data))  
for (row in 1:length(row_heights)) 
{setRowHeights(wb, sheet = "Combined Data", rows = row, heights = row_heights[row])}

# Get the column names
column_names <- names(combined_data)

# Save the workbook
saveWorkbook(wb, "combined_data.xlsx", overwrite = TRUE)

# Moving on to data visualization, we create a histogram of the "Year" variable using ggplot() and geom_histogram(). We set the fill color, border color, number of bins, and remove non-finite values using na.rm = TRUE.
# We then remove non-finite values from the "Hectares" variable in combined_data and create another histogram for this variable.
# Lastly, we create a boxplot to visualize the distribution of "Hectares" across different years using ggplot() and geom_boxplot(). We set the fill color and border color for the boxplot.
  
# Histogram:
# In this code, we use the ggplot() function to create the base plot object and specify the data and aesthetic mappings. For the histogram, we use geom_histogram() to visualize the distribution of the variables. The fill argument sets the color of the bars, and the color argument sets the color of the borders of the bars. The labs() function is used to set the title and axis labels for the plot.
# Setting na.rm = TRUE will remove the non-finite values before creating the histogram.
ggplot(combined_data, aes(x = Year)) +
geom_histogram(fill = "lightblue", color = "black", bins = 10, na.rm = TRUE) +
labs(x = "Year", y = "Frequency", title = "Histogram of Year")


# Remove non-finite values
cleaned_data <- combined_data[is.finite(combined_data$Hectares), ]  


# Histogram:
# This code creates a histogram of the "Hectares" variable after removing non-finite values.
ggplot(data = cleaned_data, aes(x = Hectares)) +
geom_histogram(fill = "steelblue", color = "black") +
labs(title = "Histogram of Hectares", x = "Hectors", y = "Frequency")


# Boxplot
# This code creates a boxplot to visualize the distribution of "Hectares" variable across different years.
ggplot(cleaned_data, aes(x = as.factor(Year), y = Hectares, group = as.factor(Year))) +
geom_boxplot(fill = "lightblue", color = "black") +
labs(x = "Year", y = "Hectares", title = "Boxplot of Year by Hectares")
