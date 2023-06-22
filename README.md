# ANA-515P-Practicum-Assignment
Here is a short description of the assignment steps:

The necessary packages are installed and loaded.
The Excel file path is set, and the data from two different sheets in the file are read into separate data frames (sheet1_data and sheet2_data).
The "Year" column in one of the sheets is converted to a consistent numeric data type to match the other sheet.
Specific rows in the "Year" column are converted to numeric to handle character values in those cells.
The two data frames are combined into one (combined_data).
Blank cells in specified columns are replaced with NA values to handle missing data.
A new Excel workbook is created, and the combined data is added to a new worksheet.
Column widths and row heights are adjusted for better visualization in the Excel file.
The modified workbook is saved as "combined_data.xlsx".
A histogram of the "Year" variable is created using ggplot2, visualizing the frequency distribution of years.
Non-finite values are removed from the "Hectares" variable, and another histogram is created to visualize its distribution.
Finally, a boxplot is created to compare the "Hectares" variable across different years.
Overall, the code demonstrates data manipulation, handling missing data, and creating visualizations to explore and analyze the data from the Excel file.
