# Insurance_stats_analyzer
## Description
This script is a tool for analyzing key operational indicators (KO) of insurance companies. It processes complex Excel files provided by various insurers, cleans and structures the data, and then analyzes the key operational indicators over different periods. This allows for the comparison of the performance of various companies in the insurance market. The output of the script is a well-structured, easily interpretable report that highlights the most important analytical data for assessing the market’s state.

## Functional Description
The script performs the following key functions:
1. **File Discovery and Selection**: 
   - It scans specified directories for Excel files related to insurance companies and identifies the relevant ones based on their naming conventions or file structure.
2. **Data Cleaning**: 
   - Standardizes column names and formats.
   - Removes unnecessary rows and columns, ensuring only relevant data is retained.
   - Handles missing or empty values by filling them with zeros or other predefined values for further analysis.
3. **Data Transformation**: 
   - Converts values into appropriate formats (e.g., numerical values from text representations).
   - Applies transformations specific to certain companies or business lines (e.g., multiplying specific columns by 1000 for certain insurers).
4. **KPI Extraction**: 
   - Extracts key performance indicators (KPIs) like Net Profit (NP), Gross Profit (ZP), Acquisition Costs, and other relevant metrics from each company’s data.
   - Organizes these KPIs into a structured format to enable comparisons across companies and time periods.
5. **Report Generation**: 
   - Generates a structured Excel report that compares the extracted KPIs across different companies and periods.
   - The report includes conditional formatting, adjusted column widths, merged cells, and clear headings for ease of use and better readability.
   - The output is a comprehensive, visually appealing, and easy-to-understand summary of the analyzed data.

## How It Works
1. **File Processing**: 
   - The script identifies the necessary Excel files from specified directories and loads them into memory.
   - It reads the data from the selected files, handling both `.xls` and `.xlsx` formats.
2. **Data Extraction and Cleaning**: 
   - The script scans the Excel sheets for relevant data, filtering out any irrelevant information (such as metadata or empty columns).
   - It standardizes the column names to ensure consistency across multiple files and companies.
   - Missing data is handled by either removing or filling in gaps with default values like zeroes.
3. **KPI Calculation**: 
   - Key performance indicators (KPIs) are extracted based on predefined column patterns and business rules.
   - These KPIs are normalized across companies to enable valid comparisons, regardless of different data structures.
4. **Report Generation**: 
   - The script uses the `openpyxl` library to generate a new Excel report that is both readable and well-structured.
   - It applies formatting like column width adjustment, bold text for headers, and colored cells to highlight important metrics.
   - The generated report allows stakeholders to quickly assess the performance of multiple companies and periods at a glance.

## Input Structure
To run the script, the following parameters need to be provided:
1. **Directory Path**: Path to the directory containing the Excel files of the insurance companies.
2. **File Naming Patterns**: Patterns that help identify relevant files for analysis.
3. **Sheet Name Patterns**: Specific patterns that help select the relevant sheet(s) within each file.
4. **Company-Specific Adjustments**: Define transformations or adjustments needed for certain companies' data (e.g., multiplying columns by specific values).

The script is designed to handle subfolders containing multiple Excel files, and each file may contain multiple sheets with different data structures.

## Technical Requirements
To run the script, the following are required:
1. Python 3.x
2. Required libraries:
   - `pandas` for data manipulation.
   - `xlrd` for reading `.xls` files.
   - `openpyxl` for working with `.xlsx` files and generating reports.
   - `numpy` for numerical operations.
3. Directory with Excel files containing insurance data for analysis.

## Usage
1. Set the path to the directory containing the insurance data files.
2. Define any necessary file naming patterns and sheet name filters.
3. Run the script. The script will:
   - Scan the directory for relevant files.
   - Clean, organize, and process the data.
   - Generate a comprehensive Excel report comparing KPIs across different companies and periods.

## Example Output
The output will be an Excel file with:
- A well-organized table that compares KPIs across insurance companies.
- Conditional formatting to highlight key performance changes.
- Adjusted columns and merged cells for better readability.

## Conclusion
This tool significantly simplifies the process of analyzing insurance data by automating data extraction, cleaning, transformation, and report generation. It helps insurance companies and analysts make better data-driven decisions by providing them with clear, comparative insights into key business metrics across time periods and companies.
