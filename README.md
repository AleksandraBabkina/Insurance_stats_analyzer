# Insurance_stats_analyzer
## Description  
This script is a tool for analyzing the key operational indicators (KO) of insurance companies. It processes complex Excel files provided by various insurers, cleans and structures the data, and then analyzes key operational indicators over different periods. This allows for a comparative assessment of various companies in the insurance market. The script produces a structured and easy-to-read report that highlights the most important analytical insights for market evaluation.

## Functional Description  
The script performs the following key tasks:  
1. **Data Collection**: Processes multiple Excel files from various insurance companies, each containing detailed operational data.  
2. **Data Cleaning**: Removes irrelevant or erroneous entries, ensuring the accuracy and integrity of the information.  
3. **Data Structuring**: Standardizes and organizes data into a structured format for easier analysis.  
4. **Key Operational Indicator Analysis**: Examines operational indicators (KO) of insurance companies over different periods, focusing on metrics required for regulatory reporting.  
5. **Data Comparison**: Compares operational indicators across multiple insurance companies to evaluate their performance.  
6. **Report Generation**: Produces a structured and readable report that highlights critical insights, facilitating a clear understanding of market performance.  

## How It Works  
1. The script scans a specified directory containing Excel reports from different insurance companies.  
2. It identifies and extracts the relevant sheets based on predefined naming patterns.  
3. The extracted data undergoes a cleaning process to remove inconsistencies and errors.  
4. The script selects and standardizes the necessary columns related to key operational indicators.  
5. It aggregates the data across different periods and companies for comparative analysis.  
6. A final report is generated in an Excel file, summarizing the key operational indicators for each company in a structured manner.  

## Input Structure  
To run the script, the following input parameters need to be provided:  
1. **Directory Path**: The path to the folder containing insurance company reports.  
2. **Sheet Identification Rules**: Patterns used to recognize relevant report sheets.  
3. **Column Selection Rules**: Keywords to identify columns with key operational indicators.  
The script is designed to work with official reporting formats issued by the Central Bank, ensuring compliance with regulatory requirements.  

## Technical Requirements  
To run the script, the following dependencies are required:  
1. **Python 3.x**  
2. **Required libraries**:  
   - `pandas` – for data processing  
   - `numpy` – for numerical operations  
   - `openpyxl` – for working with Excel files  
   - `xlrd` – for reading older Excel formats  
   - `re` – for pattern recognition in sheet names and columns  
   - `os` – for file and directory operations  
   - `datetime` – for handling date-related operations  

## Usage  
1. Place all Excel reports in the specified folder.  
2. Modify the `path` variable in the script to point to the correct directory.  
3. Run the script. It will:  
   - Extract data from the relevant sheets.  
   - Standardize and clean the data.  
   - Analyze key operational indicators.  
   - Generate a consolidated report in Excel format.  

## Example Output  
- **Structured Excel Report**:  
  - Contains key operational indicators for each insurance company.  
  - Displays comparative performance analysis.  
  - Summarizes insights for market evaluation.  

## Conclusion  
This tool automates the collection, cleaning, and analysis of key operational indicators from multiple insurance companies. It enables comparative evaluation and market performance assessment based on official regulatory reporting. The structured output simplifies further analysis and decision-making in the insurance sector.  
