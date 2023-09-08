# Automation Script for Data Processing

## Overview
This automation script is designed to process data from a CSV file, perform various operations on it, and generate a final output in the form of a well-formatted Excel file. The script includes the following features:

- Filtering specific columns from the CSV file.
- Formatting the timestamp column from UTC to GST format.
- Applying borders and alignment to cells in the Excel file.
- Creating a zip archive of the resulting Excel file.

## Getting Started

### Prerequisites
- Python 3.x
- Required Python packages (you can install them using `pip install package-name`):
  - pandas
  - openpyxl

### Installation
1. Clone this repository to your local machine:

   ```bash
   git clone https://github.com/your-username/your-repo-name.git
   cd your-repo-name

### Install the required Python packages
pip install pandas openpyxl

### Usage
Running the Script
Run the script with the following command:
python automation.py --input input.csv --columns column1,column2,column3

Replace input.csv with your CSV file's name.
Replace column1,column2,column3 with the names of the columns you want to filter, separated by commas.
Output
The script will generate the following files:

output.xlsx: The processed data in an Excel file.
output.zip: A zip archive containing the Excel file.

### Customization
You can modify the script to suit your specific data processing needs. Refer to the code comments for details on each function's functionality.

### License
This project is licensed under the MIT License - see the LICENSE file for details.

### Acknowledgments
Openpyxl - for working with Excel files in Python.
Pandas - for data manipulation and analysis in Python.


Please make sure to replace the placeholders (e.g., `your-username`, `your-repo-name`, `input.csv`, `column1,column2,column3`) with the appropriate values for your project.

