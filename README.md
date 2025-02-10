# MySQL to Excel Converter

This script loads a MySQL database from an SQL dump file, processes the data, expands JSON fields into separate columns, and exports all tables into Excel files for each tenant. The script ensures that all relational data is retained and properly formatted.

## Features
- ✅ **Automatic MySQL database creation and data import**  
- ✅ **Supports tenant-based multi-database structure**  
- ✅ **Expands JSON fields into human-readable separate columns**  
- ✅ **Handles both global and tenant-specific tables**  
- ✅ **Optimized for performance (avoids DataFrame fragmentation warnings)**  
- ✅ **Auto-adjusts column widths in Excel for better readability**  
- ✅ **Exports each tenant's data to a separate `.xlsx` file**  

## Installation

### Prerequisites
- Python 3.x
- MySQL installed and running
- Required Python packages:
  ```sh
  pip install mysql-connector-python pandas openpyxl
  ```

## Usage

1. **Clone the Repository**
   ```sh
   git clone https://github.com/NikitaVrnv/HIKO-MySQL-to-Excel.git
   cd mysql-to-excel
   ```

2. **Update Configuration**  
   Open `convert_sql_to_xlsx.py` and modify:
   - `SQL_FILE_PATH` to point to your SQL dump file.
   - `OUTPUT_DIR` to specify where Excel files should be saved.
   - `MYSQL_CONFIG` to match your MySQL credentials.

3. **Run the Script**
   ```sh
   python convert_sql_to_xlsx.py
   ```

4. **Find Your Data in Excel Files**  
   The script generates `.xlsx` files in the specified `OUTPUT_DIR`, one per tenant.

## Example Output
For a tenant named `blekastad`, the script generates:
```
output_excel/
├── blekastad.xlsx
├── brezina.xlsx
├── deml.xlsx
├── global.xlsx
└── ...
```

## JSON Expansion Example
A column containing this JSON:
```json
{
  "cs": "Český text",
  "en": "English text"
}
```
Becomes:
```
| abstract/cs   | abstract/en    |
|--------------|---------------|
| Český text   | English text  |
```

## Troubleshooting
- If you get `PerformanceWarning: DataFrame is highly fragmented`, the script has been optimized to fix this.
- If some tables are missing, ensure they are properly created in MySQL before running the script.

## License
This project is open-source and licensed under the MIT License.
