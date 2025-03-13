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

### Setup
1. **Clone the Repository**
   ```sh
   git clone https://github.com/NikitaVrnv/HIKO-MySQL-to-Excel.git
   cd HIKO-MySQL-to-Excel
   ```

2. **Create a Virtual Environment**
   ```sh
   python3 -m venv venv
   source venv/bin/activate
   ```

3. **Install Required Packages**
   ```sh
   pip install -r requirements.txt
   ```

4. **Configure Environment**
   ```sh
   cp .env.example .env
   nano .env
   ```
   Edit the `.env` file to match your MySQL credentials and other settings.

## Usage

1. **Place Your SQL Dump File**
   Place your SQL dump file in the `input` directory (created automatically) or specify a custom path in the `.env` file.

2. **Run the Script**
   ```sh
   python3 convert_sql_to_xlsx.py
   ```

3. **Find Your Data**
   Excel files will be generated in the `output` directory, one per tenant.

## Example Output
For a tenant named `blekastad`, the script generates:
```
output/
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
