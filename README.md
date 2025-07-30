# Fintech-Data-Parser

**Financial Data Parser** is a high-performance Python pipeline for parsing, analyzing, and querying financial Excel data. Built with intelligent type detection and multi-format parsing, it supports end-to-end data ingestion, cleaning, and structured querying for financial workflows.

---

## 🔍 Overview

Financial Data Parser automates:

- 🧠 Column type detection (String, Number, Date)
- 💸 Multi-format financial amount parsing
- 📆 Date recognition across multiple regional and Excel formats
- ⚡ Fast lookups and aggregations using in-memory storage
- ✅ Clean handling of null values and format inconsistencies

---

## ⚙️ Features

### 📊 Excel Intelligence
- Automatic detection of sheet info, column headers, dimensions
- Removal of empty columns

### 💱 Format Support
- **Currencies**: USD, EUR, INR, GBP, HUF, etc.
- **Negative formats**: (1,234.56), 1234.56-
- **Abbreviations**: 2.5K, 3.6M, 1.2B
- **Dates**: `DD-MM-YYYY`, `MM/DD/YYYY`, `Q1-24`, `Mar 2024`, `44927` (Excel serial)

### 🔍 Query Engine
- Range queries on parsed date/amount columns
- Group-by and aggregation
- Simple, memory-efficient design using Pandas

---

## 🛠️ Tech Stack

| Layer       | Tools Used                 |
|-------------|----------------------------|
| Core        | Python 3.12, Pandas, NumPy |
| Excel I/O   | openpyxl, pandas           |
| Parsing     | `dateutil`, `re`           |
| Storage     | In-memory via dictionaries |
| Optional    | SQLite, Streamlit (future)

---

## 🚀 Performance Summary

- ✅ Processes multiple Excel sheets with varied formats
- ✅ Handles up to ~1 million cells with confidence-based typing
- ✅ Efficient parsing pipeline using lazy parsing and sampling
- ✅ Query response in milliseconds on filtered data

---

## 🧪 Example Use Case

```python
# Load Excel
processor = ExcelProcessor(["KH_Bank.XLSX"])
processor.load_files()

# Drop empty columns
df = processor.drop_all_null_columns(processor.workbooks["KH_Bank.XLSX"].parse("Sheet1"))

# Parse amounts/dates
for col in amount_cols:
    df[col + "_parsed"] = df[col].apply(fp.parse_amount)

# Query
store = FinancialDataStore()
store.add_dataset("bank", df)
results = store.query_range("bank", "Statement.Entry.ValueDate.Date_parsed", "2024-01-01", "2024-06-30")
