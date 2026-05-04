# 📊 Purchase-to-Pay Matching Script (SQLite → Excel)

## 🔍 Overview

This project automates the process of populating missing fields in an Excel file (`Destination.xlsx`) using data stored in a SQLite database (`myCategories.db`).

It performs:

* ✅ Exact matching
* ✅ Fuzzy matching (using RapidFuzz)
* ✅ Fallback handling (`na` when no match found)

---

## 📁 Project Structure

```
project-folder/
│
├── main.py
├── Destination.xlsx
├── sourceDataForSQLite.xlsx
├── myCategories.db (auto-created)
└── README.md
```

---

## ⚙️ Requirements

Install required Python packages:

```bash
pip install pandas openpyxl rapidfuzz xlwings
```

> Note: `xlwings` is optional. If Microsoft Excel is not installed, the script will fallback to standard file writing.

---

## 📄 Input Files

### 1. `sourceDataForSQLite.xlsx`

Contains the master data with the following columns:

* Client category
* Account Description
* QoE Main Category
* QoE Subcategory

---

### 2. `Destination.xlsx`

Contains partially filled data:

* Client category
* Account Description
* QoE Main Category (empty)
* QoE Subcategory (empty)
* Match type (exact or fuzzy) (empty)

---

## 🚀 How to Run

1. Place all files in the same folder
2. Open terminal in that folder
3. Run:

```bash
python main.py
```

---

## 🧠 Matching Logic

### 🔹 Exact Match

Matches records where:

* Client category is identical
* Account Description is identical

→ Marked as **"Exact"**

---

### 🔹 Fuzzy Match

If no exact match:

* Uses `RapidFuzz` to find similar Account Descriptions
* Applies threshold-based similarity

→ Marked as **"Fuzzy"**

---

### 🔹 No Match

If no suitable match found:

* QoE Main Category = `na`
* QoE Subcategory = `na`
* Match type = `na`

---

## ✨ Output

* `Destination.xlsx` is updated with:

  * QoE Main Category
  * QoE Subcategory
  * Match type

---

## 🔄 Excel Compatibility

| Method  | Requirement            |
| ------- | ---------------------- |
| xlwings | Microsoft Excel needed |
| pandas  | Works with any system  |

The script automatically falls back to pandas if Excel is unavailable.

---

## 🛠️ Notes

* Ensure Excel file is **closed** while running (if using pandas)
* Column names must match exactly or be standardized in code
* Fuzzy threshold can be adjusted in `main.py`

---

## 📈 Future Improvements

* Add logging
* Improve fuzzy accuracy with synonyms
* GUI interface for non-technical users

---

## 👨‍💻 Author

Python automation script for data matching and Excel processing.
