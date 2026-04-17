# 📊 Excel VBA Data Cleaning Project

<img src="https://github.com/amolhatwar/excel-vba-data-cleaning-project/blob/8a166921b6f46b524052bc5567a66cdeb4e3eda5/Clean%20Messy%20Excel%20Data%20using%20VBA%20%20Real%20Data%20Cleaning%20Project%20(Step-by-Step).png" width="600">

---

## 🎥 Video Explanation

A complete step-by-step explanation of this project is available here:
👉 (https://youtu.be/Km_536lu6sI)

A practical project demonstrating how messy raw data can be cleaned and transformed into a structured, analysis-ready format using VBA in Microsoft Excel.

---

## 📌 Project Overview

In real-world scenarios, data is often unstructured, inconsistent, and difficult to analyze.
This project focuses on automating the data cleaning process using VBA, reducing manual effort and improving data quality.

---

## ⚠️ Problem

The raw dataset contained multiple issues:

* All data stored in a single column using "|" delimiter
* NULL and missing values
* Extra spaces and non-printable characters
* Inconsistent text formatting
* Incorrect interest rate values
* Non-standard date formats

---

## 🎯 Objective

To build an automated VBA solution that:

* Converts raw data into structured columns
* Cleans and standardizes data
* Handles missing values effectively
* Prepares the dataset for analysis

---

## 🔧 Solution (What the Script Does)

### 1. Data Splitting

* Splits pipe-separated values into multiple columns
* Removes unnecessary blank columns

### 2. Data Cleaning

* Applies Trim to remove extra spaces
* Uses Clean to remove non-printable characters
* Deletes unwanted rows

### 3. Formatting & Standardization

* Formats header row (bold + structured)
* Converts text fields into proper case
* Renames columns for clarity
* Standardizes date format (DD-MM-YYYY)

### 4. Interest Rate Normalization

* Removes "%" symbols
* Converts inconsistent values (e.g., 850 → 8.50, 0.085 → 8.50)
* Applies consistent numeric formatting

### 5. Handling Missing Values

* Identifies NULL or blank cells
* Replaces values using column averages
* Applies appropriate rounding

---

## 🛠 Tools & Technologies

* Microsoft Excel
* VBA (Visual Basic for Applications)

---

## 📁 Project Structure

```
excel-vba-data-cleaning-project/
│
├── Raw_Data.xlsx
├── Cleaned_Data.xlsx
├── VBA_Code.bas
├── README.md
```

---

## 📊 Output

After processing, the dataset becomes:

* Clean and structured
* Consistent in format
* Ready for analysis or reporting

---

## 💼 Use Case

This project is useful for:

* Data analytics beginners
* Students building portfolio projects
* Excel users learning automation
* Anyone dealing with messy datasets

---

## 🚀 Key Learnings

* Automating repetitive tasks using VBA
* Handling real-world data quality issues
* Importance of preprocessing before analysis
* Writing structured and reusable VBA scripts

---

## 👨‍💻 Created By

Amol Hatwar
Excel | VBA | Data Analytics

---

⭐ If you found this project useful, consider giving it a star.
