# Financial Collection & Reconciliation ETL Pipeline

## 📌 Business Problem
In appliance service operations, technician cash collections, net job values, and advance payments involve complex tracking. Discrepancies often occur between system-recorded net collections (`Net Collection`) and physically received cash amounts (`Amount Received`). Manual reconciliation is tedious, prone to human error, and risks losing custom remarks during data refreshes.

---

## 💡 The Solution & Architecture
Developed an automated **Financial Reconciliation & ETL Pipeline** using **Microsoft Excel Power Query (M Code)** that:
* **Automates Data Ingestion:** Pulls and transforms collection details dynamically from a designated folder repository (`COMPLAINT_DETAIL_COLLECTION`).
* **Automated Reconciliation Logic:** Compares system-calculated `Net Collection` against manual collection records (`Amount Received`) to instantly flag records as **Matched**, **Discrepancy**, or **Collection Pending**.
* **Self-Referential State Preservation:** Employs workbook table referencing (`Excel.CurrentWorkbook`) to anchor user-entered inputs (`CM NO` and `Current Remarks`), ensuring manual remarks survive automated data refreshes.

---

## 🛠️ Tech Stack & Skills Demonstrated
* **Tool:** Microsoft Excel, Power Query
* **Language:** Power Query Formula Language (M Code)
* **Concepts:** Financial Data Reconciliation, ETL Automation, Left Outer Joins, Conditional Logic, State Preservation.

---

## 📂 Repository Structure
* `collection_reconciliation_script.pq` - Contains the full Power Query M-code.
* `README.md` - Project documentation.

---

## 🚀 How to Implement
1. Clone this repository.
2. Update your local folder source path in the M-code:
   `Source = Folder.Files("C:\Users\PC\Desktop\COMPLAINT_DETAIL_COLLECTION")`
3. Open Excel, paste the M-code into the Advanced Editor, and ensure workbook queries match your table names (`Master_Collection_Table` and `Manual_Collection_Query`).