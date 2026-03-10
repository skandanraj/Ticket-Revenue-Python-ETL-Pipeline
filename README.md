# 📊 Ticket Revenue Reconciliation Pipeline

A Python-based **ETL pipeline** designed to reconcile **support ticket data with revenue records**.  
This automation identifies whether a **support ticket resulted in a revenue-generating service** by matching operational ticket data with billing information.

---

# 📌 Project Overview

Organizations often need to understand whether **customer support interactions lead to revenue generation**.  
Manually reconciling ticket records with billing data can be time-consuming and error-prone.

This project automates the process by:

- Loading ticket and revenue datasets
- Cleaning and transforming the data
- Matching records using **UHID or phone number**
- Validating results with automated Excel checks
- Exporting a structured reconciliation report

---

# ⚙️ Workflow

## 1️⃣ Data Extraction

The pipeline loads two Excel datasets:

```
Tickets.xlsx
Revenue.xlsx
```

These contain:

| Dataset | Description |
|-------|-------------|
| 🎫 Tickets | Customer service tickets |
| 💰 Revenue | Billing and invoice records |

---

## 2️⃣ Data Transformation

Before matching, the script performs several preprocessing steps:

- Clean column names
- Convert date fields to **datetime**
- Select only relevant columns
- Combine multiple tests belonging to the **same ticket**

### Example

If a ticket contains multiple diagnostic tests:

| Ticket Id | Test |
|----------|------|
| 1001 | Blood Test |
| 1001 | MRI |

It becomes:

| Ticket Id | Tests |
|----------|------|
| 1001 | Blood Test, MRI |

---

# 🔍 Revenue Matching Logic

The pipeline uses a **two-stage matching strategy**.

## 🥇 Primary Match — UHID

The script first attempts to match using:

- **UHID**
- Same **branch**
- **Invoice date ≥ ticket creation date**

UHID is prioritized because it uniquely identifies patients.

---

## 🥈 Fallback Match — Phone Number

If UHID is missing, the script attempts to match using:

- **Phone Number**
- Same **branch**
- **Invoice date ≥ ticket creation date**

---

## 🎯 Best Match Selection

If multiple revenue records satisfy the conditions:

➡ The **earliest invoice after the ticket creation date** is selected.

---

# 📂 Output

The script generates an Excel file:

```
App_Tickets_With_Revenue.xlsx
```

This file contains **two sheets**.

---

# 📊 Output Sheets

## 📑 Raw Data

Contains the **full matched dataset**.

Key fields include:

```
Ticket Id
Contact Name
UHID
Ticket Phone No
Branch
Priority
Status
Test IDs
Test Names
Created At
InvoiceNo
RegistrationNo
PatientName
InvoiceDate
Service Name
Gross Amount
Match_Type
```

### Match Types

| Match Type | Meaning |
|------------|--------|
| ✅ UHID | Matched using UHID |
| 📞 PHONE | Matched using phone number |
| ❌ NO MATCH | No revenue record found |

---

## ✔ Validation Sheet

This sheet verifies whether the match is correct.

Columns include:

```
Ticket Id
UHID
RegistrationNo
Ticket Phone No
Revenue Phone No
Branch
Source_Branch
Created At
InvoiceDate
Gross Amount
```

Excel formulas automatically check:

| Validation | Logic |
|------------|------|
| Test UHID | Ticket UHID = Revenue UHID |
| Test Phone | Ticket phone = Revenue phone |
| Test Branch | Ticket branch = Revenue branch |
| Test Date | Invoice date ≥ Ticket date |

Results return:

```
YES / NO
```

---

# 🛠 Technologies Used

- 🐍 **Python**
- 📊 **Pandas**
- 📄 **OpenPyXL**
- 📈 **Microsoft Excel**

---

# 📁 Project Structure

```
ticket-revenue-reconciliation
│
├── input
│   ├── Tickets.xlsx
│   └── Revenue.xlsx
│
├── output
│   └── App_Tickets_With_Revenue.xlsx
│
├── ticket_revenue_pipeline.py
└── README.md
```

---

# ▶ How to Run

## 1️⃣ Install dependencies

```bash
pip install pandas openpyxl
```

---

## 2️⃣ Update file paths in the script

```python
ticket_input = "input folder path/Tickets.xlsx"
revenue_input = "input folder path/Revenue.xlsx"
final_output = "output folder path/App_Tickets_With_Revenue.xlsx"
```

---

## 3️⃣ Run the script

```bash
python ticket_revenue_pipeline.py
```

---

# 💡 Use Case

This pipeline helps operations and analytics teams:

- 📊 Track **which support tickets generate revenue**
- 🔍 Perform **data reconciliation between operational and billing systems**
- ⚡ Reduce **manual verification work**
- ✔ Improve **data accuracy and auditing**

---

# 📝 Notes

- Dummy or anonymized data can be used for demonstration.
- Matching prioritizes **UHID over phone number** for higher accuracy.
- Revenue invoices with dates **earlier than ticket creation are excluded**.

---

# 👨‍💻 Author

**Skanda N Raj**

Data Analytics & Automation Project  
Built using **Python for operational data reconciliation**
