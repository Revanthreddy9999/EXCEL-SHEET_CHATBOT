# 📊 InsightSheet – Conversational Excel Analyzer

**InsightSheet** is a smart AI-powered chatbot built with Streamlit that lets you upload Excel files and interactively query your data using plain English. Whether you're looking to summarize, filter, compare, or visualize your data, InsightSheet gives you direct answers and insights — not just Python code.

---

## 🚀 Features

- 🗂 **Upload Excel Files** (`.xlsx`, `.xls`)
- 💬 **Ask Questions in Plain English**, like:
  - “What is the average age?”
  - “How many employees are from Canada?”
  - “Show a bar chart of sales by region”
- 📈 **Generates Instant Insights**:
  - One-line summaries (like max/min/average)
  - Filtered DataFrames
  - Grouped aggregations
  - Visualizations (bar, histogram, etc.)
- 🧠 **Understands Any Table Structure** — no need to know column names
- 🛡️ **Error Resilient**: Handles missing values, filters that return nothing, and unsafe code generation
- 🔍 Powered by the **Mistral API** for high-quality, code-generating LLM responses

---

## 🧠 Use Case

This app was developed for the **NeoStats AI Engineer Internship Challenge**:

> Build a Streamlit chatbot that can:
> - Accept Excel uploads
> - Analyze tabular data
> - Answer natural language questions
> - Provide insights via charts and text

---

## 🛠️ Tech Stack

- **Python 3.8+**
- **Streamlit** – interactive web app framework
- **pandas** – data wrangling
- **matplotlib** & **seaborn** – charts
- **openpyxl** – Excel file reading
- **Mistral API** – LLM for code generation

---

## 🧪 Example Questions to Try

| Question                                  | Output Type         |
|-------------------------------------------|----------------------|
| What is the maximum age?                  | One-line answer      |
| Show all employees from Canada            | Filtered table       |
| Average salary by department              | Grouped aggregation  |
| Bar chart of sales by region              | Bar chart            |
| Histogram of ages                         | Histogram plot       |

---

## 📦 Installation

Install required packages:

```bash
pip install -r requirements.txt
