# 📊 Natural Excel Chatbot Assistant

This project is a conversational AI assistant that allows users to upload Excel files and ask questions about their data using plain English. The assistant automatically analyzes the data and returns summaries, filters, comparisons, or charts depending on the query.

--- 

## 🚀 Features

- 📁 Upload any `.xlsx` Excel file with structured data
- 🧠 Ask questions like:
  - "What is the average salary?"
  - "How many employees are under 30?"
  - "Show a chart of sales by region"
- 📊 Returns:
  - One-line summaries
  - Filtered data tables
  - Comparison results
  - Bar, line, or histogram charts
- ✅ Works on any schema without hardcoded column names
- ♻️ Handles missing values, inconsistent column names, and mixed data types

---

## 🧠 Use Case

Designed for NeoStats’ AI Engineer Intern challenge:

> Build a Streamlit-based chatbot that can:
> - Accept Excel uploads
> - Analyze tabular data
> - Answer questions in natural language
> - Show insights as text or charts

---

## 🛠️ Tech Stack

- **Python**
- **Streamlit**
- **OpenAI GPT-3.5**
- **pandas**
- **matplotlib**
- **openpyxl**

---

## 📦 Requirements

Install required packages with:

```bash
pip install -r requirements.txt
