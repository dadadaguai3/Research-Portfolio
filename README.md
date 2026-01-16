# Research Portfolio

Code, documentation, and sample data for large-scale economic database construction. 
**Projects include:** China County-level Fiscal Database, Rubber Plantation Remote Sensing, and Local Gov Debt Cleaning.

---

### 👋 About Me
I am a Master's candidate in **Public Finance** with a strong focus on **Data Engineering** and **Empirical Economics**. 

My core competitiveness lies in constructing large-scale administrative datasets from unstructured sources (PDFs, Web Portals, Satellite Imagery) to support causal inference in political economy and regional development.

**Status:** Open to **PhD / Research Assistant (RA)** positions (Fall 2026/2027).

---

## 📂 Featured Projects

### 1. [Construction of China's Nationwide County-level Fiscal Database](./01_Fiscal_Budget_Database)
**The Challenge:** Official fiscal data at the county level is often fragmented across thousands of local government websites in various formats.
* **My Solution:** Led a team to collect and clean budget reports for **3,109 administrative divisions** (2015-2024). Used **LLM API** to handle unstructured text.
* **📂 Files Included:** * `Fiscal_Data_Technical_Report.pdf`: Detailed display of raw data specifics (includes Chinese and machine-translated English versions).
    * `Sample_City_Level_Data.xlsx`: A sample of the structured dataset (prefectural-level aggregation).
    * `kimi_api_parser.py`: Python script demonstrating LLM integration for text extraction.
* **Tech Stack:** Python (Pandas), RPA (ShadowBot), LLM API.

### 2. [Local Government Debt Database Cleaning](./02_Local_Debt_Analysis)
**The Scenario:** Data is sourced from China's unified local government debt disclosure platform. Given the relatively consistent formatting of these public reports, **Regular Expressions (Regex)** were deployed for efficient batch processing.
* **My Solution:** Developed automated scripts to parse semi-structured tables into a clean panel dataset.
* **📂 Files Included:**
    * `debt_cleaning_pipeline.py`: Regex pipelines for extracting debt indicators.
    * `Sample_Special_Debt.csv`: A sample dataset focusing on **Special Debt (专项债)**.
* **Tech Stack:** Python (Regex), Stata.

---

## 🛠️ Technical Skills
* **Data Processing:** Python (Pandas, NumPy, PyPDF2), Regular Expressions, Web Scraping.
* **Econometrics:** Stata, Causal Inference.
* **Spatial Analysis:** ArcGIS.
* **AI Tools:** LLM API integration for data cleaning.

---

## 📫 Contact
* **Email:** w103065lain@163.com
