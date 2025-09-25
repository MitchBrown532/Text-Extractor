# Text Extractor

A Python-based utility for extracting structured data from large, inconsistently formatted `.txt` files.  
Originally developed for **RBC Bank** to automate the processing of extensive text records, replacing manual sorting with a faster, regex-driven workflow.

---

## 🔍 Overview
Manually parsing large text files can be error-prone and time-consuming, especially when formatting varies.  
This project leverages **Python** and **regular expressions (regex)** to:  
- Detect and handle formatting inconsistencies.  
- Extract relevant fields into a structured format.  
- Automate sorting and filtering for downstream processing.  

⚠️ **Note**: This extractor was designed for a **specific set of confidential RBC files**.  
Due to an **NDA**, those files cannot be shared. As a result, the provided regex patterns and parsing logic are tailored to that dataset and **may not work on arbitrary `.txt` files** without modification.  

---

## ⚡ Features
- Handles **extensively large `.txt` files** without crashing.  
- Regex-driven parsing of irregular text structures.  
- Outputs clean, structured results for analysis.  
- Configurable for similar text record formats (with regex adjustments).  

---

## 🛠️ Tech Stack
- **Python 3**  
- **Regex (re module)**  
- [Optional] Pandas for structured data output (CSV/Excel).  

---

## 🚀 Usage
1. Clone the repo:
   ```bash
   git clone https://github.com/your-username/text-extractor.git
   cd text-extractor
   ```
2. Place your input .txt files in the data/ directory.
3. Run the extractor:
   ```bash
   python extractor.py
   ```
---
