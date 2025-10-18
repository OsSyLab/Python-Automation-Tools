# Web to Excel Data Automation

This Python project demonstrates how to extract a numeric value from a webpage using `requests` and `BeautifulSoup`, process it, and then write the result to a specific cell in an Excel file using `openpyxl`.

It is ideal for automating tasks like tracking online financial or statistical data and recording it directly into Excel spreadsheets.

---

## 📌 Features

- Fetches numerical data from a webpage.
- Parses HTML tables and extracts values.
- Converts annual values to monthly averages (customizable).
- Writes the result into a specified Excel cell.

---

## 🚀 Requirements

Install required Python packages:

```bash
pip install requests beautifulsoup4 openpyxl
📂 File Structure
bash
Kodu kopyala
project-folder/
├── web_to_excel.py      # Main script file
└── README.md            # Documentation
🧩 How to Use
Update the script:

In the web_to_excel.py file, go to the following line:

python
Kodu kopyala
url = "ENTER_THE_WEBPAGE_URL_TO_SCRAPE"
Replace it with the actual URL of the website you want to scrape data from.

Customize HTML parsing logic
If the structure of the table is different, you may need to update how the table is parsed using BeautifulSoup.

Set the correct Excel path and sheet name:

Update the following section with your file path and sheet name:

python
Kodu kopyala
file_path = r"PATH_TO_YOUR_EXCEL_FILE.xlsx"
sheet_name = "YourSheetName"
Run the script:

bash
Kodu kopyala
python web_to_excel.py
⚠️ Notes
Ensure the Excel file is closed before running the script.

This script assumes that the target data is in the first HTML table of the page.

The numeric format uses commas (,) for decimals. If your data uses dots (.), you may adjust the conversion logic.

📄 License
MIT License
You are free to use, modify, and distribute this code with attribution.

© 2025 Data Solutions Lab. by Osman Uluhan – All rights reserved.