# 🚀 AutoBI – PowerBI Automation & Excel Query Assistant

**AutoBI** is a powerful automation system developed during an internship at **Bajaj Mukand Ltd.** It simplifies the creation of **individual and combined PowerBI dashboards** from SAP-exported Excel data, and includes a **natural language query assistant** — all accessible through a clean, user-friendly web application.

Designed for both technical and non-technical users, **AutoBI** reduces manual effort, accelerates reporting, and makes insights accessible to everyone.

---

## ✨ Features

- 📊 **Individual Dashboard Automation**  
  Generate PowerBI dashboards for each heat number using standardized templates with minimal clicks.

- 📁 **Combined Dashboard Generation**  
  Merge multiple Excel files and refresh a unified PowerBI dashboard in seconds.

- 🤖 **Natural Language Query Assistant**  
  Query Excel data using simple English inputs like:  
  `"Heat - N6491, Material = Steel, Quantities > 10"`

- 🖥 **Web App Interface**  
  Easy-to-use Python web app that handles data combining, file linking, and natural language queries.

- 📂 **Structured Workflow**  
  Folder-based design ensures consistency and minimizes errors during automation.

---

## 💡 Use Cases

- **Production Reporting**: Monitor key data points such as material usage, heat numbers, and production metrics.
- **Cost Analysis**: Generate reports for actual quantities, pricing, and unit-level costs across batches.
- **Section-Wise Performance**: Analyze production sections (e.g., UHPF, CON, LRF) individually or in a combined view.
- **Self-Service Insights**: Allow non-technical users to retrieve specific data using natural language — no Excel or PowerBI expertise needed.

---

## ✅ Advantages

- ⏱ **Time-Saving** – Eliminates repetitive dashboard setup and manual data linking.
- 🧠 **Non-Technical Friendly** – Designed for operational users with no coding experience.
- 🧩 **Modular & Scalable** – Easy to extend for new templates, sections, or departments.
- 🔄 **Fully Integrated** – SAP → Excel → PowerBI → Query Assistant — all connected in one flow.
- 🛡 **Error-Minimizing** – Ensures clean, repeatable processes using naming conventions and automation.

---

## ⚙️ Setup Instructions

To ensure seamless execution:

1. 📦 **Install Dependencies**  
   ```bash
   pip install -r requirements.txt
📖 Follow the Full Guide
Setup, folder structure, usage flow, and best practices are detailed in instructions.md.

📂 Folder Overview

FINAL INTERNSHIP PROJECT/

├── indivisual dashboards/

├── Templates for Indivisual Dashboards/

├── Final Web Application/

  ├── combined_streamlit_app_FINAL/
         ├──app.py
         ├──query_parser
         ├──requirements.txt
         ├──run_app.bat(Main executor)
         ├──utils
         
├── Combined_outputs/

├── Combined Copied Worksheet/

├──VSCode Scripts

├── instructions.md

└── README.md



📝 Notes:

Always save files using the correct heat number as the filename (e.g., N6491.xlsx, N6491.pbix)

Use only the designated folders for input/output to ensure automation runs correctly

Refresh the PowerBI files manually after combining data if prompted

Make changes for your desired File path in app.py in Final Web Application as it is hardcoded.

Do the same output file path change for Combined_to_Connected.exe, you will find thier code in VSCode Scripts Folder.


📣 Contributions
This project was developed during an internship at Bajaj Mukand Ltd. and is tailored for internal use, but contributions or adaptations for general use are welcome. Feel free to fork or open issues for suggestions.
