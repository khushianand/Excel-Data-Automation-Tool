# 🛡️ Cybersec Excel Automation Suite

### 🚀 Tired of manual data wrangling? This is the tool for you.

This **Python-based desktop application** is an all-in-one solution designed for **cybersecurity professionals and analysts** drowning in vulnerability data.  

It transforms fragmented, raw scan reports into a **clean, actionable, and centralized source of truth**.  
Stop wasting time on manual entry, tedious filtering, and repetitive copy-pasting — and start focusing on what really matters: **threat remediation**. ⚡

---

### ✨ Key Features at a Glance

| Feature | What It Does |
| :--- | :--- |
| **Intelligent Compare & Split** | 🔍 Automatically compares your latest scan data against your existing tracking file. It intelligently identifies **new, old, and recurring vulnerabilities** and separates them into distinct sheets. No more hunting for new threats. |
| **Dynamic Row Highlighting** | 🎨 Instantly highlight rows that match across different files. Use this to quickly verify remediation efforts or visually track specific vulnerabilities. **See your progress at a glance.** |
| **Automated Data Enrichment** | 📍 Enrich your raw data with critical context. This feature automatically maps IP addresses to their corresponding **sites and products**, giving you immediate clarity on asset ownership and location. |
| **Seamless Data Integration** | ➡️ Effortlessly append new vulnerabilities to your master tracking file. The tool automatically adds new rows and **highlights them in a color of your choice**, ensuring your vulnerability log is always up-to-date and easy to audit. |

---



## ⚙️ Installation

### 1️⃣ Clone the Repository
```sh
git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name
```

### 2️⃣ Setup Virtual Environment (Recommended)
```sh
python -m venv .venv

# Windows
.venv\Scripts\activate

# Linux/Mac
source .venv/bin/activate
```

### 3️⃣ Install Dependencies
```sh
pip install -r requirements.txt
# or manually:
pip install pandas openpyxl
```

### 4️⃣ Run the App
```sh
python main.py
```

✅ The intuitive GUI will launch 🚀

---

## 🧭 How It Works (Step by Step)

1. 📂 **Select Tracking File** (master workbook)  
2. 📂 **Select Raw Scan File** (latest Nessus/Qualys export)  
3. 🖱️ **Choose sheets to process** (via dropdown in GUI)  
4. ⚖️ **Run Compare** → Splits into *New, Existing, Duplicates*  
5. 🌍 *(Optional)* **Apply Enrichment** with site/product mapping  
6. 🎨 **Append results** to tracking file with color highlights  

---

## 📁 Example Output

After running compare, you’ll get:

- ✅ **New_Vulnerabilities**  
- ♻️ **Existing_Vulnerabilities**  
- 📌 **Duplicate_Entries**  
- 🌍 **Enriched_Data**  
- 📊 **Summary**  

---

## 🧾 License

Feel free to open an issue or submit a pull request.

Use, modify, and share freely — just give credit. 🙌

---

## 👨‍💻 Developer

<div align="center">

💡 Developed with ❤️ by **Khushi Anand**  

📧 Email: *khushianand0911@gmail.com*  

🌐 GitHub: [github.com/khushianand](https://github.com/khushianand)  

🔗 LinkedIn: [linkedin.com/in/khushianand091101](www.linkedin.com/in/khushianand091101)  

</div>

---

## 🤝 Contribute & Collaborate

- 🐛 Found a bug? **Open an issue**  
- 💡 Have an idea? **Raise a feature request**  
- 🔥 Want to improve? **Submit a PR**  

Let’s make this the go-to **Cybersecurity Excel Automation Toolkit** together. 🚀
