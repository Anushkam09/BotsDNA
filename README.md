## 🤖 BotsDNA - Automation Tasks Repository  

### 🌟 What is Automation?  
Automation is the process of using **software bots** to handle repetitive tasks, making life easier by saving **time⏳, reducing errors ❌, and improving efficiency ⚡**. It can be used in various fields, such as **web automation, email handling, data extraction, and file management**.  

This repo contains **real-world automation tasks**, so you can **learn, practice, and master automation!** 🔥  

---

### 🔗 About **BotsDNA**  
[BotsDNA](https://www.botsdna.com/) is a platform that provides **hands-on automation tasks** to help developers improve their skills.  

Here, you'll find scripts for various automation challenges, including:  

✅ **File Handling** 📂  
   - Creating PDFs and writing into them 📝  
   - Reading/Writing Excel files 📊  

✅ **Web Automation** 🌍  
   - Automating website interactions using bots 🤖  
   - Uploading/downloading files 📤📥  
   - Searching Google automatically 🔍  

✅ **Email Automation** 📧  
   - Sending & receiving emails 📬  

✅ **Data Handling** 📜  
   - Extracting unstructured data 🕵️  
   - Scraping information from websites 🌐  

---

### 🛠️ **Getting Started**  
#### **1️⃣ Install Required Libraries**  
First, install the necessary Python packages using:  

```bash
pip install rpaframework openpyxl python-dotenv python-docx docx2pdf
```

This will install:  
- `RPA.Browser.Selenium` → Web automation  
- `RPA.Email.ImapSmtp` → Email handling  
- `openpyxl` → Excel file operations  
- `dotenv` → Secure credentials handling  
- `python-docx` → For creating and modifying Word documents
- `docx2pdf` → For converting Word files into PDFs

---

### 🔑 **2️⃣ Securely Storing Credentials**  
⚠️ **Never hardcode passwords in scripts!** Instead, use a **`.env` file**.  

#### ✅ **Step 1: Create a `.env` file**  
In the project folder, create a new file named `.env` and add:  

```env
EMAIL_USER=your-email@gmail.com
EMAIL_PASS=your-secure-password
```

#### ✅ **Step 2: Load `.env` in Your Script**  
(just for reference, code is already there where required 😉)
Use the following Python code to **safely** read credentials:  

```python
from dotenv import load_dotenv
import os

load_dotenv()

email = os.getenv("EMAIL_USER")
password = os.getenv("EMAIL_PASS")
```

📌 **Make sure to add `.env` to `.gitignore`** so it doesn’t get uploaded to GitHub!  

---

### 🚀 **3️⃣ Running the Automation Scripts**  
Simply run any Python script from the respective task folder:  

```bash
python your_script.py
```

Each subfolder contains **a different automation task**, and you can modify them as needed!  

---

### ❓ **Want to Contribute?**  
If you have new automation ideas, feel free to **fork & contribute**! 💡✨  

📩 **Have questions?** Open an issue or reach out! 🤗  

---

**Hope you enjoy automating! 💻⚡ Let’s build smarter workflows together!** 🚀  
