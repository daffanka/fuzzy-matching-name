# 🔍 Fuzzy Name Matching for Excel Users

This project helps **match similar but not identical names** between two sheets in Excel, solving a common limitation faced when using `VLOOKUP` or `XLOOKUP`, which only work on exact matches.

Powered by [RapidFuzz](https://github.com/maxbachmann/RapidFuzz), this script intelligently matches names even when they differ slightly — such as `"zura.mohamad"` vs `"Zura Mohamad"`.

---

## 📌 Why Use This?

Excel users often struggle when names come from different sources or systems and don’t match exactly. This script helps:
- Map usernames (e.g., from Azure) to a reference list
- Clean inconsistent name formats (dots, dashes, parentheses, etc.)
- Automate matching without manual checking

---

## ✅ Example Use Case

| UserName            | Name                     | Matched? |
|---------------------|--------------------------|----------|
| `natalie.tan`       | `Natalie Tan`            | ✅       |
| `(NATALIE) CHRSTIE` | `Natalie Christie`       | ✅       |
| `M. Ahmad Fulan`    | `Muhammad Ahmad Fulan`   | ✅       |
| `jane-doe`          | `Jane Doe`               | ✅       |

---

## 🔧 How to Use

### 💻 No Python? No Problem!

You can run everything directly in **Google Colab** — no setup needed on your device.

👉 [**Try it on Google Colab**](https://colab.research.google.com/drive/1ISGoQRRvp-3abWnX9f8SVCo7j64mZSgN?usp=sharing)

⚠️ **Important:** Please **make a copy first**  
Before editing, go to `File > Save a copy in Drive` to create your own version. This ensures the original notebook stays unchanged.

🧾 **Steps:**
1. Click the link above  
2. Go to `File` > `Save a copy in Drive`  
3. Upload your `fuzzy_match_template.xlsx` file  
4. Run all cells in the notebook  
5. Download `matched_result.xlsx` from the left-side files panel

---

### 📝 Steps:

1. Click the **Google Colab** link above  
2. Upload your `fuzzy_match_template.xlsx` file when prompted  
   _(modify it first to contain your actual data in the two sheets)_
3. Run all cells in Colab  
4. Download the `matched_result.xlsx` file from the left-side **Files** panel

---

### 🧾 Your Excel Template Format

Create or edit your `fuzzy_match_template.xlsx` to include:

#### 🟦 Sheet 1: `Names_To_Match`

| UserName            |
|---------------------|
| natalie.tan         |
| (NATALIE) CHRSTIE   |
| M. Ahmad Fulan      |

#### 🟩 Sheet 2: `Reference_Names`

| Name                     |
|--------------------------|
| Natalie Tan              |
| Natalie Christie         |
| Muhammad Ahmad Fulan     |

---

### 📤 Output

After running the script, Colab will generate:

📁 `matched_result.xlsx`  

This file includes:

| UserName            | Name                     |
|---------------------|--------------------------|
| `natalie.tan`       | `Natalie Tan`            |
| `(NATALIE) CHRSTIE` | `Natalie Christie`       |
| `M. Ahmad Fulan`    | `Muhammad Ahmad Fulan`   |

You can open this file to verify the matches and use it for further data analysis or reporting.

---

## 📁 Files in This Repo

| File                    | Description                              |
|-------------------------|------------------------------------------|
| `fuzzy_match_template.xlsx` | Sample input data in 2 sheets             |
| `match_fuzzy_names.py`      | Python script for matching names         |
| `matched_result.xlsx`       | Output file showing match results        |
| `README.md`                 | This documentation                       |

---

## 🤝 Contribute

Found a better cleaning method? More name anomalies to match?  
Fork this repo and submit a pull request — **all contributions are welcome!**

---

## 📬 Contact

If this tool helps you, feel free to ⭐ the repo or share it on **[LinkedIn](https://www.linkedin.com/in/nauvaldaffanka/)**!
