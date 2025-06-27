# Oracle Fusion Statement Cleaner - Excel VBA Macro

## 📌 Project Summary

This Excel VBA macro automates the cleanup and validation of raw financial statement data exported from **Oracle Fusion (Feogen)**. The tool solved a critical operational bottleneck for our team, reducing hours of manual work each month.

### ✅ Key Achievements
- Cut down statement processing time from hours to **seconds**
- Improved **data accuracy** and reduced risk of manual errors
- Helped standardize input formats across **finance operations**
- Enhanced team visibility into **duplicate or suspicious entries**

---

## 🧠 Problem It Solved

Our monthly statements from Oracle Fusion came in a raw, inconsistent format:
- PO and Invoice numbers were merged and unstructured (e.g. `912345/INV`)
- Some values were numeric, others alphanumeric or incomplete
- Many rows required validation and cleanup before use
- Rows with zero values or quantity needed to be cleaned out
- Invoice validation (e.g., offsetting positives/negatives) was manual
- Formatting inconsistencies made review slow and error-prone

---

## 🚀 The VBA Macro Solution

The macro, named `Split_PO_INV_WithSlash`, performs the following:

🔹 **Splits** PO and Invoice numbers into structured columns  
🔹 **Identifies and flags** inconsistent or suspicious entries  
🔹 **Highlights duplicates** using conditional formatting  
🔹 **Validates reversals** with logic based on value signs  
🔹 **Removes noise**: rows with zero quantity and amount  
🔹 **Auto-formats and filters** the result for quick review

This logic was fully customized to our business rules and statement structure.

---

## 🔒 Note on Confidentiality

Due to the sensitivity of financial data and internal formatting rules:
- The actual VBA code is **not shared publicly**
- No sample Excel files or outputs are included

This repository serves as a **portfolio case study** to demonstrate the business value and technical thinking behind the solution.

---

## 🛠 Skills Used

- ✅ Excel VBA (advanced macros)
- ✅ String parsing, type-checking, and conditional logic
- ✅ Dictionary objects for duplicate tracking
- ✅ Conditional formatting via code
- ✅ Finance data logic and ERP integration awareness

---

## 👤 Author

**Badriah Jaber**  
Finance Data Automation | Excel VBA Specialist | Process Optimizer  

📎 [LinkedIn](https://www.linkedin.com/in/badriah-jaber)  

---

## 💬 Let’s Connect

If you're solving similar data automation or financial cleanup challenges, feel free to reach out! I'm passionate about using smart tools (like VBA, Power Query, and Power BI) to solve real problems.
