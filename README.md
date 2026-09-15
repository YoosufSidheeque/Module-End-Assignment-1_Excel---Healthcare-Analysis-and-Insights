# Healthcare Data Analysis and Insights Pipeline (Excel Engine)

An end-to-end data pipeline, transformation framework, and executive reporting engine built entirely within Microsoft Excel. This repository showcases the rigorous engineering required to take a noisy, unvetted clinical and billing dataset of 250 longitudinal patient registers and transform it into an analytical data product.

## 📌 Project Overview
* **Project Title:** Healthcare Data Analysis and Insights
* **Problem Statement:** The healthcare sector scales massive transaction and diagnostic volumes daily. This project implements programmatic quality control, feature engineering, structural schema unification, and multi-axis analytical dashboards. These steps translate raw transactional fields into high-utility patient profiles and operational cost vectors.

---

## 🛠️ Data Pipeline Architecture & Methodologies

### Phase 1: Data Cleansing & Quality Control
Isolating and safely resolving non-numeric anomalies (`?` indicators and blank cells) to fix dataset skew without introducing bias.

* **Null Audit Baseline:** Applied literal criteria pattern matching using the tilde (`~`) escape character to handle wildcards and measure data sparsity.
  ```excel
  =COUNTIF(Raw_Dataset!A:A, "~?")
  ```
* **Imputation Strategy for Continuous Fields (Year):** Replaced missing time vectors with a rounded arithmetic mean using array-safe exclusion bounds.
  ```excel
  =ROUND(AVERAGEIF(K2:K251, "<>?", K2:K251), 0)
  ```
* **Imputation Strategy for Text Categories (Month, Smoker, Tiers, State ID):** Set conditional defaults. Used `"Sep"` for months, `"Unknown"` for missing State IDs, and applied the mathematical **Mode** via a matrix look-up array for regional identifiers:
  ```excel
  =INDEX(E2:E251, MODE(IF(E2:E251<>"?", MATCH(E2:E251, E2:E251, 0))))
  ```

### Phase 2: Structural Transformation & Feature Engineering
Converting strings and raw values into standardized, queryable analytics variables.

* **Schema Splitting (Customer Names):** Used the Excel **Text to Columns Wizard** (delimited by spaces) to break full name strings into structured components: `[Title]`, `[First Name]`, and `[Last Name]`.
* **String-to-Numeric Casting:** Sanitized non-numeric operational fields (e.g., converting text strings like `"None"`, `"One"`, or `"Two"` into mathematical integers):
  ```excel
  =IFS(OR(F2="No", F2="None", F2="?", ISBLANK(F2)), 0, F2="One", 1, F2="Two", 2, TRUE, VALUE(F2))
  ```
* **Clinical Group Tiers:** Implemented logical cascading tiers to group complex physiological markers into clinical categories:
  * **BMI Weight Status:**
    ```excel
    =IFS(B2<18.5, "Underweight", B2<=24.9, "Normal Weight", B2<=29.9, "Overweight", B2>=30, "Obesity")
    ```
  * **Diabetes Status (HbA1C Levels):**
    ```excel
    =IFS(E2<5.7, "Normal", E2<=6.4, "Prediabetes", E2>=6.5, "Diabetes")
    ```
* **Time-Series Serialization (Date of Birth):** Unified split year, month, and date columns into actual serial numbers, formatted as `dd-mmm-yyyy`:
  ```excel
  =DATEVALUE(I2 & "-" & J2 & "-" & K2)
  ```
* **Demographic Age Generation:** Calculated completed patient lifespans using a fixed historical data collection checkpoint date (`8th June 2023`):
  ```excel
  =DATEDIF(K2, DATE(2023, 6, 8), "Y")
  ```

### Phase 3: Relational Unification (`Healthcare` Sheet)
Created a consolidated tabular core on a new sheet named `Healthcare`. Using `Customer ID` as the primary key, we joined variables across three source worksheets via exact-match relational arrays:
```excel
=VLOOKUP(\$A2, 'Customer Names'!A:D, 3, FALSE)
```
*Note: All data values were locked using a **Paste as Values** pass to separate the database from its calculation dependencies and optimize performance.*

---

## 📈 Multi-Axis Executive Dashboards

The consolidated sheet feeds automated pivot arrays, visualized with custom presentation charts:

### 1. Epidemiological Distribution: Cancer History Among Smokers
* **Pivot Settings:** `Rows: smoker` | `Columns: Cancer history` | `Values: Count of Customer ID`. Values were normalized to `% of Row Total` to accurately evaluate structural differences between categories.
* **Visual Representation:** Two localized side-by-side **Donut Charts** (or a unified **100% Stacked Bar**), illustrating how cancer rates shift between groups.

### 2. Operational Loading: Surgeries & Glycemic Profiles vs. Transplant History
* **Pivot Settings:** `Rows: Any Transplants` | `Values: Sum of NumberOfMajorSurgeries`, `Average of HBA1C`.
* **Visual Representation:** A **Dual-Axis Combo Chart**. This uses a *Clustered Column Chart* on the primary axis to measure the large scale of surgery counts, and overlays a *Trend Line Chart* on the secondary vertical axis to display decimal-scale HbA1C variations without breaking the visual scale.

---

## 📂 Sheet Hierarchy Map
* 📁 `Raw_Dataset` → The raw baseline containing initial data and anomalies.
* 📁 `Evaluation_Instructions` → Standard operating criteria and grading targets.
* 📁 `Healthcare` → The unified data warehouse sheet (250 cleaned records, 17 columns).
* 📁 `Analysis_Visuals` → The presentation layer with pivot structures and charts.

---
### 🛠️ Execution Requirements
* Microsoft Excel 2019, 2021, or Microsoft 365.
* Macros are not required (the pipeline runs entirely on native formula functions).
