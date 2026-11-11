# Campus Placement Prediction & Package Regression

![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)
![Scikit-Learn](https://img.shields.io/badge/Scikit--Learn-F7931E.svg)
![XGBoost](https://img.shields.io/badge/XGBoost-11B5E5.svg)
![Pandas](https://img.shields.io/badge/Pandas-150458.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

A dual-objective Machine Learning system solving both **Placement Status Classification** (Placed vs. Not Placed) and **CTC Package Regression** (predicting offered annual salary in LPA) based on 3,000 student academic profiles, technical proficiencies, project records, and internship histories.

---

## 📌 Project Overview

1. **Classification Task:** Predicts whether a candidate will secure a campus placement using **Random Forest** and **XGBoost** classifiers with 5-Fold Stratified Cross-Validation and hyperparameter optimization (`GridSearchCV`).
2. **Regression Task:** Predicts the continuous expected CTC (Cost to Company in LPA) using **Random Forest Regressor** and **Ridge Regression**, evaluated via $R^2$, MAE, and RMSE.
3. **Feature Importance & Insights:** Identifies key factors driving recruitment decisions (CGPA, coding test performance, internship count, technical projects, backlog count).
4. **Documentation Engineering:** Includes automated reporting pipelines and the complete project report (`Campus_Placement_Project_Report.pdf`).

---

## 📊 Dataset Features (`campus_placement_dataset.csv`)

| Feature Name | Description | Type |
| :--- | :--- | :--- |
| `Gender` | Student gender | Categorical |
| `Branch` | Academic stream (CSE, IT, ECE, ME, etc.) | Categorical |
| `SSC_Percentage` | 10th grade score | Continuous (%) |
| `HSC_Percentage` | 12th grade score | Continuous (%) |
| `Degree_CGPA` | Undergraduate academic CGPA | Continuous (0–10) |
| `Internships` | Number of completed industry internships | Discrete integer |
| `Projects` | Number of software/hardware projects completed | Discrete integer |
| `Certifications` | Number of technical certifications | Discrete integer |
| `Coding_Test_Score` | Technical assessment score | Continuous (0–100) |
| `Soft_Skills_Rating` | Behavioral and communication rating | Continuous (0–5) |
| `Backlogs` | Number of active or cleared backlogs | Discrete integer |
| `Placement_Status` | Target 1: Placed (1) / Not Placed (0) | Binary label |
| `Salary_Package_LPA` | Target 2: Offered annual package in LPA | Continuous |

---

## 📁 Repository Structure

```
├── Campus_Placement_Prediction.ipynb       # End-to-end Jupyter Notebook (EDA, models, evaluation)
├── campus_placement_dataset.csv            # 3,000-record augmented student dataset
├── Campus_Placement_Project_Report.pdf     # Full academic project report artifact (964 KB)
├── Campus_Placement_Project_Report.docx    # Formatted Word documentation (1.5 MB)
├── build_guidelines_report.py              # Script generating guideline-compliant documentation
├── build_old_styled_report.py              # Custom report generator
├── generate_screenshots.py                 # Automated screenshot & figure generator
├── requirements.txt                        # Python dependencies
└── report_images/                          # Generated ROC curves, confusion matrices & feature importance charts
```

---

## 🚀 How to Run

### 1. Clone the Repository
```bash
git clone https://github.com/aashbiii/Campus-Placement-Prediction-ML.git
cd Campus-Placement-Prediction-ML
```

### 2. Set Up Virtual Environment & Dependencies
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Open the Jupyter Notebook
```bash
jupyter notebook Campus_Placement_Prediction.ipynb
```

---

## 👤 Author
**Ashbi Saju**  
GitHub: [@aashbiii](https://github.com/aashbiii)
