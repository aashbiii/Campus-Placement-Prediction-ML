import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_report():
    # Backup existing docx if not backed up
    if os.path.exists('Campus_Placement_Project_Report.docx') and not os.path.exists('Campus_Placement_Project_Report_old.docx'):
        shutil.copyfile('Campus_Placement_Project_Report.docx', 'Campus_Placement_Project_Report_old.docx')
        print("Backed up old report to Campus_Placement_Project_Report_old.docx")

    doc = docx.Document()

    # Configure 1-inch margins
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Configure Normal Style: Times New Roman, 12pt, Justified
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0, 0, 0)
    
    # Ensure XML rFonts is Times New Roman
    rPr = style.element.get_or_add_rPr()
    rFonts = parse_xml(r'<w:rFonts %s w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' % nsdecls('w'))
    rPr.append(rFonts)

    def set_run_tnr(run, size_pt=12, bold=False, italic=False, color=(0, 0, 0)):
        run.font.name = 'Times New Roman'
        run.font.size = Pt(size_pt)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = RGBColor(*color)
        rPr = run._r.get_or_add_rPr()
        rFonts = parse_xml(r'<w:rFonts %s w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' % nsdecls('w'))
        rPr.append(rFonts)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, line_spacing=1.15):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            set_run_tnr(run, size_pt=12)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        set_run_tnr(run, size_pt=14, bold=True)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        set_run_tnr(run, size_pt=12.5, bold=True)
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            run1 = p.add_run(bold_prefix + ": ")
            set_run_tnr(run1, size_pt=12, bold=True)
        run2 = p.add_run(text)
        set_run_tnr(run2, size_pt=12)
        return p

    def add_figure(image_path, caption_text, width_inches=6.0):
        if os.path.exists(image_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(10)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.paragraph_format.keep_with_next = True
            run_img = p_img.add_run()
            run_img.add_picture(image_path, width=Inches(width_inches))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(12)
            run_cap = p_cap.add_run(caption_text)
            set_run_tnr(run_cap, size_pt=11, bold=True, italic=True)
        else:
            print(f"Warning: image path not found: {image_path}")

    def format_table(table, col_widths, headers, data):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Border XML
        tblPr = table._tbl.tblPr
        tblBorders = parse_xml(r'''
            <w:tblBorders %s>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="888888"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="888888"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>
        ''' % nsdecls('w'))
        tblPr.append(tblBorders)

        # Header Row
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(4)
            p.paragraph_format.space_after = Pt(4)
            for r in p.runs:
                set_run_tnr(r, size_pt=11, bold=True)
            shd = parse_xml(r'<w:shd %s w:fill="E8EEF5"/>' % nsdecls('w'))
            hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

        # Data Rows
        for row_idx, row_data in enumerate(data):
            row_cells = table.add_row().cells
            for col_idx, cell_value in enumerate(row_data):
                row_cells[col_idx].text = str(cell_value)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx <= 1 else WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    set_run_tnr(r, size_pt=10.5)
                if row_idx % 2 == 1:
                    shd = parse_xml(r'<w:shd %s w:fill="F9FBFD"/>' % nsdecls('w'))
                    row_cells[col_idx]._tc.get_or_add_tcPr().append(shd)

        # Widths
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)

    print("Helper functions initialized. Generating report content...")

    # -------------------------------------------------------------
    # COVER / HEADER BLOCK
    # -------------------------------------------------------------
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(4)
    run_t = p_title.add_run("CAMPUS PLACEMENT & PACKAGE PREDICTION SYSTEM USING MACHINE LEARNING")
    set_run_tnr(run_t, size_pt=16, bold=True)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    run_sub = p_sub.add_run("Machine Learning Mini Project Report\nSubmitted by: Ashbi Saju | Roll No: 16 | Date: October 2026")
    set_run_tnr(run_sub, size_pt=12, italic=True)

    # -------------------------------------------------------------
    # 1. Project Title / Topic
    # -------------------------------------------------------------
    add_h1("1. Project Title / Topic")
    
    p = add_p()
    r = p.add_run("Project Title: ")
    set_run_tnr(r, size_pt=12, bold=True)
    r2 = p.add_run("Campus Placement & Package Prediction System Using Machine Learning")
    set_run_tnr(r2, size_pt=12)

    add_p(
        "Campus recruitment drives serve as the primary career launchpad for university students and provide corporate organizations "
        "with qualified entry-level technical talent. However, academic placement departments often evaluate student employability "
        "through simplistic, single-dimensional criteria - predominantly undergraduate Cumulative Grade Point Average (CGPA) or historical "
        "examination percentages. This traditional approach overlooks essential technical competencies such as coding proficiency, "
        "practical software projects, professional certifications, aptitude test scores, internship experiences, and interpersonal "
        "communication abilities."
    )
    add_p(
        "This reliance on marks alone creates two critical operational challenges: first, institutional placement cells fail to "
        "identify at-risk students sufficiently early to deliver targeted remedial interventions; second, students lack quantitative, "
        "data-driven clarity regarding how their non-academic qualifications influence both their probability of getting placed and their "
        "expected salary package. To systematically resolve these limitations, this project implements an end-to-end Machine Learning system "
        "that addresses placement prediction as a dual-stage predictive pipeline: binary classification of placement status (Placed vs. Not Placed), "
        "followed by conditional continuous regression of the expected salary package in Lakhs Per Annum (LPA) for placed candidates."
    )

    # -------------------------------------------------------------
    # 2. Objectives
    # -------------------------------------------------------------
    add_h1("2. Objectives")
    add_p(
        "The primary goal of this mini project is to build, evaluate, and deploy an automated, verifiable machine learning framework "
        "that provides holistic, data-driven evaluation of campus recruitment readiness. The project addresses four core objectives:"
    )

    add_bullet(
        "Objective 1 (Comprehensive Data Integration)",
        "To formulate and synthesize a multi-dimensional campus recruitment dataset comprising 3,000 student records and 16 features that holistically integrate secondary school marks, higher secondary percentages, undergraduate CGPA, internship counts, technical projects, certifications, programming scores, aptitude test performance, communication ratings, and active backlogs."
    )
    add_bullet(
        "Objective 2 (Placement Status Classification)",
        "To develop and validate supervised ensemble machine learning classification models - specifically Random Forest Classifier and Extreme Gradient Boosting (XGBoost) Classifier - to predict whether an individual student will successfully secure a campus placement offer (Placed vs. Not Placed) along with calibrated prediction confidence probabilities."
    )
    add_bullet(
        "Objective 3 (Salary Package Estimation via Regression)",
        "To develop and validate predictive regression models - specifically Random Forest Regressor and L2-regularized Ridge Regression - to estimate the continuous annual compensation package in Lakhs Per Annum (LPA) for placed candidates, capturing both linear domain trends and complex non-linear feature interactions."
    )
    add_bullet(
        "Objective 4 (Interpretability, Verification, and Pipeline Deployment)",
        "To perform rigorous 5-fold cross-validation across all models, extract quantitative feature importance rankings to identify the primary drivers of placement success and compensation, and package the complete workflow into an end-to-end Python inference function capable of generating real-time predictions for newly submitted student profiles."
    )

    # -------------------------------------------------------------
    # 3. Description of the Dataset
    # -------------------------------------------------------------
    add_h1("3. Description of the Dataset")
    add_p(
        "A rigorous empirical study requires a dataset containing both comprehensive technical/soft-skill features and dual prediction targets "
        "(placement status and salary package). Because publicly accessible recruitment datasets on repositories like Kaggle and HuggingFace "
        "typically provide either academic percentages with salary or skill scores without salary, a realistic dataset of 3,000 student records "
        "was synthesized programmatically using NumPy and Pandas with a fixed random seed (seed = 42). The statistical distributions, correlations, "
        "and class boundaries mirror published literature on engineering campus placements."
    )

    add_bullet("Dataset Name", "Campus Placement and Package Prediction Dataset (campus_placement_dataset.csv)")
    add_bullet("Source of Dataset", "Synthesized programmatically to mirror empirical distributions and feature correlations from peer-reviewed campus recruitment benchmarks.")
    add_bullet("Number of Records / Rows", "3,000 individual student profiles")
    add_bullet("Number of Attributes / Columns", "16 attributes (14 input predictor features and 2 target variables)")
    add_bullet("Target Variables", "1) Placement_Status: Categorical target for binary classification ('Placed' vs. 'Not Placed'); 2) Package_LPA: Continuous numerical target for regression (annual salary in Lakhs Per Annum, defined conditionally for placed candidates).")

    add_p("A complete description of all 16 attributes, including their data types, domain definitions, and observed ranges, is detailed in Table 1 below:")

    # Table 1: Attribute description
    headers_t1 = ["Attribute Name", "Data Type", "Role / Category", "Description & Value Range"]
    data_t1 = [
        ["Gender", "object (string)", "Predictor", "Student's biological gender: 'Male' (58.0%), 'Female' (42.0%)."],
        ["Branch", "object (string)", "Predictor", "Undergraduate engineering degree: CS (30%), IT (20%), EC (20%), ME (12%), CE (9%), EE (9%)."],
        ["SSC_Percentage", "float64", "Predictor", "Secondary School Certificate (10th grade) percentage. Range: 40.00% to 98.00% (Mean: 74.88%)."],
        ["HSC_Percentage", "float64", "Predictor", "Higher Secondary Certificate (12th grade) percentage. Range: 35.00% to 98.00% (45 missing values)."],
        ["CGPA", "float64", "Predictor", "Cumulative Grade Point Average in engineering degree. Scale: 4.00 to 10.00 (Mean: 7.21)."],
        ["Internships", "int64", "Predictor", "Count of completed industry internships. Range: 0 to 4 internships (Mean: 1.12)."],
        ["Projects", "int64", "Predictor", "Count of technical projects completed. Range: 0 to 6 projects (Mean: 2.01)."],
        ["Certifications", "int64", "Predictor", "Count of verified technical certifications/workshops. Range: 0 to 5 certifications (Mean: 1.31)."],
        ["Coding_Skill_Score", "float64", "Predictor", "Technical coding and algorithmic test score. Scale: 0.0 to 100.0 (Mean: 53.64)."],
        ["Aptitude_Test_Score", "float64", "Predictor", "Score in placement aptitude assessment (quant/verbal/logic). Scale: 20.0 to 100.0 (90 missing values)."],
        ["Communication_Score", "float64", "Predictor", "Soft-skills, interview, and communication rating. Scale: 1.0 to 10.0 (75 missing values)."],
        ["Backlogs", "float64 / int", "Predictor", "Number of pending academic course backlogs. Range: 0 to 5 backlogs (60 missing values)."],
        ["Extracurricular_Activities", "object (string)", "Predictor", "Active participation in extracurricular activities: 'Yes' (45.0%), 'No' (55.0%)."],
        ["Placement_Training", "object (string)", "Predictor", "Completion of formal placement preparation training: 'Yes' (55.0%), 'No' (45.0%)."],
        ["Placement_Status", "object (string)", "Target 1 (Clf)", "Binary classification target: 'Placed' (1,498 rows, 49.9%) vs. 'Not Placed' (1,502 rows, 50.1%)."],
        ["Package_LPA", "float64", "Target 2 (Reg)", "Continuous salary package in LPA (placed candidates only). Range: 8.79 to 29.80 LPA (Mean: 17.51 LPA)."]
    ]
    t1 = doc.add_table(rows=1, cols=4)
    format_table(t1, [1.6, 1.0, 1.1, 2.8], headers_t1, data_t1)

    p_cap_t1 = doc.add_paragraph()
    p_cap_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t1.paragraph_format.space_before = Pt(4)
    p_cap_t1.paragraph_format.space_after = Pt(12)
    run_cap_t1 = p_cap_t1.add_run("Table 1: Comprehensive Dataset Attribute Dictionary and Data Types")
    set_run_tnr(run_cap_t1, size_pt=11, bold=True, italic=True)

    # -------------------------------------------------------------
    # 4. Summary of Procedures
    # -------------------------------------------------------------
    add_h1("4. Summary of Procedures")
    add_p(
        "To ensure reproducible and scientifically verifiable outcomes, the project was executed in a systematic pipeline adhering "
        "to industry-standard Machine Learning lifecycle guidelines. The procedure comprises fifteen clearly demarcated stages:"
    )

    add_h2("4.1 Loading the Required Libraries")
    add_p(
        "The project workspace was configured in Python 3.13 by importing core data-science and scientific libraries: NumPy (v2.2.6) for "
        "numerical array computation; Pandas (v2.3.1) for tabular dataframe manipulation and schema alignment; Matplotlib (v3.10.8) and "
        "Seaborn (v0.13.2) for exploratory visual analytics; Scikit-Learn (v1.6.1) for preprocessing, pipeline execution, ensemble modeling, "
        "and evaluation metrics; and XGBoost (v3.2.0) for extreme gradient boosting classification."
    )

    add_h2("4.2 Importing the Dataset")
    add_p(
        "The primary dataset was loaded from 'campus_placement_dataset.csv' into a Pandas DataFrame using pd.read_csv(). A fixed random seed "
        "(RANDOM_STATE = 42) was initialized across all NumPy and Scikit-Learn modules to guarantee exact determinism across data splitting, "
        "tree induction, cross-validation partitioning, and test evaluations."
    )

    add_h2("4.3 Understanding the Structure of the Dataset")
    add_p(
        "Initial structural exploration was conducted by inspecting df.shape (3,000 rows × 16 columns), df.info() for column data types and "
        "non-null counts, and df.describe(include='all').T for statistical distributional properties. This confirmed that the dataset contains "
        "8 floating-point numerical features, 3 integer counts, and 5 object-type categorical variables. Class distribution for Placement_Status "
        "revealed 1,502 'Not Placed' (50.1%) and 1,498 'Placed' (49.9%), establishing an exceptionally balanced distribution ideal for classification."
    )

    add_h2("4.4 Checking and Handling Missing Values")
    add_p(
        "Inspection via df.isnull().sum() identified realistic data-collection gaps in four numerical features: Aptitude_Test_Score (90 missing values, 3.0%), "
        "Communication_Score (75 missing values, 2.5%), Backlogs (60 missing values, 2.0%), and HSC_Percentage (45 missing values, 1.5%). Because these "
        "attributes follow skewed and count-oriented distributions, Scikit-Learn's SimpleImputer with strategy='median' was applied. Imputing with the "
        "median prevents outlier distortion. Post-imputation validation verified zero remaining missing values across all predictor columns."
    )

    add_h2("4.5 Identifying and Correcting Incorrect or Inconsistent Data")
    add_p(
        "Range checks and integrity constraints were executed across all variables: percentages (SSC, HSC) were checked to ensure values strictly "
        "reside within [0, 100]; CGPA was bounded within the standard academic scale [4.0, 10.0]; internship and project counts were validated to be "
        "non-negative integers; and backlogs were bounded between 0 and 5. Categorical fields were checked for string inconsistencies or trailing "
        "whitespaces, confirming full semantic uniformity."
    )

    add_h2("4.6 Removing Duplicate Records")
    add_p(
        "Duplicate checking was executed using df.duplicated().sum(), which returned exactly 0 duplicate rows across all 3,000 records. Every student "
        "profile represents a distinct combination of academic, technical, and extracurricular metrics, preserving sample independence."
    )

    add_h2("4.7 Performing Data Preprocessing")
    add_p(
        "Categorical variables were converted into machine-readable numeric formats: binary features (Gender, Extracurricular_Activities, "
        "Placement_Training) were label-encoded using Scikit-Learn's LabelEncoder (e.g., Female=0, Male=1; No=0, Yes=1). The multi-class nominal "
        "variable Branch (6 disciplines) was one-hot encoded using pd.get_dummies() with drop_first=True, producing 5 dummy indicator columns "
        "and eliminating the dummy-variable trap (multicollinearity). The primary classification target Placement_Status was label-encoded "
        "into binary integers ('Not Placed'=0, 'Placed'=1)."
    )

    add_h2("4.8 Performing Exploratory Data Analysis (EDA)")
    add_p(
        "Exploratory analysis examined univariate distributions and bivariate relationships. Summary statistics established that placed students "
        "exhibit substantially higher mean CGPAs (7.78 vs. 6.64) and coding skill scores (67.4 vs. 39.9) compared to unplaced students. The salary "
        "distribution for placed candidates was observed to be approximately normal with a slight right skew, peaking around 17 LPA."
    )

    add_h2("4.9 Creating Appropriate Visualizations")
    add_p(
        "Informative visualization figures were generated to unpack feature relationships: countplots and histograms for target distributions; "
        "a Pearson correlation heatmap across all 10 numeric features; a 2×5 panel of boxplots comparing feature spreads across placement categories; "
        "and a multi-dimensional scatter plot visualizing CGPA versus salary package colored by coding skill score."
    )

    add_h2("4.10 Selecting Relevant Features")
    add_p(
        "The model feature matrix was constructed by removing the two target columns (Placement_Status and Package_LPA), yielding an 18-dimensional "
        "feature space (13 numerical/binary predictors plus 5 one-hot encoded branch indicators). Rather than relying on aggressive heuristic filtering, "
        "all 18 features were retained, and post-hoc tree-ensemble feature importance analyses were employed to confirm that no spurious features dominated."
    )

    add_h2("4.11 Applying the Appropriate Analytical or Machine Learning Technique")
    add_p(
        "Algorithms were selected based on the mathematical properties of the data: for classification, Random Forest Classifier (bagging ensemble) "
        "and XGBoost Classifier (gradient boosting) were selected because they handle mixed categorical/numerical inputs, capture non-linear interaction "
        "effects without manual feature expansion, and inherently resist overfitting. For regression, Random Forest Regressor was selected to model "
        "complex non-linear salary drivers, paired with L2-regularized Ridge Regression as an interpretable, robust linear benchmark."
    )

    add_h2("4.12 Training the Model")
    add_p(
        "For classification, the 3,000 records were partitioned into an 80% training set (2,400 records) and a 20% test set (600 records) using "
        "stratified sampling (stratify=y_clf). For regression, the placed subset (1,498 candidates) was partitioned into an 80% training set "
        "(1,198 candidates) and a 20% test set (300 candidates). StandardScaler was fitted on the training split to scale features for Ridge Regression. "
        "All models were trained through explicit .fit() calls, and wall-clock training times were systematically logged."
    )

    add_h2("4.13 Making Predictions or Generating Clusters")
    add_p(
        "Trained classification models generated discrete class labels (predict) and continuous probability estimates (predict_proba) on the held-out "
        "600-sample test set. Trained regression models produced salary predictions in LPA on the held-out 300-sample placed test set. Both model families "
        "exhibited well-calibrated, monotonic predictions across all performance percentiles."
    )

    add_h2("4.14 Evaluating the Results")
    add_p(
        "Model evaluation followed standard statistical metrics: classification performance was assessed using Accuracy, Precision, Recall, F1-Score, "
        "Confusion Matrix, Receiver Operating Characteristic Area Under the Curve (ROC-AUC), and 5-Fold Stratified Cross-Validation. Regression performance "
        "was assessed using Mean Absolute Error (MAE), Root Mean Squared Error (RMSE), Coefficient of Determination (R² Score), and 5-Fold Cross-Validation."
    )

    add_h2("4.15 Interpreting the Findings")
    add_p(
        "Feature importances and regression coefficients were extracted to establish real-world insights. CGPA, Coding Skill Score, and Aptitude Test Score "
        "emerged as the three dominant predictors across both placement status and salary package. Finally, an interactive function predict_new_student() "
        "was verified to execute the full pipeline on new student profiles end-to-end."
    )

    # -------------------------------------------------------------
    # 5. Evaluation and Results
    # -------------------------------------------------------------
    add_h1("5. Evaluation and Results")
    add_p(
        "This section presents the empirical evaluation metrics obtained from testing the trained classification and regression models on the unseen "
        "test splits, accompanied by cross-validation stability checks and substantive technical interpretations."
    )

    add_h2("5.1 Classification Results  -  Placement Status")
    add_p(
        "The binary classification task evaluated whether a student is placed (Class 1) or not placed (Class 0) on the 600-sample held-out test split "
        "(300 Not Placed, 300 Placed). The comparative results for Random Forest Classifier and XGBoost Classifier are presented in Table 2 below:"
    )

    # Table 2: Classification metrics
    headers_t2 = ["Model Name", "Accuracy", "Precision", "Recall", "F1-Score", "ROC-AUC", "5-Fold CV Accuracy (Mean ± Std)"]
    data_t2 = [
        ["Random Forest Classifier", "81.83%", "80.32%", "84.33%", "82.28%", "0.9142", "81.40% (± 1.73%)"],
        ["XGBoost Classifier", "84.33%", "83.44%", "85.67%", "84.54%", "0.9250", "82.63% (± 1.17%)"]
    ]
    t2 = doc.add_table(rows=1, cols=7)
    format_table(t2, [1.6, 0.7, 0.7, 0.7, 0.7, 0.7, 1.4], headers_t2, data_t2)

    p_cap_t2 = doc.add_paragraph()
    p_cap_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t2.paragraph_format.space_before = Pt(4)
    p_cap_t2.paragraph_format.space_after = Pt(10)
    run_cap_t2 = p_cap_t2.add_run("Table 2: Comparative Performance Metrics for Placement Status Classification")
    set_run_tnr(run_cap_t2, size_pt=11, bold=True, italic=True)

    add_p(
        "Confusion Matrix Breakdown:\n"
        "• Random Forest Classifier: True Negatives (TN) = 238, False Positives (FP) = 62, False Negatives (FN) = 47, True Positives (TP) = 253.\n"
        "• XGBoost Classifier: True Negatives (TN) = 249, False Positives (FP) = 51, False Negatives (FN) = 43, True Positives (TP) = 257."
    )
    add_p(
        "Interpretation of Classification Results:\n"
        "The XGBoost Classifier achieved the strongest overall classification performance, obtaining an accuracy of 84.33% and an ROC-AUC of 0.9250, "
        "outperforming the Random Forest Classifier (81.83% accuracy, 0.9142 ROC-AUC). Crucially, XGBoost achieved a higher recall (85.67%) while simultaneously "
        "lowering false positive predictions from 62 down to 51. In a campus placement setting, minimizing false positives is critical so that institutions "
        "do not mistakenly assume an at-risk student is safe. Furthermore, 5-fold cross-validation confirmed consistent generalization with low variance "
        "(82.63% ± 1.17%), proving that performance is robust across distinct data splits."
    )

    add_h2("5.2 Regression Results  -  Salary Package (LPA)")
    add_p(
        "The regression task estimated the continuous salary package in Lakhs Per Annum (LPA) for placed candidates on the 300-sample placed test split. "
        "The comparative evaluation between the non-linear Random Forest Regressor and the regularized linear Ridge Regression model is summarized in Table 3 below:"
    )

    # Table 3: Regression metrics
    headers_t3 = ["Model Name", "MAE (LPA)", "MSE (LPA²)", "RMSE (LPA)", "R² Score", "5-Fold CV R² (Mean ± Std)"]
    data_t3 = [
        ["Random Forest Regressor", "0.811 LPA", "1.034", "1.017 LPA", "0.9032", "0.8853 (± 0.013)"],
        ["Ridge Regression (Scaled)", "0.647 LPA", "0.666", "0.816 LPA", "0.9377", "0.9361 (± 0.007)"]
    ]
    t3 = doc.add_table(rows=1, cols=6)
    format_table(t3, [1.8, 0.9, 0.9, 0.9, 0.9, 1.4], headers_t3, data_t3)

    p_cap_t3 = doc.add_paragraph()
    p_cap_t3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t3.paragraph_format.space_before = Pt(4)
    p_cap_t3.paragraph_format.space_after = Pt(10)
    run_cap_t3 = p_cap_t3.add_run("Table 3: Comparative Evaluation Metrics for Salary Package Regression (Placed Candidates)")
    set_run_tnr(run_cap_t3, size_pt=11, bold=True, italic=True)

    add_p(
        "Interpretation of Regression Results:\n"
        "Both regression models exhibited strong predictive accuracy, with R² scores exceeding 0.90. Ridge Regression achieved superior overall precision, "
        "yielding a Mean Absolute Error (MAE) of only 0.647 LPA (approximately Rs. 64,700 per year error on an average package of Rs. 17.5 Lakhs), a Root Mean "
        "Squared Error (RMSE) of 0.816 LPA, and an R² of 0.9377. Random Forest Regressor achieved an MAE of 0.811 LPA and an R² of 0.9032. The fact that "
        "Ridge Regression slightly outperformed the non-linear ensemble indicates that conditional on being placed, corporate compensation structures "
        "behave with strong regularized linear additive properties, where base academic percentage, coding prowess, and branch multipliers compound "
        "predictably. The 5-fold cross-validation R² of 0.9361 (± 0.007) confirms exceptional generalization across all data subsets."
    )

    add_h2("5.3 End-to-End Prediction Demonstration Results")
    add_p(
        "To validate the operational readiness of the system, a simulated profile representing an unseen candidate was passed to the pipeline:\n"
        "• Candidate Profile: Female, Computer Science, 10th = 88.0%, 12th = 85.0%, CGPA = 8.6, 2 Internships, 4 Projects, 3 Certifications, "
        "Coding Skill Score = 82.0, Aptitude Score = 78.0, Communication Score = 8.0, 0 Backlogs, Extracurriculars = Yes, Placement Training = Yes.\n"
        "• Model Output:\n"
        "  - Predicted Placement Status: Placed\n"
        "  - Placement Confidence Probability: 97.6%\n"
        "  - Predicted Annual Salary Package: Rs. 27.12 LPA\n"
        "This confirms that the end-to-end inference pipeline successfully coordinates preprocessing, feature alignment, classification gating, and "
        "conditional regression to deliver realistic, actionable predictions."
    )

    # -------------------------------------------------------------
    # 6. Output Screenshots
    # -------------------------------------------------------------
    add_h1("6. Output Screenshots")
    add_p(
        "In accordance with the project guidelines, this section includes all essential output screenshots and visualizations generated "
        "throughout the execution of the project. Each figure is clearly captioned and illustrates a core phase of the pipeline - ranging from initial "
        "dataset inspection and exploratory data analysis to training execution, evaluation matrices, feature importances, and end-to-end predictions."
    )

    # Figures
    add_figure("report_images/fig1_dataset_head.png", "Figure 1: Dataset Overview and Sample Student Records (df.head())", width_inches=6.2)
    add_figure("report_images/fig2_dataset_info.png", "Figure 2: Dataset Schema, Column Types, and Non-Null Counts (df.info())", width_inches=5.8)
    add_figure("report_images/fig3_preprocessing_summary.png", "Figure 3: Missing Value Detection and Median Imputation Summary", width_inches=5.8)
    add_figure("report_images/fig4_summary_statistics.png", "Figure 4: Summary Statistics of Key Numeric Predictors and Target Features (df.describe())", width_inches=6.2)
    add_figure("report_images/fig5_placement_package_distribution.png", "Figure 5: Placement Status Distribution (Left) and Salary Package Distribution Among Placed Students (Right)", width_inches=6.2)
    add_figure("report_images/fig6_correlation_heatmap.png", "Figure 6: Correlation Heatmap Across Numeric Academic and Technical Skill Features", width_inches=5.2)
    add_figure("report_images/fig7_numeric_features_by_status.png", "Figure 7: Boxplot Distributions of Numeric Features Grouped by Placement Status", width_inches=6.3)
    add_figure("report_images/fig8_cgpa_vs_package.png", "Figure 8: Scatter Plot of CGPA vs. Salary Package Coloured by Coding Skill Score", width_inches=5.2)
    add_figure("report_images/fig9_model_training_logs.png", "Figure 9: Model Training Execution Times, Tree Counts, and Verification Logs", width_inches=5.8)
    add_figure("report_images/fig10_classification_metrics.png", "Figure 10: Classification Performance Metrics and 5-Fold Stratified Cross-Validation", width_inches=5.8)
    add_figure("report_images/fig11_confusion_matrices.png", "Figure 11: Confusion Matrices for Random Forest (Left) and XGBoost (Right) Classifiers", width_inches=6.2)
    add_figure("report_images/fig12_roc_curves.png", "Figure 12: Receiver Operating Characteristic (ROC) Curves for Placement Classifiers", width_inches=5.0)
    add_figure("report_images/fig13_feature_importances_clf.png", "Figure 13: Top 10 Feature Importances for Placement Status Classification (Random Forest)", width_inches=5.4)
    add_figure("report_images/fig14_regression_metrics.png", "Figure 14: Regression Model Performance Summary (MAE, RMSE, R²) and Cross-Validation", width_inches=5.8)
    add_figure("report_images/fig15_actual_vs_predicted_regression.png", "Figure 15: Actual vs. Predicted Package for Random Forest Regressor (Left) and Ridge Regression (Right)", width_inches=6.2)
    add_figure("report_images/fig16_feature_importances_reg.png", "Figure 16: Top 10 Feature Importances for Salary Package Regression (Random Forest)", width_inches=5.4)
    add_figure("report_images/fig17_prediction_demo.png", "Figure 17: End-to-End Prediction Pipeline Output for Unseen Candidate Profile (predict_new_student())", width_inches=5.8)

    # -------------------------------------------------------------
    # 7. Code
    # -------------------------------------------------------------
    add_h1("7. Code")
    add_p(
        "The following complete, executable, and fully commented Python code implements all procedures detailed in Section 4. "
        "It follows the exact sequence of data generation, exploratory analysis, median imputation, categorical encoding, "
        "train-test splitting, model training, cross-validation, and deployment demonstration."
    )

    # Read extracted_code.py and add as formatted code block paragraphs
    with open('extracted_code.py', 'r', encoding='utf-8') as f:
        code_text = f.read()

    # Add code in clean, indented single-spaced Times New Roman 10pt
    p_code = doc.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_code.paragraph_format.space_before = Pt(6)
    p_code.paragraph_format.space_after = Pt(6)
    p_code.paragraph_format.line_spacing = 1.0

    # Add border/shading to code block
    pBdr = parse_xml(r'''
        <w:pBdr %s>
            <w:left w:val="single" w:sz="12" w:space="8" w:color="2B4C7E"/>
        </w:pBdr>
    ''' % nsdecls('w'))
    p_code._p.get_or_add_pPr().append(pBdr)
    shd_code = parse_xml(r'<w:shd %s w:fill="F8F9FA"/>' % nsdecls('w'))
    p_code._p.get_or_add_pPr().append(shd_code)

    run_code = p_code.add_run(code_text)
    set_run_tnr(run_code, size_pt=9.5)

    # -------------------------------------------------------------
    # 8. Conclusion
    # -------------------------------------------------------------
    add_h1("8. Conclusion")
    add_p(
        "This project successfully formulated, developed, and validated an automated Machine Learning pipeline for Campus Placement and "
        "Salary Package Prediction, strictly fulfilling all academic and technical requirements outlined in the project guidelines."
    )

    add_bullet(
        "Major Findings from the Analysis",
        "The exploratory data analysis established that placement outcomes are fundamentally multi-factorial. Traditional marks-based screening (10th/12th percentages) showed low direct correlation with final placement status. Instead, undergraduate CGPA, algorithmic coding proficiency, and aptitude test scores emerged as the most critical determinants of recruitment success and compensation package magnitude. Furthermore, the analysis confirmed strong non-linear interactions: exceptional coding skills can compensate for a moderate CGPA, whereas high academic scores without technical competency frequently lead to unplaced outcomes."
    )
    add_bullet(
        "Performance of the Machine Learning Models",
        "Both classification and regression models delivered high predictive precision and exceptional cross-validated stability. For placement classification, the XGBoost Classifier achieved the highest overall accuracy of 84.33%, a precision of 83.44%, a recall of 85.67%, and an ROC-AUC of 0.9250, supported by a 5-fold cross-validation accuracy of 82.63% (± 1.17%). For salary package estimation, Ridge Regression achieved an outstanding R² score of 0.9377 with a Mean Absolute Error of only 0.647 LPA (RMSE = 0.816 LPA), closely followed by Random Forest Regressor (R² = 0.9032, MAE = 0.811 LPA). The consistency between hold-out test results and 5-fold cross-validation confirms that the models generalize reliably without overfitting."
    )
    add_bullet(
        "Important Insights Obtained from the Dataset",
        "Feature importance evaluations revealed that CGPA, Coding Skill Score, and Aptitude Test Score constitute over 70% of the total predictive power for placement classification. For salary package regression among placed students, department specialization (Branch multipliers favoring Computer Science and Information Technology) alongside CGPA and Coding Skill Score dictate compensation tiers. Extracurricular activities and formal placement preparation training provide secondary positive boosts, whereas pending academic backlogs significantly diminish placement likelihood."
    )
    add_bullet(
        "Addressing the Given Problem and Real-World Utility",
        "The implemented system directly solves the two practical dilemmas faced by campus placement cells: early identification of vulnerable students and objective expectation setting. By integrating the model into university management systems prior to the final academic year, placement directors can proactively flag candidates with low placement probabilities and enroll them into targeted coding bootcamps, resume-building workshops, and communication practice. Concurrently, the conditional regression model provides realistic salary expectations that assist students in career counseling and offer evaluations. The complete pipeline was packaged into a reusable predict_new_student() function, confirming its immediate feasibility for real-world academic deployment."
    )

    output_path = 'Campus_Placement_Project_Report.docx'
    doc.save(output_path)
    print(f"Report generated successfully and saved to: {output_path}")

if __name__ == '__main__':
    create_report()
