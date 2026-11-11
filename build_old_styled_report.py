import os
import shutil
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_report():
    print("Building report in the authentic look and aesthetic of the old report...")
    doc = docx.Document()

    # Section configuration: 0.75 in margins, Letter size, page borders, titlePg
    sec = doc.sections[0]
    sec.top_margin = Inches(0.75)
    sec.bottom_margin = Inches(0.75)
    sec.left_margin = Inches(0.75)
    sec.right_margin = Inches(0.75)
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11.0)

    secPr = sec._sectPr
    # Page border XML matching old document
    pgBorders = parse_xml(r'''
        <w:pgBorders %s w:offsetFrom="page">
            <w:top w:val="single" w:sz="4" w:space="31" w:color="auto"/>
            <w:left w:val="single" w:sz="4" w:space="31" w:color="auto"/>
            <w:bottom w:val="single" w:sz="4" w:space="31" w:color="auto"/>
            <w:right w:val="single" w:sz="4" w:space="31" w:color="auto"/>
        </w:pgBorders>
    ''' % nsdecls('w'))
    secPr.append(pgBorders)

    # Different first page (no header/footer on cover)
    secPr.append(parse_xml(r'<w:titlePg %s/>' % nsdecls('w')))

    # Header: right-aligned, 8pt Calibri, #888888
    header = sec.header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    hp.paragraph_format.space_after = Pt(4)
    hrun = hp.add_run("Campus Placement & Package Prediction  -  Project Report")
    hrun.font.name = "Calibri"
    hrun.font.size = Pt(8)
    hrun.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    # Footer: centered, 9pt Calibri, #888888, with PAGE field
    footer = sec.footer
    fld_xml = r'''
        <w:p %s>
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                    <w:color w:val="888888"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:fldChar w:fldCharType="begin"/>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                    <w:color w:val="888888"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:instrText>PAGE</w:instrText>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                    <w:color w:val="888888"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:fldChar w:fldCharType="separate"/>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                    <w:color w:val="888888"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:t>1</w:t>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
                    <w:color w:val="888888"/>
                    <w:sz w:val="18"/>
                </w:rPr>
                <w:fldChar w:fldCharType="end"/>
            </w:r>
        </w:p>
    ''' % nsdecls('w')
    footer._element.clear()
    footer._element.append(parse_xml(fld_xml))

    # Helper function to format runs in Calibri
    def set_font(run, size_pt=12, bold=False, italic=False, color_rgb=(0, 0, 0)):
        run.font.name = 'Calibri'
        run.font.size = Pt(size_pt)
        run.bold = bold
        run.italic = italic
        run.font.color.rgb = RGBColor(*color_rgb)
        rPr = run._r.get_or_add_rPr()
        rFonts = parse_xml(r'<w:rFonts %s w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>' % nsdecls('w'))
        rPr.append(rFonts)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, line_spacing=1.15):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            set_font(run, size_pt=12, color_rgb=(0, 0, 0))
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(20)
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.keep_with_next = True

        # Navy blue bottom accent rule line (#1F4E79)
        pBdr = parse_xml(r'''
            <w:pBdr %s>
                <w:bottom w:val="single" w:sz="6" w:space="4" w:color="1F4E79"/>
            </w:pBdr>
        ''' % nsdecls('w'))
        p._p.get_or_add_pPr().append(pBdr)

        run = p.add_run(text)
        set_font(run, size_pt=16, bold=True, color_rgb=(0x1F, 0x4E, 0x79))
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        set_font(run, size_pt=13, bold=True, color_rgb=(0x1F, 0x4E, 0x79))
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        set_font(run, size_pt=11.5, bold=True, italic=True, color_rgb=(0x33, 0x33, 0x33))
        return p

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15

        rb = p.add_run("●  ")
        set_font(rb, size_pt=8, color_rgb=(0x1F, 0x4E, 0x79))

        if bold_prefix:
            r1 = p.add_run(bold_prefix + ": ")
            set_font(r1, size_pt=12, bold=True, color_rgb=(0, 0, 0))
        r2 = p.add_run(text)
        set_font(r2, size_pt=12, color_rgb=(0, 0, 0))
        return p

    def add_figure(image_path, caption_num, caption_text, width_inches=6.0):
        if os.path.exists(image_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(12)
            p_img.paragraph_format.space_after = Pt(4)
            p_img.paragraph_format.keep_with_next = True
            run_img = p_img.add_run()
            run_img.add_picture(image_path, width=Inches(width_inches))

            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(14)

            run_fig = p_cap.add_run(f"Figure {caption_num}. ")
            set_font(run_fig, size_pt=10.5, bold=True, color_rgb=(0x1F, 0x4E, 0x79))

            run_cap = p_cap.add_run(caption_text)
            set_font(run_cap, size_pt=10.5, italic=True, color_rgb=(0x33, 0x33, 0x33))
        else:
            print(f"Warning: image path not found: {image_path}")

    def format_table(table, col_widths, headers, data):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Old report border style
        tblPr = table._tbl.tblPr
        tblBorders = parse_xml(r'''
            <w:tblBorders %s>
                <w:top w:val="single" w:sz="6" w:space="0" w:color="1F4E79"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="1F4E79"/>
                <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D7DE"/>
                <w:insideV w:val="none"/>
                <w:left w:val="none"/>
                <w:right w:val="none"/>
            </w:tblBorders>
        ''' % nsdecls('w'))
        tblPr.append(tblBorders)

        # Header Row: #1F4E79 Navy Blue background, white bold text
        hdr_cells = table.rows[0].cells
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(5)
            p.paragraph_format.space_after = Pt(5)
            for r in p.runs:
                set_font(r, size_pt=10.5, bold=True, color_rgb=(255, 255, 255))
            shd = parse_xml(r'<w:shd %s w:fill="1F4E79"/>' % nsdecls('w'))
            hdr_cells[i]._tc.get_or_add_tcPr().append(shd)

        # Data Rows: alternating white and light blue-gray #F2F5F8
        for row_idx, row_data in enumerate(data):
            row_cells = table.add_row().cells
            for col_idx, cell_value in enumerate(row_data):
                row_cells[col_idx].text = str(cell_value)
                p = row_cells[col_idx].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if col_idx <= 1 else WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(4)
                p.paragraph_format.line_spacing = 1.15
                for r in p.runs:
                    set_font(r, size_pt=10, color_rgb=(0, 0, 0))
                if row_idx % 2 == 1:
                    shd = parse_xml(r'<w:shd %s w:fill="F2F5F8"/>' % nsdecls('w'))
                    row_cells[col_idx]._tc.get_or_add_tcPr().append(shd)

        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)

    # -------------------------------------------------------------
    # COVER PAGE (Authentic Old Report Design)
    # -------------------------------------------------------------
    p_sp1 = doc.add_paragraph()
    p_sp1.paragraph_format.space_before = Pt(40)

    p_t1 = doc.add_paragraph()
    p_t1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_t1.paragraph_format.space_before = Pt(20)
    p_t1.paragraph_format.space_after = Pt(0)
    r_t1 = p_t1.add_run("CAMPUS PLACEMENT & PACKAGE")
    set_font(r_t1, size_pt=24, bold=True, color_rgb=(0, 0, 0))

    p_t2 = doc.add_paragraph()
    p_t2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_t2.paragraph_format.space_before = Pt(0)
    p_t2.paragraph_format.space_after = Pt(14)
    r_t2 = p_t2.add_run("PREDICTION SYSTEM USING MACHINE LEARNING")
    set_font(r_t2, size_pt=24, bold=True, color_rgb=(0, 0, 0))

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(60)
    r_sub = p_sub.add_run("Machine Learning Mini Project Report")
    set_font(r_sub, size_pt=13, italic=True, color_rgb=(0x33, 0x33, 0x33))

    p_sub_by = doc.add_paragraph()
    p_sub_by.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_sub_by.paragraph_format.space_before = Pt(40)
    p_sub_by.paragraph_format.space_after = Pt(8)
    r_sb = p_sub_by.add_run("Submitted by")
    set_font(r_sb, size_pt=12, color_rgb=(0x33, 0x33, 0x33))

    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_name.paragraph_format.space_before = Pt(0)
    p_name.paragraph_format.space_after = Pt(4)
    r_nl = p_name.add_run("Name: ")
    set_font(r_nl, size_pt=14, bold=True)
    r_nv = p_name.add_run("Ashbi Saju")
    set_font(r_nv, size_pt=14, bold=True)

    p_roll = doc.add_paragraph()
    p_roll.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_roll.paragraph_format.space_before = Pt(0)
    p_roll.paragraph_format.space_after = Pt(4)
    r_rl = p_roll.add_run("Roll No: ")
    set_font(r_rl, size_pt=14, bold=True)
    r_rv = p_roll.add_run("16")
    set_font(r_rv, size_pt=14, bold=True)

    p_date = doc.add_paragraph()
    p_date.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_date.paragraph_format.space_before = Pt(0)
    p_date.paragraph_format.space_after = Pt(0)
    r_dl = p_date.add_run("Submission Date: ")
    set_font(r_dl, size_pt=14, bold=True)
    r_dv = p_date.add_run("October 5, 2026")
    set_font(r_dv, size_pt=14, bold=True)

    # Page Break after Cover Page
    doc.add_page_break()

    # -------------------------------------------------------------
    # PAGE 2: TABLE OF CONTENTS (Authentic Old Report Design)
    # -------------------------------------------------------------
    add_h1("Table of Contents")

    toc_items = [
        ("1. Project Title / Topic", False),
        ("2. Objectives", False),
        ("3. Description of the Dataset", False),
        ("4. Summary of Procedures", False),
        ("    4.1 Loading the Required Libraries", True),
        ("    4.2 Importing the Dataset", True),
        ("    4.3 Understanding the Structure of the Dataset", True),
        ("    4.4 Checking and Handling Missing Values", True),
        ("    4.5 Identifying and Correcting Inconsistent Data", True),
        ("    4.6 Removing Duplicate Records", True),
        ("    4.7 Performing Data Preprocessing", True),
        ("    4.8 Performing Exploratory Data Analysis (EDA)", True),
        ("    4.9 Creating Appropriate Visualizations", True),
        ("    4.10 Selecting Relevant Features", True),
        ("    4.11 Machine Learning Technique Selection & Justification", True),
        ("    4.12 Model Training & Verification", True),
        ("    4.13 Generating Predictions", True),
        ("    4.14 Evaluating the Results", True),
        ("    4.15 Interpreting the Findings", True),
        ("5. Evaluation and Results", False),
        ("    5.1 Classification Results  -  Placement Status", True),
        ("    5.2 Regression Results  -  Salary Package (LPA)", True),
        ("    5.3 End-to-End Prediction Demonstration Results", True),
        ("6. Output Screenshots", False),
        ("7. Code", False),
        ("8. Conclusion", False)
    ]

    for item_text, is_sub in toc_items:
        p_toc = doc.add_paragraph()
        p_toc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_toc.paragraph_format.space_before = Pt(1)
        p_toc.paragraph_format.space_after = Pt(2)
        r_toc = p_toc.add_run(item_text)
        if not is_sub:
            set_font(r_toc, size_pt=12, bold=True, color_rgb=(0x1F, 0x4E, 0x79))
        else:
            set_font(r_toc, size_pt=11, color_rgb=(0x33, 0x33, 0x33))

    # Page Break after Table of Contents
    doc.add_page_break()

    # -------------------------------------------------------------
    # 1. Project Title / Topic
    # -------------------------------------------------------------
    add_h1("1. Project Title / Topic")

    p = add_p()
    r = p.add_run("Project Title: ")
    set_font(r, size_pt=12, bold=True, color_rgb=(0x1F, 0x4E, 0x79))
    r2 = p.add_run("Campus Placement & Package Prediction System Using Machine Learning")
    set_font(r2, size_pt=12, bold=True)

    add_p(
        "During campus placement drives, academic institutions and students lack data-driven insights to evaluate job readiness effectively. "
        "Traditional evaluations often rely solely on academic marks (such as CGPA or 10th/12th percentages), ignoring technical coding skills, "
        "practical project experience, professional certifications, and soft-skill proficiency. This creates two practical operational gaps: "
        "placement cells cannot identify at-risk students early enough to intervene with targeted training, and students have no quantitative sense "
        "of what is actually driving (or hurting) their placement chances and expected salary."
    )
    add_p(
        "To systematically address this problem, this project implements an end-to-end Machine Learning system that formulates campus placement "
        "as two complementary predictive tasks: a binary classification problem to predict whether an individual student will successfully secure "
        "a placement offer (Placed vs. Not Placed), followed by a continuous regression problem to estimate the expected annual salary package in "
        "Lakhs Per Annum (LPA) for placed candidates."
    )

    # -------------------------------------------------------------
    # 2. Objectives
    # -------------------------------------------------------------
    add_h1("2. Objectives")
    add_p(
        "The project aims to provide comprehensive, data-driven decision support for campus placement readiness through the following four concise objectives:"
    )

    add_bullet(
        "Objective 1 (Holistic Student Profiling)",
        "To synthesize and structure a multi-dimensional campus recruitment dataset of 3,000 student records that captures academic percentages, cumulative GPA, coding proficiencies, aptitude test performance, soft skills, internships, projects, certifications, and academic backlogs."
    )
    add_bullet(
        "Objective 2 (Placement Status Classification)",
        "To train, tune, and evaluate supervised ensemble classification models  -  specifically Random Forest Classifier and Extreme Gradient Boosting (XGBoost) Classifier  -  to predict student placement status (Placed vs. Not Placed) along with calibrated prediction confidence probabilities."
    )
    add_bullet(
        "Objective 3 (Salary Package Estimation via Regression)",
        "To develop and evaluate predictive regression models  -  specifically Random Forest Regressor and L2-regularized Ridge Regression  -  to estimate the continuous annual salary package in Lakhs Per Annum (LPA) for placed candidates, capturing both linear domain dynamics and non-linear feature interactions."
    )
    add_bullet(
        "Objective 4 (Interpretability, Verification & Real-World Utility)",
        "To conduct 5-fold cross-validation, extract feature importance rankings to identify what credentials drive hiring outcomes, and implement an end-to-end Python prediction function (predict_new_student) for real-time evaluation of unseen student profiles."
    )

    # -------------------------------------------------------------
    # 3. Description of the Dataset
    # -------------------------------------------------------------
    add_h1("3. Description of the Dataset")
    add_p(
        "Public campus-recruitment datasets on repositories such as Kaggle and HuggingFace typically provide either (a) rich technical/soft-skill features "
        "with a placement label but no salary column, or (b) a small set of academic percentages with a salary column but no technical-skill features. "
        "No single publicly available dataset covers both halves of this project's objective together. To directly satisfy the problem statement, a dataset "
        "was generated programmatically (fixed random seed for reproducibility) following the exact statistical structure  -  same feature list, value ranges, "
        "class balance, realistic noise and missing values  -  used across well-known public campus-recruitment datasets referenced in the assignment's background papers."
    )

    add_bullet("Dataset Name", "Campus Placement and Package Prediction Dataset (campus_placement_dataset.csv)")
    add_bullet("Source", "Programmatically synthesized based on empirical distributions and domain relationships from published recruitment benchmarks.")
    add_bullet("Number of Records / Rows", "3,000 student records")
    add_bullet("Number of Attributes / Columns", "16 columns (14 input features and 2 target variables)")
    add_bullet("Target Variables", "1) Placement_Status (classification target: 'Placed' vs. 'Not Placed'); 2) Package_LPA (regression target: salary package in LPA, defined only for placed candidates).")

    add_p(
        "Class balance: 1,502 Not Placed (50.1%) vs. 1,498 Placed (49.9%)  -  an exceptionally well-balanced target, favourable for classification. "
        "Table 1 provides a comprehensive feature dictionary detailing the 16 attributes, data types, and observed value ranges:"
    )

    # Table 1: Dataset Attributes
    headers_t1 = ["Feature", "Data Type", "Role", "Description & Value Range"]
    data_t1 = [
        ["Gender", "object (string)", "Predictor", "Student's gender: 'Male' (58.0%), 'Female' (42.0%)."],
        ["Branch", "object (string)", "Predictor", "Undergraduate degree branch: CS (30%), IT (20%), EC (20%), ME (12%), CE (9%), EE (9%)."],
        ["SSC_Percentage", "float64", "Predictor", "10th grade (Secondary School) percentage. Range: 40.00% to 98.00% (Mean: 74.88%)."],
        ["HSC_Percentage", "float64", "Predictor", "12th grade (Higher Secondary) percentage. Range: 35.00% to 98.00% (45 missing values)."],
        ["CGPA", "float64", "Predictor", "Cumulative GPA in undergraduate degree. Scale: 4.00 to 10.00 (Mean: 7.21)."],
        ["Internships", "int64", "Predictor", "Number of industry internships completed. Range: 0 to 4 (Mean: 1.12)."],
        ["Projects", "int64", "Predictor", "Number of academic/personal projects completed. Range: 0 to 6 (Mean: 2.01)."],
        ["Certifications", "int64", "Predictor", "Number of technical certifications/workshops attended. Range: 0 to 5 (Mean: 1.31)."],
        ["Coding_Skill_Score", "float64", "Predictor", "Standardized coding proficiency score. Scale: 0.0 to 100.0 (Mean: 53.64)."],
        ["Aptitude_Test_Score", "float64", "Predictor", "Score in placement aptitude assessment. Scale: 20.0 to 100.0 (90 missing values)."],
        ["Communication_Score", "float64", "Predictor", "Soft-skill / communication interview rating. Scale: 1.0 to 10.0 (75 missing values)."],
        ["Backlogs", "float64 / int", "Predictor", "Number of pending academic backlogs. Range: 0 to 5 (60 missing values)."],
        ["Extracurricular_Activities", "object (string)", "Predictor", "Active participation in extracurricular activities: 'Yes' (45.0%), 'No' (55.0%)."],
        ["Placement_Training", "object (string)", "Predictor", "Underwent pre-placement training: 'Yes' (55.0%), 'No' (45.0%)."],
        ["Placement_Status", "object (string)", "Target 1 (Clf)", "Primary classification target: 'Placed' (49.9%) vs. 'Not Placed' (50.1%)."],
        ["Package_LPA", "float64", "Target 2 (Reg)", "Offered salary package in LPA (placed students only). Range: 8.79 to 29.80 LPA (Mean: 17.51 LPA)."]
    ]
    t1 = doc.add_table(rows=1, cols=4)
    format_table(t1, [1.6, 1.0, 1.1, 2.8], headers_t1, data_t1)

    p_cap_t1 = doc.add_paragraph()
    p_cap_t1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t1.paragraph_format.space_before = Pt(4)
    p_cap_t1.paragraph_format.space_after = Pt(14)
    run_cap_t1 = p_cap_t1.add_run("Table 1: Dataset Feature Dictionary and Value Ranges")
    set_font(run_cap_t1, size_pt=10.5, bold=True, italic=True, color_rgb=(0x1F, 0x4E, 0x79))

    # -------------------------------------------------------------
    # 4. Summary of Procedures
    # -------------------------------------------------------------
    add_h1("4. Summary of Procedures")
    add_p(
        "The project was executed following the fifteen standard procedure stages defined in the academic project guidelines, "
        "ensuring transparency, verifiability, and rigorous machine learning engineering."
    )

    add_h2("4.1 Loading the Required Libraries")
    add_p(
        "The development environment was configured by importing core Python scientific computing libraries: NumPy for vectorized numeric arrays; "
        "Pandas for tabular data structures and data manipulation; Matplotlib and Seaborn for statistical visual analytics; Scikit-Learn for preprocessing, "
        "cross-validation, and model estimation; and XGBoost for optimized extreme gradient boosting."
    )

    add_h2("4.2 Importing the Dataset")
    add_p(
        "The synthesized dataset was read from 'campus_placement_dataset.csv' into a Pandas DataFrame using pd.read_csv(). A master seed "
        "(RANDOM_STATE = 42) was set across NumPy and Scikit-Learn routines to ensure exact determinism and reproducibility."
    )

    add_h2("4.3 Understanding the Structure of the Dataset")
    add_p(
        "Dataset dimensions were examined via df.shape (3,000 rows × 16 columns), df.info() for column data types and memory footprints, and "
        "df.describe(include='all').T for statistical distributions (mean, standard deviation, quartiles, and min/max boundaries). Target distribution "
        "confirmed 1,502 unplaced and 1,498 placed students, demonstrating balanced target proportions."
    )

    add_h2("4.4 Checking and Handling Missing Values")
    add_p(
        "Missing value inspection using df.isnull().sum() identified realistic missingness in four features: Aptitude_Test_Score (90 missing rows, 3.0%), "
        "Communication_Score (75 missing rows, 2.5%), Backlogs (60 missing rows, 2.0%), and HSC_Percentage (45 missing rows, 1.5%). Median imputation "
        "via Scikit-Learn's SimpleImputer was selected over mean imputation due to its robust resistance to distribution outliers. Post-imputation "
        "inspection verified zero remaining missing values across all feature columns."
    )

    add_h2("4.5 Identifying and Correcting Incorrect or Inconsistent Data")
    add_p(
        "Domain bounds were validated across all numeric attributes: examination percentages (40% to 98%), CGPA (4.0 to 10.0), scores (0 to 100), "
        "and backlogs (0 to 5). No out-of-range values, anomalies, or negative entries were observed, confirming complete data integrity."
    )

    add_h2("4.6 Removing Duplicate Records")
    add_p(
        "A duplicate check via df.duplicated().sum() yielded 0 duplicate rows across all 3,000 entries. Every candidate profile represents a unique "
        "combination of metrics, preserving sample independence."
    )

    add_h2("4.7 Performing Data Preprocessing")
    add_p(
        "Binary categorical variables (Gender, Extracurricular_Activities, Placement_Training) were label-encoded. The nominal Branch column (6 branches) "
        "was one-hot encoded using pd.get_dummies() with drop_first=True to avoid multicollinearity (dummy variable trap). The classification target "
        "Placement_Status was label-encoded ('Not Placed'=0, 'Placed'=1)."
    )

    add_h2("4.8 Performing Exploratory Data Analysis (EDA)")
    add_p(
        "EDA established that placed students show noticeably higher CGPA (mean 7.78 vs. 6.64) and coding skill scores (mean 67.4 vs. 39.9) compared to "
        "unplaced students. Salary packages among placed students follow a bell-shaped distribution peaking around 17 LPA."
    )

    add_h2("4.9 Creating Appropriate Visualizations")
    add_p(
        "Comprehensive plots were constructed: target countplots and package histograms; a Pearson correlation heatmap across all 10 numeric features; "
        "boxplots illustrating feature distributions split by placement outcome; and a multi-attribute scatter plot of CGPA vs. package colored by coding skill."
    )

    add_h2("4.10 Selecting Relevant Features")
    add_p(
        "The model input matrix was constructed by excluding target columns (Placement_Status and Package_LPA), yielding an 18-dimensional feature matrix. "
        "All features were retained and validated post-hoc using tree-ensemble feature importance rankings."
    )

    add_h2("4.11 Applying the Appropriate Analytical or Machine Learning Technique")
    add_p(
        "For classification, Random Forest Classifier and XGBoost Classifier were selected because ensemble trees naturally capture non-linear interactions "
        "(e.g., strong coding skill offsetting an average CGPA) without manual feature engineering, resist overfitting, and provide clear feature importances. "
        "For regression, Random Forest Regressor was paired with L2-regularized Ridge Regression as an interpretable linear baseline to test whether "
        "salary dynamics are primarily linear or non-linear."
    )

    add_h2("4.12 Training the Model")
    add_p(
        "For classification, an 80% / 20% stratified train-test split was used (2,400 training samples, 600 test samples). For regression, only placed candidates "
        "(1,498 rows) were used, split into 1,198 training and 300 test samples. Numeric features were scaled via StandardScaler for Ridge Regression. "
        "Every model was trained using genuine .fit() calls, logging wall-clock training times."
    )

    add_h2("4.13 Making Predictions or Generating Clusters")
    add_p(
        "Trained classifiers produced binary labels and confidence probabilities on the test split. Trained regressors predicted continuous salary packages in LPA. "
        "Both models showed stable, well-calibrated predictions."
    )

    add_h2("4.14 Evaluating the Results")
    add_p(
        "Models were evaluated using formal metrics: classification via Accuracy, Precision, Recall, F1-Score, Confusion Matrix, ROC-AUC, and 5-fold cross-validation; "
        "regression via MAE, RMSE, R² Score, and 5-fold cross-validation."
    )

    add_h2("4.15 Interpreting the Findings")
    add_p(
        "Feature importances revealed that CGPA, Coding Skill Score, and Aptitude Test Score are the dominant drivers for placement and package. An end-to-end "
        "predict_new_student() function was implemented and verified on an unseen student profile."
    )

    # -------------------------------------------------------------
    # 5. Evaluation and Results
    # -------------------------------------------------------------
    add_h1("5. Evaluation and Results")
    add_p(
        "This section presents the quantitative performance metrics obtained on the held-out test sets, accompanied by 5-fold cross-validation "
        "and detailed interpretations."
    )

    add_h2("5.1 Classification Results  -  Placement Status")
    add_p(
        "Classification models were evaluated on the 600-sample test set (300 Not Placed, 300 Placed). The results are summarized in Table 2:"
    )

    # Table 2: Classification Metrics
    headers_t2 = ["Model", "Accuracy", "Precision", "Recall", "F1-Score", "ROC-AUC", "5-Fold CV Accuracy"]
    data_t2 = [
        ["Random Forest Classifier", "81.83%", "80.32%", "84.33%", "82.28%", "0.9142", "81.40% (\u00b1 1.73%)"],
        ["XGBoost Classifier", "84.33%", "83.44%", "85.67%", "84.54%", "0.9250", "82.63% (\u00b1 1.17%)"]
    ]
    t2 = doc.add_table(rows=1, cols=7)
    format_table(t2, [1.6, 0.7, 0.7, 0.7, 0.7, 0.7, 1.4], headers_t2, data_t2)

    p_cap_t2 = doc.add_paragraph()
    p_cap_t2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t2.paragraph_format.space_before = Pt(4)
    p_cap_t2.paragraph_format.space_after = Pt(10)
    run_cap_t2 = p_cap_t2.add_run("Table 2: Classification Performance Metrics (Test Set: N = 600)")
    set_font(run_cap_t2, size_pt=10.5, bold=True, italic=True, color_rgb=(0x1F, 0x4E, 0x79))

    add_p(
        "Confusion Matrix Breakdown:\n"
        "● Random Forest Classifier: True Negatives (TN) = 238, False Positives (FP) = 62, False Negatives (FN) = 47, True Positives (TP) = 253.\n"
        "● XGBoost Classifier: True Negatives (TN) = 249, False Positives (FP) = 51, False Negatives (FN) = 43, True Positives (TP) = 257."
    )
    add_p(
        "Interpretation: The XGBoost Classifier achieved the strongest overall performance with 84.33% accuracy and 0.9250 ROC-AUC, outperforming "
        "Random Forest (81.83% accuracy, 0.9142 ROC-AUC). XGBoost reduced false positive predictions from 62 to 51 while boosting recall to 85.67%. "
        "In placement screening, reducing false positives is critical so that placement officers do not overlook students needing remedial help. "
        "Both models generalized consistently across all 5 cross-validation folds."
    )

    add_h2("5.2 Regression Results  -  Salary Package (LPA)")
    add_p(
        "Regression models were evaluated on the 300 placed test samples. Comparative results are summarized in Table 3:"
    )

    # Table 3: Regression Metrics
    headers_t3 = ["Model", "MAE (LPA)", "RMSE (LPA)", "R\u00b2 Score", "5-Fold CV R\u00b2"]
    data_t3 = [
        ["Random Forest Regressor", "0.811 LPA", "1.017 LPA", "0.9032", "0.8853 (\u00b1 0.013)"],
        ["Ridge Regression (Scaled)", "0.647 LPA", "0.816 LPA", "0.9377", "0.9361 (\u00b1 0.007)"]
    ]
    t3 = doc.add_table(rows=1, cols=5)
    format_table(t3, [1.8, 1.1, 1.1, 1.0, 1.5], headers_t3, data_t3)

    p_cap_t3 = doc.add_paragraph()
    p_cap_t3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap_t3.paragraph_format.space_before = Pt(4)
    p_cap_t3.paragraph_format.space_after = Pt(10)
    run_cap_t3 = p_cap_t3.add_run("Table 3: Regression Performance Metrics (Placed Students: N = 300)")
    set_font(run_cap_t3, size_pt=10.5, bold=True, italic=True, color_rgb=(0x1F, 0x4E, 0x79))

    add_p(
        "Interpretation: Both models achieved strong explanatory power exceeding R\u00b2 = 0.90. Ridge Regression achieved superior precision, "
        "yielding a Mean Absolute Error of 0.647 LPA (approximately Rs. 64,700 error on an average package of Rs. 17.5 Lakhs) and an R\u00b2 of 0.9377, "
        "outperforming Random Forest Regressor (MAE = 0.811 LPA, R\u00b2 = 0.9032). This confirms that once a student secures placement, corporate "
        "salary offers scale with strong linear additive characteristics driven by academic percentages, coding skills, and branch multipliers."
    )

    add_h2("5.3 End-to-End Prediction Demonstration Results")
    add_p(
        "A sample unseen student profile was evaluated through the predict_new_student() function:\n"
        "● Input Profile: Female, Computer Science, 10th = 88.0%, 12th = 85.0%, CGPA = 8.6, 2 Internships, 4 Projects, 3 Certifications, "
        "Coding Skill Score = 82.0, Aptitude Score = 78.0, Communication Score = 8.0, 0 Backlogs, Extracurriculars = Yes, Placement Training = Yes.\n"
        "● Model Output:\n"
        "  - Predicted Placement Status: Placed\n"
        "  - Placement Probability: 97.6%\n"
        "  - Predicted Salary Package: Rs. 27.12 LPA\n"
        "This confirms that the full pipeline  -  preprocessing, encoding, alignment with training features, classification, and conditional regression  -  "
        "works correctly end-to-end on new data, exactly as a placement cell would use it in practice."
    )

    # -------------------------------------------------------------
    # 6. Output Screenshots
    # -------------------------------------------------------------
    add_h1("6. Output Screenshots")
    add_p(
        "This section includes all seventeen output screenshots and visualizations generated throughout the execution of the project, "
        "illustrating every stage from data inspection and EDA to training verification, evaluation matrices, feature rankings, and prediction outputs."
    )

    # 17 Figures with Old Report Captions and Styling
    add_figure("report_images/fig1_dataset_head.png", 1, "Dataset overview showing first 6 records and feature columns (df.head()).", width_inches=6.2)
    add_figure("report_images/fig2_dataset_info.png", 2, "Dataset schema, non-null counts, and data types (df.info()).", width_inches=5.8)
    add_figure("report_images/fig3_preprocessing_summary.png", 3, "Missing value detection and median imputation summary.", width_inches=5.8)
    add_figure("report_images/fig4_summary_statistics.png", 4, "Summary statistics of key numeric attributes (df.describe()).", width_inches=6.2)
    add_figure("report_images/fig5_placement_package_distribution.png", 5, "Placement status distribution (left) and package distribution among placed students (right).", width_inches=6.2)
    add_figure("report_images/fig6_correlation_heatmap.png", 6, "Correlation heatmap of numeric features.", width_inches=5.2)
    add_figure("report_images/fig7_numeric_features_by_status.png", 7, "Distribution of each numeric feature split by placement status.", width_inches=6.3)
    add_figure("report_images/fig8_cgpa_vs_package.png", 8, "CGPA vs. Package among placed students, coloured by Coding Skill Score.", width_inches=5.2)
    add_figure("report_images/fig9_model_training_logs.png", 9, "Model training execution times, fitted trees, and verification logs.", width_inches=5.8)
    add_figure("report_images/fig10_classification_metrics.png", 10, "Classification performance metrics and 5-fold cross-validation summary.", width_inches=5.8)
    add_figure("report_images/fig11_confusion_matrices.png", 11, "Confusion matrices  -  Random Forest (left) and XGBoost (right).", width_inches=6.2)
    add_figure("report_images/fig12_roc_curves.png", 12, "ROC curves for both classifiers.", width_inches=5.0)
    add_figure("report_images/fig13_feature_importances_clf.png", 13, "Top 10 feature importances  -  Random Forest (classification).", width_inches=5.4)
    add_figure("report_images/fig14_regression_metrics.png", 14, "Regression model evaluation metrics (MAE, RMSE, R\u00b2) and cross-validation.", width_inches=5.8)
    add_figure("report_images/fig15_actual_vs_predicted_regression.png", 15, "Actual vs. predicted package  -  Random Forest Regressor (left) and Ridge Regression (right).", width_inches=6.2)
    add_figure("report_images/fig16_feature_importances_reg.png", 16, "Top 10 feature importances  -  Random Forest (regression).", width_inches=5.4)
    add_figure("report_images/fig17_prediction_demo.png", 17, "End-to-end prediction pipeline demonstration output (predict_new_student()).", width_inches=5.8)

    # -------------------------------------------------------------
    # 7. Code
    # -------------------------------------------------------------
    add_h1("7. Code")
    add_p(
        "Below is the complete, formatted, and commented Python code used to execute the project end-to-end, following the exact sequence of procedures."
    )

    with open('extracted_code.py', 'r', encoding='utf-8') as f:
        code_text = f.read()

    p_code = doc.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_code.paragraph_format.space_before = Pt(6)
    p_code.paragraph_format.space_after = Pt(6)
    p_code.paragraph_format.line_spacing = 1.0

    pBdr = parse_xml(r'''
        <w:pBdr %s>
            <w:left w:val="single" w:sz="12" w:space="8" w:color="1F4E79"/>
        </w:pBdr>
    ''' % nsdecls('w'))
    p_code._p.get_or_add_pPr().append(pBdr)
    shd_code = parse_xml(r'<w:shd %s w:fill="F8F9FA"/>' % nsdecls('w'))
    p_code._p.get_or_add_pPr().append(shd_code)

    run_code = p_code.add_run(code_text)
    set_font(run_code, size_pt=9.5, color_rgb=(0x22, 0x22, 0x22))

    # -------------------------------------------------------------
    # 8. Conclusion
    # -------------------------------------------------------------
    add_h1("8. Conclusion")
    add_p(
        "This project successfully formulates campus placement as two complementary Machine Learning problems  -  classification of placement status "
        "and regression of the offered package  -  and implements a complete, verifiable pipeline covering problem definition, dataset construction, "
        "preprocessing, algorithm selection with justification, training and testing, rigorous evaluation (including cross-validation), and interpretation "
        "of results for a genuine real-world use case."
    )

    add_bullet(
        "Major Findings from the Analysis",
        "Both Random Forest and XGBoost achieve strong, comparable placement-classification performance (82-84% accuracy, ROC-AUC 0.91-0.93), "
        "confirming that CGPA, coding-skill proficiency, internships, aptitude score and communication ability are, together, far more informative "
        "than academic marks alone."
    )
    add_bullet(
        "Model Performance",
        "Both classification models exceed 80% accuracy with strong ROC-AUC, and both regression models exceed 0.90 R\u00b2 (Ridge Regression achieving "
        "0.938 R\u00b2 and 0.647 LPA MAE). Cross-validation confirms all reported metrics are stable across different data splits and not the result of a favourable train-test split."
    )
    add_bullet(
        "Important Insights Obtained from the Dataset",
        "Feature-importance analysis consistently highlights the same handful of features (CGPA, Coding Skill Score, Aptitude Score, Internships, "
        "Communication Score, Branch) across both tasks, giving a clear, actionable picture of what drives placement success and package. Placed students "
        "consistently show higher CGPA, coding skill, internships, aptitude and communication ratings, and fewer backlogs."
    )
    add_bullet(
        "Real-World Utility and Application",
        "For institutions / placement cells: the classification model can flag at-risk students (low predicted placement probability) before the recruitment "
        "drive, enabling targeted intervention programs  -  coding bootcamps, soft-skill workshops, additional aptitude practice, mock interviews  -  in time "
        "to improve outcomes. For students: the model quantifies which attributes most affect their placement chances and expected package, helping them "
        "prioritize effort. For expectation-setting: the regression model gives a realistic package estimate conditional on placement, useful for career counselling."
    )

    output_path = 'Campus_Placement_Project_Report.docx'
    doc.save(output_path)
    print(f"Report in old report look successfully saved to: {output_path}")

if __name__ == '__main__':
    create_report()
