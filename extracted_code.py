

# ======================================================================import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.preprocessing import StandardScaler, LabelEncoder
from sklearn.impute import SimpleImputer
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.linear_model import Ridge
from sklearn.metrics import (accuracy_score, precision_score, recall_score, f1_score,
                              confusion_matrix, classification_report, roc_auc_score, roc_curve,
                              mean_absolute_error, mean_squared_error, r2_score)
from xgboost import XGBClassifier

RANDOM_STATE = 42
np.random.seed(RANDOM_STATE)

sns.set_theme(style="whitegrid")
plt.rcParams['figure.figsize'] = (8, 5)

pd.set_option('display.max_columns', None)
print("Libraries imported successfully.")

def generate_dataset(n=3000, seed=42):
    rng = np.random.default_rng(seed)

    gender = rng.choice(['Male', 'Female'], size=n, p=[0.58, 0.42])
    branch = rng.choice(
        ['Computer Science', 'Information Technology', 'Electronics', 'Mechanical', 'Civil', 'Electrical'],
        size=n, p=[0.30, 0.20, 0.20, 0.12, 0.09, 0.09]
    )

    ssc = np.clip(rng.normal(75, 10, n), 40, 98)
    hsc = np.clip(ssc + rng.normal(0, 8, n), 35, 98)
    cgpa = np.clip(rng.normal(7.2, 1.0, n), 4.0, 10.0)

    internships = np.clip(rng.poisson(1.1, n), 0, 4)
    projects = np.clip(rng.poisson(2.0, n), 0, 6)
    certifications = np.clip(rng.poisson(1.3, n), 0, 5)

    def norm(x):
        return (x - x.min()) / (x.max() - x.min())

    coding_raw = 0.55 * cgpa * 10 + rng.normal(0, 10, n) + 3 * internships + 2 * projects
    coding_skill = norm(coding_raw) * 100

    aptitude = np.clip(rng.normal(65, 15, n), 20, 100)
    communication = np.clip(rng.normal(6.5, 1.5, n), 1, 10)
    backlogs = np.clip(rng.poisson(0.4, n), 0, 5)

    extracurricular = rng.choice(['Yes', 'No'], size=n, p=[0.45, 0.55])
    placement_training = rng.choice(['Yes', 'No'], size=n, p=[0.55, 0.45])

    # An explicit interaction term: strong coding skill can offset an average CGPA, and vice-versa
    interaction = norm(cgpa) * norm(coding_skill)

    logit = (
        2.8 * norm(cgpa) +
        2.6 * norm(coding_skill) +
        2.0 * interaction +
        1.3 * norm(internships.astype(float)) +
        1.0 * norm(projects.astype(float)) +
        1.3 * norm(aptitude) +
        1.1 * norm(communication) +
        0.7 * (placement_training == 'Yes').astype(float) +
        0.5 * norm(certifications.astype(float)) -
        1.8 * norm(backlogs.astype(float)) +
        rng.normal(0, 0.35, n)
    )
    logit = logit - np.median(logit)
    prob = 1 / (1 + np.exp(-3.2 * logit))
    placed = (rng.random(n) < prob).astype(int)

    branch_mult = pd.Series(branch).map({
        'Computer Science': 1.35, 'Information Technology': 1.22, 'Electronics': 1.08,
        'Electrical': 1.04, 'Mechanical': 0.96, 'Civil': 0.90
    }).values

    package_raw = (3.5 + (7.5 * norm(cgpa) + 8.5 * norm(coding_skill) + 3.0 * interaction
                   + 2.0 * norm(internships.astype(float)) + 1.2 * norm(aptitude)
                   + 1.4 * norm(communication)) * branch_mult + rng.normal(0, 0.7, n))
    package = np.clip(package_raw, 3.0, 35.0)
    package_lpa = np.where(placed == 1, package, np.nan)

    df = pd.DataFrame({
        'Gender': gender,
        'Branch': branch,
        'SSC_Percentage': np.round(ssc, 2),
        'HSC_Percentage': np.round(hsc, 2),
        'CGPA': np.round(cgpa, 2),
        'Internships': internships,
        'Projects': projects,
        'Certifications': certifications,
        'Coding_Skill_Score': np.round(coding_skill, 1),
        'Aptitude_Test_Score': np.round(aptitude, 1),
        'Communication_Score': np.round(communication, 1),
        'Backlogs': backlogs,
        'Extracurricular_Activities': extracurricular,
        'Placement_Training': placement_training,
        'Placement_Status': np.where(placed == 1, 'Placed', 'Not Placed'),
        'Package_LPA': np.round(package_lpa, 2)
    })

    # Inject a small % of realistic missing values (as in real-world data collection)
    for col, frac in [('Aptitude_Test_Score', 0.03), ('Communication_Score', 0.025),
                       ('Backlogs', 0.02), ('HSC_Percentage', 0.015)]:
        miss_idx = rng.choice(df.index, size=int(frac * len(df)), replace=False)
        df.loc[miss_idx, col] = np.nan

    return df


df = generate_dataset(n=3000, seed=RANDOM_STATE)
df.to_csv('campus_placement_dataset.csv', index=False)
print("Dataset shape:", df.shape)
df.head()

df.info()

df.describe(include='all').T

print("Missing values per column:\n")
print(df.isnull().sum())
print("\nPlacement status distribution:")
print(df['Placement_Status'].value_counts())
print(df['Placement_Status'].value_counts(normalize=True).round(3) * 100)

fig, axes = plt.subplots(1, 2, figsize=(13, 5))

sns.countplot(data=df, x='Placement_Status', hue='Placement_Status', palette='viridis',
              legend=False, ax=axes[0])
axes[0].set_title('Placement Status Distribution')

sns.histplot(df.loc[df['Placement_Status'] == 'Placed', 'Package_LPA'], bins=25, kde=True,
             color='seagreen', ax=axes[1])
axes[1].set_title('Package Distribution (Placed Students, LPA)')
axes[1].set_xlabel('Package (LPA)')

plt.tight_layout()
plt.show()

numeric_cols = ['SSC_Percentage', 'HSC_Percentage', 'CGPA', 'Internships', 'Projects',
                 'Certifications', 'Coding_Skill_Score', 'Aptitude_Test_Score',
                 'Communication_Score', 'Backlogs']

fig, axes = plt.subplots(2, 5, figsize=(20, 8))
axes = axes.flatten()
for i, col in enumerate(numeric_cols):
    sns.boxplot(data=df, x='Placement_Status', y=col, hue='Placement_Status',
                palette='Set2', legend=False, ax=axes[i])
    axes[i].set_title(col)
plt.tight_layout()
plt.show()

plt.figure(figsize=(10, 8))
corr = df[numeric_cols].corr()
sns.heatmap(corr, annot=True, fmt='.2f', cmap='coolwarm', center=0)
plt.title('Correlation Heatmap of Numeric Features')
plt.tight_layout()
plt.show()

plt.figure(figsize=(8, 6))
sns.scatterplot(data=df[df['Placement_Status'] == 'Placed'], x='CGPA', y='Package_LPA',
                 hue='Coding_Skill_Score', palette='viridis', alpha=0.6)
plt.title('CGPA vs Package (colored by Coding Skill Score) - Placed Students')
plt.tight_layout()
plt.show()

df_clean = df.copy()

# 4.1 Handle missing values (median imputation for numeric columns)
num_impute_cols = ['Aptitude_Test_Score', 'Communication_Score', 'Backlogs', 'HSC_Percentage']
imputer = SimpleImputer(strategy='median')
df_clean[num_impute_cols] = imputer.fit_transform(df_clean[num_impute_cols])

print("Missing values after imputation:")
print(df_clean.isnull().sum())

# 4.2 Encode categorical variables
df_model = df_clean.copy()

binary_map_cols = ['Extracurricular_Activities', 'Placement_Training', 'Gender']
for col in binary_map_cols:
    le = LabelEncoder()
    df_model[col] = le.fit_transform(df_model[col])   # e.g. No=0, Yes=1 / Female=0, Male=1

df_model = pd.get_dummies(df_model, columns=['Branch'], prefix='Branch', drop_first=True)

# Encode target for classification: Placed = 1, Not Placed = 0
target_le = LabelEncoder()
df_model['Placement_Status_Encoded'] = target_le.fit_transform(df_model['Placement_Status'])
print("Classes:", dict(zip(target_le.classes_, target_le.transform(target_le.classes_))))

df_model.head()

feature_cols = [c for c in df_model.columns
                if c not in ['Placement_Status', 'Placement_Status_Encoded', 'Package_LPA']]

X = df_model[feature_cols]
y_clf = df_model['Placement_Status_Encoded']

print("Feature columns used for modeling:")
print(feature_cols)
print("\nFeature matrix shape:", X.shape)

X_train, X_test, y_train, y_test = train_test_split(
    X, y_clf, test_size=0.2, random_state=RANDOM_STATE, stratify=y_clf
)
print("Train shape:", X_train.shape, "Test shape:", X_test.shape)

import time

# Random Forest Classifier
rf_clf = RandomForestClassifier(
    n_estimators=300, max_depth=8, min_samples_leaf=3,
    random_state=RANDOM_STATE, n_jobs=-1
)

start = time.time()
rf_clf.fit(X_train, y_train)
train_time_rf = time.time() - start

rf_pred = rf_clf.predict(X_test)
rf_proba = rf_clf.predict_proba(X_test)[:, 1]

print(f"Random Forest Classifier trained in {train_time_rf:.3f} seconds.")
print(f"Number of trees in the fitted forest : {len(rf_clf.estimators_)}")
print(f"Number of features seen during fit   : {rf_clf.n_features_in_}")

# XGBoost Classifier
xgb_clf = XGBClassifier(
    n_estimators=300, max_depth=4, learning_rate=0.05,
    subsample=0.8, colsample_bytree=0.8,
    eval_metric='logloss', random_state=RANDOM_STATE
)

start = time.time()
xgb_clf.fit(X_train, y_train)
train_time_xgb = time.time() - start

xgb_pred = xgb_clf.predict(X_test)
xgb_proba = xgb_clf.predict_proba(X_test)[:, 1]

print(f"XGBoost Classifier trained in {train_time_xgb:.3f} seconds.")
print(f"Number of boosting rounds (trees) fitted: {xgb_clf.get_booster().num_boosted_rounds()}")

from sklearn.model_selection import cross_val_score, StratifiedKFold

cv = StratifiedKFold(n_splits=5, shuffle=True, random_state=RANDOM_STATE)

start = time.time()
rf_cv_scores = cross_val_score(rf_clf, X, y_clf, cv=cv, scoring='accuracy', n_jobs=-1)
xgb_cv_scores = cross_val_score(xgb_clf, X, y_clf, cv=cv, scoring='accuracy', n_jobs=-1)
cv_time = time.time() - start

print(f"5-fold cross-validation completed in {cv_time:.2f} seconds (10 model fits total: 5 folds x 2 models).")
print(f"Random Forest  - CV accuracy per fold: {np.round(rf_cv_scores, 4)}")
print(f"Random Forest  - Mean CV accuracy    : {rf_cv_scores.mean():.4f} (+/- {rf_cv_scores.std():.4f})")
print(f"XGBoost        - CV accuracy per fold: {np.round(xgb_cv_scores, 4)}")
print(f"XGBoost        - Mean CV accuracy    : {xgb_cv_scores.mean():.4f} (+/- {xgb_cv_scores.std():.4f})")

def classification_report_summary(name, y_true, y_pred, y_proba):
    acc = accuracy_score(y_true, y_pred)
    prec = precision_score(y_true, y_pred)
    rec = recall_score(y_true, y_pred)
    f1 = f1_score(y_true, y_pred)
    auc = roc_auc_score(y_true, y_proba)
    print(f"--- {name} ---")
    print(f"Accuracy : {acc:.4f}")
    print(f"Precision: {prec:.4f}")
    print(f"Recall   : {rec:.4f}")
    print(f"F1-Score : {f1:.4f}")
    print(f"ROC-AUC  : {auc:.4f}")
    print()
    return {'Model': name, 'Accuracy': acc, 'Precision': prec, 'Recall': rec, 'F1': f1, 'ROC_AUC': auc}

results_rf = classification_report_summary("Random Forest Classifier", y_test, rf_pred, rf_proba)
results_xgb = classification_report_summary("XGBoost Classifier", y_test, xgb_pred, xgb_proba)

clf_comparison = pd.DataFrame([results_rf, results_xgb]).set_index('Model')
clf_comparison

fig, axes = plt.subplots(1, 2, figsize=(13, 5))

for ax, (name, pred) in zip(axes, [("Random Forest", rf_pred), ("XGBoost", xgb_pred)]):
    cm = confusion_matrix(y_test, pred)
    sns.heatmap(cm, annot=True, fmt='d', cmap='Blues',
                xticklabels=target_le.classes_, yticklabels=target_le.classes_, ax=ax)
    ax.set_title(f'Confusion Matrix - {name}')
    ax.set_xlabel('Predicted')
    ax.set_ylabel('Actual')

plt.tight_layout()
plt.show()

print("Random Forest classification report:\n", classification_report(y_test, rf_pred, target_names=target_le.classes_))
print("XGBoost classification report:\n", classification_report(y_test, xgb_pred, target_names=target_le.classes_))

plt.figure(figsize=(7, 6))
for name, proba in [("Random Forest", rf_proba), ("XGBoost", xgb_proba)]:
    fpr, tpr, _ = roc_curve(y_test, proba)
    auc_val = roc_auc_score(y_test, proba)
    plt.plot(fpr, tpr, label=f"{name} (AUC = {auc_val:.3f})")

plt.plot([0, 1], [0, 1], 'k--', label='Random Guess')
plt.xlabel('False Positive Rate')
plt.ylabel('True Positive Rate')
plt.title('ROC Curve - Placement Classification')
plt.legend()
plt.tight_layout()
plt.show()

importances = pd.Series(rf_clf.feature_importances_, index=feature_cols).sort_values(ascending=False)

plt.figure(figsize=(9, 6))
sns.barplot(x=importances.values[:10], y=importances.index[:10], hue=importances.index[:10],
            palette='viridis', legend=False)
plt.title('Top 10 Feature Importances - Random Forest (Placement Classification)')
plt.xlabel('Importance')
plt.tight_layout()
plt.show()

importances.head(10)

df_placed = df_model[df_model['Placement_Status'] == 'Placed'].copy()
X_reg = df_placed[feature_cols]
y_reg = df_placed['Package_LPA']

print("Placed-students feature matrix shape:", X_reg.shape)

Xr_train, Xr_test, yr_train, yr_test = train_test_split(
    X_reg, y_reg, test_size=0.2, random_state=RANDOM_STATE
)

# Scale features for Ridge Regression
scaler = StandardScaler()
Xr_train_scaled = scaler.fit_transform(Xr_train)
Xr_test_scaled = scaler.transform(Xr_test)

print("Regression train shape:", Xr_train.shape, "test shape:", Xr_test.shape)

# Random Forest Regressor (uses unscaled features)
rf_reg = RandomForestRegressor(
    n_estimators=300, max_depth=8, min_samples_leaf=3,
    random_state=RANDOM_STATE, n_jobs=-1
)
start = time.time()
rf_reg.fit(Xr_train, yr_train)
train_time_rfreg = time.time() - start
rf_reg_pred = rf_reg.predict(Xr_test)

# Ridge Regression (uses scaled features)
ridge_reg = Ridge(alpha=1.0, random_state=RANDOM_STATE)
start = time.time()
ridge_reg.fit(Xr_train_scaled, yr_train)
train_time_ridge = time.time() - start
ridge_reg_pred = ridge_reg.predict(Xr_test_scaled)

print(f"Random Forest Regressor trained in {train_time_rfreg:.3f} seconds "
      f"({len(rf_reg.estimators_)} trees, {rf_reg.n_features_in_} features).")
print(f"Ridge Regression trained in {train_time_ridge:.5f} seconds "
      f"({ridge_reg.coef_.shape[0]} learned coefficients).")

from sklearn.model_selection import KFold

cv_reg = KFold(n_splits=5, shuffle=True, random_state=RANDOM_STATE)
X_reg_scaled = StandardScaler().fit_transform(X_reg)

start = time.time()
rf_reg_cv = cross_val_score(rf_reg, X_reg, y_reg, cv=cv_reg, scoring='r2', n_jobs=-1)
ridge_cv = cross_val_score(ridge_reg, X_reg_scaled, y_reg, cv=cv_reg, scoring='r2', n_jobs=-1)
cv_reg_time = time.time() - start

print(f"5-fold cross-validation (regression) completed in {cv_reg_time:.2f} seconds.")
print(f"Random Forest Regressor - CV R2 per fold: {np.round(rf_reg_cv, 4)}")
print(f"Random Forest Regressor - Mean CV R2    : {rf_reg_cv.mean():.4f} (+/- {rf_reg_cv.std():.4f})")
print(f"Ridge Regression        - CV R2 per fold: {np.round(ridge_cv, 4)}")
print(f"Ridge Regression        - Mean CV R2    : {ridge_cv.mean():.4f} (+/- {ridge_cv.std():.4f})")

def regression_report_summary(name, y_true, y_pred):
    mae = mean_absolute_error(y_true, y_pred)
    rmse = np.sqrt(mean_squared_error(y_true, y_pred))
    r2 = r2_score(y_true, y_pred)
    print(f"--- {name} ---")
    print(f"MAE  : {mae:.3f} LPA")
    print(f"RMSE : {rmse:.3f} LPA")
    print(f"R2   : {r2:.4f}")
    print()
    return {'Model': name, 'MAE': mae, 'RMSE': rmse, 'R2': r2}

res_rf_reg = regression_report_summary("Random Forest Regressor", yr_test, rf_reg_pred)
res_ridge_reg = regression_report_summary("Ridge Regression", yr_test, ridge_reg_pred)

reg_comparison = pd.DataFrame([res_rf_reg, res_ridge_reg]).set_index('Model')
reg_comparison

best_reg_model = reg_comparison['R2'].idxmax()
print(f"Best-performing regression model on this test split (by R2): {best_reg_model}")

fig, axes = plt.subplots(1, 2, figsize=(13, 5))

for ax, (name, pred) in zip(axes, [("Random Forest Regressor", rf_reg_pred), ("Ridge Regression", ridge_reg_pred)]):
    ax.scatter(yr_test, pred, alpha=0.5, color='teal')
    lims = [min(yr_test.min(), pred.min()), max(yr_test.max(), pred.max())]
    ax.plot(lims, lims, 'r--', label='Ideal fit')
    ax.set_xlabel('Actual Package (LPA)')
    ax.set_ylabel('Predicted Package (LPA)')
    ax.set_title(f'{name}: Actual vs Predicted')
    ax.legend()

plt.tight_layout()
plt.show()

reg_importances = pd.Series(rf_reg.feature_importances_, index=feature_cols).sort_values(ascending=False)

plt.figure(figsize=(9, 6))
sns.barplot(x=reg_importances.values[:10], y=reg_importances.index[:10], hue=reg_importances.index[:10],
            palette='mako', legend=False)
plt.title('Top 10 Feature Importances - Random Forest (Package Regression)')
plt.xlabel('Importance')
plt.tight_layout()
plt.show()

reg_importances.head(10)

def predict_new_student(student_dict, rf_classifier=rf_clf, rf_regressor=rf_reg,
                         feature_columns=feature_cols, target_encoder=target_le):
    """
    student_dict: dict with keys matching original df_clean columns
    (Gender, Branch, SSC_Percentage, HSC_Percentage, CGPA, Internships, Projects,
     Certifications, Coding_Skill_Score, Aptitude_Test_Score, Communication_Score,
     Backlogs, Extracurricular_Activities, Placement_Training)
    """
    row = pd.DataFrame([student_dict])

    for col in ['Extracurricular_Activities', 'Placement_Training']:
        row[col] = row[col].map({'No': 0, 'Yes': 1})
    row['Gender'] = row['Gender'].map({'Female': 0, 'Male': 1})

    row = pd.get_dummies(row, columns=['Branch'], prefix='Branch')

    # Align columns with training feature set (fills any missing dummy columns with 0)
    row = row.reindex(columns=feature_columns, fill_value=0)

    placement_pred = rf_classifier.predict(row)[0]
    placement_proba = rf_classifier.predict_proba(row)[0][1]
    status = target_encoder.inverse_transform([placement_pred])[0]

    result = {'Predicted_Status': status, 'Placement_Probability': round(float(placement_proba), 3)}

    if status == 'Placed':
        predicted_package = rf_regressor.predict(row)[0]
        result['Predicted_Package_LPA'] = round(float(predicted_package), 2)

    return result


sample_student = {
    'Gender': 'Female',
    'Branch': 'Computer Science',
    'SSC_Percentage': 88.0,
    'HSC_Percentage': 85.0,
    'CGPA': 8.6,
    'Internships': 2,
    'Projects': 4,
    'Certifications': 3,
    'Coding_Skill_Score': 82.0,
    'Aptitude_Test_Score': 78.0,
    'Communication_Score': 8.0,
    'Backlogs': 0,
    'Extracurricular_Activities': 'Yes',
    'Placement_Training': 'Yes'
}

predict_new_student(sample_student)

print("=== Classification (Placement Status) ===")
display(clf_comparison)
print()
print("=== Regression (Package in LPA, placed students only) ===")
display(reg_comparison)