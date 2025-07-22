# 🚀 Pump Health Prediction using Sensor-Based Time-Series & Machine Learning Techniques

This repository presents a comprehensive end-to-end pipeline for assessing pump health based on sensor-generated time-series data. The workflow includes raw data ingestion, missing value handling, outlier detection, feature extraction using FFT and recurrence methods, temporal pattern analysis, and machine learning-based health classification. The aim is to provide an automated, interpretable framework for detecting early anomalies and estimating health stages without any predefined thresholds or labels.

---

## 📁 Repository Structure

```bash
pump-health-prediction/
│
├── Data/
│   ├── Raw/                            # Original sensor data (categorized by operating conditions)
│   │   ├── Off condition/
│   │   ├── On condition/
│   │   ├── Op condition/
│   │   ├── Op condition-1/
│   │   ├── Op condition-2/
│   │   └── Op condition-3/
│   └── Processed/                      # Optional: preprocessed outputs and intermediate files
│
├── Reports/
│   ├── CERTIFICATE.pdf                 # Internship completion certificate
│   └── PUMP HEALTH PREDICTION.pdf      # Final project report
│
├── Scripts/                            # All preprocessing and model scripts
│   ├── 01_convert_json_to_excel.py
│   ├── 02_convert_timestamp_date-time.py
│   ├── 03_convert_g_mps2.py
│   ├── 04_detect_missing_values.py
│   ├── 05_missing_value_reports.py
│   ├── 06_handle_missing_values_using_rollingmean.py
│   ├── 07_impute_missing_values.py
│   ├── 08_time_series.py
│   ├── 09_outlier_detection_box_plot.py
│   ├── 10_outlier_detection_z-score.py
│   ├── 11_outlier_classification_01.py
│   ├── 12_outlier_classification_02.py
│   ├── 13_outlier_classification_03.py
│   ├── 14_FFT_feature.py
│   ├── 15_final_score_label.py
│   ├── dataclean.py
│   └── handle_outlier_values_using_rolling_mean.py
│
├── Utils/                              # (Optional) Helper utilities or configs
│   └── (Place logger.py, config.py if required)
│
├── LICENSE
├── README.md
├── requirements.txt                    # Python dependencies
└── .gitignore                          # Files to exclude from Git versioning

📊 Dataset Files (Required)

To run this project, ensure the following directory structure under Data/Raw:
Data/
└── Raw/
    ├── Off condition/
    ├── On condition/
    ├── Op condition/
    ├── Op condition-1/
    ├── Op condition-2/
    └── Op condition-3/

⚙️ Setup & Execution

1. Clone the Repository

git clone https://github.com/Ashmit-workplace/pump-health-prediction.git
cd pump-health-prediction

2. Install Dependencies

We recommend creating a virtual environment.
pip install -r requirements.txt

3. Pipeline Execution (Script Order)

Run the scripts in the following order for complete execution:
python Scripts/01_convert_json_to_excel.py
python Scripts/02_convert_timestamp_date-time.py
python Scripts/03_convert_g_mps2.py
python Scripts/04_detect_missing_values.py
python Scripts/05_missing_value_reports.py
python Scripts/06_handle_missing_values_using_rollingmean.py
python Scripts/07_impute_missing_values.py
python Scripts/08_time_series.py
python Scripts/09_outlier_detection_box_plot.py
python Scripts/10_outlier_detection_z-score.py
python Scripts/11_outlier_classification_01.py
python Scripts/12_outlier_classification_02.py
python Scripts/13_outlier_classification_03.py
python Scripts/14_FFT_feature.py
python Scripts/15_final_score_label.py

This will generate fully preprocessed and labeled datasets ready for ML modeling.

ML Model Training
Feature vectors extracted include: FFT coefficients, recurrence counts, temporal flags, and contextual anomaly scores.

Models used:

Random Forest Classifier

Support Vector Machine (RBF Kernel)

Decision Tree

Performance metrics including accuracy, precision, recall, and F1-score are included in the final report.






