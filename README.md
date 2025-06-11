# 🚀 Pump Health Prediction via Vibration Signal Analysis

This repository contains the early-stage development of a predictive maintenance system that monitors the health of industrial pumps using vibration data collected from accelerometer sensors placed on key components — the **blower**, **gearbox**, and **motor**.

---

## ✅ Completed So Far

### 🔁 1. **Raw Data Extraction and Conversion**
- Original vibration sensor data was available in `.json` format.
- The data was **converted from JSON to CSV**, and subsequently to **`.xlsx` format** to facilitate easier data manipulation and compatibility with Excel and Python libraries.

### 📆 2. **Timestamp Conversion**
- All data files contained timestamps in **Unix epoch (milliseconds)** format.
- A batch script was developed to:
  - **Traverse folders recursively**
  - Detect `.xlsx` files containing a `timestamp` column
  - Convert these timestamps to **human-readable datetime** using Python’s `pandas.to_datetime()`
  - Save the updated files in place or with modified names

✅ Example:
Raw: 1712734753750
Converted: 2024-04-10 07:39:13.750
