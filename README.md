# Niagara Point Validation Tool – IN PROGRESS

## Overview
This repository contains a Python-based tool for validating **Tridium Niagara** point configurations. The tool performs automated quality checks on alarms, trends, and naming conventions to ensure compliance with design standards and consistency across systems.

Alarms in a **Tridium Niagara System** are typically configured by dragging and dropping alarm extensions from the **Alarm Palette**. Common alarm extensions include:

- **OutOfRangeAlarmExt**  
  Used for **Analog Inputs (AI)** and **Analog Values (AV)** such as:
  - Cold Aisle Temperature Sensors  
  - Pressure Sensors  
  - Supply Air Temperature Sensors  

- **BooleanChangeOfStateAlarmExt**  
  Used for **Binary Inputs (BI)** and **Binary Values (BV)** such as:
  - Power Fail Alarm  
  - Fan Trip Alarm  
  - Chiller Trip Alarm  

### Key Alarm Configurations
- **Notification Class**: `Critical`, `High`, `Medium`, `Low`  
- **Delays**  
- **High and Low Limits** (for temperature, pressure, etc.)  

These configurations are critical to ensure **Building Operators trust the alarm system**.

---

## Validation Workflow

Follow these steps to validate Tridium Niagara alarms and trends:

### 1. Database Setup
Ensure that **control point names** in the Tridium N4 database include the strings from the **Point Name** column in the **Construction Data Exchange List**.

### 2. Extract Configuration Data
- Write **BQL Queries** to pull **Alarm** and **Trend** configuration data from the Tridium Niagara database.
- Export the **as-built alarm and trend configuration data** as **CSV files**.

### 3. Run Validation Script
- Feed the extracted CSV files into the **Niagara Validation Script**.
- The script generates an **Excel report** with color-coded results:

| Color   | Meaning                                                                 |
|---------|-------------------------------------------------------------------------|
| **Red**    | Alarm attribute does **NOT** match design requirements (e.g., Notification Class) |
| **Yellow** | Alarm attribute could **NOT** be validated by the tool             |
| **Green**  | Alarm attribute is validated and matches design requirements        |

---

## Example Output
The validation report will help identify:
- Misconfigured alarms
- Attributes that need manual review
- Fully compliant configurations

---

## Outputs

After execution, the tool generates the following reports:

- **Niagara_Point_Validation_Report.xlsx** – Consolidated report listing all points validated, with highlights for mismatches and missing alarms.
- **Niagara_Validation_Overview.xlsx** – Summary sheet with total, compliant, and non-compliant points, along with pie chart visualization.

---

## Example Niagara Hierarchy
```
Niagara Station
├── AHU
│   ├── AHU01
│   │   ├── SupplyTemp (AI)
│   │   ├── ReturnTemp (AI)
│   │   ├── FanStatus (BI)
│   │   ├── SupplyFanCmd (BO)
│   │   └── Alarms
├── CRAH
│   ├── CRAH01
│   ├── CRAH02
├── ER
│   ├── ER01
│   │   ├── SmokeDetStatus
│   │   ├── TempSensor
│   │   └── Alarms
└── CHWS
    ├── Pump01
    ├── Pump02
    └── Valves
```

---

## Requirements

```bash
pip install pandas openpyxl xlsxwriter PySimpleGUI matplotlib
```

Python 3.8 or later is recommended.

---

## Example Usage

```bash
python niagara_point_validation.py
```

1. Launch the GUI window.
2. Select the Siemens or Schneider option.
3. Browse and select the Alarm, Trend, and CDE Point List files.
4. Choose an output folder.
5. Run the validation process.

The tool will automatically generate summary reports in the selected directory.

---

## Summary
This tool simplifies the validation of multi-vendor Niagara databases used in data centers. It ensures configuration integrity, trend alignment, and alarm consistency between Siemens, Schneider, and Tridium systems — helping engineers reduce commissioning effort and maintain compliance with data center operational standards.
