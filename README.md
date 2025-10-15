# Niagara Point Validation Tool – Data Center Edition

## Overview
This repository contains a Python-based tool for validating **Tridium Niagara**, **Siemens Desigo CC**, and **Schneider EcoStruxure** point configurations in mission-critical **data center environments**. The tool performs automated quality checks on alarms, trends, and naming conventions to ensure compliance with design standards and consistency across systems.

It supports both **CSV** and **Excel** report formats exported from BMS systems and outputs consolidated reports highlighting compliance, mismatches, and missing points.

---

## Features

- **Supports Multiple Vendors:** Siemens Desigo CC, Schneider EcoStruxure, and Tridium Niagara.
- **Comprehensive Validation:** Checks alarm extensions, notification levels, delays, and trend intervals against the CDE (Construction Data Exchange) point list.
- **Automatic Compliance Reporting:** Generates Excel-based reports with color-coded highlights for mismatched or missing configurations.
- **GUI-Based Operation:** Built using PySimpleGUI, allowing non-technical users to run validations easily.
- **Visualization:** Produces pie charts summarizing compliance percentages.

---

## Typical Niagara Database Configurations in Data Centers

### Siemens (Desigo CC / PX Controllers)

#### Architecture
- PX series controllers (PXC, PXM) connect via **BACnet/IP**.
- Niagara integrates as a supervisory or aggregation layer.
- Data points: **AI**, **AO**, **BI**, **BO**, and **MV** objects.

#### Niagara Station Organization
| Component | Description |
|------------|-------------|
| **Folder Structure** | Organized by equipment: `AHU`, `CRAH`, `CHWS`, `ER`, `PUMP`, etc. |
| **Point Naming Convention** | `<System>_<Equipment>_<PointName>` (e.g., `AHU01_SupplyTemp`, `CRAH03_ValveCmd`). |
| **Trends** | Standard Niagara trend extensions, typically 5–15 minutes or COV. |
| **Alarms** | Inherited from PX metadata; managed through Niagara’s `BAlarm` service. |
| **Schedules** | Managed with `BSchedule` objects and PX mappings. |

#### Integration Notes
- Ensure BACnet Network Numbers are unique between Niagara and Desigo CC.
- Use Priority Levels 8–10 for Niagara supervisory writes.
- Align trend intervals and units across PX and Niagara databases.

---

### Schneider (EcoStruxure / SmartX Controllers)

#### Architecture
- SmartX AS-P and AS-B controllers communicate over **BACnet/IP**.
- Schneider’s EBO holds the native configuration; Niagara serves as a secondary supervisory or analytics layer.

#### Niagara Station Organization
| Component | Description |
|------------|-------------|
| **Folder Structure** | Grouped by zone or function (e.g., `AHU`, `CRAH`, `UPS`, `ER`). |
| **Point Naming Convention** | `<Equipment>_<Function>` (e.g., `CRAH02_SupplyTemp`, `UPS01_Status`). |
| **Alarms** | Imported using BACnet Notification Class mapping or local alarm service. |
| **Trends** | Configurable per site; typically 1–5 minutes. |
| **Schedules** | Usually managed from EBO or DCIM. |

#### Integration Notes
- Validate Notification Class IDs and COV increments during import.
- Niagara can poll via BACnet/IP or Modbus TCP depending on exposed objects.
- Mirror or subscribe to Schneider alarms to ensure synchronization.

---

## Common Niagara Best Practices for Data Centers

| Area | Best Practice |
|------|----------------|
| **Tagging** | Use Haystack tags (`equip`, `zone`, `point`, `siteRef`) for semantic consistency. |
| **Alarming** | Define standard levels: Critical (CR), Major (MJ), Minor (MN), Info (IN). |
| **Trend Retention** | Store local histories (180 days) before archiving to SQL or cloud. |
| **Network Design** | Place Niagara and controllers on a dedicated OT VLAN for isolation. |
| **Validation Scripts** | Use `niagara_point_validation.py` to compare points and detect discrepancies. |

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
