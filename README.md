# Niagara Alarm & Trend Validation Tool

The Niagara Alarm & Trend Validation Tool automatically checks that the **alarm and trend configuration** built in **Tridium Niagara (Siemens / Schneider)** matches the project's **Control Point List (CPL)**.

Instead of manually comparing exported Niagara configurations against design documentation point-by-point, this tool loads both sources, matches the points, applies validation rules, and produces a colour-coded Excel report showing exactly where the as-built system agrees with — or deviates from — the design intent.

It runs as a lightweight **PySimpleGUI** desktop application built around a simple, deterministic **pipeline architecture**.

## Introduction

In Niagara N4, alarms are configured directly on control points using **Alarm Palettes** and various **Alarm Extensions**. The tool understands the most common alarm types:

- **Boolean Alarm Extension** – triggers when a boolean point evaluates to a specific state.
- **Numeric Alarm Extension** – evaluates thresholds on numeric values.
- **OutOfRange / Limit Alarms** – high/low conditions tied to analog values.
- **Fault / Status Alarms** – based on device health or communication status.

Because these settings are entered by hand during commissioning, they can drift from the approved design. This tool provides a **repeatable, engineering-focused** way to catch those differences before handover.

For **existing sites**, the tool also lets you compare the current as-built configuration against the latest iteration of the design standard (CPL), identifying gaps and allowing the configurations to be updated so as to align with the current standard.

> **Note:** This validation depends on the **Control Point List following a standard template**, so that the tool can reliably parse the exported output and perform the comparison. If the CPL does not conform to the expected template structure, the parsing and validation steps will not produce accurate results.

# System Architecture

The tool operates as a simple, deterministic pipeline. It takes two inputs:

- **As-built alarm & trend configuration** exported from Niagara (CSV)
- **Design intent** defined in the Control Point List (CPL, Excel)

and produces a single **validated configuration report** (Excel).

The block diagram below shows the main layers and how data flows through them:

```
+-----------------------------+         +-----------------------------+
|  Control Point List (CPL)   |         |    Niagara JACE/Supervisor  |
|  - Design requirements      |         |    - Alarm configuration    |
|  - Point names & metadata   |         |    - Trend configuration    |
|  (Excel)                    |         |    (CSV via BQL/Reports)    |
+--------------+--------------+         +--------------+--------------+
               \                                      /
                \                                    /
                 \                                  /
                  v                                v
              +------------------------------------------+
              |           Data Extraction Layer          |
              |  - Read CPL (Excel)                      |
              |  - Read Alarm/Trend CSVs                |
              +----------------------+-------------------+
                                     |
                                     v
              +------------------------------------------+
              |           Normalization Layer            |
              |  - Clean & uppercase point names         |
              |  - Remove symbols / brackets             |
              |  - Normalize delays and formats          |
              +----------------------+-------------------+
                                     |
                                     v
              +------------------------------------------+
              |          Point Matching Engine           |
              |  - Tokenize names (Niagara & CPL)        |
              |  - Match as-built points to CPL entries  |
              |  - Detect missing / extra points         |
              +----------------------+-------------------+
                                     |
                                     v
              +------------------------------------------+
              |            Rule Evaluation Layer         |
              |  - Alarm Class vs Notification Level     |
              |  - Delay vs CPL Delay                    |
              |  - Trend enablement / intervals          |
              +----------------------+-------------------+
                                     |
                                     v
              +------------------------------------------+
              |             Reporting Layer              |
              |  - Generate Excel reports                |
              |  - Red = mismatch, Green = match         |
              |  - Yellow = manual review required       |
              +----------------------+-------------------+
                                     |
                                     v
              +------------------------------------------+
              |       Final Compliance Report (Excel)    |
              +------------------------------------------+
```

## Pipeline Stages

1. **Data Extraction** – Load the Alarm and Trend CSVs exported from the Niagara JACE/Supervisor, plus the CPL (Excel) from the project documentation.

2. **Normalization** – Standardize point names, strip unwanted characters and brackets, uppercase text, and normalize alarm/trend attributes such as delays so the two sources can be compared reliably.

3. **Point Naming & Matching** – Tokenize the Niagara and CPL names (which usually share prefixes and tokens) and match each as-built point to its corresponding CPL entry, flagging any that are missing or extra.

4. **Rule Evaluation** – Compare Alarm Class, Notification Level, Delay, and Trend settings for every matched point to identify mismatches, missing points, and fully compliant entries.

5. **Report Generation** – Write Excel reports with colour highlighting for mismatches, matches, and items that need manual review.

---

# Process Flow Diagrams

## Niagara Configuration Export Flow

This process describes how the as-built configuration is pulled out of Niagara. Using the Workbench, the alarm and trend extensions are queried via the Report Service and BQL, then exported into CSV files (Boolean alarms, Numeric alarms, and Trends). These files become the raw input consumed by the validation tool.

```
+-----------------------------+
|     Niagara Workbench       |
|  (JACE / Supervisor)        |
+--------------+--------------+
               |
               v
+-----------------------------+
|      Report Service / BQL   |
|  - Query alarm extensions   |
|  - Query trend extensions   |
+--------------+--------------+
               |
               v
+-----------------------------+
|      Export to CSV          |
|  - Alarm File 1 (Boolean)   |
|  - Alarm File 2 (Numeric)   |
|  - Trend File               |
+--------------+--------------+
               |
               v
+-----------------------------+
|   Files Ready for Tool      |
+-----------------------------+
```

## Validation Workflow

This process describes how the exported data is checked against design intent. The Alarm/Trend CSVs and the Control Point List (CPL) are loaded and normalized, as-built points are matched to their CPL entries, and each match is evaluated against the configured rules (alarm class, notification level, delays, and trend settings). The result is a colour-coded Excel report flagging mismatches, compliant items, and entries needing manual review.

```
+-----------------------------+     +-----------------------------+
|   Alarm / Trend CSVs        |     |   Control Point List (CPL)  |
+--------------+--------------+     +--------------+--------------+
               \                                   /
                \                                 /
                 v                               v
              +------------------------------------+
              |        Load & Normalize Data       |
              +------------------+-----------------+
                                 |
                                 v
              +------------------------------------+
              |      Match Points to CPL Entries   |
              +------------------+-----------------+
                                 |
                                 v
              +------------------------------------+
              |        Evaluate Rules              |
              |  - Alarm Class / Notification      |
              |  - Delay values                    |
              |  - Trend enablement / intervals    |
              +------------------+-----------------+
                                 |
                                 v
              +------------------------------------+
              |     Generate Excel Report          |
              |  Red / Green / Yellow highlights   |
              +------------------------------------+
```

---

# PySimpleGUI Interface

The application provides a simple desktop interface so no command line is required:

- **File inputs** for the data sources:
  - Alarm File 1 (Boolean alarms)
  - Alarm File 2 (Numeric alarms)
  - Trend File
  - CPL (Control Point List)
- **Vendor-specific validation buttons:**
  - Point Validation (Siemens)
  - Point Validation (Schneider)
- **Preview buttons** to inspect the loaded data directly in the GUI before running
- **Progress bar** to track long validation runs

---

# Requirements

## Environment

- **Python 3.8+** (Windows recommended)
- **Microsoft Excel** installed \u2014 required by `xlwings` for report generation and formatting

## Python Packages

The tool depends on the following third-party packages:

| Package        | Purpose                                             |
| -------------- | --------------------------------------------------- |
| `pandas`       | Load and manipulate CSV / Excel data                |
| `numpy`        | Numeric operations and array handling               |
| `PySimpleGUI`  | Desktop graphical user interface                    |
| `xlwings`      | Read/write and format Excel reports                 |
| `openpyxl`     | Read/write `.xlsx` files and apply cell fills       |
| `xlsxwriter`   | Write formatted Excel output                        |
| `matplotlib`   | Plotting / visualization support                    |

Standard-library modules (`os`, `re`, `time`, `threading`, `warnings`, `pathlib`) require no installation.

## Installation

Install all dependencies with:

```bash
pip install -r requirements.txt
```

---

# Usage

1. Launch the GUI.
2. Select the Alarm/Trend CSVs and the CPL file.
3. Choose Siemens or Schneider validation to match the project vendor.
4. Select the output folder for the report.
5. Run the validation.
6. The Excel report is generated automatically when processing completes.

---

# Output Summary

The generated Excel report uses colour coding so issues stand out at a glance:

- **Red** → Configuration does **not** match the CPL
- **Green** → Fully compliant
- **Yellow** → Partially validated (manual review needed)

Together these give a fast, repeatable, engineering-focused view of how closely the as-built Niagara alarm and trend configuration follows the approved design.
