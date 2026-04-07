# Automated Resistograph Processing Suite

This repository contains an automated R-based pipeline for processing, visualizing, merging, and validating Resistograph (`.rgp`) drilling data. It is designed to handle large-scale forestry and tree-breeding experiments, taking raw machine outputs and transforming them into analysis-ready datasets alongside detailed visual diagnostic reports.

## 🚀 Key Features
* **Automated Trace Detrending:** Mathematically corrects raw resistance amplitudes for feed and drill speeds.
* **Feature Extraction:** Automatically detects bark entry/exit, calculates DBH, identifies the pith, estimates ring counts, and flags internal cracks.
* **Smart Visual Reporting:** Generates PDF reports plotting the traces and detected features. Includes dynamic file naming and the ability to process targeted subdirectories without overwriting master reports.
* **Messy Field Data Merging:** Uses spatial "BLOCK" anchors to read highly unstructured Excel field layouts and accurately join them with machine-generated scan data.
* **Benchmarking & Validation:** Built-in testing suites to compare automated algorithm outputs against historically manual-processed data (both 1:1 paired tests and macro-distribution tests).

## 📂 Repository Structure

### 1. `automated_resitograph_calculations.R` (Core Pipeline)
The engine of the suite. It scans a defined directory for raw `.rgp` files and applies the core mathematical processing algorithms. 
* **Phase 1:** Exports a multi-sheet Excel workbook containing detrended metrics for every scan, organized by subdirectory.
* **Phase 2:** Generates a multipage PDF visual report of the traces. Supports targeted batching (processing a single subfolder) with smart file-naming to prevent accidental overwrites.

### 2. `merge_with_messy_fielddata.R` (Metadata Joiner)
A robust data-wrangling script that maps the raw machine outputs back to the physical field design.
* Reads messy field assessment sheets by searching for spatial anchors (e.g., "BLOCK 1").
* Extracts Tree IDs, plot labels, and surveyor notes based on column offsets.
* Automatically audits the data, flagging missing trees, swept trees, dead trees, or typos before performing an inner join with the processed Resistograph Excel file.

### 3. `automated_benchmark_tests.R` (Validation Suite)
Ensures data integrity and algorithm accuracy by comparing automated outputs against manually processed legacy data.
* **Test 1 (Paired Validation):** Matches new automated outputs 1:1 with manual outputs for the exact same files, calculating bias and Mean Absolute Error (MAE) for metrics like DBH and resistance.
* **Test 2 (Macro Benchmarking):** Generates density plots comparing the overall distribution of new trials against historical baselines to check for drift or systemic errors.

## 📦 Prerequisites & Dependencies

This suite requires R and the following packages. You can install missing packages using `install.packages()`.

* **Data Wrangling:** `dplyr`, `tidyr`, `purrr`, `stringr`
* **File & Path Management:** `here`, `tools`
* **Excel Handling:** `readxl`, `openxlsx`, `writexl`, `tidyxl` (Crucial for the spatial cell-mapping in the merging script)
* **Visualization:** `ggplot2`, `gridExtra`

## 🛠️ Usage Workflow

### Step 1: Process Raw Data
Open `automated_resitograph_calculations.R`. Go to **SECTION 6: MAIN EXECUTION BLOCK**. Update your `experiment_name` and `root_folder` paths. Run the script to generate your raw metric Excel file and visual PDF report.
> **Note:** To process only a specific batch (e.g., "Block 16"), pass the folder name to the `target_subdir` argument in the execution block. 

### Step 2: Merge with Field Layout
Open `merge_with_messy_fielddata.R`. Update the paths in **SECTION 3: MAIN EXECUTION BLOCK** to point to your newly generated Excel file and your raw field map Excel sheet. Run the script to generate the `_Combined_Final.xlsx` output.

### Step 3: Validate (Optional)
If you are tweaking the detection algorithms or testing against older datasets, open `automated_benchmark_tests.R`. Update the paths to point to your historical manual `.xlsx` files and run the benchmarking visualizations.

## ⚠️ Important Formatting Notes for Field Data
The merging script (`merge_with_messy_fielddata.R`) relies on specific formatting in the field layout Excel sheets:
1.  There must be text anchors starting with **"BLOCK"** (e.g., "BLOCK 1", "Block 12") to define the spatial bounding boxes.
2.  The folder/subdirectory name corresponding to that block should be in the cell immediately to the right of the "BLOCK" cell.
3.  Column headers for tree identification must contain **"Plot"** or **"Tree"**.
