# Supply Chain Inventory & Automated Procurement Engine

An interactive, end-to-end Excel analytics tool designed to streamline supply chain visibility and eliminate the manual friction of reordering stock. This project addresses a common business bottleneck: the lag time between identifying inventory deficits and alerting the purchasing team. 

By utilizing native Excel features (Power Query, Data Modeling, and VBA), this solution runs entirely locally without requiring external database plugins or administrative IT overrides.

---

## 📊 Dashboard Preview
![Dashboard Layout Preview](dashboard_preview.jpg)
*(Note: Replace this image file in your repository with your actual dashboard screenshot to display it here).*

---

## 🛠️ System Architecture & Mechanics

### 1. Extraction & Data Transformation (Power Query)
* **ETL Pipeline:** Connects directly to raw, multi-source logistics spreadsheets, bypassing manual data compilation.
* **Schema Standardization:** Removes structural inconsistencies, reformats changing regional metrics, and removes duplicate operational rows.
* **Data Modeling:** Produces optimized, indexed flat tables ready for pivot analysis, keeping the overall workbook file size lightweight and highly responsive.

### 2. User Interface & Dynamic Controls
* **Cross-Filtering:** Integrates synchronized visual slicers across parallel pivot tables, allowing users to drill down by region, status, and timeline simultaneously without breaking the master layout.
* **One-Click State Reset:** Features a dedicated control macro that instantly clears all active UI filters, returning the entire screen to the default global view smoothly.

### 3. Automated Reorder Engine (VBA)
* **Exception Scanning Loop:** A targeted VBA loop dynamically text-scans active rows to instantly evaluate current stock tallies against defined target safety thresholds.
* **Data Aggregation:** Collects all flagged deficit items, SKUs, and replenishment numbers, compiling them into a clean, human-readable text string.
* **MAPI System Integration:** Automatically opens a pre-formatted Outlook email draft populated with the target purchasing team's address, a clear subject line, and the itemized deficit summary—removing copy-paste workflows entirely.

---

## 📋 Operational Workflow

1. **Refresh Data:** Open the workbook and click "Refresh All" to pull the latest logistics records through the Power Query pipeline.
2. **Analyze Trends:** Use the synchronized dashboard slicers to evaluate regional performance metrics and high-level stockout alerts.
3. **Execute Alert:** Click the **"Generate Alerts"** interface button. The background script instantly audits the rows, isolates the shortages, and generates your structured email draft instantly.

---

## ⚙️ Repository Setup Instructions

To interact with the dashboard and run the automated workflows locally:

1. Download the `Nexus_Inventory_Command_Center.xlsm` file from this repository.
2. Open the file using desktop Microsoft Excel.
3. **Important:** Click **"Enable Macros"** at the top yellow security bar when prompted, as the email automation and filter-reset utilities depend entirely on the embedded VBA modules.

---
*Developed as a practical portfolio project to demonstrate clean data structure design, functional UI implementation, and localized business process automation.*
