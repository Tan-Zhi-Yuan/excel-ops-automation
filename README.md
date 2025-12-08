# 🚀 VBA Business Operations Toolkit

## Overview
A library of high-performance VBA modules designed to automate complex supply chain and reporting tasks in Excel. This toolkit focuses on data sanitization, ETL (Extract, Transform, Load) processes, and business logic enforcement.

---

## 📂 Script Library

<details>
<summary><b>1. Unstructured Data Sanitizer (Click to Expand)</b></summary>
<br>

> **File:** [TableAutoAutomation.bas](./TableAutoAutomation.bas)
>
> **The Problem:**
> Raw reports from legacy systems often dump data without fixed ranges or headers, making analysis impossible.
>
> **The Solution:**
> * **Dynamic Detection:** Scans for keywords (e.g., "Item Code") to define table boundaries dynamically.
> * **Context-Aware Filtering:** Reads the *titles* above the data (e.g., "Sales < Forecast") to determine how to filter the table *below*.
> * **Sanitization:** Renames inconsistent headers (e.g., "%" → "Percentage") for standard reporting.

</details>

<details>
<summary><b>2. Logistics Sync Engine (Click to Expand)</b></summary>
<br>

> **File:** [LogisticsSyncEngine.bas](./LogisticsSyncEngine.bas)
>
> **The Problem:**
> Syncing a daily "Running List" with a master "Total List" (10k+ rows) using nested loops caused Excel to freeze. Business rules regarding "Line 0" summaries were often missed.
>
> **The Solution:**
> * **O(N) Performance:** Implemented `Scripting.Dictionary` (Hash Maps) to replace nested loops, reducing runtime by 90%+.
> * **Business Logic:** Enforces a hierarchical rule where "Line 0" entries are protected from manual overwrites.
> * **Data Integrity:** Automatically zeroes out financial columns for subsidiary lines to prevent double-counting.

</details>

<details>
<summary><b>3. Client-Facing Report Extractor (Click to Expand)</b></summary>
<br>

> **File:** [Export_Regional_Reports.bas](./Export_Regional_Reports.bas)
>
> **The Problem:**
> Sharing internal workbooks with external stakeholders is risky. Formulas linked to backend databases break when the file is moved, and "cleaning" a file manually (Paste Special > Values) is tedious and error-prone.
>
> **The Solution:**
> * **Hybrid Copy Logic:** Intelligently copies "Dashboard" sheets *with* full formatting/charts, while converting "Raw Data" sheets to *static values only*.
> * **Security:** Ensures no backend formulas or external file links exist in the output file.
> * **Maintainability:** Uses centralized `Const` variables to map source/target sheet names, making it easy to adapt to different reporting periods.

</details>

<details>
<summary><b>4. Dynamic Sheet Importer (Click to Expand)</b></summary>
<br>

> **File:** [SheetImporter.bas](./SheetImporter.bas)
>
> **The Problem:**
> Monthly reporting required opening a massive system export file and manually copying 5 specific sheets (e.g., "OTH", "SMT") into the main dashboard. This was repetitive and prone to "copy-paste" errors.
>
> **The Solution:**
> * **User-Configurable:** Uses Excel "Named Ranges" to let non-technical users define the File Path and Sheet List directly in the grid, without touching VBA.
> * **Smart Refresh:** Automatically deletes the old versions of the sheets before importing the new ones to ensure data is never duplicated.
> * **Resilient:** Includes checks to ensure the source file exists and handles missing sheets gracefully.

</details>

<details>
<summary><b>5. Automated Monthly Reporting Suite (Click to Expand)</b></summary>
<br>

> **File:** [Historical_Report_and_Consolidation_Tools.bas](./Historical_Report_and_Consolidation_Tools.bas)
>
> **The Problem:**
> Generating a 6-month accuracy trend report required two manual bottlenecks: 1) Hunting through network folders to find specific monthly files, and 2) Manually copying/pasting rows matching a specific "Item List" while aligning inconsistent columns.
>
> **The Solution:**
> * **Recursive File Retrieval:** Uses `DateAdd` logic to calculate the last 6 months of directory paths (e.g., `2025-10`, `2025-09`) and fetch the correct reports automatically.
> * **High-Speed Filtering:** Loads the "Master Item List" into a `Scripting.Dictionary` for O(1) lookup speed during the consolidation loop.
> * **In-Memory ETL:** Performs column remapping (dropping "Supplier", adding "Sales Month") and data cleaning ("NA" -> 0) within arrays before writing to the sheet.

</details>

---
## ⚙️ Engineering Philosophy & AI-Augmented Workflow

This repository demonstrates a **Modern "Hybrid" Development Strategy**.
In an era where syntax is cheap but logic is expensive, my focus is on **Architectural Design** and **Business Value**. I utilize Large Language Models (LLMs) as a "force multiplier" to accelerate development across my tech stack.

### 🧠 The Division of Labor
* **Human Architect (Me):**
    * **Business Logic:** Defining the complex rules (e.g., FIFO Valuation, Line 0 Protection, File Routing Matrix).
    * **System Architecture:** Designing how disparate systems (Excel, Outlook, File System) interact securely.
    * **Validation:** Reviewing, debugging, and stress-testing all code for accuracy and edge cases.
* **AI Assistant (Tooling):**
    * **Syntax Generation:** Rapidly translating logic into specific syntax for **VBA**, **M Code (Power Query)**, and **DAX**.
    * **Pattern Optimization:** Refactoring nested loops into $O(N)$ Hash Maps (`Scripting.Dictionary`) for performance.
    * **Boilerplate:** Generating standard error-handling blocks and UI elements.

### 🛠️ Tech Stack & Methodology
This workflow allows me to maintain high standards of code quality across multiple domains:
* **VBA:** used for Event-Driven Automation and Object Model manipulation.
* **Power Automate (RPA):** used for OS-level orchestration and "Low-Code" integration.
* **Power Query (M Code):** used for robust ETL data transformation and cleaning steps.

---

## 💡 Technical Highlights
* **Performance:** Shifted from $O(n^2)$ nested loops to $O(n)$ Hash Maps for large datasets.
* **Modularity:** All scripts use helper functions to ensure the Single Responsibility Principle.
* **Defensive Coding:** Extensive use of `On Error GoTo` handlers and object validation to prevent runtime crashes.
