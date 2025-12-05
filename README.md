# 🚀 VBA Business Operations Toolkit

## Overview
A library of high-performance VBA modules designed to automate complex supply chain and reporting tasks in Excel. This toolkit focuses on data sanitization, ETL (Extract, Transform, Load) processes, and business logic enforcement.

---

## 📂 Script Library

<details>
<summary><b>1. Unstructured Data Sanitizer (Click to Expand)</b></summary>
<br>

> **File:** `TableAutoAutomation.bas`
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

> **File:** `LogisticsSyncEngine.bas`
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
<summary><b>3. (Placeholder for Next Script)</b></summary>
<br>

> **File:** `ComingSoon.bas`
>
> **The Problem:**
> Description of the problem goes here.
>
> **The Solution:**
> * Feature 1
> * Feature 2

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
