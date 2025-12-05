# Excel Automation: Dynamic Table Detection & Context-Aware Filtering with help of AI

## Overview
This VBA module automates the standardization of unstructured Excel reports. It scans multiple worksheets for data blocks, converts them into structured Excel Tables (`ListObjects`), and applies intelligent filtering logic based on the text found in the report headers.

## The Challenge
The raw data source produces reports where:
1.  Multiple tables exist on a single sheet without defined ranges.
2.  Column headers are inconsistent (e.g., using symbols like "%" instead of text).
3.  Filtering requirements change based on the section title (e.g., "Sales > Forecast" requires different logic than "Sales < Forecast").

## The Solution
The code is structured into modular helper functions to ensure maintainability and readability. 

**Key Features:**
* **Dynamic Range Detection:** Automatically identifies the start and end of data blocks based on content, not fixed cell addresses.
* **Header Sanitization:** Renames columns to standard naming conventions to prevent downstream errors.
* **Contextual Logic:** Scans the "human-readable" titles *above* the grid to determine how the data *inside* the grid should be filtered.
* **Robust Error Handling:** Uses `On Error GoTo` handlers and object checks to prevent runtime crashes on empty or malformed sheets.

## Technical Details
* **Language:** VBA (Visual Basic for Applications)
* **Concepts Used:** Modular programming, ListObjects, Range manipulation, String parsing.
