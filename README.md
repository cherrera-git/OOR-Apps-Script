# **OOR Apps Script Automation**

**Author:** Christianne Herrera, Senior Supervisor

**Organization:** Mountain Technologies Limited

**Last Updated:** August 2026

## **Overview**

This Google Apps Script project automates the inventory and order tracking workflow for Mountain Technologies Limited. It acts as a robust integration layer, synchronizing data between external Infor SyteLine ERP exports and internal Google Sheets trackers (OOR, Stock Items, and New Orders).

The system reduces manual data entry, ensures data integrity through strict safety gates, features automated FIFO-based shortage calculations, and prevents cell bloat with a self-healing text deduplication engine.

## **System Architecture & File Structure**

The project is highly modularized into distinct functional areas to separate core logic, UI, and data integrity engines:

* **Code.gs**: Main logic dispatchers for the Cleanup and Update workflows. Contains the global CONFIG object.  
* **Helpers.gs**: The data-processing engine. Handles caching headers, fetching sheet data, and executing FIFO (First-In, First-Out) material allocations for the shortage list.  
* **PCNotesEngine.gs**: A specialized self-healing text parser. It deduplicates changelogs and intelligently merges automated history with manual user notes to maintain QMS data integrity and prevent cell bloat.  
* **RealStartDate.gs**: Event-driven trigger module containing handleMaterialClearance(e), which monitors sheet edits and automatically stamps timestamps when parts are cleared.  
* **RetroactiveCleanup.gs**: A one-time macro to instantly execute the Self-Healing deduplication across all legacy PC Notes.  
* **UI.gs**: Encapsulates server-side UI methods, including lock services to prevent concurrent script executions and safety gates requiring exact phrase matches.  
* **dialog.html & confirmation.html**: Front-end user interfaces (modals) for secure CSV file uploads and operation confirmations.  
* **appsscript.json**: Project manifest defining time zones, OAuth scopes, and runtime versions.

## **Core Workflows**

The script adds a custom menu **"Update Tools"** to the Google Sheet with two primary workflows:

### **1\. Step 1: Import & Update (Synchronization)**

* **Trigger:** Update Tools \> Step 1: Import & Update  
* **Safety Phrases:** IMPORT (for file upload) and UPDATE (for processing).  
* **Actions:**  
  * **Ingestion:** Parses uploaded .csv, .txt, or .tsv SyteLine exports via a secure HTML modal into hidden system sheets.  
  * **Shortage Calculation:** Generates a "Shortage List" using FIFO logic to allocate Purchase Order (PO) supplies to Job Material demands.  
  * **Tracker Updates:** Syncs "MTL Due Date", fills blank "Project Coordinators", and intelligently updates "End Date Notes" across the OOR, Stock Items, and New Orders sheets.  
  * **Auditing:** Employs fuzzy logic to flag SyteLine jobs missing from tracking sheets to ensure no active orders slip through the cracks. Dispatches a summary email to the administrator.

### **2\. Step 2: Cleanup (Maintenance)**

* **Trigger:** Update Tools \> Step 2: Cleanup (Move & Archive)  
* **Safety Phrase:** CLEANUP  
* **Actions:**  
  * **Move:** Safely transfers active WIP (Work in Progress) valid jobs from "New Orders" to "OOR".  
  * **Archive:** Sweeps and moves jobs marked "Closed" or "Invalid" from tracking sheets to a temporary archive sheet Archive(temp).  
  * **Sort:** Applies a multi-level priority sort to the OOR sheet (Job Order \> PO \> MTL Due Date).

## **Key Technical Features**

* **Automated Shortage Notes (P-Dates):** Aggregates material shortages into the "End Date Notes" / "PC Notes" column. Formatted as P-MM/DD (Item A); P-MM/DD (Item B) and sorted chronologically.  
* **Self-Healing PC Notes Engine:** Parses cell history from the bottom up, prioritizing the newest status for specific part numbers or schedule metrics. It drops outdated legacy data before merging everything back seamlessly with manual user notes.  
* **Event-Driven Timestamping:** Automatically stamps the current date and time into the "Real Start Date" column when a job is marked as "Picked" or "No Picking" with zero shortages.  
* **Concurrency Control:** Employs Google's LockService to prevent multiple users from running structural updates simultaneously.

## **Required Import Files**

The system requires the following SyteLine exports. File names must match the target sheet names exactly:

1. **ToExcel\_JobOrders**: Job details, dates, and status.  
2. **ToExcel\_JobMaterialsListing**: Material demands and sub-assemblies.  
3. **ToExcel\_PurchaseOrderListing**: PO supply data.  
4. **ToExcel\_CustomerPart**: CSP status percentages.

## **Configuration & Setup**

Global settings are managed at the top of Code.gs under the CONFIG object. Key settings include:

* ADMIN\_EMAIL: Email address to receive automated run summaries and missing job alerts.  
* STOCK\_SHEET\_NAME: Target sheet name for stock items.  
* SAFETY\_ENABLED: Toggles the requirement for modal safety phrases.

### **Trigger Setup**

To enable the automated timestamping feature (RealStartDate.gs), a Google Apps Script installable trigger must be created:

1. Open the Apps Script Editor.  
2. Go to **Triggers** (clock icon on the left).  
3. Click **Add Trigger**.  
4. Set the function to run as handleMaterialClearance.  
5. Set the event source to From spreadsheet and the event type to On edit.