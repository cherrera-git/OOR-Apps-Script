OOR Apps Script Automation
Author: Christianne Herrera, Senior Supervisor
Organization: Mountain Technologies Limited
Last Updated: February 2026
Overview
This Google Apps Script project automates the inventory and order tracking workflow for Mountain Technologies. It synchronizes data between external SyteLine system exports and internal Google Sheets trackers (OOR, Stock Items, and New Orders).
The system ensures data integrity through safety gates, detailed change logging, and automated shortage calculations.
Core Workflows
The script adds a custom menu "Update Tools" to the Google Sheet with two primary workflows:
1. Cleanup (Maintenance)
Trigger: Step 1: Cleanup (Move & Archive)
Safety Phrase: CLEANUP
Actions:
Moves valid jobs from "New Orders" to "OOR".
Moves jobs marked "Closed" or "Invalid" from "OOR" and "STOCK ITEMS" to "Archive(temp)".
2. Import & Update (Synchronization)
Trigger: Step 2: Import & Update
Safety Phrases: IMPORT (for file upload) and UPDATE (for processing).
Actions:
Ingestion: Parses uploaded .csv or .txt exports into hidden system sheets.
Shortage Calculation: Generates a "Shortage List" using FIFO logic to allocate Purchase Order (PO) supplies to Job Material demands.
Tracker Updates: syncs "MTL Due Date", fills blank "Project Coordinators", and updates "End Date Notes".
Auditing: Flags SyteLine jobs that are missing from tracking sheets.
Required Import Files
The system requires the following SyteLine exports. File names must match the target sheet names (extensions .csv, .txt, .tsv supported):
ToExcel_JobOrders: Job details, dates, and status.
ToExcel_JobMaterialsListing: Material demands and sub-assemblies.
ToExcel_PurchaseOrderListing: PO supply data.
ToExcel_CustomerPart: CSP status percentages.
Technical Logic
Shortage Notes (P-Dates)
The system aggregates material shortages into the "End Date Notes" column.
Logic: If a job has multiple shortages, the script concatenates all shortage dates.
Format: P-MM/DD (Item A); P-MM/DD (Item B)
Sorting: Dates are listed chronologically (earliest to latest).
Safety Gates
To prevent accidental data corruption, critical actions require user confirmation via modal dialogs. Users must type specific safety phrases (CLEANUP, IMPORT, UPDATE) to unlock the execution buttons.
Logging
All changes are recorded in an external Update Change Log spreadsheet defined in CONFIG.LOG_SHEET_ID. Logs include:
Run ID (Timestamped unique identifier)
Previous vs. New values for Dates and Notes
SyteLine Audit warnings
Configuration
Global settings are managed in Code.gs under the CONFIG object:
LOG_SHEET_ID: ID of the external audit log sheet.
STOCK_SHEET_NAME: Target sheet for stock items.
SAFETY_ENABLED: Toggles the requirement for safety phrases.
Generated for internal use at Mountain Technologies Limited.
