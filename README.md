# Solid Edge Manufacturing Automation Add-In

## Overview
This is a custom Add-In for Siemens Solid Edge. It automates the preparation of manufacturing files, drawing exports, and Bill of Materials (BOM) generation. 

The main purpose of this tool is to quickly export production-ready files directly from 3D CAD assemblies and to generate Excel reports. This reduces the time spent on manual file conversion and prevents human errors during documentation.

## Core Mechanisms
The program is built to execute tasks quickly while keeping your computer's memory clean. It does this through three main rules:

*   **Process Management:** The application tightly controls Solid Edge and Excel background operations. It forces both programs to release their data immediately after a task is finished. This prevents hidden processes from running in the background and slowing down the computer.
*   **Data Transfer:** Instead of writing data into Excel one cell at a time, the program builds the entire spreadsheet in its memory and pastes it into Excel as a single, large block of data. This method drastically reduces the time required to create long reports.
*   **Export Optimization:** Before generating a new DXF or PDF file, the program reads the file's custom properties, such as the exact date of the last DXF export. If the file is already up to date with the 3D model, the program skips it. This avoids unnecessary file processing and speeds up batch exports.

## Features
The Add-In adds a custom menu to Solid Edge with the following specific commands:

**Drafting & Formats**
*   **Save PDF & DXF:** Takes the currently active Draft document and saves exact copies in both PDF and DXF formats directly to the original file's folder.
*   **Save Flat Pattern:** Unfolds the active sheet metal part and saves its flat 2D shape as a DXF file.
*   **Save STEP:** Converts the currently open 3D part or sheet metal model into a standard STEP file.

**Assembly & BOM Automation**
*   **Export DXFs (Batch):** Scans every level of the active assembly to find all sheet metal components. It generates a flat pattern DXF for any part that needs one, and then creates a formatted Excel BOM listing all the processed parts.
*   **Export Parts List:** Opens a temporary Draft file in the background to extract the assembly's parts list to Excel. It also generates image thumbnails of each 3D model and inserts them directly into the spreadsheet rows.
*   **Set Count Property:** Calculates the exact total quantity of specific components inside an assembly. It then writes this number directly into the custom properties of each CAD file, multiplying the result by a user-defined number if necessary.
*   **Clear DXF Date:** Deletes the specific property that records when a file was last exported. This forces the program to ignore the optimization rule and create a brand-new export file during the next batch process.

**Manufacturing Preparation**
*   **Copy Drawings:** Reads a previously generated Excel BOM file, searches the main project folder for the exact PDF and DXF files listed, and copies them into a separate folder created specifically for the production team.
*   **Organize Drawings:** Scans the main project directory for existing PDF and DXF files and sorts them into dedicated subfolders based on the name of the main assembly.