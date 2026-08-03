# Solid Edge Manufacturing Automation Add-In

## Overview
This is a custom tool made for Siemens Solid Edge[cite: 61]. Its main job is to automate boring and repetitive tasks related to preparing files for manufacturing, exporting drawings, and creating parts lists (BOMs). 

The goal is simple: help you export production files much faster, create Excel reports directly from your 3D models, and avoid manual mistakes.

## How It Works (The Benefits)
This application is designed to be fast, reliable, and invisible to the user:
*   **Computer Friendly:** It cleanly opens and closes Solid Edge and Excel in the background so it won't freeze your computer or leave hidden processes running.
*   **Super Fast Excel:** Instead of writing data cell by cell, it builds the entire report in memory and pastes it into Excel all at once, making it incredibly fast.
*   **Smart Exporting:** Before creating new DXF or PDF files, it checks if they already exist and are up to date. Skipping files you don't need to overwrite saves a massive amount of time during bulk exports.

## What Can It Do? (Features)
The tool adds a new menu inside Solid Edge with the following easy-to-use buttons:

**Drafting & Formats**
*   **Save PDF & DXF:** Instantly saves your current drawing as both a PDF and a DXF in your project folder.
*   **Save Flat Pattern:** Takes a sheet metal part and exports its flat shape directly to a DXF file.
*   **Save STEP:** Exports the 3D model you are currently looking at into a universal STEP file.

**Assembly & Parts Lists Automation**
*   **Export DXFs (Batch):** It looks through your entire assembly, finds all the sheet metal parts, creates DXF files for them, and puts everything into a neat Excel report.
*   **Export Parts List:** Creates a complete list of parts in Excel and even takes small pictures (thumbnails) of your 3D models to put inside the spreadsheet.
*   **Set Count Property:** Automatically counts how many times a part is used in the assembly and saves that exact number into the file's properties.
*   **Clear DXF Date:** Erases the old creation dates from files, forcing the system to generate brand-new DXFs next time.

**Manufacturing Preparation**
*   **Copy Drawings:** It reads your Excel parts list, searches your computer for the matching PDFs and DXFs, and copies them all into one dedicated package folder for the factory.
*   **Organize Drawings:** Automatically sorts and cleans up your project folders by putting drawings into their proper places.