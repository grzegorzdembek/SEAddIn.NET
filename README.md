# Solid Edge Manufacturing Automation Add-In

## Overview
This is a custom Add-In for Siemens Solid Edge, developed in C# and the .NET Framework. It automates repetitive tasks related to manufacturing preparation, drawing exports, and Bill of Materials (BOM) generation. 

The goal of this tool is to speed up the export of production files and generate Excel reports directly from 3D CAD assemblies while reducing manual errors.

## Technical Details
This application was built with a strong focus on memory safety and execution speed when interacting with Solid Edge and Excel COM interfaces:

*   **COM Memory Management:** Solid Edge and Excel APIs are prone to leaving background processes running if objects are not handled correctly. This codebase wraps COM object calls in `try-finally` blocks and explicitly releases every instance using `Marshal.ReleaseComObject` to prevent memory leaks.
*   **Bulk Excel Operations:** To avoid the performance bottleneck of cell-by-cell manipulation, the Add-In reads and writes spreadsheet data in bulk using 2D object arrays (`object[,]`). 
*   **Export Optimization:** Before invoking the Solid Edge API to generate DXFs or PDFs, the Add-In checks the local directory and reads custom CAD properties (like DXF generation dates). It skips files that are already up to date, which significantly reduces batch export times.

## Features & Commands
The Add-In adds a custom Ribbon UI to Solid Edge (available in Assembly, Part, Sheet Metal, and Draft environments) with the following commands:

**Drafting & Formats**
*   **Save PDF & DXF:** Saves the active Draft document as both a PDF and a DXF in the source directory.
*   **Save Flat Pattern:** Extracts the sheet metal flat pattern and saves it as a DXF.
*   **Save STEP:** Exports the active part or sheet metal model to the STEP format.

**Assembly & BOM Automation**
*   **Export DXFs (Batch):** Traverses the active assembly tree, identifies sheet metal components, generates missing DXF flat patterns, and outputs a formatted Excel BOM.
*   **Export Parts List (Thumbnails):** Opens a background Draft environment to extract the BOM to Excel, generating and embedding thumbnail images of the 3D models into the spreadsheet.
*   **Set Count Property:** Recursively calculates and updates custom quantity properties across specific component types based on a user-defined multiplier.
*   **Clear DXF Date:** Clears generation dates from component metadata to force a clean re-export of specific files.

**Manufacturing Preparation**
*   **Add Drawings:** Reads a generated Excel BOM, locates the corresponding PDFs and DXFs in the project directory, and copies them into a dedicated production package folder.

## Tech Stack
*   **Language:** C#
*   **Framework:** .NET Framework
*   **APIs:** Solid Edge COM API, Microsoft Office Interop (Excel)

## References
*   **Solid Edge Community AddIn:** The foundation for registering COM Add-Ins is based on design patterns from the [SolidEdge.Community.AddIn](https://github.com/SolidEdgeCommunity/SolidEdge.Community.AddIn) repository.
*   **Solid Edge .NET Programmer's Guide:** Used as the primary reference for interacting with the Solid Edge API.