# Solid Edge Manufacturing Automation Add-In ⚙️

## 📖 Overview
This project is a high-performance, custom-built Add-In for Siemens Solid Edge, developed in C# and the .NET Framework. It deeply integrates with the Solid Edge API to automate tedious, time-consuming tasks related to manufacturing preparation, technical documentation, and Bill of Materials (BOM) generation. 

The primary goal of this software is to eliminate human error during the export of production files and to generate comprehensive, automated Excel reports directly from complex 3D CAD assemblies.

## 🚀 Core Architecture & Performance
This application is designed with a strict focus on system stability, execution speed, and memory safety when interacting with older COM interfaces:

*   **Smart Caching & State Synchronization:** To prevent redundant processing, the application utilizes a "Two-Pass" architecture. It maps the assembly tree directly into fast RAM and cross-references it with the physical hard drive state. It only invokes the heavy Solid Edge application layer to generate DXFs or PDFs if the files are physically missing or explicitly flagged for an update.
*   **Strict COM Memory Management:** Solid Edge and Excel APIs are notorious for generating memory leaks ("zombie processes") in .NET. This architecture strictly enforces a pattern of capturing and releasing every single COM instance via `Marshal.ReleaseComObject` inside isolated `try-finally` blocks to guarantee a zero-leak footprint.
*   **Bulk RAM Excel Interop:** To bypass the severe bottleneck of cell-by-cell Excel manipulation, the Add-In reads and writes spreadsheet data via bulk 2D object arrays (`object[,]`) entirely in RAM. 
*   **Fail-Fast Defensive Programming:** The codebase utilizes aggressive error catching, pre-validation of file paths, and safe skipping of uninitialized or unsaved CAD geometries to ensure the main application thread never crashes the host CAD environment.

## 🛠️ Features & Commands
The Add-In injects a custom Ribbon UI into Solid Edge environments (Assembly, Part, Sheet Metal, Draft). It provides the following automated toolset:

**Drafting & Formats**
*   **Save PDF & DXF:** Instantly saves the active Draft document simultaneously as both PDF and DXF in the source directory.
*   **Save Flat Pattern:** Extracts and saves the sheet metal flat pattern directly to DXF.
*   **Save STEP:** One-click export of the active part or sheet metal model to the STEP format.

**Assembly & BOM Automation**
*   **Export DXFs (Batch):** Recursively traverses the active assembly tree, identifies all sheet metal components, bypasses cached files, generates missing DXF flat patterns, and outputs a highly formatted Excel BOM summarizing the batch.
*   **Export Parts List (Thumbnails):** Automatically spins up a background Draft environment, extracts the BOM to Excel, and utilizes the Solid Edge `View` API to generate and embed real-time thumbnail screenshots of 3D models directly into the spreadsheet.
*   **Set Count Property:** A bulk-editing utility that recursively calculates and updates custom quantitative properties ("Ilość") across multiple specific component types (A, B, C) based on a user-defined multiplier.
*   **Clear DXF Date:** Safely purges cached generation dates from component metadata to force a clean re-export of specific assembly batches.

**Manufacturing Preparation**
*   **Add Drawings (Package Compilation):** Reads a previously generated Excel BOM, actively searches the project directory for corresponding PDFs and DXFs, and dynamically copies them into an organized "Production Package" folder for the manufacturing floor.

## 💻 Tech Stack
*   **Language:** C#
*   **Framework:** .NET Framework
*   **APIs:** Solid Edge COM API, Microsoft Office Interop (Excel)
*   **Architecture:** Object-Oriented Programming (Commands, Helpers, Wrappers, Ribbon Controllers)

## 📚 References
This project was developed utilizing the following materials and resources:
1.  **Solid Edge Community AddIn:** The technical foundation and solution structure for registering COM Add-Ins are based on the design patterns available in the [SolidEdge.Community.AddIn](https://github.com/SolidEdgeCommunity/SolidEdge.Community.AddIn) repository.
2.  **Solid Edge .NET Programmer's Guide:** API interactions were structured utilizing the official .NET Programmer's Guide for Solid Edge with Synchronous Technology.
