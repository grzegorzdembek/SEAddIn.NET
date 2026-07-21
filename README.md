# \# Solid Edge Manufacturing Automation Add-In ⚙️ \[IN PROGRESS]

# 

# \## 📖 Overview

# This project is a robust, custom-built Add-In for Siemens Solid Edge, developed in C# and the .NET Framework\[cite: 16]. It is designed to deeply integrate with the Solid Edge API to automate tedious, time-consuming tasks related to manufacturing preparation, technical documentation, and Bill of Materials (BOM) generation\[cite: 33]. 

# 

# The primary goal of this software is to eliminate human error during the export of production files and to generate comprehensive, highly optimized Excel reports directly from complex 3D CAD assemblies\[cite: 33].

# 

# \## 🚀 Technical Highlights (Why it's built this way)

# As a developer, my focus during this project was not just on functionality, but on \*\*performance, stability, and architectural best practices\*\*:

# 

# \*   \*\*COM Object Memory Management:\*\* Solid Edge is built on older COM technology, which is notorious for causing memory leaks and "zombie processes" in .NET. To combat this, the entire application strictly enforces a pattern of releasing COM objects via `Marshal.ReleaseComObject` inside `try-finally` blocks (implemented via custom `CoreUtils` generic methods)\[cite: 35].

# \*   \*\*Extreme Excel Optimization (RAM-based Processing):\*\* Instead of interacting with the Excel COM API cell-by-cell (which is extremely slow), the application reads and writes bulk data using 2D arrays (`object\[,]`) directly in RAM\[cite: 37, 40]. Excel runs completely in the background (`Visible = false`, `ScreenUpdating = false`) to maximize speed\[cite: 25, 26].

# \*   \*\*Reduced Disk I/O via LINQ:\*\* When searching for matching PDF or DXF drawings on the hard drive, the app reads the directory contents once into RAM and utilizes `LINQ` to filter and map file names. This drastically reduces the execution time compared to iterative disk querying\[cite: 31, 40].

# \*   \*\*Recursive Assembly Traversal:\*\* Implemented a robust `AssemblyTreeWalker` that recursively scans deep nested CAD assemblies, dynamically bypassing excluded components and gathering custom properties (Material, Thickness) directly from the parts\[cite: 41].

# 

# \## 🛠️ Features \& Commands

# The Add-In injects a custom Ribbon UI into Solid Edge environments (Assembly, Part, Sheet Metal, Draft) with the following automated tools\[cite: 33]:

# 

# 1\.  \*\*Export DXFs (Assemblies):\*\* Recursively traverses the active assembly, identifies sheet metal parts, dynamically generates DXF flat patterns, and names them based on thickness and material. It outputs a formatted Excel BOM summarizing the batch\[cite: 33, 41].

# 2\.  \*\*Export Parts List (BOM with Thumbnails):\*\* Automatically generates a Draft environment in the background, extracts the BOM, and exports it to Excel. It even generates real-time thumbnail screenshots (shots) of the 3D models using the Solid Edge `View` API and embeds them into the Excel spreadsheet\[cite: 26, 40].

# 3\.  \*\*Add Drawings (Automated Packaging):\*\* Reads a previously generated Excel BOM and automatically searches the project directory for corresponding technical drawings (PDFs/DXFs), copying them into a clean "Production Package" folder for the manufacturing floor\[cite: 31].

# 4\.  \*\*Set Count Property:\*\* A bulk-editing tool that recursively updates custom quantitative properties ("Ilość") across multiple components in an assembly based on a user-defined multiplier\[cite: 30].

# 5\.  \*\*Quick Save Formats:\*\* Single-click utilities for individual files to quickly save parts to STEP, flat patterns to DXF, or drawings to PDF + DXF simultaneously\[cite: 27, 28, 29].

# 

# \## 💻 Tech Stack

# \*   \*\*Language:\*\* C#

# \*   \*\*Framework:\*\* .NET Framework

# \*   \*\*APIs:\*\* Solid Edge COM API, Microsoft Office Interop (Excel)\[cite: 38]

# \*   \*\*Architecture:\*\* Object-Oriented Programming (Commands, Helpers, Wrappers, Dependency isolation)\[cite: 17, 18, 32]

# 

# \## 📚 References

# This project was developed based on the following materials and resources:

# 1\.  \*\*Solid Edge Community AddIn:\*\* The technical foundation and solution structure are based on the library and design patterns available in the \[SolidEdge.Community.AddIn](https://github.com/SolidEdgeCommunity/SolidEdge.Community.AddIn) repository\[cite: 15].

# 2\.  \*\*Solid Edge .NET Programmer's Guide:\*\* Interactions with the Solid Edge API were developed utilizing the knowledge and examples found in the official .NET Programmer's Guide for Solid Edge with Synchronous Technology API\[cite: 15].

