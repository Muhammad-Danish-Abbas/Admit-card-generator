# Admit Card Generator

A browser tool that makes exam admit cards from student Excel files and photos.
No server, no login. Everything stays in the browser.

## Features
- Upload one or many Excel files (.xlsx, .xls, .csv). 3000+ students at once works.
- Photos are matched by Student ID (file name like `2724981.png`).
- Class navigation for Class 11 and Class 12 (C1, EE6B, EM3, M1 ...).
- Morning and Evening shifts with their own entry time and signature.
- A5 landscape, one card per page.
- Download a PDF for one class, all classes in one PDF, or a ZIP with one PDF per class.
- Single card mode, browser print, and a missing-photo list.

## Excel columns
Required: `Student ID`, `Student Name`, `Class` (or `SIS /Fee Class` such as `11-C1-HSA`).
Optional: `Father Name`, `Shift`, `Class-1` (11 or 12).

## Run locally
Open `index.html` in a browser, or use the VS Code Live Server extension.

## Publish on GitHub Pages
Settings -> Pages -> Deploy from a branch -> `main` -> `/ (root)` -> Save.

## Privacy
Student data and photos are stored only in the visitor's own browser
(localStorage and IndexedDB). Do not commit Excel files or photos to this repository.

## Libraries (loaded from CDN)
SheetJS (Excel), jsPDF (PDF), JSZip (ZIP), Google Fonts.
