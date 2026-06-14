# Logistics Email Automation (Excel VBA to Outlook)

## Overview
This VBA macro provides a robust, zero-touch solution for generating daily logistics dispatch notifications (Avisos). It bridges Microsoft Excel and Microsoft Outlook, enabling the automation of partner-specific email generation.

**Author:** Nándor Magyar  
**Business Impact:** Successfully deployed in a corporate environment, saving logistics coordinators approximately **1 hour of manual administrative work per day**.

## Key Features
- **Dynamic Environment Parsing:** Uses `Environ("USERNAME")` to automatically map local network and cloud sync (OneDrive) paths. This allows the script to be distributed and run on any team member's computer without modifying the source code.
- **Smart Template Fallback Logic:** Scans multiple predefined directories to locate specific Outlook template files (`.oft`). If a template is not found in the primary folder, it cascades down to alternative department folders.
- **Data to HTML Parsing:** Extracts specific dynamic data ranges from Excel and converts them on-the-fly into clean, CSS-styled HTML tables, preserving crucial formatting like bold fonts and cell background colors.
- **Automated String Replacement:** Injects the generated HTML tables directly into the Outlook `.oft` body by targeting placeholder tags (e.g., `{Tábla1}`).
- **Execution Summary:** Gracefully handles missing templates or corrupted files, providing the user with a clean, end-of-run `MsgBox` report detailing successful generations and error logs.
