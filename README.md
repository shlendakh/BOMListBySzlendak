# Simple BOM List By Szlendak (Fusion 360 Script)

## Overview
This Fusion 360 script generates a simple Bill of Materials (BOM) table for the active design. It lists each component’s name, quantity, and dimensions (X, Y, Z). The script also allows you to export the BOM as a CSV file.

## Installation
1. Copy this folder into your Fusion 360 Scripts directory:
   - macOS: `~/Library/Application Support/Autodesk/Autodesk Fusion 360/API/Scripts/`
   - Windows: `%APPDATA%/Autodesk/Autodesk Fusion 360/API/Scripts/`
2. Restart Fusion 360 (or open the **Scripts and Add-Ins** dialog to refresh).

## Usage
1. Open a design in Fusion 360.
2. Go to **Tools → Scripts and Add-Ins**.
3. Select **BOMListBySzlendak** (this script) and click **Run**.
4. In the dialog:
   - Choose a **Thickness parameter** (optional).
   - Toggle **Export CSV** if you want a file.
   - Set the output path (or click **Browse...**).
5. Click **OK** to generate the table and (optionally) export the CSV.

## Notes
- Thickness logic:
  - If a parameter is selected, its value is used as **Z**.
  - Otherwise, **Z** is the smallest dimension of the component.
- Units are shown in the header and match the design’s default length units.

## License
Licensed under Creative Commons Attribution 4.0 International. See `LICENSE`.
