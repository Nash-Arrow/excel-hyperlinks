# Excel Hyperlinks Add-in

This repository contains an Excel Add-in that opens URLs from selected cells in an active sheet using OfficeJS. 

## Features
- Adds a custom "Open URLs" button to the Excel ribbon.
- Opens all URLs contained in the selected cell range in your default browser, each in a new tab.
- Prompts the user if the number of URLs exceeds 20, requiring confirmation to proceed.

## Setup Instructions

1. **Download the manifest file**: The add-in's manifest file is `manifest.xml`. You need to upload it to Excel to sideload the add-in.
2. **Upload the Add-in**:
   - Open Excel and go to **File** > **Options** > **Trust Center** > **Trust Center Settings**.
   - Enable **Sideloading of Office Add-ins**.
   - Go to **Insert** > **My Add-ins** > **Manage My Add-ins** > **Upload My Add-in**.
   - Select the `manifest.xml` file from this repository.
3. The add-in will be available in the **Home** ribbon under the "Open URLs" group.

## Folder Structure

```plaintext
excel-hyperlinks/
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   └── icon-80.png
├── manifest.xml
├── index.html
├── taskpane.css
└── taskpane.js
```

## Usage

1. Select a range of cells containing URLs in your Excel sheet.
2. Click the "Open URLs" button in the Home tab.
3. The URLs will open in your default browser, with a prompt for confirmation if more than 20 URLs are selected.

## Support

For issues or further assistance, please visit the [Support Page](https://github.com/Nash-Arrow/excel-hyperlinks).
