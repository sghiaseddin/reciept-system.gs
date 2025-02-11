# Google Forms to Google Docs Automation Script

## Overview
This script automates the process of injecting data from Google Sheets into a Google Docs template when a Google Form is submitted. It retrieves environmental parameters from a "dev" sheet and generates a formatted document based on the form submission.

## Features
- Extracts submitted form data from a Google Sheet.
- Retrieves additional parameters from a "dev" sheet (e.g., receipt number, payment methods).
- Formats currency and dates for consistency.
- Populates a predefined Google Docs template with the extracted data.
- Saves the generated document in a specified Google Drive folder.
- Automatically increments receipt numbers for tracking.

## Dependencies
This script runs within Google Apps Script and requires:
- Google Sheets (to store form responses and configurations)
- Google Docs (template for document generation)
- Google Drive (for storing generated documents)

## Configuration
Modify the following constants in the script to match your environment:
```javascript
const TEMPLATE_FILE_ID = 'YOUR_TEMPLATE_FILE_ID'; // Google Docs template ID
const DESTINATION_FOLDER_ID = 'YOUR_DESTINATION_FOLDER_ID'; // Target Google Drive folder
const CURRENCY_SIGN = 'CAD '; // Default currency format
```

## Functions
### `toCurrency(num)`
Converts a numeric value to a formatted currency string.

### `toDateFmt(dt_string)`
Formats a date string to `YYYY-MM-DD`.

### `parseFormData(values, header)`
Extracts and processes form submission data, calculates totals, and retrieves additional parameters.

### `populateTemplate(document, response_data)`
Replaces placeholders in the Google Docs template with extracted data.

### `createDocFromForm()`
Main function that:
1. Retrieves form submission data.
2. Parses and formats the data.
3. Copies the Google Docs template.
4. Populates the template with data.
5. Saves the final document to Google Drive.

## Usage
1. Deploy the script within Google Apps Script.
2. Ensure the script is triggered on form submission.
3. Set up the Google Sheet and Google Docs template with the required structure.
4. Update the necessary configurations (template ID, folder ID, currency format).
5. Submit a form and check Google Drive for the generated document.

## Acknowledgments
This script was originally created by [automagictv](https://gist.github.com/automagictv/48bc3dd1bc785601422e80b2de98359e) and further modified by [Shayan Ghiaseddin](https://sghiaseddin.com).

