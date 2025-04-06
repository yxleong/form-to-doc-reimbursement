# Google Form Reimbursement Automation
This project automates the generation of reimbursement documents based on Google Form responses using Google Apps Script, Google Docs, and Google Drive. It’s designed to streamline the reimbursement process by dynamically generating documents, handling uploaded receipts, and reducing manual administrative work.

## Motivation
This project was implemented during my service as the **Minister of the Department of Finance** at the **National Taiwan University of Science and Technology (NTUST) Student Association**. Our goal was to simplify and digitize the reimbursement workflow, reduce human error, and improve efficiency across our finance operations.

## Features
- Collect and process Google Form responses automatically
- Generate formatted Google Docs for each submission
- Dynamically replace template placeholders (e.g., name, date, item details)
- Format dates and budget codes (e.g., `1-1-1-1`) to meet local standards
- Append all receipts as images on additional pages
- Insert and resize uploaded receipt images
- Count uploaded receipts and reflect in the document
- Send confirmation emails with the generated document

## Demo
![image](https://github.com/user-attachments/assets/95ac5998-f260-4e1f-838a-ffb68feb3eb5)

## Getting Started
To use this project, you'll need the following setup:
1. **A Google Form & Response Sheet**  
   Used to collect reimbursement submissions.
   - [Sample Google Form](https://github.com/yxleong/form-to-doc-reimbursement/blob/main/SampleGoogleForm.pdf)
   - The linked response **Google Sheet** will serve as the data source for the script.
2. **A Google Docs Template**  
   This is the template used to generate the reimbursement document.
   - Use `{{ }}` curly braces to define placeholders for dynamic replacement.
    ![image](https://github.com/user-attachments/assets/61e54cdd-5cb7-4786-8e10-914ff873d4b7)

## Usage
Follow these steps to set up the script:
1. **Open the Google Sheets linked to your form**  
   Go to `Extensions > Apps Script` to open the script editor.
2. **Paste the Code**  
   - Copy and paste the script from [`form2doc.gs`](form2doc.gs) into your Apps Script editor.
   - Be sure to **replace** the following placeholder variables in the script with your own IDs:
     ```js
     const templateId = 'REPLACE_DOC_ID';            // ID of your Google Docs template
     const folderId = 'REPLACE_NEW_DOC_FOLDER_ID';   // ID of the folder where generated docs will be stored
     ```
   - You can find these IDs in the URL of the corresponding Google Docs and Drive folder.
3. **Replace the Manifest File**  
   - In the Apps Script editor, go to `Project Settings` and enable **"Show 'appsscript.json' manifest file."**  
   - Then replace your existing `appsscript.json` with [`appsscript.json`](appsscript.json) for proper OAuth scopes and advanced service configuration.
4. **Set up triggers**  
   In the Apps Script editor:
   - Click on `Triggers` (the clock icon)
   - Add a new trigger:
     - Function to run: `onFormSubmit`
     - Event source: `From spreadsheet`
     - Event type: `On form submit`

## Example Placeholders in Template
| Placeholder     | Description                           |
|----------------|---------------------------------------|
| `{{承辦人}}`     | Person submitting the form            |
| `{{日期}}`       | Automatically formatted date (`YYYY/MM/DD`) |
| `{{K}} {{X}} {{M}} {{J}}` | Extracted from budget code like `1-1-1-1` |
| `{{HT}} ... {{Z}}` | Amount in traditional Chinese numerals |
| `{{QTY}}`        | Total number of receipt images uploaded |
| `{{image}}`      | Placeholder for appending images     |
| ...      | ...     |

## Contact
For questions, support, or feedback, you can reach us at [yx0829leong@gmail.com](mailto:yx0829leong@gmail.com). Feel free to [open an issue](../../issues) as well.
