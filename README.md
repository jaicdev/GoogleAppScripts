# GoogleAppScripts

## Personalized Certificate Sender & Document Generator Scripts

This repository provides flexible Google Apps Scripts for automating document generation and email-based certificate distribution. These scripts simplify workflows by integrating Google Workspace apps like Sheets, Docs, Drive, and Gmail. 

---

## **Features**

### **1. Personalized Certificate Sender Script** (`sendCertificates.gs`)
- **Sheet Integration**: Fetches participant data from a Google Sheet with configurable column mappings.
- **Certificate Matching**: Matches participants to certificates stored in a Google Drive folder.
- **Customizable Emails**: Supports personalized email templates in both plain text and HTML formats.
- **Status Tracking**: Logs the status of each email sent in the Google Sheet for easy tracking.
- **Error Handling**: Skips rows with missing data and logs errors for unmatched certificates or failed email sends.

### **2. Document Generator Script** (`DocGeneration.gs`)
- **Dynamic Placeholders**: Replaces placeholders (e.g., `{{Name}}`, `{{Role}}`) in a Google Doc template with data from a Google Sheet.
- **Flexible Configurations**: Adaptable to various document types like certificates, contracts, or letters.
- **Automated File Creation**: Generates documents with unique names and stores them in a specified Google Drive folder.
- **Seamless Integration**: Fetches data directly from a Google Sheet for automated processing.

---

## **Setup Instructions**

### **1. Clone or Download**
Clone this repository or download the source code.

---

### **2. Prepare Google Sheets**

#### For Certificate Sender:
- Create a Google Sheet with the following columns (customizable in the script):
  - `Full Name`
  - `Category`
  - `Affiliation`
  - `Email`
  - `Status` (optional; created automatically if missing).

#### For Document Generator:
- Create a Google Sheet with column names matching placeholders in your template, such as:
  - `Name`
  - `Role`
  - `Date`

---

### **3. Create Templates**

#### Certificate Sender:
- Certificates should be stored in a Google Drive folder, with filenames matching participant names (case-insensitive).

#### Document Generator:
- Design a Google Doc template with placeholders in `{{Placeholder}}` format.
- Example:
  ```
  Certificate of Achievement

  This is to certify that {{Name}} has successfully completed the role of {{Role}} on {{Date}}.
  ```

---

### **4. Configure the Scripts**

#### For Certificate Sender (`sendCertificates.gs`):
1. Open your Google Spreadsheet.
2. Go to `Extensions > Apps Script`.
3. Copy the script into the Apps Script editor.
4. Update the `CONFIG` object with:
   - `folderId`: The Google Drive folder ID for certificates.
   - `sheetName`: The name of the sheet with participant data.
   - `emailSubject` and `senderName`: Customize email settings.
   - `columns`: Map your sheet’s columns to expected fields.
   - `statusColumn`: Column name for logging email statuses.

#### For Document Generator (`DocGeneration.gs`):
1. Copy the script into the Apps Script editor.
2. Update the `CONFIG` object with:
   - `templateId`: The Google Doc template ID.
   - `sheetName`: The name of the sheet with data.
   - `folderId`: The Google Drive folder ID for generated documents.
   - `placeholders`: List of placeholders in your template.

---

### **5. Authorize the Scripts**
Run the scripts for the first time and grant the required permissions, such as access to Google Drive, Gmail, and Sheets.

---

### **6. Execute the Scripts**

#### Certificate Sender:
- Run the `sendCertificates` function to start sending personalized emails with attached certificates.

#### Document Generator:
- Run the `generateDocumentsFromTemplate` function to generate personalized documents.

---

## **Examples**

### **Certificate Sender Configuration**
```javascript
const CONFIG = {
  folderId: "your-folder-id-here",
  sheetName: "Your-Sheet-Name",
  emailSubject: "Thank you for your contribution",
  senderName: "Your Organization Name",
  columns: {
    fullName: "Full Name",
    category: "Category",
    affiliation: "Affiliation",
    email: "Email",
  },
  statusColumn: "Status",
  emailTemplate: {
    plainText: createEmailBody,
    htmlText: createHtmlBody,
  },
};
```

### **Document Generator Configuration**
```javascript
const CONFIG = {
  templateId: "your-template-document-id-here",
  sheetName: "Data",
  folderId: "your-destination-folder-id-here",
  placeholders: ["{{Name}}", "{{Role}}", "{{Date}}"],
};
```

---

## **Email Templates**

### **Plain Text Template**
```javascript
function createEmailBody(fullName, category, affiliation) {
  return `
Dear ${fullName},

Thank you for your contribution as ${category}. Your efforts, representing ${affiliation}, have been instrumental to our success.

Best regards,
Your Organization Name
  `;
}
```

### **HTML Template**
```javascript
function createHtmlBody(fullName, category, affiliation) {
  return `
<p>Dear <strong>${fullName}</strong>,</p>

<p>Thank you for your contribution as <strong>${category}</strong>. Your efforts, representing <strong>${affiliation}</strong>, have been instrumental to our success.</p>

<p>Best regards,<br/>
Your Organization Name</p>
  `;
}
```

---

## **License**
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## **Contributing**

Contributions are welcome! Feel free to submit a pull request or report an issue to help improve these scripts.

- For guidelines, see [CONTRIBUTING.md](CONTRIBUTING.md).
- For bug reports, please open an issue on GitHub.

---

## **Disclaimer**

These scripts are provided "as-is" without warranty of any kind. Use responsibly and ensure compliance with data privacy regulations.

---

This combined and enhanced README ensures clarity and usability for both the `sendCertificates.gs` and `DocGeneration.gs` scripts, making it ready for open-source sharing on GitHub.
