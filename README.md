# GoogleAppScripts

## Personlised Certificate Sender Script

This repository provides a flexible Google Apps Script that automates the process of sending certificates to participants listed in a Google Sheet. The certificates are retrieved from a specified Google Drive folder and emailed with personalized messages. The script is fully configurable, making it adaptable to various use cases.

## Features

- **Sheet Integration**: Fetches participant data from a Google Sheet with configurable column mappings.
- **Certificate Matching**: Matches participants to certificates stored in a Google Drive folder.
- **Customizable Emails**: Supports personalized email templates in both plain text and HTML formats.
- **Status Tracking**: Logs the status of each email sent in the Google Sheet for easy tracking.
- **Error Handling**: Skips rows with missing data and logs errors for unmatched certificates or failed email sends.

---

## Setup Instructions

### 1. **Clone or Download**
Clone this repository or download the source code.

---

### 2. **Prepare Google Sheet**
Create a Google Sheet with the following columns (customizable in the script):
- `Full Name`
- `Category`
- `Affiliation`
- `Email`
- A `Status` column (optional, will be created automatically if missing).

---

### 3. **Organize Certificates**
- Upload all certificates to a folder in Google Drive.
- Name each certificate file to match the participant’s full name (case-insensitive).

---

### 4. **Configure the Script**
1. Open your Google Spreadsheet.
2. Go to `Extensions > Apps Script`.
3. Copy the script (`sendCertificates.gs`) into the Apps Script editor.
4. Update the `CONFIG` object with:
   - `folderId`: The ID of the Google Drive folder containing certificates.
   - `sheetName`: The name of the sheet with participant data.
   - `emailSubject` and `senderName`: For email customization.
   - `columns`: Map your sheet’s columns to expected fields.
   - `statusColumn`: Name of the column to log email statuses.

---

### 5. **Authorize the Script**
Run the script for the first time and authorize the required permissions, such as access to Google Drive, Gmail, and Sheets.

---

### 6. **Execute the Script**
- Run the `sendCertificates` function from the Apps Script editor.
- Monitor the logs for real-time updates.
- Check the "Status" column in the sheet for results.

---

## Configuration Example

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

---

## Email Templates

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

## Logs and Debugging

- Use the `Logger.log()` statements in the script to monitor progress and debug issues.
- Status updates will be logged in the "Status" column of the Google Sheet.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

## Contributing

Contributions are welcome! Feel free to submit a pull request or report an issue for improvements.

---

## Disclaimer

This script is provided "as-is" without warranty of any kind. Use it responsibly and ensure compliance with data privacy regulations.
