function sendCertificates() {
  const CONFIG = {
    folderId: "your-folder-id-here", // Replace with your folder ID
    sheetName: "Your-Sheet-Name", // Replace with your sheet name
    emailSubject: "Thank you for your contribution", // Customize email subject
    senderName: "Your Organization Name", // Customize sender name
    columns: { // Map your sheet's columns
      fullName: "Full Name",
      category: "Category",
      affiliation: "Affiliation",
      email: "Email",
    },
    statusColumn: "Status", // Status column name in the sheet
    emailTemplate: {
      plainText: createEmailBody,
      htmlText: createHtmlBody,
    },
  };

  const folder = getFolderById(CONFIG.folderId);
  const certificatesMap = loadCertificates(folder);

  const sheet = getSheetByName(CONFIG.sheetName);
  const { data, header } = getSheetData(sheet);

  const indices = getColumnIndices(header, Object.values(CONFIG.columns));
  const statusIndex = ensureStatusColumn(sheet, header, CONFIG.statusColumn);

  Logger.log(`Processing ${data.length - 1} entries.`);
  data.slice(1).forEach((row, index) => {
    processRecipient(row, indices, certificatesMap, CONFIG, sheet, statusIndex, index + 2);
  });
}

// Helper Functions

function getFolderById(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    Logger.log(`Folder accessed: ${folder.getName()}`);
    return folder;
  } catch (error) {
    Logger.log(`Error accessing folder: ${error}`);
    throw new Error("Folder not found or inaccessible. Please check the folder ID.");
  }
}

function loadCertificates(folder) {
  const certificatesMap = new Map();
  const files = folder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName().split(".")[0].toLowerCase().trim();
    certificatesMap.set(fileName, file);
  }
  Logger.log(`Loaded certificates for: ${Array.from(certificatesMap.keys()).join(", ")}`);
  return certificatesMap;
}

function getSheetByName(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet named '${sheetName}' not found.`);
  Logger.log(`Sheet accessed: ${sheet.getName()}`);
  return sheet;
}

function getSheetData(sheet) {
  const data = sheet.getDataRange().getValues();
  const header = data[0];
  Logger.log(`Header columns: ${header.join(", ")}`);
  return { data, header };
}

function getColumnIndices(header, requiredColumns) {
  const indices = {};
  requiredColumns.forEach(column => {
    const index = header.indexOf(column);
    if (index === -1) throw new Error(`Sheet must have a '${column}' column.`);
    indices[column] = index;
  });
  return indices;
}

function ensureStatusColumn(sheet, header, statusColumnName) {
  let statusIndex = header.indexOf(statusColumnName);
  if (statusIndex === -1) {
    statusIndex = header.length;
    sheet.getRange(1, statusIndex + 1).setValue(statusColumnName);
  }
  return statusIndex;
}

function processRecipient(row, indices, certificatesMap, config, sheet, statusIndex, rowIndex) {
  const fullName = String(row[indices[config.columns.fullName]]).trim();
  const category = String(row[indices[config.columns.category]]).trim();
  const affiliation = String(row[indices[config.columns.affiliation]]).trim();
  const email = String(row[indices[config.columns.email]]).trim();

  if (!fullName || !category || !affiliation || !email) {
    Logger.log(`Skipping row ${rowIndex}: Missing required data.`);
    updateStatus(sheet, statusIndex, rowIndex, "Skipped: Missing data");
    return;
  }

  const certificateKey = fullName.toLowerCase();
  const file = certificatesMap.get(certificateKey);
  if (!file) {
    Logger.log(`Certificate not found for: ${fullName}`);
    updateStatus(sheet, statusIndex, rowIndex, "Certificate not found");
    return;
  }

  const emailBody = config.emailTemplate.plainText(fullName, category, affiliation);
  const htmlBody = config.emailTemplate.htmlText(fullName, category, affiliation);

  try {
    GmailApp.sendEmail(email, config.emailSubject, emailBody, {
      attachments: [file.getAs(MimeType.PDF)],
      name: config.senderName,
      htmlBody: htmlBody,
    });
    Logger.log(`Email sent to ${email} with certificate for ${fullName}`);
    updateStatus(sheet, statusIndex, rowIndex, `Sent on ${new Date()}`);
  } catch (error) {
    Logger.log(`Failed to send email to ${email}: ${error}`);
    updateStatus(sheet, statusIndex, rowIndex, `Failed: ${error.message}`);
  }
}

function createEmailBody(fullName, category, affiliation) {
  return `
Dear ${fullName},

Thank you for your contribution as ${category}. Your efforts, representing ${affiliation}, have been instrumental to our success.

Best regards,
Your Organization Name
  `;
}

function createHtmlBody(fullName, category, affiliation) {
  return `
<p>Dear <strong>${fullName}</strong>,</p>

<p>Thank you for your contribution as <strong>${category}</strong>. Your efforts, representing <strong>${affiliation}</strong>, have been instrumental to our success.</p>

<p>Best regards,<br/>
Your Organization Name</p>
  `;
}

function updateStatus(sheet, statusIndex, rowIndex, status) {
  sheet.getRange(rowIndex, statusIndex + 1).setValue(status);
}
