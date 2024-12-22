/**
 * Document Automation Script
 * Version: 1.0.0
 * License: MIT
 * Author: Open Source Contributor
 */

function generateDocumentsFromTemplate() {
  const CONFIG = {
    templateId: "your-template-document-id-here", // Replace with your Google Doc template ID
    sheetName: "Data", // Replace with the name of your Google Sheet
    folderId: "your-destination-folder-id-here", // Replace with your Google Drive folder ID
    placeholders: ["{{Name}}", "{{Role}}", "{{Date}}"], // Define placeholders used in the template
  };

  try {
    // Load Google Sheet
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.sheetName);
    if (!sheet) throw new Error(`Sheet named '${CONFIG.sheetName}' not found.`);
    const data = sheet.getDataRange().getValues();
    const header = data[0];

    // Validate and map column indices
    const columnIndices = mapColumnIndices(header, CONFIG.placeholders.map(p => p.replace(/{{|}}/g, "")));

    // Load Google Doc template
    const templateDoc = DriveApp.getFileById(CONFIG.templateId);
    const destinationFolder = DriveApp.getFolderById(CONFIG.folderId);

    // Generate documents for each row in the sheet
    data.slice(1).forEach((row, index) => {
      const rowData = mapRowData(row, columnIndices);
      const docName = `${rowData.Name || "Document"}_${new Date().toISOString().slice(0, 10)}`;
      const newDoc = createDocument(templateDoc, rowData, CONFIG.placeholders, docName);
      destinationFolder.addFile(newDoc);
      Logger.log(`Generated document: ${docName}`);
    });

    Logger.log("All documents generated successfully.");
  } catch (error) {
    Logger.log(`Error: ${error.message}`);
    throw error;
  }
}

/**
 * Maps column names to their indices in the sheet header.
 * @param {Array} header - The header row of the sheet.
 * @param {Array} requiredColumns - The column names to map.
 * @returns {Object} - An object mapping column names to indices.
 */
function mapColumnIndices(header, requiredColumns) {
  const indices = {};
  requiredColumns.forEach(column => {
    const index = header.indexOf(column);
    if (index === -1) throw new Error(`Column '${column}' not found in the sheet.`);
    indices[column] = index;
  });
  return indices;
}

/**
 * Maps data from a row to an object using column indices.
 * @param {Array} row - The data row from the sheet.
 * @param {Object} columnIndices - An object mapping column names to indices.
 * @returns {Object} - A key-value pair of column names and row data.
 */
function mapRowData(row, columnIndices) {
  const rowData = {};
  for (const column in columnIndices) {
    rowData[column] = row[columnIndices[column]];
  }
  return rowData;
}

/**
 * Creates a document from a template by replacing placeholders with data.
 * @param {GoogleAppsScript.Drive.File} templateDoc - The template document.
 * @param {Object} rowData - The data for the current row.
 * @param {Array} placeholders - The placeholders in the template.
 * @param {String} newDocName - The name for the generated document.
 * @returns {GoogleAppsScript.Drive.File} - The generated document file.
 */
function createDocument(templateDoc, rowData, placeholders, newDocName) {
  const newDoc = templateDoc.makeCopy(newDocName);
  const doc = DocumentApp.openById(newDoc.getId());
  const body = doc.getBody();

  placeholders.forEach(placeholder => {
    const key = placeholder.replace(/{{|}}/g, "");
    const value = rowData[key] || "";
    body.replaceText(placeholder, value);
  });

  doc.saveAndClose();
  return newDoc;
}
