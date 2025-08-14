/**
 * Bar Employee CRM - Google Sheets Add-on
 * This creates a modal popup CRM directly in Google Sheets
 * 
 * Setup Instructions:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Replace Code.gs content with this script
 * 4. Save and run the onOpen function
 * 5. Refresh your Google Sheet
 * 6. You'll see "Employee CRM" in the menu bar
 */

// Configuration
const SHEET_NAME = 'Sheet2';
const HEADER_ROW = 1;
const DATA_START_ROW = 2;

// Headers in order
const HEADERS = ['Emp Id', 'First Name', 'Last Name', 'Phone', 'Email', 'Position', 'Note', 'Photo ID'];

/**
 * Runs when the spreadsheet is opened - adds the CRM menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Employee CRM')
    .addItem('Open CRM Manager', 'openCRMDialog')
    .addSeparator()
    .addItem('Initialize Sheet2', 'initializeSheet')
    .addToUi();
}

/**
 * Opens the CRM modal dialog
 */
function openCRMDialog() {
  const html = HtmlService.createTemplateFromFile('CRMDialog');
  const htmlOutput = html.evaluate()
    .setWidth(1200)
    .setHeight(700)
    .setTitle('Bar Employee CRM Manager');
  
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Bar Employee CRM Manager');
}

/**
 * Include external files (for CSS/JS in HTML template)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Ensure the header row exists and matches HEADERS. Upgrades older sheets safely.
 */
function ensureHeaders(sheet) {
  const targetSheet = sheet || getSheet();
  const headerRange = targetSheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
  const existing = targetSheet.getRange(HEADER_ROW, 1, 1, Math.max(HEADERS.length, targetSheet.getLastColumn())).getValues()[0];

  let needsUpdate = false;
  // If any required header mismatches, or headers are fewer than expected, update
  for (let i = 0; i < HEADERS.length; i++) {
    if ((existing[i] || '') !== HEADERS[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    targetSheet.getRange(HEADER_ROW, 1, 1, HEADERS.length).setValues([HEADERS]);
    const hr = targetSheet.getRange(HEADER_ROW, 1, 1, HEADERS.length);
    hr.setBackground('#f8f9fa');
    hr.setFontWeight('bold');
    hr.setBorder(true, true, true, true, true, true);
  }

  return targetSheet;
}

/**
 * Initialize the sheet with headers if they don't exist
 */
function initializeSheet() {
  const sheet = getSheet();
  
  // Always ensure headers are up-to-date (idempotent)
  ensureHeaders(sheet);
  
  SpreadsheetApp.getUi().alert('Success', 'Sheet2 has been initialized/updated with employee headers!', SpreadsheetApp.getUi().ButtonSet.OK);
  Logger.log('Sheet initialized/updated with headers');
  
  return sheet;
}

/**
 * Get or create the target sheet
 */
function getSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    Logger.log(`Created new sheet: ${SHEET_NAME}`);
  }
  
  return sheet;
}

/**
 * Get all employees from the sheet
 */
function getAllEmployees() {
  try {
    const sheet = ensureHeaders(getSheet());
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: true, employees: [] };
    }
    
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, HEADERS.length);
    const values = dataRange.getValues();
    
    const employees = values
      .filter(row => row[0] !== '') // Filter out empty rows (Emp Id is required)
      .map(row => ({
        empId: row[0] || '',
        firstName: row[1] || '',
        lastName: row[2] || '',
        phone: row[3] || '',
        email: row[4] || '',
        position: row[5] || '',
        note: row[6] || '',
        photoId: row[7] || ''
      }));
    
    Logger.log(`Retrieved ${employees.length} employees`);
    return { success: true, employees: employees };
    
  } catch (error) {
    Logger.log('Error in getAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save all employees to the sheet (replaces existing data)
 */
function saveAllEmployees(employees) {
  try {
    const sheet = ensureHeaders(getSheet());
    
    // Clear existing data (keep headers)
    const lastRow = sheet.getLastRow();
    if (lastRow >= DATA_START_ROW) {
      sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW + 1, HEADERS.length).clear();
    }
    
    if (employees && employees.length > 0) {
      // Prepare data rows
      const dataRows = employees.map(emp => [
        emp.empId || '',
        emp.firstName || '',
        emp.lastName || '',
        emp.phone || '',
        emp.email || '',
        emp.position || '',
        emp.note || '',
        emp.photoId || ''
      ]);
      
      // Write data to sheet
      const range = sheet.getRange(DATA_START_ROW, 1, dataRows.length, HEADERS.length);
      range.setValues(dataRows);
      
      Logger.log(`Saved ${employees.length} employees`);
    }
    
    return { success: true, message: `Saved ${employees ? employees.length : 0} employees` };
    
  } catch (error) {
    Logger.log('Error in saveAllEmployees: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Add a single employee
 */
function addEmployee(employee) {
  try {
    const sheet = ensureHeaders(getSheet());
    
    // Check for duplicate Emp ID
    const existingData = getAllEmployees();
    if (existingData.success) {
      const duplicate = existingData.employees.find(emp => emp.empId === employee.empId);
      if (duplicate) {
        return { success: false, error: 'Employee ID already exists' };
      }
    }
    
    // Add to the end of the sheet
    const newRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.note || '',
      employee.photoId || ''
    ];
    
    sheet.appendRow(newRow);
    
    Logger.log(`Added employee: ${employee.empId}`);
    return { success: true, message: 'Employee added successfully' };
    
  } catch (error) {
    Logger.log('Error in addEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Update an existing employee
 */
function updateEmployee(employee, originalEmpId) {
  try {
    const sheet = ensureHeaders(getSheet());
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1);
    const empIds = dataRange.getValues().flat();
    const rowIndex = empIds.findIndex(id => id === originalEmpId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Check for duplicate Emp ID if it's being changed
    if (employee.empId !== originalEmpId) {
      const duplicate = empIds.findIndex(id => id === employee.empId);
      if (duplicate !== -1) {
        return { success: false, error: 'New Employee ID already exists' };
      }
    }
    
    // Update the row
    const targetRow = DATA_START_ROW + rowIndex;
    const updatedRow = [
      employee.empId || '',
      employee.firstName || '',
      employee.lastName || '',
      employee.phone || '',
      employee.email || '',
      employee.position || '',
      employee.note || '',
      employee.photoId || ''
    ];
    
    sheet.getRange(targetRow, 1, 1, HEADERS.length).setValues([updatedRow]);
    
    Logger.log(`Updated employee: ${originalEmpId} -> ${employee.empId}`);
    return { success: true, message: 'Employee updated successfully' };
    
  } catch (error) {
    Logger.log('Error in updateEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Delete an employee
 */
function deleteEmployee(empId) {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow < DATA_START_ROW) {
      return { success: false, error: 'No employees found' };
    }
    
    // Find the employee row
    const dataRange = sheet.getRange(DATA_START_ROW, 1, lastRow - HEADER_ROW, 1);
    const empIds = dataRange.getValues().flat();
    const rowIndex = empIds.findIndex(id => id === empId);
    
    if (rowIndex === -1) {
      return { success: false, error: 'Employee not found' };
    }
    
    // Delete the row
    const targetRow = DATA_START_ROW + rowIndex;
    sheet.deleteRow(targetRow);
    
    Logger.log(`Deleted employee: ${empId}`);
    return { success: true, message: 'Employee deleted successfully' };
    
  } catch (error) {
    Logger.log('Error in deleteEmployee: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Save an employee photo (data URL) to Drive and return file info
 */
function saveEmployeePhoto(dataUrl, empId) {
  try {
    if (!dataUrl) {
      return { success: false, error: 'No image data provided' };
    }

    // Parse data URL: data:image/png;base64,AAAA
    const matches = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
    if (!matches) {
      return { success: false, error: 'Invalid image data format' };
    }

    const contentType = matches[1];
    const base64 = matches[2];
    const bytes = Utilities.base64Decode(base64);

    // Derive extension from contentType
    const extension = contentType.split('/')[1] || 'png';
    const safeEmpId = (empId || 'employee').toString().replace(/[^a-zA-Z0-9_-]/g, '_');
    const fileName = `${safeEmpId}_${Date.now()}.${extension}`;

    const folder = getOrCreatePhotosFolder();
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);

    // Build a view URL that works for Drive-hosted images (user must have access)
    const photoId = file.getId();
    const viewUrl = `https://drive.google.com/uc?export=view&id=${photoId}`;

    return { success: true, photoId: photoId, viewUrl: viewUrl };
  } catch (error) {
    Logger.log('Error in saveEmployeePhoto: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Get or create the folder used to store employee photos
 */
function getOrCreatePhotosFolder() {
  const folderName = 'Bar Employee CRM Photos';
  const it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) {
    return it.next();
  }
  return DriveApp.createFolder(folderName);
}
