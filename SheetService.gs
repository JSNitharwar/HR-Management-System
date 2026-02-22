/**
 * Google Sheets Database Service
 * Handles all spreadsheet operations
 */

// Get Master Spreadsheet
function getMasterSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
  } catch(error) {
    Logger.log('Error opening spreadsheet: ' + error.toString());
    throw new Error('Cannot open spreadsheet. Check MASTER_SPREADSHEET_ID in Config.gs');
  }
}

// Get Sheet by Name
function getSheet(sheetName) {
  try {
    const ss = getMasterSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log('Sheet not found: ' + sheetName + '. Running initialization...');
      initializeSheets();
      sheet = ss.getSheetByName(sheetName);
    }
    
    return sheet;
  } catch(error) {
    Logger.log('Error getting sheet ' + sheetName + ': ' + error.toString());
    return null;
  }
}

// Get all data from sheet as array of objects
function getSheetData(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    if (!sheet) {
      Logger.log('Sheet not found: ' + sheetName);
      return [];
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1 || lastCol === 0) {
      return [];
    }
    
    const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = data[0];
    const rows = data.slice(1);
    
    const result = [];
    
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      
      // Skip empty rows
      const hasData = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
      if (!hasData) continue;
      
      const obj = { rowIndex: i + 2 };
      
      for (let j = 0; j < headers.length; j++) {
        let value = row[j];
        
        // Convert Date objects to strings
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        
        obj[headers[j]] = value;
      }
      
      result.push(obj);
    }
    
    return result;
    
  } catch(error) {
    Logger.log('Error getting data from ' + sheetName + ': ' + error.toString());
    return [];
  }
}

// Add new row to sheet
function addRowToSheet(sheetName, rowData) {
  try {
    const sheet = getSheet(sheetName);
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const row = headers.map(header => {
      const value = rowData[header];
      if (value === undefined || value === null) return '';
      return value;
    });
    
    sheet.appendRow(row);
    SpreadsheetApp.flush();
    
    return sheet.getLastRow();
    
  } catch(error) {
    Logger.log('Error adding row to ' + sheetName + ': ' + error.toString());
    throw error;
  }
}

// Update existing row in sheet
function updateRowInSheet(sheetName, rowIndex, rowData) {
  try {
    const sheet = getSheet(sheetName);
    if (!sheet) {
      throw new Error('Sheet not found: ' + sheetName);
    }
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (rowData.hasOwnProperty(header)) {
        sheet.getRange(rowIndex, i + 1).setValue(rowData[header]);
      }
    }
    
    SpreadsheetApp.flush();
    return true;
    
  } catch(error) {
    Logger.log('Error updating row in ' + sheetName + ': ' + error.toString());
    throw error;
  }
}

// Find single row by column value
function findRowByValue(sheetName, columnName, value) {
  try {
    const data = getSheetData(sheetName);
    const found = data.find(row => row[columnName] == value);
    return found || null;
  } catch(error) {
    Logger.log('Error finding row: ' + error.toString());
    return null;
  }
}

// Find multiple rows by column value
function findRowsByValue(sheetName, columnName, value) {
  try {
    const data = getSheetData(sheetName);
    return data.filter(row => row[columnName] == value);
  } catch(error) {
    Logger.log('Error finding rows: ' + error.toString());
    return [];
  }
}

// Generate unique ID with prefix
function generateId(prefix) {
  const timestamp = new Date().getTime();
  const random = Math.floor(Math.random() * 10000);
  return prefix + timestamp + random;
}

// ============================================
// LEGAL ENTITY FUNCTIONS
// ============================================

// Get Active Legal Entities
function getLegalEntities() {
  try {
    const entities = getSheetData(SHEETS.LEGAL_ENTITIES);
    return entities.filter(e => e['Status'] === 'Active');
  } catch(error) {
    Logger.log('Error getting legal entities: ' + error.toString());
    return [];
  }
}

// Get All Legal Entities (including inactive)
function getAllLegalEntities() {
  try {
    return getSheetData(SHEETS.LEGAL_ENTITIES);
  } catch(error) {
    Logger.log('Error getting all legal entities: ' + error.toString());
    return [];
  }
}

// Get Legal Entity by ID
function getLegalEntityById(entityId) {
  try {
    return findRowByValue(SHEETS.LEGAL_ENTITIES, 'Entity ID', entityId);
  } catch(error) {
    Logger.log('Error getting entity: ' + error.toString());
    return null;
  }
}

// Add New Legal Entity
function addLegalEntity(entityData) {
  try {
    const entityId = generateId('LE');
    
    const entity = {
      'Entity ID': entityId,
      'Entity Name': entityData.entityName || '',
      'Entity Code': (entityData.entityCode || '').toUpperCase(),
      'Registered Address': entityData.registeredAddress || '',
      'Contact Email': entityData.contactEmail || '',
      'Contact Phone': entityData.contactPhone || '',
      'GST Number': (entityData.gstNumber || '').toUpperCase(),
      'CIN Number': (entityData.cinNumber || '').toUpperCase(),
      'Status': 'Active'
    };
    
    addRowToSheet(SHEETS.LEGAL_ENTITIES, entity);
    logAction('CREATE', 'LegalEntity', entityId, '', JSON.stringify(entity));
    
    return {success: true, entityId: entityId, message: 'Legal entity added successfully'};
    
  } catch(error) {
    Logger.log('Error adding legal entity: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Update Legal Entity
function updateLegalEntity(entityId, entityData) {
  try {
    const entity = getLegalEntityById(entityId);
    if (!entity) {
      return {success: false, error: 'Entity not found'};
    }
    
    const updateData = {
      'Entity Name': entityData.entityName || entity['Entity Name'],
      'Entity Code': (entityData.entityCode || entity['Entity Code']).toUpperCase(),
      'Registered Address': entityData.registeredAddress || '',
      'Contact Email': entityData.contactEmail || '',
      'Contact Phone': entityData.contactPhone || '',
      'GST Number': (entityData.gstNumber || '').toUpperCase(),
      'CIN Number': (entityData.cinNumber || '').toUpperCase(),
      'Status': entityData.status || entity['Status']
    };
    
    updateRowInSheet(SHEETS.LEGAL_ENTITIES, entity.rowIndex, updateData);
    logAction('UPDATE', 'LegalEntity', entityId, JSON.stringify(entity), JSON.stringify(updateData));
    
    return {success: true, message: 'Legal entity updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating legal entity: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Delete (Deactivate) Legal Entity
function deleteLegalEntity(entityId) {
  try {
    const entity = getLegalEntityById(entityId);
    if (!entity) {
      return {success: false, error: 'Entity not found'};
    }
    
    // Check for associated employees
    const employees = findRowsByValue(SHEETS.EMPLOYEES, 'Legal Entity', entity['Entity Name']);
    if (employees.length > 0) {
      return {success: false, error: 'Cannot delete. ' + employees.length + ' employees are associated with this entity.'};
    }
    
    updateRowInSheet(SHEETS.LEGAL_ENTITIES, entity.rowIndex, {'Status': 'Inactive'});
    logAction('DELETE', 'LegalEntity', entityId, 'Active', 'Inactive');
    
    return {success: true, message: 'Legal entity deactivated successfully'};
    
  } catch(error) {
    Logger.log('Error deleting legal entity: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Reactivate Legal Entity
function reactivateLegalEntity(entityId) {
  try {
    const entity = getLegalEntityById(entityId);
    if (!entity) {
      return {success: false, error: 'Entity not found'};
    }
    
    updateRowInSheet(SHEETS.LEGAL_ENTITIES, entity.rowIndex, {'Status': 'Active'});
    logAction('REACTIVATE', 'LegalEntity', entityId, 'Inactive', 'Active');
    
    return {success: true, message: 'Legal entity reactivated successfully'};
    
  } catch(error) {
    Logger.log('Error reactivating legal entity: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}
