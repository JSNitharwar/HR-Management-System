/**
 * Utility Functions
 */

// Generate Random Token
function generateToken() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let token = '';
  for (let i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token;
}

// Save Token to Sheet
function saveToken(token, type, referenceId, email) {
  try {
    const now = new Date();
    const expiry = new Date(now.getTime() + (7 * 24 * 60 * 60 * 1000)); // 7 days
    
    const tokenData = {
      'Token': token,
      'Type': type,
      'Reference ID': referenceId,
      'Email': email,
      'Created Date': now,
      'Expiry Date': expiry,
      'Used': 'No',
      'Used Date': ''
    };
    
    addRowToSheet(SHEETS.TOKENS, tokenData);
    
  } catch(error) {
    Logger.log('Error saving token: ' + error.toString());
  }
}

// Validate Token
function validateToken(token, expectedType) {
  try {
    const tokenData = findRowByValue(SHEETS.TOKENS, 'Token', token);
    
    if (!tokenData) return false;
    if (tokenData['Type'] !== expectedType) return false;
    if (tokenData['Used'] === 'Yes') return false;
    
    const expiry = new Date(tokenData['Expiry Date']);
    if (new Date() > expiry) return false;
    
    return true;
    
  } catch(error) {
    Logger.log('Error validating token: ' + error.toString());
    return false;
  }
}

// Log Action to Audit Log
function logAction(action, entityType, entityId, oldValue, newValue) {
  try {
    const logData = {
      'Timestamp': new Date(),
      'User': Session.getActiveUser().getEmail() || 'System',
      'Action': action,
      'Entity Type': entityType,
      'Entity ID': entityId,
      'Old Value': oldValue || '',
      'New Value': newValue || '',
      'IP Address': ''
    };
    
    addRowToSheet(SHEETS.AUDIT_LOG, logData);
    
  } catch(error) {
    Logger.log('Error logging action: ' + error.toString());
  }
}

// Parse CSV Data
function parseCSVData(csvContent) {
  try {
    const lines = csvContent.split('\n');
    const headers = lines[0].split(',').map(h => h.trim());
    const data = [];
    
    for (let i = 1; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;
      
      const values = line.split(',');
      const row = {};
      
      for (let j = 0; j < headers.length; j++) {
        row[headers[j]] = values[j] ? values[j].trim() : '';
      }
      
      data.push(row);
    }
    
    return data;
    
  } catch(error) {
    Logger.log('Error parsing CSV: ' + error.toString());
    return [];
  }
}

// Format Date
function formatDate(date, format) {
  if (!date) return '';
  
  try {
    const d = new Date(date);
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    
    if (format === 'dd/MM/yyyy') {
      return day + '/' + month + '/' + year;
    }
    
    return year + '-' + month + '-' + day;
    
  } catch(error) {
    return '';
  }
}

// Validate Email
function isValidEmail(email) {
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

// Validate Phone
function isValidPhone(phone) {
  const regex = /^[0-9]{10}$/;
  return regex.test(phone.replace(/[\s-]/g, ''));
}
