/**
 * HR Management System - Main Entry Point
 * Version: 3.0 (Fixed)
 */

// Web App Entry Points
function doGet(e) {
  try {
    const page = e.parameter.page || 'dashboard';
    const token = e.parameter.token || '';
    const jobId = e.parameter.jobId || '';
    
    switch(page) {
      case 'dashboard':
        return createHRDashboard();
      case 'joining':
        return createJoiningForm(token);
      case 'apply':
        return createCandidateApplicationForm(jobId);
      case 'interview':
        return createInterviewForm(token);
      default:
        return createHRDashboard();
    }
  } catch(error) {
    Logger.log('Error in doGet: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.toString() + '</p>');
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    
    switch(action) {
      case 'addEmployee':
        return ContentService.createTextOutput(JSON.stringify(addNewEmployee(data.employee)));
      case 'submitJoiningForm':
        return ContentService.createTextOutput(JSON.stringify(submitJoiningForm(data)));
      case 'submitApplication':
        return ContentService.createTextOutput(JSON.stringify(submitCandidateApplication(data)));
      default:
        return ContentService.createTextOutput(JSON.stringify({success: false, error: 'Unknown action'}));
    }
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: error.toString()}));
  }
}

// Create HR Dashboard
function createHRDashboard() {
  try {
    const template = HtmlService.createTemplateFromFile('HRDashboard');
    template.legalEntities = getLegalEntities() || [];
    template.webAppUrl = getWebAppUrl();
    
    return template.evaluate()
      .setTitle('HR Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(error) {
    Logger.log('Error creating dashboard: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error loading dashboard</h1><p>' + error.toString() + '</p>');
  }
}

// Create Joining Form for Employees
function createJoiningForm(token) {
  try {
    if (!token || !validateToken(token, 'joining')) {
      return HtmlService.createHtmlOutput(`
        <html>
        <head><style>body{font-family:Arial;text-align:center;padding:50px;}</style></head>
        <body>
          <h1>Invalid or Expired Link</h1>
          <p>Please contact HR for a new joining form link.</p>
        </body>
        </html>
      `);
    }
    
    const template = HtmlService.createTemplateFromFile('EmployeeJoiningForm');
    template.token = token;
    template.employeeData = getEmployeeByToken(token) || {};
    
    return template.evaluate()
      .setTitle('Employee Joining Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(error) {
    Logger.log('Error creating joining form: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.toString() + '</p>');
  }
}

// Create Candidate Application Form
function createCandidateApplicationForm(jobId) {
  try {
    const template = HtmlService.createTemplateFromFile('CandidateApplicationForm');
    template.jobDetails = jobId ? getJobDetails(jobId) : null;
    template.openJobs = getOpenJobs() || [];
    
    return template.evaluate()
      .setTitle('Career Application')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(error) {
    Logger.log('Error creating application form: ' + error.toString());
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.toString() + '</p>');
  }
}

// Create Interview Form
function createInterviewForm(token) {
  try {
    const template = HtmlService.createTemplateFromFile('CandidateApplicationForm');
    template.jobDetails = null;
    template.openJobs = getOpenJobs() || [];
    
    return template.evaluate()
      .setTitle('Complete Your Application')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch(error) {
    return HtmlService.createHtmlOutput('<h1>Error</h1><p>' + error.toString() + '</p>');
  }
}

// Include HTML files
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Get Web App URL
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch(error) {
    return '';
  }
}

// Test function to verify setup
function testSetup() {
  const results = {
    config: {
      spreadsheetId: CONFIG.MASTER_SPREADSHEET_ID,
      folderId: CONFIG.HR_DOCUMENTS_FOLDER_ID
    },
    spreadsheetAccess: false,
    folderAccess: false,
    sheets: {}
  };
  
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    results.spreadsheetAccess = true;
    results.spreadsheetName = ss.getName();
    
    const sheetNames = [SHEETS.EMPLOYEES, SHEETS.LEGAL_ENTITIES, SHEETS.CANDIDATES, 
                        SHEETS.JOBS, SHEETS.INTERVIEWS, SHEETS.TOKENS, SHEETS.AUDIT_LOG];
    
    sheetNames.forEach(name => {
      const sheet = ss.getSheetByName(name);
      results.sheets[name] = {
        exists: !!sheet,
        rows: sheet ? sheet.getLastRow() : 0
      };
    });
  } catch(e) {
    results.spreadsheetError = e.toString();
  }
  
  try {
    const folder = DriveApp.getFolderById(CONFIG.HR_DOCUMENTS_FOLDER_ID);
    results.folderAccess = true;
    results.folderName = folder.getName();
  } catch(e) {
    results.folderError = e.toString();
  }
  
  Logger.log(JSON.stringify(results, null, 2));
  return results;
}
