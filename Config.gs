/**
 * Configuration and Constants
 * UPDATE THE IDs BELOW WITH YOUR OWN
 */

const CONFIG = {
  MASTER_SPREADSHEET_ID: '1kidR05EnXdeL7uvtxnMEa6mcQ5MZ19-RuaTXk0UAJBw',
  HR_DOCUMENTS_FOLDER_ID: '1k3Gu32yfnwV_iqTm46gwBo47qs854ZO7',
  HR_EMAIL: 'hr@sarthihrms.co.in',
  COMPANY_NAME: 'Sarthi HRMS',
  COMPANY_LOGO_URL: '',
  COMPANY_ADDRESS: '123 Business Park, City, State - 123456',
  COMPANY_WEBSITE: 'www.sarthihrms.co.in',
  COMPANY_PHONE: '+91 9261116667'
};

const SHEETS = {
  EMPLOYEES: 'Employees',
  LEGAL_ENTITIES: 'LegalEntities',
  CANDIDATES: 'Candidates',
  JOBS: 'Jobs',
  INTERVIEWS: 'Interviews',
  TOKENS: 'Tokens',
  AUDIT_LOG: 'AuditLog',
  OFFERS: 'Offers'  // NEW
};

const EMPLOYEE_STATUS = {
  INVITED: 'Invited',
  FORM_PENDING: 'Form Pending',
  DOCUMENTS_PENDING: 'Documents Pending',
  ACTIVE: 'Active',
  INACTIVE: 'Inactive',
  TERMINATED: 'Terminated'
};

const CANDIDATE_STATUS = {
  APPLIED: 'Applied',
  SCREENING: 'Screening',
  INTERVIEW_SCHEDULED: 'Interview Scheduled',
  INTERVIEWED: 'Interviewed',
  SELECTED: 'Selected',
  REJECTED: 'Rejected',
  OFFER_SENT: 'Offer Sent',
  OFFER_ACCEPTED: 'Offer Accepted',
  OFFER_REJECTED: 'Offer Rejected',
  JOINED: 'Joined'
};

const OFFER_STATUS = {
  DRAFT: 'Draft',
  SENT: 'Sent',
  ACCEPTED: 'Accepted',
  REJECTED: 'Rejected',
  EXPIRED: 'Expired',
  WITHDRAWN: 'Withdrawn'
};

const EMPLOYMENT_TYPES = ['Full-time', 'Part-time', 'Contract', 'Intern', 'Trainee', 'Consultant', 'Temporary'];
const SHIFT_TYPES = ['General Shift', 'Morning Shift', 'Day Shift', 'Evening Shift', 'Night Shift', 'Rotational Shift', 'Flexible'];

// Initialize all required sheets
function initializeSheets() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    
    // Employees Sheet
    createSheetWithHeaders(ss, SHEETS.EMPLOYEES, [
      'Employee ID', 'Legal Entity', 'First Name', 'Last Name', 'Personal Email', 
      'Official Email', 'Phone', 'Alternate Phone', 'Aadhar Number', 'PAN Number',
      'Date of Birth', 'Gender', 'Blood Group', 'Marital Status', 'Nationality',
      'Current Address', 'Permanent Address', 'Emergency Contact Name', 
      'Emergency Contact Phone', 'Emergency Contact Relation',
      'Department', 'Designation', 'Date of Joining', 'Reporting Manager',
      'Employment Type', 'Work Location', 'Shift', 'Probation End Date',
      'Bank Name', 'Account Number', 'IFSC Code', 'Account Holder Name',
      'UAN Number', 'ESI Number', 
      'CTC Annual', 'CTC Monthly', 'Basic Salary', 'HRA', 'Conveyance Allowance',
      'Medical Allowance', 'Special Allowance', 'Other Allowances', 'PF Employer',
      'ESI Employer', 'Gratuity', 'Gross Salary', 'Net Salary', 'Salary Notes',
      'Documents Folder ID', 'Form Completed', 'Official Info Completed',
      'Status', 'Created Date', 'Modified Date', 'Created By'
    ]);
    
    // Legal Entities Sheet
    createSheetWithHeaders(ss, SHEETS.LEGAL_ENTITIES, [
      'Entity ID', 'Entity Name', 'Entity Code', 'Registered Address',
      'Contact Email', 'Contact Phone', 'GST Number', 'CIN Number', 'Status'
    ]);
    
    // Candidates Sheet
    createSheetWithHeaders(ss, SHEETS.CANDIDATES, [
      'Candidate ID', 'Job ID', 'First Name', 'Last Name', 'Email', 'Phone',
      'Current Company', 'Current Designation', 'Total Experience', 
      'Current CTC', 'Expected CTC', 'Notice Period', 'Location',
      'Resume Folder ID', 'Application Date', 'Source', 'Status',
      'Screening Score', 'HR Notes', 'Created Date', 'Modified Date'
    ]);
    
    // Jobs Sheet
    createSheetWithHeaders(ss, SHEETS.JOBS, [
      'Job ID', 'Legal Entity', 'Job Title', 'Department', 'Description',
      'Requirements', 'Experience Required', 'Location', 'Employment Type',
      'Salary Range', 'Number of Positions', 'Status', 'Posted Date',
      'Closing Date', 'Created By'
    ]);
    
    // Interviews Sheet
    createSheetWithHeaders(ss, SHEETS.INTERVIEWS, [
      'Interview ID', 'Candidate ID', 'Job ID', 'Round', 'Interview Date',
      'Interview Time', 'Duration', 'Interview Type', 'Meeting Link',
      'Interviewer Email', 'Interviewer Name', 'Status', 'Feedback',
      'Rating', 'Recommendation', 'Created Date'
    ]);
    
    // Offers Sheet - NEW
    createSheetWithHeaders(ss, SHEETS.OFFERS, [
      'Offer ID', 'Candidate ID', 'Job ID', 'Legal Entity',
      'Designation', 'Department', 'Work Location', 'Employment Type',
      'Reporting Manager', 'Joining Date', 'Probation Period',
      'CTC Annual', 'CTC Monthly', 'Basic Salary', 'HRA', 'Other Allowances',
      'Offer Date', 'Validity Date', 'Acceptance Deadline',
      'Status', 'Sent Date', 'Response Date', 'Response Notes',
      'Offer Letter URL', 'Acceptance Token', 'Created By', 'Created Date'
    ]);
    
    // Tokens Sheet
    createSheetWithHeaders(ss, SHEETS.TOKENS, [
      'Token', 'Type', 'Reference ID', 'Email', 'Created Date', 
      'Expiry Date', 'Used', 'Used Date'
    ]);
    
    // Audit Log Sheet
    createSheetWithHeaders(ss, SHEETS.AUDIT_LOG, [
      'Timestamp', 'User', 'Action', 'Entity Type', 'Entity ID',
      'Old Value', 'New Value', 'IP Address'
    ]);
    
    Logger.log('All sheets initialized successfully');
    return {success: true, message: 'All sheets initialized successfully'};
    
  } catch(error) {
    Logger.log('Error initializing sheets: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

function createSheetWithHeaders(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    Logger.log('Created new sheet: ' + sheetName);
  }
  
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  } else {
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const newHeaders = headers.filter(h => !existingHeaders.includes(h));
    
    if (newHeaders.length > 0) {
      newHeaders.forEach((header, index) => {
        sheet.getRange(1, lastCol + index + 1).setValue(header);
        sheet.getRange(1, lastCol + index + 1).setFontWeight('bold');
        sheet.getRange(1, lastCol + index + 1).setBackground('#f3f3f3');
      });
    }
  }
  
  return sheet;
}

function getDepartments() {
  try {
    const employees = getSheetData(SHEETS.EMPLOYEES);
    const departments = [...new Set(employees.map(e => e['Department']).filter(d => d && d.trim()))];
    return departments.sort();
  } catch(error) { return []; }
}

function getDesignations() {
  try {
    const employees = getSheetData(SHEETS.EMPLOYEES);
    const designations = [...new Set(employees.map(e => e['Designation']).filter(d => d && d.trim()))];
    return designations.sort();
  } catch(error) { return []; }
}

function getWorkLocations() {
  try {
    const employees = getSheetData(SHEETS.EMPLOYEES);
    const locations = [...new Set(employees.map(e => e['Work Location']).filter(l => l && l.trim()))];
    return locations.sort();
  } catch(error) { return []; }
}

function getManagers() {
  try {
    const employees = getSheetData(SHEETS.EMPLOYEES);
    return employees
      .filter(e => e['Status'] === 'Active')
      .map(e => ({
        id: e['Employee ID'],
        name: ((e['First Name'] || '') + ' ' + (e['Last Name'] || '')).trim(),
        designation: e['Designation'] || ''
      }))
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch(error) { return []; }
}
