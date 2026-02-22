/**
 * Employee Management Service
 */

// Add New Employee
function addNewEmployee(employeeData) {
  try {
    const employeeId = generateId('EMP');
    const token = generateToken();
    const now = new Date();
    
    // Create employee folder
    let folderId = '';
    try {
      const folderResult = createEmployeeFolder(employeeId, employeeData['First Name'] || '', employeeData['Last Name'] || '');
      folderId = folderResult.folderId || '';
    } catch(e) {
      Logger.log('Could not create folder: ' + e.toString());
    }
    
    const employee = {
      'Employee ID': employeeId,
      'Legal Entity': employeeData['Legal Entity'] || '',
      'First Name': employeeData['First Name'] || '',
      'Last Name': employeeData['Last Name'] || '',
      'Personal Email': employeeData['Personal Email'] || '',
      'Official Email': employeeData['Official Email'] || '',
      'Phone': employeeData['Phone'] || '',
      'Alternate Phone': '',
      'Aadhar Number': employeeData['Aadhar Number'] || '',
      'PAN Number': (employeeData['PAN Number'] || '').toUpperCase(),
      'Date of Birth': '',
      'Gender': '',
      'Blood Group': '',
      'Marital Status': '',
      'Nationality': 'Indian',
      'Current Address': '',
      'Permanent Address': '',
      'Emergency Contact Name': '',
      'Emergency Contact Phone': '',
      'Emergency Contact Relation': '',
      'Department': employeeData['Department'] || '',
      'Designation': employeeData['Designation'] || '',
      'Date of Joining': employeeData['Date of Joining'] || '',
      'Reporting Manager': employeeData['Reporting Manager'] || '',
      'Employment Type': employeeData['Employment Type'] || '',
      'Work Location': employeeData['Work Location'] || '',
      'Shift': employeeData['Shift'] || '',
      'Probation End Date': employeeData['Probation End Date'] || '',
      'Bank Name': '',
      'Account Number': '',
      'IFSC Code': '',
      'Account Holder Name': '',
      'UAN Number': '',
      'ESI Number': '',
      'CTC Annual': employeeData['CTC Annual'] || '',
      'CTC Monthly': employeeData['CTC Monthly'] || '',
      'Basic Salary': employeeData['Basic Salary'] || '',
      'HRA': employeeData['HRA'] || '',
      'Conveyance Allowance': employeeData['Conveyance Allowance'] || '',
      'Medical Allowance': employeeData['Medical Allowance'] || '',
      'Special Allowance': employeeData['Special Allowance'] || '',
      'Other Allowances': employeeData['Other Allowances'] || '',
      'PF Employer': employeeData['PF Employer'] || '',
      'ESI Employer': employeeData['ESI Employer'] || '',
      'Gratuity': employeeData['Gratuity'] || '',
      'Gross Salary': employeeData['Gross Salary'] || '',
      'Net Salary': employeeData['Net Salary'] || '',
      'Salary Notes': employeeData['Salary Notes'] || '',
      'Documents Folder ID': folderId,
      'Form Completed': 'No',
      'Official Info Completed': employeeData['Department'] ? 'Yes' : 'No',
      'Status': EMPLOYEE_STATUS.INVITED,
      'Created Date': now,
      'Modified Date': now,
      'Created By': Session.getActiveUser().getEmail() || 'System'
    };
    
    addRowToSheet(SHEETS.EMPLOYEES, employee);
    saveToken(token, 'joining', employeeId, employeeData['Personal Email'] || '');
    
    // Send welcome email
    try {
      if (employeeData['Personal Email']) {
        sendWelcomeEmail(employeeData['Personal Email'], employeeData['First Name'] || 'Employee', token);
      }
    } catch(e) {
      Logger.log('Could not send email: ' + e.toString());
    }
    
    logAction('CREATE', 'Employee', employeeId, '', JSON.stringify(employee));
    
    return {success: true, employeeId: employeeId, message: 'Employee added successfully'};
    
  } catch(error) {
    Logger.log('Error adding employee: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get All Employees
function getAllEmployees(filters) {
  try {
    filters = filters || {};
    let employees = getSheetData(SHEETS.EMPLOYEES);
    
    if (filters.legalEntity && filters.legalEntity !== '') {
      employees = employees.filter(e => e['Legal Entity'] === filters.legalEntity);
    }
    if (filters.status && filters.status !== '') {
      employees = employees.filter(e => e['Status'] === filters.status);
    }
    if (filters.department && filters.department !== '') {
      employees = employees.filter(e => e['Department'] === filters.department);
    }
    
    return employees;
    
  } catch(error) {
    Logger.log('Error getting employees: ' + error.toString());
    return [];
  }
}

// Get Employee by ID
function getEmployeeById(employeeId) {
  try {
    return findRowByValue(SHEETS.EMPLOYEES, 'Employee ID', employeeId);
  } catch(error) {
    Logger.log('Error getting employee: ' + error.toString());
    return null;
  }
}

// Get Employee by Token
function getEmployeeByToken(token) {
  try {
    const tokenData = findRowByValue(SHEETS.TOKENS, 'Token', token);
    if (tokenData && tokenData['Type'] === 'joining') {
      return getEmployeeById(tokenData['Reference ID']);
    }
    return null;
  } catch(error) {
    Logger.log('Error getting employee by token: ' + error.toString());
    return null;
  }
}

// Update Employee Basic Info
function updateEmployee(employeeId, updateData) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    updateData['Modified Date'] = new Date();
    updateRowInSheet(SHEETS.EMPLOYEES, employee.rowIndex, updateData);
    
    logAction('UPDATE', 'Employee', employeeId, JSON.stringify(employee), JSON.stringify(updateData));
    
    return {success: true, message: 'Employee updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating employee: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Update Employee Official Information (New Function)
function updateEmployeeOfficialInfo(employeeId, officialData) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    const updateData = {
      'Official Email': officialData.officialEmail || employee['Official Email'] || '',
      'Department': officialData.department || '',
      'Designation': officialData.designation || '',
      'Date of Joining': officialData.dateOfJoining || '',
      'Reporting Manager': officialData.reportingManager || '',
      'Employment Type': officialData.employmentType || '',
      'Work Location': officialData.workLocation || '',
      'Shift': officialData.shift || '',
      'Probation End Date': officialData.probationEndDate || '',
      'Official Info Completed': 'Yes',
      'Modified Date': new Date()
    };
    
    // Update status to Active if form is completed and official info is filled
    if (employee['Form Completed'] === 'Yes' && officialData.department) {
      updateData['Status'] = EMPLOYEE_STATUS.ACTIVE;
    }
    
    updateRowInSheet(SHEETS.EMPLOYEES, employee.rowIndex, updateData);
    
    logAction('UPDATE_OFFICIAL_INFO', 'Employee', employeeId, JSON.stringify(employee), JSON.stringify(updateData));
    
    return {success: true, message: 'Official information updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating official info: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Update Employee Salary Information (New Function)
function updateEmployeeSalary(employeeId, salaryData) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    // Calculate values
    const ctcAnnual = parseFloat(salaryData.ctcAnnual) || 0;
    const ctcMonthly = ctcAnnual / 12;
    
    const updateData = {
      'CTC Annual': ctcAnnual,
      'CTC Monthly': Math.round(ctcMonthly * 100) / 100,
      'Basic Salary': parseFloat(salaryData.basicSalary) || 0,
      'HRA': parseFloat(salaryData.hra) || 0,
      'Conveyance Allowance': parseFloat(salaryData.conveyanceAllowance) || 0,
      'Medical Allowance': parseFloat(salaryData.medicalAllowance) || 0,
      'Special Allowance': parseFloat(salaryData.specialAllowance) || 0,
      'Other Allowances': parseFloat(salaryData.otherAllowances) || 0,
      'PF Employer': parseFloat(salaryData.pfEmployer) || 0,
      'ESI Employer': parseFloat(salaryData.esiEmployer) || 0,
      'Gratuity': parseFloat(salaryData.gratuity) || 0,
      'Gross Salary': parseFloat(salaryData.grossSalary) || 0,
      'Net Salary': parseFloat(salaryData.netSalary) || 0,
      'Salary Notes': salaryData.salaryNotes || '',
      'Modified Date': new Date()
    };
    
    updateRowInSheet(SHEETS.EMPLOYEES, employee.rowIndex, updateData);
    
    logAction('UPDATE_SALARY', 'Employee', employeeId, 'Old Salary Data', JSON.stringify(updateData));
    
    return {success: true, message: 'Salary information updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating salary: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Update Employee Status
function updateEmployeeStatus(employeeId, newStatus) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    const oldStatus = employee['Status'];
    updateRowInSheet(SHEETS.EMPLOYEES, employee.rowIndex, {
      'Status': newStatus,
      'Modified Date': new Date()
    });
    
    logAction('UPDATE_STATUS', 'Employee', employeeId, oldStatus, newStatus);
    
    return {success: true, message: 'Status updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating status: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get Employee Statistics
function getEmployeeStatistics(legalEntity) {
  try {
    let employees = getSheetData(SHEETS.EMPLOYEES);
    
    if (legalEntity && legalEntity !== '') {
      employees = employees.filter(e => e['Legal Entity'] === legalEntity);
    }
    
    return {
      total: employees.length,
      active: employees.filter(e => e['Status'] === EMPLOYEE_STATUS.ACTIVE).length,
      pending: employees.filter(e => 
        e['Status'] === EMPLOYEE_STATUS.INVITED ||
        e['Status'] === EMPLOYEE_STATUS.FORM_PENDING ||
        e['Status'] === EMPLOYEE_STATUS.DOCUMENTS_PENDING
      ).length,
      inactive: employees.filter(e => e['Status'] === EMPLOYEE_STATUS.INACTIVE).length,
      officialInfoPending: employees.filter(e => e['Official Info Completed'] !== 'Yes').length
    };
    
  } catch(error) {
    Logger.log('Error getting statistics: ' + error.toString());
    return {total: 0, active: 0, pending: 0, inactive: 0, officialInfoPending: 0};
  }
}

// Submit Joining Form
function submitJoiningForm(formData) {
  try {
    const token = formData.token;
    const tokenData = findRowByValue(SHEETS.TOKENS, 'Token', token);
    
    if (!tokenData || tokenData['Used'] === 'Yes') {
      return {success: false, error: 'Invalid or used token'};
    }
    
    const employeeId = tokenData['Reference ID'];
    const employee = getEmployeeById(employeeId);
    
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    const updateData = {
      'Date of Birth': formData.dateOfBirth || '',
      'Gender': formData.gender || '',
      'Blood Group': formData.bloodGroup || '',
      'Marital Status': formData.maritalStatus || '',
      'Nationality': formData.nationality || 'Indian',
      'Current Address': formData.currentAddress || '',
      'Permanent Address': formData.permanentAddress || '',
      'Alternate Phone': formData.alternatePhone || '',
      'Emergency Contact Name': formData.emergencyContactName || '',
      'Emergency Contact Phone': formData.emergencyContactPhone || '',
      'Emergency Contact Relation': formData.emergencyContactRelation || '',
      'Bank Name': formData.bankName || '',
      'Account Number': formData.accountNumber || '',
      'IFSC Code': (formData.ifscCode || '').toUpperCase(),
      'Account Holder Name': formData.accountHolderName || '',
      'UAN Number': formData.uanNumber || '',
      'ESI Number': formData.esiNumber || '',
      'Form Completed': 'Yes',
      'Status': EMPLOYEE_STATUS.DOCUMENTS_PENDING,
      'Modified Date': new Date()
    };
    
    // If official info is also completed, set status to Active
    if (employee['Official Info Completed'] === 'Yes') {
      updateData['Status'] = EMPLOYEE_STATUS.ACTIVE;
    }
    
    updateRowInSheet(SHEETS.EMPLOYEES, employee.rowIndex, updateData);
    
    // Upload documents
    if (formData.documents && formData.documents.length > 0 && employee['Documents Folder ID']) {
      try {
        uploadEmployeeDocuments(employee['Documents Folder ID'], formData.documents);
        if (employee['Official Info Completed'] === 'Yes') {
          updateRowInSheet(SHEETS.EMPLOYEES, employee.rowIndex, {'Status': EMPLOYEE_STATUS.ACTIVE});
        }
      } catch(e) {
        Logger.log('Document upload error: ' + e.toString());
      }
    }
    
    // Mark token as used
    updateRowInSheet(SHEETS.TOKENS, tokenData.rowIndex, {
      'Used': 'Yes',
      'Used Date': new Date()
    });
    
    // Send emails
    try {
      sendJoiningFormConfirmationEmail(tokenData['Email'], employee['First Name']);
      sendHRNotification('Joining Form Submitted', employee['First Name'] + ' ' + employee['Last Name'] + ' submitted their joining form.');
    } catch(e) {
      Logger.log('Email error: ' + e.toString());
    }
    
    return {success: true, message: 'Form submitted successfully'};
    
  } catch(error) {
    Logger.log('Error submitting form: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Bulk Upload Employees
function bulkUploadEmployees(employeesData, legalEntity) {
  const results = {success: [], failed: [], total: employeesData.length};
  
  for (let i = 0; i < employeesData.length; i++) {
    try {
      const emp = employeesData[i];
      emp['Legal Entity'] = legalEntity;
      
      const result = addNewEmployee(emp);
      
      if (result.success) {
        results.success.push({row: i + 1, employeeId: result.employeeId});
      } else {
        results.failed.push({row: i + 1, error: result.error});
      }
    } catch(e) {
      results.failed.push({row: i + 1, error: e.toString()});
    }
  }
  
  return results;
}

// Send Reminder Email
function sendReminderEmail(employeeId, reminderType) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    // Find existing valid token or create new one
    let token;
    const existingToken = findRowByValue(SHEETS.TOKENS, 'Reference ID', employeeId);
    
    if (existingToken && existingToken['Used'] !== 'Yes') {
      token = existingToken['Token'];
    } else {
      token = generateToken();
      saveToken(token, 'joining', employeeId, employee['Personal Email']);
    }
    
    sendEmployeeReminderEmail(
      employee['Personal Email'],
      employee['First Name'],
      reminderType || 'form',
      token
    );
    
    return {success: true, message: 'Reminder sent'};
    
  } catch(error) {
    Logger.log('Error sending reminder: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Send Bulk Reminders
function sendBulkReminders(employeeIds, reminderType) {
  const results = {success: 0, failed: 0};
  
  for (let i = 0; i < employeeIds.length; i++) {
    const result = sendReminderEmail(employeeIds[i], reminderType);
    if (result.success) {
      results.success++;
    } else {
      results.failed++;
    }
  }
  
  return results;
}

// Get Employee for Edit Form
function getEmployeeForEdit(employeeId) {
  try {
    const employee = getEmployeeById(employeeId);
    if (!employee) {
      return {success: false, error: 'Employee not found'};
    }
    
    // Get additional data for dropdowns
    const departments = getDepartments();
    const designations = getDesignations();
    const workLocations = getWorkLocations();
    const managers = getManagers().filter(m => m.id !== employeeId); // Exclude self
    
    return {
      success: true,
      employee: employee,
      departments: departments,
      designations: designations,
      workLocations: workLocations,
      managers: managers,
      employmentTypes: EMPLOYMENT_TYPES,
      shiftTypes: SHIFT_TYPES
    };
    
  } catch(error) {
    Logger.log('Error getting employee for edit: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}
