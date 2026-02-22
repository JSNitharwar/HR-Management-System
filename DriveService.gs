/**
 * Google Drive Service - File and Folder Management
 */

// Get HR Documents Folder
function getHRFolder() {
  try {
    return DriveApp.getFolderById(CONFIG.HR_DOCUMENTS_FOLDER_ID);
  } catch(error) {
    Logger.log('Error accessing HR folder: ' + error.toString());
    throw new Error('Cannot access HR folder. Check HR_DOCUMENTS_FOLDER_ID in Config.gs');
  }
}

// Create Employee Folder
function createEmployeeFolder(employeeId, firstName, lastName) {
  try {
    const hrFolder = getHRFolder();
    
    // Get or create Employees folder
    let employeesFolder;
    const folders = hrFolder.getFoldersByName('Employees');
    if (folders.hasNext()) {
      employeesFolder = folders.next();
    } else {
      employeesFolder = hrFolder.createFolder('Employees');
    }
    
    // Create employee folder
    const folderName = employeeId + '_' + firstName + '_' + lastName;
    const employeeFolder = employeesFolder.createFolder(folderName);
    
    // Create subfolders
    employeeFolder.createFolder('Personal Documents');
    employeeFolder.createFolder('Educational Documents');
    employeeFolder.createFolder('Employment Documents');
    employeeFolder.createFolder('Other Documents');
    
    return {
      success: true,
      folderId: employeeFolder.getId(),
      folderUrl: employeeFolder.getUrl()
    };
    
  } catch(error) {
    Logger.log('Error creating employee folder: ' + error.toString());
    return {success: false, folderId: '', error: error.toString()};
  }
}

// Create Candidate Folder
function createCandidateFolder(candidateId, firstName, lastName) {
  try {
    const hrFolder = getHRFolder();
    
    // Get or create Candidates folder
    let candidatesFolder;
    const folders = hrFolder.getFoldersByName('Candidates');
    if (folders.hasNext()) {
      candidatesFolder = folders.next();
    } else {
      candidatesFolder = hrFolder.createFolder('Candidates');
    }
    
    // Create candidate folder
    const folderName = candidateId + '_' + firstName + '_' + lastName;
    const candidateFolder = candidatesFolder.createFolder(folderName);
    
    return {
      success: true,
      folderId: candidateFolder.getId(),
      folderUrl: candidateFolder.getUrl()
    };
    
  } catch(error) {
    Logger.log('Error creating candidate folder: ' + error.toString());
    return {success: false, folderId: '', error: error.toString()};
  }
}

// Upload Employee Documents
function uploadEmployeeDocuments(folderId, documents) {
  const results = [];
  
  try {
    const folder = DriveApp.getFolderById(folderId);
    
    for (let i = 0; i < documents.length; i++) {
      const doc = documents[i];
      
      try {
        // Determine subfolder
        let subfolderName = 'Other Documents';
        if (['Aadhar Card', 'PAN Card', 'Photo', 'Address Proof'].includes(doc.type)) {
          subfolderName = 'Personal Documents';
        } else if (['10th Certificate', '12th Certificate', 'Degree Certificate'].includes(doc.type)) {
          subfolderName = 'Educational Documents';
        } else if (['Experience Letter', 'Relieving Letter', 'Last 3 Payslips'].includes(doc.type)) {
          subfolderName = 'Employment Documents';
        }
        
        // Get or create subfolder
        let subfolder;
        const subfolders = folder.getFoldersByName(subfolderName);
        if (subfolders.hasNext()) {
          subfolder = subfolders.next();
        } else {
          subfolder = folder.createFolder(subfolderName);
        }
        
        // Create file from base64
        const blob = Utilities.newBlob(
          Utilities.base64Decode(doc.content),
          doc.mimeType,
          doc.type + '_' + doc.fileName
        );
        
        const file = subfolder.createFile(blob);
        
        results.push({
          success: true,
          type: doc.type,
          fileId: file.getId()
        });
        
      } catch(e) {
        results.push({
          success: false,
          type: doc.type,
          error: e.toString()
        });
      }
    }
    
  } catch(error) {
    Logger.log('Error uploading documents: ' + error.toString());
  }
  
  return results;
}

// Upload Candidate Resume
function uploadCandidateResume(folderId, resumeData, candidateId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    
    const blob = Utilities.newBlob(
      Utilities.base64Decode(resumeData.content),
      resumeData.mimeType,
      'Resume_' + candidateId + '_' + resumeData.fileName
    );
    
    const file = folder.createFile(blob);
    
    return {
      success: true,
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };
    
  } catch(error) {
    Logger.log('Error uploading resume: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get Employee Documents
function getEmployeeDocuments(folderId) {
  try {
    const folder = DriveApp.getFolderById(folderId);
    const documents = [];
    
    const subfolders = folder.getFolders();
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      const files = subfolder.getFiles();
      
      while (files.hasNext()) {
        const file = files.next();
        documents.push({
          id: file.getId(),
          name: file.getName(),
          folder: subfolder.getName(),
          url: file.getUrl(),
          mimeType: file.getMimeType()
        });
      }
    }
    
    return documents;
    
  } catch(error) {
    Logger.log('Error getting documents: ' + error.toString());
    return [];
  }
}

// Check if required documents are uploaded
function allRequiredDocumentsUploaded(folderId) {
  const documents = getEmployeeDocuments(folderId);
  const requiredDocs = ['Aadhar Card', 'PAN Card', 'Photo'];
  
  return requiredDocs.every(reqDoc => 
    documents.some(doc => doc.name.includes(reqDoc))
  );
}
