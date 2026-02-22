/**
 * Email Service - All Email Functions
 */

// Send Welcome Email to New Employee
function sendWelcomeEmail(email, firstName, token) {
  try {
    const joiningFormUrl = getWebAppUrl() + '?page=joining&token=' + token;
    
    const subject = 'Welcome to ' + CONFIG.COMPANY_NAME + ' - Complete Your Joining Formalities';
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#2196F3;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '.button{display:inline-block;padding:12px 30px;background:#4CAF50;color:white;text-decoration:none;border-radius:5px;margin:20px 0;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>Welcome to ' + CONFIG.COMPANY_NAME + '!</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + firstName + ',</p>' +
      '<p>We are delighted to welcome you to our team!</p>' +
      '<p>Please complete your joining form by clicking the button below:</p>' +
      '<p style="text-align:center;"><a href="' + joiningFormUrl + '" class="button">Complete Joining Form</a></p>' +
      '<p>Please have these documents ready:</p>' +
      '<ul><li>Passport photo</li><li>Aadhar Card</li><li>PAN Card</li><li>Educational certificates</li><li>Bank details</li></ul>' +
      '<p>Contact HR at ' + CONFIG.HR_EMAIL + ' for any questions.</p>' +
      '<p>Best regards,<br>HR Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(email, subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending welcome email: ' + error.toString());
    throw error;
  }
}

// Send Joining Form Confirmation
function sendJoiningFormConfirmationEmail(email, firstName) {
  try {
    const subject = 'Joining Form Received - ' + CONFIG.COMPANY_NAME;
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#4CAF50;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>Form Submitted Successfully!</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + firstName + ',</p>' +
      '<p>Thank you for completing your joining form.</p>' +
      '<p>Our HR team will review your submission and contact you soon.</p>' +
      '<p>Best regards,<br>HR Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(email, subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending confirmation: ' + error.toString());
  }
}

// Send Reminder Email to Employee
function sendEmployeeReminderEmail(email, firstName, reminderType, token) {
  try {
    const joiningFormUrl = getWebAppUrl() + '?page=joining&token=' + token;
    
    const subject = 'Reminder: Complete Your Joining Form - ' + CONFIG.COMPANY_NAME;
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#FF9800;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '.button{display:inline-block;padding:12px 30px;background:#2196F3;color:white;text-decoration:none;border-radius:5px;margin:20px 0;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>Reminder</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + firstName + ',</p>' +
      '<p>We noticed you haven\'t completed your joining form yet. Please complete it soon.</p>' +
      '<p style="text-align:center;"><a href="' + joiningFormUrl + '" class="button">Complete Now</a></p>' +
      '<p>Best regards,<br>HR Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(email, subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending reminder: ' + error.toString());
  }
}

// Send HR Notification
function sendHRNotification(subject, message) {
  try {
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.alert{background:#E3F2FD;border-left:4px solid #2196F3;padding:15px;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<h2>' + subject + '</h2>' +
      '<div class="alert"><p>' + message + '</p></div>' +
      '<p><a href="' + getWebAppUrl() + '">Open HR Dashboard</a></p>' +
      '</div></body></html>';
    
    GmailApp.sendEmail(CONFIG.HR_EMAIL, '[HR System] ' + subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending HR notification: ' + error.toString());
  }
}

// Send Application Confirmation to Candidate
function sendApplicationConfirmationEmail(email, firstName, jobId) {
  try {
    const job = getJobDetails(jobId);
    const jobTitle = job ? job['Job Title'] : 'the position';
    
    const subject = 'Application Received - ' + CONFIG.COMPANY_NAME;
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#673AB7;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>Application Received!</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + firstName + ',</p>' +
      '<p>Thank you for applying for <strong>' + jobTitle + '</strong> at ' + CONFIG.COMPANY_NAME + '.</p>' +
      '<p>We will review your application and contact you soon.</p>' +
      '<p>Best regards,<br>Talent Acquisition Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(email, subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending application confirmation: ' + error.toString());
  }
}

// Send Candidate Invite Email
function sendCandidateInviteEmail(email, firstName, jobId, token) {
  try {
    const applicationUrl = getWebAppUrl() + '?page=apply&jobId=' + jobId;
    const job = getJobDetails(jobId);
    const jobTitle = job ? job['Job Title'] : 'the position';
    
    const subject = 'You\'re Invited to Apply - ' + CONFIG.COMPANY_NAME;
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#9C27B0;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '.button{display:inline-block;padding:12px 30px;background:#4CAF50;color:white;text-decoration:none;border-radius:5px;margin:20px 0;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>You\'re Invited!</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + firstName + ',</p>' +
      '<p>We believe you could be a great fit for <strong>' + jobTitle + '</strong> at ' + CONFIG.COMPANY_NAME + '.</p>' +
      '<p style="text-align:center;"><a href="' + applicationUrl + '" class="button">Apply Now</a></p>' +
      '<p>Best regards,<br>Talent Acquisition Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(email, subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending invite: ' + error.toString());
  }
}

// Send Interview Invite to Candidate
function sendInterviewInviteToCandidate(candidate, interview, job) {
  try {
    const subject = 'Interview Scheduled - ' + (job ? job['Job Title'] : 'Position') + ' at ' + CONFIG.COMPANY_NAME;
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#00BCD4;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '.details{background:white;padding:15px;border-radius:5px;margin:15px 0;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>Interview Scheduled!</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + candidate['First Name'] + ',</p>' +
      '<p>You have been scheduled for an interview.</p>' +
      '<div class="details">' +
      '<p><strong>Date:</strong> ' + interview['Interview Date'] + '</p>' +
      '<p><strong>Time:</strong> ' + interview['Interview Time'] + '</p>' +
      '<p><strong>Duration:</strong> ' + interview['Duration'] + '</p>' +
      '<p><strong>Type:</strong> ' + interview['Interview Type'] + '</p>' +
      '<p><strong>Interviewer:</strong> ' + interview['Interviewer Name'] + '</p>' +
      (interview['Meeting Link'] ? '<p><strong>Meeting Link:</strong> <a href="' + interview['Meeting Link'] + '">' + interview['Meeting Link'] + '</a></p>' : '') +
      '</div>' +
      '<p>Best of luck!</p>' +
      '<p>Best regards,<br>Talent Acquisition Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(candidate['Email'], subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending interview invite to candidate: ' + error.toString());
  }
}

// Send Interview Invite to Interviewer
function sendInterviewInviteToInterviewer(interview, candidate, job) {
  try {
    const subject = 'Interview Assignment: ' + candidate['First Name'] + ' ' + candidate['Last Name'];
    
    const htmlBody = '<!DOCTYPE html><html><head><style>' +
      'body{font-family:Arial,sans-serif;line-height:1.6;color:#333;}' +
      '.container{max-width:600px;margin:0 auto;padding:20px;}' +
      '.header{background:#607D8B;color:white;padding:20px;text-align:center;}' +
      '.content{padding:20px;background:#f9f9f9;}' +
      '.details{background:white;padding:15px;border-radius:5px;margin:15px 0;}' +
      '</style></head><body>' +
      '<div class="container">' +
      '<div class="header"><h1>Interview Assignment</h1></div>' +
      '<div class="content">' +
      '<p>Dear ' + interview['Interviewer Name'] + ',</p>' +
      '<p>You have been assigned to conduct an interview.</p>' +
      '<div class="details">' +
      '<h3>Interview Details</h3>' +
      '<p><strong>Candidate:</strong> ' + candidate['First Name'] + ' ' + candidate['Last Name'] + '</p>' +
      '<p><strong>Position:</strong> ' + (job ? job['Job Title'] : 'N/A') + '</p>' +
      '<p><strong>Date:</strong> ' + interview['Interview Date'] + '</p>' +
      '<p><strong>Time:</strong> ' + interview['Interview Time'] + '</p>' +
      '<p><strong>Duration:</strong> ' + interview['Duration'] + '</p>' +
      (interview['Meeting Link'] ? '<p><strong>Meeting Link:</strong> <a href="' + interview['Meeting Link'] + '">' + interview['Meeting Link'] + '</a></p>' : '') +
      '</div>' +
      '<div class="details">' +
      '<h3>Candidate Info</h3>' +
      '<p><strong>Email:</strong> ' + candidate['Email'] + '</p>' +
      '<p><strong>Phone:</strong> ' + candidate['Phone'] + '</p>' +
      '<p><strong>Experience:</strong> ' + (candidate['Total Experience'] || 'N/A') + '</p>' +
      '</div>' +
      '<p>Best regards,<br>HR Team</p>' +
      '</div></div></body></html>';
    
    GmailApp.sendEmail(interview['Interviewer Email'], subject, '', {htmlBody: htmlBody});
    
  } catch(error) {
    Logger.log('Error sending interview invite to interviewer: ' + error.toString());
  }
}
