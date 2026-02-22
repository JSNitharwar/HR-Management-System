/**
 * Offer Letter Management Service
 */

// Create Offer Letter
function createOffer(offerData) {
  try {
    const offerId = generateId('OFR');
    const now = new Date();
    const acceptanceToken = generateToken();
    
    const candidate = getCandidateById(offerData.candidateId);
    if (!candidate) {
      return {success: false, error: 'Candidate not found'};
    }
    
    const job = getJobDetails(offerData.jobId || candidate['Job ID']);
    
    // Calculate monthly from annual
    const ctcAnnual = parseFloat(offerData.ctcAnnual) || 0;
    const ctcMonthly = Math.round((ctcAnnual / 12) * 100) / 100;
    
    const offer = {
      'Offer ID': offerId,
      'Candidate ID': offerData.candidateId,
      'Job ID': offerData.jobId || candidate['Job ID'],
      'Legal Entity': offerData.legalEntity || (job ? job['Legal Entity'] : ''),
      'Designation': offerData.designation || (job ? job['Job Title'] : ''),
      'Department': offerData.department || (job ? job['Department'] : ''),
      'Work Location': offerData.workLocation || (job ? job['Location'] : ''),
      'Employment Type': offerData.employmentType || 'Full-time',
      'Reporting Manager': offerData.reportingManager || '',
      'Joining Date': offerData.joiningDate || '',
      'Probation Period': offerData.probationPeriod || '6 months',
      'CTC Annual': ctcAnnual,
      'CTC Monthly': ctcMonthly,
      'Basic Salary': parseFloat(offerData.basicSalary) || 0,
      'HRA': parseFloat(offerData.hra) || 0,
      'Other Allowances': parseFloat(offerData.otherAllowances) || 0,
      'Offer Date': now,
      'Validity Date': offerData.validityDate || '',
      'Acceptance Deadline': offerData.acceptanceDeadline || '',
      'Status': OFFER_STATUS.DRAFT,
      'Sent Date': '',
      'Response Date': '',
      'Response Notes': '',
      'Offer Letter URL': '',
      'Acceptance Token': acceptanceToken,
      'Created By': Session.getActiveUser().getEmail() || 'System',
      'Created Date': now
    };
    
    addRowToSheet(SHEETS.OFFERS, offer);
    
    // Save acceptance token
    saveToken(acceptanceToken, 'offer_response', offerId, candidate['Email']);
    
    logAction('CREATE', 'Offer', offerId, '', JSON.stringify(offer));
    
    return {
      success: true, 
      offerId: offerId, 
      message: 'Offer created successfully'
    };
    
  } catch(error) {
    Logger.log('Error creating offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get All Offers
function getAllOffers(filters) {
  try {
    filters = filters || {};
    let offers = getSheetData(SHEETS.OFFERS);
    
    if (filters.status && filters.status !== '') {
      offers = offers.filter(o => o['Status'] === filters.status);
    }
    if (filters.candidateId && filters.candidateId !== '') {
      offers = offers.filter(o => o['Candidate ID'] === filters.candidateId);
    }
    
    // Add candidate details
    const candidates = getSheetData(SHEETS.CANDIDATES);
    offers = offers.map(o => {
      const candidate = candidates.find(c => c['Candidate ID'] === o['Candidate ID']);
      o.candidateName = candidate ? (candidate['First Name'] + ' ' + candidate['Last Name']) : 'Unknown';
      o.candidateEmail = candidate ? candidate['Email'] : '';
      o.candidatePhone = candidate ? candidate['Phone'] : '';
      return o;
    });
    
    return offers;
    
  } catch(error) {
    Logger.log('Error getting offers: ' + error.toString());
    return [];
  }
}

// Get Offer by ID
function getOfferById(offerId) {
  try {
    const offer = findRowByValue(SHEETS.OFFERS, 'Offer ID', offerId);
    if (offer) {
      const candidate = getCandidateById(offer['Candidate ID']);
      offer.candidateName = candidate ? (candidate['First Name'] + ' ' + candidate['Last Name']) : 'Unknown';
      offer.candidateEmail = candidate ? candidate['Email'] : '';
    }
    return offer;
  } catch(error) {
    Logger.log('Error getting offer: ' + error.toString());
    return null;
  }
}

// Get Offer for Candidate
function getOfferForCandidate(candidateId) {
  try {
    const offers = findRowsByValue(SHEETS.OFFERS, 'Candidate ID', candidateId);
    return offers.sort((a, b) => new Date(b['Created Date']) - new Date(a['Created Date']))[0] || null;
  } catch(error) {
    return null;
  }
}

// Update Offer
function updateOffer(offerId, updateData) {
  try {
    const offer = findRowByValue(SHEETS.OFFERS, 'Offer ID', offerId);
    if (!offer) {
      return {success: false, error: 'Offer not found'};
    }
    
    updateRowInSheet(SHEETS.OFFERS, offer.rowIndex, updateData);
    
    return {success: true, message: 'Offer updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Send Offer Letter
function sendOfferLetter(offerId) {
  try {
    const offer = getOfferById(offerId);
    if (!offer) {
      return {success: false, error: 'Offer not found'};
    }
    
    const candidate = getCandidateById(offer['Candidate ID']);
    if (!candidate) {
      return {success: false, error: 'Candidate not found'};
    }
    
    // Generate offer letter PDF
    const pdfResult = generateOfferLetterPDF(offer, candidate);
    
    // Send email with offer letter
    sendOfferLetterEmail(candidate, offer, pdfResult.pdfUrl, offer['Acceptance Token']);
    
    // Update offer status
    const now = new Date();
    updateRowInSheet(SHEETS.OFFERS, offer.rowIndex, {
      'Status': OFFER_STATUS.SENT,
      'Sent Date': now,
      'Offer Letter URL': pdfResult.pdfUrl || ''
    });
    
    // Update candidate status
    updateCandidateStatus(offer['Candidate ID'], CANDIDATE_STATUS.OFFER_SENT);
    
    logAction('SEND_OFFER', 'Offer', offerId, '', 'Offer sent to ' + candidate['Email']);
    
    return {
      success: true, 
      message: 'Offer letter sent successfully to ' + candidate['Email']
    };
    
  } catch(error) {
    Logger.log('Error sending offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Generate Offer Letter PDF
function generateOfferLetterPDF(offer, candidate) {
  try {
    const html = generateOfferLetterHTML(offer, candidate);
    
    // Create a temporary HTML file
    const blob = Utilities.newBlob(html, 'text/html', 'offer_letter.html');
    
    // For now, we'll store the HTML as a document
    // In production, you might want to use a PDF service
    const folder = DriveApp.getFolderById(CONFIG.HR_DOCUMENTS_FOLDER_ID);
    
    let offersFolder;
    const folderIterator = folder.getFoldersByName('Offer Letters');
    if (folderIterator.hasNext()) {
      offersFolder = folderIterator.next();
    } else {
      offersFolder = folder.createFolder('Offer Letters');
    }
    
    const fileName = 'Offer_' + offer['Offer ID'] + '_' + candidate['First Name'] + '_' + candidate['Last Name'] + '.html';
    const file = offersFolder.createFile(blob.setName(fileName));
    
    return {
      success: true,
      pdfUrl: file.getUrl(),
      fileId: file.getId()
    };
    
  } catch(error) {
    Logger.log('Error generating PDF: ' + error.toString());
    return {success: false, pdfUrl: '', error: error.toString()};
  }
}

// Generate Offer Letter HTML
function generateOfferLetterHTML(offer, candidate) {
  const joiningDate = offer['Joining Date'] ? new Date(offer['Joining Date']).toLocaleDateString('en-IN', {day: 'numeric', month: 'long', year: 'numeric'}) : '[Date]';
  const offerDate = offer['Offer Date'] ? new Date(offer['Offer Date']).toLocaleDateString('en-IN', {day: 'numeric', month: 'long', year: 'numeric'}) : new Date().toLocaleDateString('en-IN', {day: 'numeric', month: 'long', year: 'numeric'});
  const validityDate = offer['Validity Date'] ? new Date(offer['Validity Date']).toLocaleDateString('en-IN', {day: 'numeric', month: 'long', year: 'numeric'}) : '';
  
  const ctcAnnual = parseFloat(offer['CTC Annual']) || 0;
  const ctcMonthly = parseFloat(offer['CTC Monthly']) || 0;
  
  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <title>Offer Letter - ${candidate['First Name']} ${candidate['Last Name']}</title>
  <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    
    * { margin: 0; padding: 0; box-sizing: border-box; }
    
    body {
      font-family: 'Roboto', Arial, sans-serif;
      font-size: 14px;
      line-height: 1.6;
      color: #333;
      background: #f5f5f5;
      padding: 20px;
    }
    
    .letter-container {
      max-width: 800px;
      margin: 0 auto;
      background: white;
      box-shadow: 0 5px 30px rgba(0,0,0,0.1);
    }
    
    .header {
      background: linear-gradient(135deg, #1a237e, #3949ab);
      color: white;
      padding: 40px;
      text-align: center;
    }
    
    .header h1 {
      font-size: 28px;
      font-weight: 700;
      margin-bottom: 10px;
    }
    
    .header p {
      font-size: 13px;
      opacity: 0.9;
    }
    
    .content {
      padding: 50px;
    }
    
    .date-ref {
      display: flex;
      justify-content: space-between;
      margin-bottom: 30px;
      font-size: 13px;
      color: #666;
    }
    
    .recipient {
      margin-bottom: 30px;
    }
    
    .recipient strong {
      font-size: 16px;
      color: #1a237e;
    }
    
    .subject {
      background: #e8eaf6;
      padding: 15px 20px;
      margin-bottom: 30px;
      border-left: 4px solid #3949ab;
      font-weight: 500;
    }
    
    .greeting {
      margin-bottom: 20px;
    }
    
    .body-text {
      margin-bottom: 20px;
      text-align: justify;
    }
    
    .highlight-box {
      background: linear-gradient(135deg, #e3f2fd, #e8eaf6);
      border-radius: 10px;
      padding: 25px;
      margin: 30px 0;
    }
    
    .highlight-box h3 {
      color: #1a237e;
      margin-bottom: 20px;
      font-size: 16px;
      border-bottom: 2px solid #3949ab;
      padding-bottom: 10px;
    }
    
    .detail-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 15px;
    }
    
    .detail-item {
      display: flex;
      flex-direction: column;
    }
    
    .detail-label {
      font-size: 11px;
      color: #666;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      margin-bottom: 3px;
    }
    
    .detail-value {
      font-size: 15px;
      font-weight: 500;
      color: #1a237e;
    }
    
    .compensation-box {
      background: linear-gradient(135deg, #1a237e, #3949ab);
      color: white;
      border-radius: 10px;
      padding: 25px;
      margin: 30px 0;
      text-align: center;
    }
    
    .compensation-box h3 {
      font-size: 14px;
      opacity: 0.9;
      margin-bottom: 10px;
    }
    
    .compensation-amount {
      font-size: 32px;
      font-weight: 700;
    }
    
    .compensation-subtext {
      font-size: 12px;
      opacity: 0.8;
      margin-top: 5px;
    }
    
    .terms-list {
      margin: 20px 0;
      padding-left: 20px;
    }
    
    .terms-list li {
      margin-bottom: 10px;
    }
    
    .signature-section {
      margin-top: 50px;
      display: flex;
      justify-content: space-between;
    }
    
    .signature-box {
      width: 45%;
    }
    
    .signature-line {
      border-top: 1px solid #333;
      margin-top: 60px;
      padding-top: 10px;
    }
    
    .footer {
      background: #f5f5f5;
      padding: 20px 50px;
      text-align: center;
      font-size: 12px;
      color: #666;
      border-top: 1px solid #e0e0e0;
    }
    
    .acceptance-section {
      background: #fff3e0;
      border: 2px dashed #ff9800;
      border-radius: 10px;
      padding: 25px;
      margin: 30px 0;
      text-align: center;
    }
    
    .acceptance-section h4 {
      color: #e65100;
      margin-bottom: 10px;
    }
    
    @media print {
      body { background: white; padding: 0; }
      .letter-container { box-shadow: none; }
    }
  </style>
</head>
<body>
  <div class="letter-container">
    <div class="header">
      <h1>${CONFIG.COMPANY_NAME}</h1>
      <p>${CONFIG.COMPANY_ADDRESS}</p>
      <p>${CONFIG.COMPANY_PHONE} | ${CONFIG.COMPANY_WEBSITE}</p>
    </div>
    
    <div class="content">
      <div class="date-ref">
        <span><strong>Date:</strong> ${offerDate}</span>
        <span><strong>Ref:</strong> ${offer['Offer ID']}</span>
      </div>
      
      <div class="recipient">
        <strong>${candidate['First Name']} ${candidate['Last Name']}</strong><br>
        ${candidate['Email']}<br>
        ${candidate['Phone']}
      </div>
      
      <div class="subject">
        Subject: Offer of Employment - ${offer['Designation']}
      </div>
      
      <p class="greeting">Dear <strong>${candidate['First Name']}</strong>,</p>
      
      <p class="body-text">
        We are delighted to offer you the position of <strong>${offer['Designation']}</strong> 
        at <strong>${CONFIG.COMPANY_NAME}</strong>. After careful consideration of your qualifications 
        and interview performance, we believe you will be a valuable addition to our team.
      </p>
      
      <div class="highlight-box">
        <h3>📋 Employment Details</h3>
        <div class="detail-grid">
          <div class="detail-item">
            <span class="detail-label">Position</span>
            <span class="detail-value">${offer['Designation']}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Department</span>
            <span class="detail-value">${offer['Department']}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Work Location</span>
            <span class="detail-value">${offer['Work Location']}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Employment Type</span>
            <span class="detail-value">${offer['Employment Type']}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Reporting To</span>
            <span class="detail-value">${offer['Reporting Manager'] || 'To be assigned'}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Joining Date</span>
            <span class="detail-value">${joiningDate}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Probation Period</span>
            <span class="detail-value">${offer['Probation Period']}</span>
          </div>
          <div class="detail-item">
            <span class="detail-label">Legal Entity</span>
            <span class="detail-value">${offer['Legal Entity']}</span>
          </div>
        </div>
      </div>
      
      <div class="compensation-box">
        <h3>💰 ANNUAL COMPENSATION</h3>
        <div class="compensation-amount">₹ ${ctcAnnual.toLocaleString('en-IN')}</div>
        <div class="compensation-subtext">(Rupees ${numberToWords(ctcAnnual)} Only)</div>
        <div class="compensation-subtext" style="margin-top:15px;">Monthly CTC: ₹ ${ctcMonthly.toLocaleString('en-IN')}</div>
      </div>
      
      <p class="body-text">
        The detailed salary breakup will be provided along with your appointment letter upon joining. 
        Your compensation includes statutory benefits as per applicable laws.
      </p>
      
      <p class="body-text"><strong>Terms and Conditions:</strong></p>
      <ul class="terms-list">
        <li>This offer is subject to successful completion of background verification.</li>
        <li>You will be required to submit all original educational and employment documents for verification.</li>
        <li>The probation period is ${offer['Probation Period']}, during which your performance will be evaluated.</li>
        <li>You are required to maintain confidentiality of all company information.</li>
        <li>This offer is valid until <strong>${validityDate || 'as communicated'}</strong>.</li>
      </ul>
      
      <div class="acceptance-section">
        <h4>📝 OFFER ACCEPTANCE</h4>
        <p>Please confirm your acceptance by responding to the HR team<br>or signing and returning a copy of this letter.</p>
        ${validityDate ? '<p><strong>Response Deadline: ' + validityDate + '</strong></p>' : ''}
      </div>
      
      <p class="body-text">
        We are excited about the possibility of you joining our team and look forward to your positive response. 
        If you have any questions, please don't hesitate to contact us.
      </p>
      
      <p class="body-text">Welcome to ${CONFIG.COMPANY_NAME}!</p>
      
      <div class="signature-section">
        <div class="signature-box">
          <div class="signature-line">
            <strong>For ${CONFIG.COMPANY_NAME}</strong><br>
            HR Department
          </div>
        </div>
        <div class="signature-box">
          <div class="signature-line">
            <strong>Acceptance Signature</strong><br>
            ${candidate['First Name']} ${candidate['Last Name']}
          </div>
        </div>
      </div>
    </div>
    
    <div class="footer">
      <p>This is a computer-generated document. For any queries, contact HR at ${CONFIG.HR_EMAIL}</p>
      <p>${CONFIG.COMPANY_NAME} | ${CONFIG.COMPANY_ADDRESS}</p>
    </div>
  </div>
</body>
</html>
  `;
}

// Send Offer Letter Email
function sendOfferLetterEmail(candidate, offer, pdfUrl, acceptanceToken) {
  const acceptanceUrl = getWebAppUrl() + '?page=offer_response&token=' + acceptanceToken;
  
  const joiningDate = offer['Joining Date'] ? new Date(offer['Joining Date']).toLocaleDateString('en-IN', {day: 'numeric', month: 'long', year: 'numeric'}) : 'As discussed';
  const ctcAnnual = parseFloat(offer['CTC Annual']) || 0;
  
  const subject = `Congratulations! Offer Letter from ${CONFIG.COMPANY_NAME} - ${offer['Designation']}`;
  
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; }
    .container { max-width: 600px; margin: 0 auto; }
    .header { background: linear-gradient(135deg, #1a237e, #3949ab); color: white; padding: 40px; text-align: center; }
    .header h1 { margin: 0 0 10px 0; font-size: 24px; }
    .content { padding: 40px; background: #f9f9f9; }
    .highlight { background: linear-gradient(135deg, #e3f2fd, #e8eaf6); padding: 25px; border-radius: 10px; margin: 25px 0; }
    .highlight h3 { color: #1a237e; margin: 0 0 15px 0; }
    .detail-row { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid #e0e0e0; }
    .salary-box { background: linear-gradient(135deg, #1a237e, #3949ab); color: white; padding: 25px; border-radius: 10px; text-align: center; margin: 25px 0; }
    .salary-amount { font-size: 28px; font-weight: bold; }
    .btn { display: inline-block; padding: 15px 40px; background: #4CAF50; color: white; text-decoration: none; border-radius: 8px; font-weight: bold; margin: 10px 5px; }
    .btn-secondary { background: #1976d2; }
    .btn-container { text-align: center; margin: 30px 0; }
    .footer { padding: 25px; text-align: center; font-size: 12px; color: #666; background: #f0f0f0; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🎉 Congratulations!</h1>
      <p>You have been selected to join ${CONFIG.COMPANY_NAME}</p>
    </div>
    
    <div class="content">
      <p>Dear <strong>${candidate['First Name']}</strong>,</p>
      
      <p>We are thrilled to offer you the position of <strong>${offer['Designation']}</strong> at ${CONFIG.COMPANY_NAME}. Your skills and experience impressed us, and we believe you'll be a valuable addition to our team.</p>
      
      <div class="highlight">
        <h3>📋 Your Offer Summary</h3>
        <div class="detail-row"><span>Position:</span><strong>${offer['Designation']}</strong></div>
        <div class="detail-row"><span>Department:</span><strong>${offer['Department']}</strong></div>
        <div class="detail-row"><span>Location:</span><strong>${offer['Work Location']}</strong></div>
        <div class="detail-row"><span>Joining Date:</span><strong>${joiningDate}</strong></div>
      </div>
      
      <div class="salary-box">
        <p style="margin:0 0 10px 0; font-size: 14px;">Annual Compensation (CTC)</p>
        <div class="salary-amount">₹ ${ctcAnnual.toLocaleString('en-IN')}</div>
      </div>
      
      <p>Please find the detailed offer letter attached to this email. We request you to review the terms and conditions carefully.</p>
      
      <div class="btn-container">
        ${pdfUrl ? `<a href="${pdfUrl}" class="btn btn-secondary">📄 View Offer Letter</a>` : ''}
      </div>
      
      <p><strong>Next Steps:</strong></p>
      <ol>
        <li>Review the attached offer letter carefully</li>
        <li>Respond to confirm your acceptance</li>
        <li>Submit required documents before joining</li>
      </ol>
      
      ${offer['Validity Date'] ? `<p style="background:#fff3e0;padding:15px;border-radius:5px;"><strong>⏰ Please respond by ${new Date(offer['Validity Date']).toLocaleDateString('en-IN', {day: 'numeric', month: 'long', year: 'numeric'})}</strong></p>` : ''}
      
      <p>If you have any questions, please don't hesitate to reach out to our HR team.</p>
      
      <p>We look forward to welcoming you to ${CONFIG.COMPANY_NAME}!</p>
      
      <p>Best regards,<br><strong>HR Team</strong><br>${CONFIG.COMPANY_NAME}</p>
    </div>
    
    <div class="footer">
      <p>${CONFIG.COMPANY_NAME} | ${CONFIG.COMPANY_ADDRESS}</p>
      <p>📧 ${CONFIG.HR_EMAIL} | 📞 ${CONFIG.COMPANY_PHONE}</p>
    </div>
  </div>
</body>
</html>
  `;
  
  try {
    GmailApp.sendEmail(candidate['Email'], subject, '', {
      htmlBody: htmlBody,
      name: CONFIG.COMPANY_NAME + ' HR',
      replyTo: CONFIG.HR_EMAIL
    });
  } catch(error) {
    Logger.log('Error sending offer email: ' + error.toString());
    throw error;
  }
}

// Convert number to words (Indian format)
function numberToWords(num) {
  const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten',
    'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
  
  if (num === 0) return 'Zero';
  
  function convertLessThanThousand(n) {
    if (n === 0) return '';
    if (n < 20) return ones[n];
    if (n < 100) return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
    return ones[Math.floor(n / 100)] + ' Hundred' + (n % 100 ? ' ' + convertLessThanThousand(n % 100) : '');
  }
  
  let result = '';
  
  if (num >= 10000000) {
    result += convertLessThanThousand(Math.floor(num / 10000000)) + ' Crore ';
    num %= 10000000;
  }
  if (num >= 100000) {
    result += convertLessThanThousand(Math.floor(num / 100000)) + ' Lakh ';
    num %= 100000;
  }
  if (num >= 1000) {
    result += convertLessThanThousand(Math.floor(num / 1000)) + ' Thousand ';
    num %= 1000;
  }
  if (num > 0) {
    result += convertLessThanThousand(num);
  }
  
  return result.trim();
}

// Withdraw Offer
function withdrawOffer(offerId, reason) {
  try {
    const offer = getOfferById(offerId);
    if (!offer) {
      return {success: false, error: 'Offer not found'};
    }
    
    updateRowInSheet(SHEETS.OFFERS, offer.rowIndex, {
      'Status': OFFER_STATUS.WITHDRAWN,
      'Response Date': new Date(),
      'Response Notes': 'Withdrawn: ' + (reason || 'No reason provided')
    });
    
    // Update candidate status back to Selected
    updateCandidateStatus(offer['Candidate ID'], CANDIDATE_STATUS.SELECTED);
    
    return {success: true, message: 'Offer withdrawn successfully'};
    
  } catch(error) {
    Logger.log('Error withdrawing offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Mark Offer as Accepted
function acceptOffer(offerId, notes) {
  try {
    const offer = getOfferById(offerId);
    if (!offer) {
      return {success: false, error: 'Offer not found'};
    }
    
    updateRowInSheet(SHEETS.OFFERS, offer.rowIndex, {
      'Status': OFFER_STATUS.ACCEPTED,
      'Response Date': new Date(),
      'Response Notes': notes || 'Accepted by candidate'
    });
    
    // Update candidate status
    updateCandidateStatus(offer['Candidate ID'], CANDIDATE_STATUS.OFFER_ACCEPTED);
    
    // Notify HR
    sendHRNotification('Offer Accepted', 
      offer.candidateName + ' has accepted the offer for ' + offer['Designation']);
    
    return {success: true, message: 'Offer marked as accepted'};
    
  } catch(error) {
    Logger.log('Error accepting offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Mark Offer as Rejected
function rejectOffer(offerId, notes) {
  try {
    const offer = getOfferById(offerId);
    if (!offer) {
      return {success: false, error: 'Offer not found'};
    }
    
    updateRowInSheet(SHEETS.OFFERS, offer.rowIndex, {
      'Status': OFFER_STATUS.REJECTED,
      'Response Date': new Date(),
      'Response Notes': notes || 'Rejected by candidate'
    });
    
    // Update candidate status
    updateCandidateStatus(offer['Candidate ID'], CANDIDATE_STATUS.OFFER_REJECTED);
    
    // Notify HR
    sendHRNotification('Offer Rejected', 
      offer.candidateName + ' has rejected the offer for ' + offer['Designation']);
    
    return {success: true, message: 'Offer marked as rejected'};
    
  } catch(error) {
    Logger.log('Error rejecting offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get Candidate for Offer (with job details)
function getCandidateForOffer(candidateId) {
  try {
    const candidate = getCandidateById(candidateId);
    if (!candidate) {
      return {success: false, error: 'Candidate not found'};
    }
    
    const job = getJobDetails(candidate['Job ID']);
    const departments = getDepartments();
    const designations = getDesignations();
    const workLocations = getWorkLocations();
    const managers = getManagers();
    const entities = getLegalEntities();
    
    return {
      success: true,
      candidate: candidate,
      job: job,
      departments: departments,
      designations: designations,
      workLocations: workLocations,
      managers: managers,
      entities: entities
    };
    
  } catch(error) {
    Logger.log('Error getting candidate for offer: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Convert Offer to Employee
function convertOfferToEmployee(offerId) {
  try {
    const offer = getOfferById(offerId);
    if (!offer) {
      return {success: false, error: 'Offer not found'};
    }
    
    if (offer['Status'] !== OFFER_STATUS.ACCEPTED) {
      return {success: false, error: 'Only accepted offers can be converted to employees'};
    }
    
    const candidate = getCandidateById(offer['Candidate ID']);
    if (!candidate) {
      return {success: false, error: 'Candidate not found'};
    }
    
    // Create employee from offer data
    const employeeData = {
      'Legal Entity': offer['Legal Entity'],
      'First Name': candidate['First Name'],
      'Last Name': candidate['Last Name'],
      'Personal Email': candidate['Email'],
      'Phone': candidate['Phone'],
      'Department': offer['Department'],
      'Designation': offer['Designation'],
      'Date of Joining': offer['Joining Date'],
      'Reporting Manager': offer['Reporting Manager'],
      'Employment Type': offer['Employment Type'],
      'Work Location': offer['Work Location'],
      'CTC Annual': offer['CTC Annual'],
      'CTC Monthly': offer['CTC Monthly'],
      'Basic Salary': offer['Basic Salary'],
      'HRA': offer['HRA'],
      'Other Allowances': offer['Other Allowances']
    };
    
    const result = addNewEmployee(employeeData);
    
    if (result.success) {
      // Update candidate status
      updateCandidateStatus(offer['Candidate ID'], CANDIDATE_STATUS.JOINED);
      
      return {
        success: true, 
        employeeId: result.employeeId,
        message: 'Employee created successfully from offer'
      };
    }
    
    return result;
    
  } catch(error) {
    Logger.log('Error converting offer to employee: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}
