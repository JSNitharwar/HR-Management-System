/**
 * Recruitment Management Service
 */

// Create Job Posting
function createJobPosting(jobData) {
  try {
    const jobId = generateId('JOB');
    const now = new Date();
    
    const job = {
      'Job ID': jobId,
      'Legal Entity': jobData.legalEntity || '',
      'Job Title': jobData.jobTitle || '',
      'Department': jobData.department || '',
      'Description': jobData.description || '',
      'Requirements': jobData.requirements || '',
      'Experience Required': jobData.experienceRequired || '',
      'Location': jobData.location || '',
      'Employment Type': jobData.employmentType || 'Full-time',
      'Salary Range': jobData.salaryRange || '',
      'Number of Positions': jobData.numberOfPositions || 1,
      'Status': 'Open',
      'Posted Date': now,
      'Closing Date': jobData.closingDate || '',
      'Created By': Session.getActiveUser().getEmail() || 'System'
    };
    
    addRowToSheet(SHEETS.JOBS, job);
    logAction('CREATE', 'Job', jobId, '', JSON.stringify(job));
    
    return {success: true, jobId: jobId, message: 'Job posted successfully'};
    
  } catch(error) {
    Logger.log('Error creating job: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get All Jobs
function getAllJobs(filters) {
  try {
    filters = filters || {};
    let jobs = getSheetData(SHEETS.JOBS);
    
    if (filters.status && filters.status !== '') {
      jobs = jobs.filter(j => j['Status'] === filters.status);
    }
    if (filters.legalEntity && filters.legalEntity !== '') {
      jobs = jobs.filter(j => j['Legal Entity'] === filters.legalEntity);
    }
    if (filters.department && filters.department !== '') {
      jobs = jobs.filter(j => j['Department'] === filters.department);
    }
    
    return jobs;
    
  } catch(error) {
    Logger.log('Error getting jobs: ' + error.toString());
    return [];
  }
}

// Get Open Jobs
function getOpenJobs() {
  try {
    const jobs = getSheetData(SHEETS.JOBS);
    return jobs.filter(j => j['Status'] === 'Open');
  } catch(error) {
    Logger.log('Error getting open jobs: ' + error.toString());
    return [];
  }
}

// Get Job Details
function getJobDetails(jobId) {
  try {
    return findRowByValue(SHEETS.JOBS, 'Job ID', jobId);
  } catch(error) {
    Logger.log('Error getting job: ' + error.toString());
    return null;
  }
}

// Update Job
function updateJob(jobId, updateData) {
  try {
    const job = getJobDetails(jobId);
    if (!job) {
      return {success: false, error: 'Job not found'};
    }
    
    updateRowInSheet(SHEETS.JOBS, job.rowIndex, updateData);
    
    return {success: true, message: 'Job updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating job: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Update Job Status
function updateJobStatus(jobId, newStatus) {
  try {
    const job = getJobDetails(jobId);
    if (!job) {
      return {success: false, error: 'Job not found'};
    }
    
    updateRowInSheet(SHEETS.JOBS, job.rowIndex, {'Status': newStatus});
    
    return {success: true, message: 'Job status updated'};
    
  } catch(error) {
    Logger.log('Error updating job status: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get All Candidates
function getAllCandidates(filters) {
  try {
    filters = filters || {};
    let candidates = getSheetData(SHEETS.CANDIDATES);
    
    if (filters.jobId && filters.jobId !== '') {
      candidates = candidates.filter(c => c['Job ID'] === filters.jobId);
    }
    if (filters.status && filters.status !== '') {
      candidates = candidates.filter(c => c['Status'] === filters.status);
    }
    
    // Add job title to each candidate
    const jobs = getSheetData(SHEETS.JOBS);
    candidates = candidates.map(c => {
      const job = jobs.find(j => j['Job ID'] === c['Job ID']);
      c.jobTitle = job ? job['Job Title'] : 'Unknown';
      return c;
    });
    
    return candidates;
    
  } catch(error) {
    Logger.log('Error getting candidates: ' + error.toString());
    return [];
  }
}

// Get Candidate by ID
function getCandidateById(candidateId) {
  try {
    return findRowByValue(SHEETS.CANDIDATES, 'Candidate ID', candidateId);
  } catch(error) {
    Logger.log('Error getting candidate: ' + error.toString());
    return null;
  }
}

// Submit Candidate Application
function submitCandidateApplication(applicationData) {
  try {
    const candidateId = generateId('CAND');
    const now = new Date();
    
    // Create folder for resume
    let folderId = '';
    try {
      const folderResult = createCandidateFolder(candidateId, applicationData.firstName || '', applicationData.lastName || '');
      folderId = folderResult.folderId || '';
    } catch(e) {
      Logger.log('Could not create folder: ' + e.toString());
    }
    
    const candidate = {
      'Candidate ID': candidateId,
      'Job ID': applicationData.jobId || '',
      'First Name': applicationData.firstName || '',
      'Last Name': applicationData.lastName || '',
      'Email': applicationData.email || '',
      'Phone': applicationData.phone || '',
      'Current Company': applicationData.currentCompany || '',
      'Current Designation': applicationData.currentDesignation || '',
      'Total Experience': applicationData.totalExperience || '',
      'Current CTC': applicationData.currentCtc || '',
      'Expected CTC': applicationData.expectedCtc || '',
      'Notice Period': applicationData.noticePeriod || '',
      'Location': applicationData.location || '',
      'Resume Folder ID': folderId,
      'Application Date': now,
      'Source': applicationData.source || 'Career Portal',
      'Status': CANDIDATE_STATUS.APPLIED,
      'Screening Score': '',
      'HR Notes': '',
      'Created Date': now,
      'Modified Date': now
    };
    
    addRowToSheet(SHEETS.CANDIDATES, candidate);
    
    // Upload resume
    if (applicationData.resume && folderId) {
      try {
        uploadCandidateResume(folderId, applicationData.resume, candidateId);
      } catch(e) {
        Logger.log('Resume upload error: ' + e.toString());
      }
    }
    
    // Send emails
    try {
      sendApplicationConfirmationEmail(applicationData.email, applicationData.firstName, applicationData.jobId);
      
      const job = getJobDetails(applicationData.jobId);
      sendHRNotification('New Application', 
        'New application for ' + (job ? job['Job Title'] : 'position') + 
        ' from ' + applicationData.firstName + ' ' + applicationData.lastName);
    } catch(e) {
      Logger.log('Email error: ' + e.toString());
    }
    
    return {success: true, candidateId: candidateId, message: 'Application submitted successfully'};
    
  } catch(error) {
    Logger.log('Error submitting application: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Add Candidate by HR
function addCandidateByHR(candidateData) {
  try {
    const candidateId = generateId('CAND');
    const now = new Date();
    const token = generateToken();
    
    let folderId = '';
    try {
      const folderResult = createCandidateFolder(candidateId, candidateData.firstName || '', candidateData.lastName || '');
      folderId = folderResult.folderId || '';
    } catch(e) {
      Logger.log('Could not create folder: ' + e.toString());
    }
    
    const candidate = {
      'Candidate ID': candidateId,
      'Job ID': candidateData.jobId || '',
      'First Name': candidateData.firstName || '',
      'Last Name': candidateData.lastName || '',
      'Email': candidateData.email || '',
      'Phone': candidateData.phone || '',
      'Current Company': '',
      'Current Designation': '',
      'Total Experience': '',
      'Current CTC': '',
      'Expected CTC': '',
      'Notice Period': '',
      'Location': '',
      'Resume Folder ID': folderId,
      'Application Date': now,
      'Source': 'HR Added',
      'Status': CANDIDATE_STATUS.APPLIED,
      'Screening Score': '',
      'HR Notes': '',
      'Created Date': now,
      'Modified Date': now
    };
    
    addRowToSheet(SHEETS.CANDIDATES, candidate);
    saveToken(token, 'application', candidateId, candidateData.email);
    
    try {
      sendCandidateInviteEmail(candidateData.email, candidateData.firstName, candidateData.jobId, token);
    } catch(e) {
      Logger.log('Email error: ' + e.toString());
    }
    
    return {success: true, candidateId: candidateId, message: 'Candidate added successfully'};
    
  } catch(error) {
    Logger.log('Error adding candidate: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Update Candidate Status
function updateCandidateStatus(candidateId, newStatus, notes) {
  try {
    notes = notes || '';
    const candidate = getCandidateById(candidateId);
    if (!candidate) {
      return {success: false, error: 'Candidate not found'};
    }
    
    const updateData = {
      'Status': newStatus,
      'Modified Date': new Date()
    };
    
    if (notes) {
      const existingNotes = candidate['HR Notes'] || '';
      updateData['HR Notes'] = existingNotes + '\n' + new Date().toLocaleString() + ': ' + notes;
    }
    
    updateRowInSheet(SHEETS.CANDIDATES, candidate.rowIndex, updateData);
    
    return {success: true, message: 'Status updated successfully'};
    
  } catch(error) {
    Logger.log('Error updating status: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Schedule Interview
function scheduleInterview(interviewData) {
  try {
    const interviewId = generateId('INT');
    const now = new Date();
    
    const candidate = getCandidateById(interviewData.candidateId);
    if (!candidate) {
      return {success: false, error: 'Candidate not found'};
    }
    
    const jobId = interviewData.jobId || candidate['Job ID'];
    const job = getJobDetails(jobId);
    
    const interview = {
      'Interview ID': interviewId,
      'Candidate ID': interviewData.candidateId,
      'Job ID': jobId,
      'Round': interviewData.round || '1',
      'Interview Date': interviewData.interviewDate,
      'Interview Time': interviewData.interviewTime,
      'Duration': interviewData.duration || '60 minutes',
      'Interview Type': interviewData.interviewType || 'Video Call',
      'Meeting Link': interviewData.meetingLink || '',
      'Interviewer Email': interviewData.interviewerEmail || '',
      'Interviewer Name': interviewData.interviewerName || '',
      'Status': 'Scheduled',
      'Feedback': '',
      'Rating': '',
      'Recommendation': '',
      'Created Date': now
    };
    
    addRowToSheet(SHEETS.INTERVIEWS, interview);
    updateCandidateStatus(interviewData.candidateId, CANDIDATE_STATUS.INTERVIEW_SCHEDULED);
    
    // Send invites
    try {
      sendInterviewInviteToCandidate(candidate, interview, job);
      sendInterviewInviteToInterviewer(interview, candidate, job);
    } catch(e) {
      Logger.log('Email error: ' + e.toString());
    }
    
    return {success: true, interviewId: interviewId, message: 'Interview scheduled successfully'};
    
  } catch(error) {
    Logger.log('Error scheduling interview: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get All Interviews
function getAllInterviews(filters) {
  try {
    filters = filters || {};
    let interviews = getSheetData(SHEETS.INTERVIEWS);
    
    if (filters.status && filters.status !== '') {
      interviews = interviews.filter(i => i['Status'] === filters.status);
    }
    if (filters.interviewerEmail && filters.interviewerEmail !== '') {
      interviews = interviews.filter(i => i['Interviewer Email'] === filters.interviewerEmail);
    }
    if (filters.dateFrom && filters.dateFrom !== '') {
      interviews = interviews.filter(i => new Date(i['Interview Date']) >= new Date(filters.dateFrom));
    }
    if (filters.dateTo && filters.dateTo !== '') {
      interviews = interviews.filter(i => new Date(i['Interview Date']) <= new Date(filters.dateTo));
    }
    
    // Add candidate and job details
    const candidates = getSheetData(SHEETS.CANDIDATES);
    const jobs = getSheetData(SHEETS.JOBS);
    
    interviews = interviews.map(i => {
      const candidate = candidates.find(c => c['Candidate ID'] === i['Candidate ID']);
      const job = jobs.find(j => j['Job ID'] === i['Job ID']);
      
      i.candidateName = candidate ? (candidate['First Name'] + ' ' + candidate['Last Name']) : 'Unknown';
      i.candidateEmail = candidate ? candidate['Email'] : '';
      i.jobTitle = job ? job['Job Title'] : 'Unknown';
      
      return i;
    });
    
    return interviews;
    
  } catch(error) {
    Logger.log('Error getting interviews: ' + error.toString());
    return [];
  }
}

// Submit Interview Feedback
function submitInterviewFeedback(interviewId, feedback) {
  try {
    const interview = findRowByValue(SHEETS.INTERVIEWS, 'Interview ID', interviewId);
    if (!interview) {
      return {success: false, error: 'Interview not found'};
    }
    
    updateRowInSheet(SHEETS.INTERVIEWS, interview.rowIndex, {
      'Feedback': feedback.feedback || '',
      'Rating': feedback.rating || '',
      'Recommendation': feedback.recommendation || '',
      'Status': 'Completed'
    });
    
    updateCandidateStatus(interview['Candidate ID'], CANDIDATE_STATUS.INTERVIEWED);
    
    return {success: true, message: 'Feedback submitted successfully'};
    
  } catch(error) {
    Logger.log('Error submitting feedback: ' + error.toString());
    return {success: false, error: error.toString()};
  }
}

// Get Recruitment Statistics
function getRecruitmentStatistics() {
  try {
    const jobs = getSheetData(SHEETS.JOBS);
    const candidates = getSheetData(SHEETS.CANDIDATES);
    const interviews = getSheetData(SHEETS.INTERVIEWS);
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    let todaysInterviews = 0;
    interviews.forEach(i => {
      try {
        const intDate = new Date(i['Interview Date']);
        intDate.setHours(0, 0, 0, 0);
        if (intDate.getTime() === today.getTime() && i['Status'] === 'Scheduled') {
          todaysInterviews++;
        }
      } catch(e) {}
    });
    
    return {
      openJobs: jobs.filter(j => j['Status'] === 'Open').length,
      totalCandidates: candidates.length,
      candidatesByStatus: {
        applied: candidates.filter(c => c['Status'] === CANDIDATE_STATUS.APPLIED).length,
        screening: candidates.filter(c => c['Status'] === CANDIDATE_STATUS.SCREENING).length,
        interviewing: candidates.filter(c => 
          c['Status'] === CANDIDATE_STATUS.INTERVIEW_SCHEDULED ||
          c['Status'] === CANDIDATE_STATUS.INTERVIEWED
        ).length,
        selected: candidates.filter(c => c['Status'] === CANDIDATE_STATUS.SELECTED).length,
        rejected: candidates.filter(c => c['Status'] === CANDIDATE_STATUS.REJECTED).length
      },
      scheduledInterviews: interviews.filter(i => i['Status'] === 'Scheduled').length,
      todaysInterviews: todaysInterviews
    };
    
  } catch(error) {
    Logger.log('Error getting statistics: ' + error.toString());
    return {
      openJobs: 0,
      totalCandidates: 0,
      candidatesByStatus: {applied: 0, screening: 0, interviewing: 0, selected: 0, rejected: 0},
      scheduledInterviews: 0,
      todaysInterviews: 0
    };
  }
}
