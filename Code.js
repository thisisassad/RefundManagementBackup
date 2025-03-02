// Add these constants at the top of the file
const SPREADSHEET_ID = '1j6pjLj37M5h4e_fqTqteeDjVukU6kcDhhlReTIqUxXY'; 
const SHEET_NAME = 'Refund Requests';
const DRIVE_FOLDER_ID = '1Y-EBjnyUz9TjutaICd5yTQo2BYpgoXap'; 
const MAX_FILE_SIZE = 11 * 1024 * 1024; // 11MB in bytes
const MAX_FILES = 3; // Maximum number of files allowed
const FINANCE_SHEET_NAME = 'Finance Updates';
const SETTINGS_SHEET_NAME = 'Settings';
const PRODUCT_SHEET_NAME = 'Product Access Tracking'; // New constant
const REFUND_APPROVAL_SHEET_NAME = 'Refund Approval Updates';
const BI_SHEET_NAME = 'BI_Updates';

// Add status constants
const REQUEST_STATUS = {
  PENDING: 'Pending',
  APPROVED: 'Approved',
  REJECTED: 'Rejected',
  FINANCE_COMPLETED: 'Completed',
  FINANCE_REJECTED: 'Rejected',
  FINANCE_IN_PROGRESS: 'In Progress',
  PRODUCT_COMPLETED: 'Completed',
  PRODUCT_IN_PROGRESS: 'In Progress'
};

// Add BI status constants
const BI_STATUS = {
  PENDING: 'Pending',
  IN_PROGRESS: 'In Progress',
  COMPLETED: 'Completed',
  REJECTED: 'Rejected'
};

// Add helper function for URL handling
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

// Add course data lookup variable - initialize early in the script
let courseData = { courses: [], coursePhaseMap: {} };
try {
  courseData = getAllCourseNames();
} catch (error) {
  console.error('Error pre-loading course data:', error);
}

// Add email configuration at the top
const EMAIL_CONFIG = {
  defaultRecipients: {
    to: 'contact.mdassaduzzaman@shikho.com',
    cc: [
       'md.assaduzzaman@shikho.com',  // Maintainer
      //  'afsanalima846@gmail.com',         // BI Team
    ]
  },
  conditionalRecipients: {
    revokeAccess: ['qa.shikho@gmail.com'],
    'Jessore Telesales': ['afsanalima846@gmail.com'],
    // 'Dhaka Telesales': ['afsanalima846@gmail.com'],
    // 'CX Team': ['afsanalima846@gmail.com'],
    // 'Customer': ['afsanalima846@gmail.com'],  // Add CX Team for customer requests
    // 'BI Team': ['afsanalima846@gmail.com']       // Add BI Team for BI updates
  }
};

// Modified doGet function to handle authentication
function doGet(e) {
  try {
    console.log('doGet called with parameters:', e.parameters);
    
    // Get page parameter, default to login page
    const page = e.parameter.page || 'login';
    const url = e.parameter.url || '';
    
    // Special case - if we detect userCodeAppPanel in the URL hash or parameter, force redirect to login
    if (page === 'userCodeAppPanel' || page.includes('userCodeAppPanel') || 
        url.includes('userCodeAppPanel') || e.parameter.hash === 'userCodeAppPanel') {
      console.log('User landed on Apps Script runtime URL, redirecting to login');
      return HtmlService.createHtmlOutput(`
        <script>
          window.top.location.href = "${getScriptUrl()}?page=login";
        </script>
      `);
    }
    
    // Handle logout parameter first
    if (e.parameter.logout === 'true') {
      console.log('Logout requested');
      
      // If token is provided, invalidate it
      if (e.parameter.token) {
        console.log('Invalidating token:', e.parameter.token);
        logoutUser(e.parameter.token);
      }
      
      // Return login page directly with logout=true parameter instead of redirecting
      console.log('Rendering login page after logout');
      const template = HtmlService.createTemplateFromFile('Login');
      return template
        .evaluate()
        .setTitle('Login - Refund Management System')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Check user authentication
    const sessionToken = e.parameter.token || '';
    const isAuthenticated = sessionToken ? validateSession(sessionToken) : false;
    
    // Get user access data if authenticated
    const userAccess = isAuthenticated ? getUserAccess(sessionToken) : {};
    
    console.log('Page requested:', page);
    console.log('Authentication status:', isAuthenticated);
    if (isAuthenticated) {
      console.log('User access:', userAccess);
    }
    
    // Special handling for root URL (no parameters)
    if (Object.keys(e.parameters).length === 0) {
      console.log('Root URL accessed with no parameters');
      // For root URL, directly render the login page instead of redirecting
      return HtmlService.createTemplateFromFile('Login')
        .evaluate()
        .setTitle('Login - Refund Management System')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    
    // Route to appropriate page with authentication checks
    switch(page) {
      case 'login':
        // If already authenticated, redirect to menu
        if (isAuthenticated) {
          console.log('User already authenticated, redirecting to menu');
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=menu&token=${sessionToken}";
            </script>
          `);
        }
        
        console.log('Rendering login page');
        return HtmlService.createTemplateFromFile('Login')
          .evaluate()
          .setTitle('Login - Refund Management System')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      case 'finance':
        // Verify user has Finance access
        if (isAuthenticated && hasAccess(sessionToken, 'Finance Portal')) {
          const template = HtmlService.createTemplateFromFile('FinancePortal');
          // Pass authentication data to template
          template.isAuthenticated = true;
          template.userAccess = userAccess;
          template.sessionToken = sessionToken;
          template.currentPage = 'finance'; // Add current page for active state in portal switcher
          
          return template
            .evaluate()
            .setTitle('Finance Refund Portal')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
          // If authenticated but no access, show access denied
          if (isAuthenticated) {
            return createAccessDeniedPage();
          }
          // If not authenticated, redirect to login
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=login";
            </script>
          `);
        }
      
      case 'product':
        // Verify user has Product access
        if (isAuthenticated && hasAccess(sessionToken, 'Product Portal')) {
          const template = HtmlService.createTemplateFromFile('ProductDashboard');
          // Pass authentication data to template
          template.isAuthenticated = true;
          template.userAccess = userAccess;
          template.sessionToken = sessionToken;
          template.currentPage = 'product'; // Add current page for active state in portal switcher
          
          console.log('Passing user data to ProductDashboard:', {
            isAuthenticated: true,
            userAccessKeys: Object.keys(userAccess),
            userEmail: userAccess.email,
            userId: userAccess.id,
            userName: userAccess.name
          });
          
          // Add explicit log for critical data
          console.log('USER EMAIL being passed to ProductDashboard:', userAccess.email);
          console.log('USER ID being passed to ProductDashboard:', userAccess.id);
          console.log('USER ACCESS object contains email:', userAccess.hasOwnProperty('email'));
          
          return template
            .evaluate()
            .setTitle('Product Team Dashboard')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
          // If authenticated but no access, show access denied
          if (isAuthenticated) {
            return createAccessDeniedPage();
          }
          // If not authenticated, redirect to login
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=login";
            </script>
          `);
        }
      
      case 'approver':
        // Verify user has Approver access
        if (isAuthenticated && hasAccess(sessionToken, 'Approver Portal')) {
          const template = HtmlService.createTemplateFromFile('RefundApproverPortal');
          // Pass authentication data to template
          template.isAuthenticated = true;
          template.userAccess = userAccess;
          template.sessionToken = sessionToken;
          template.currentPage = 'approver'; // Add current page for active state in portal switcher
          
          return template
            .evaluate()
            .setTitle('Refund Approver Portal')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
          // If authenticated but no access, show access denied
          if (isAuthenticated) {
            return createAccessDeniedPage();
          }
          // If not authenticated, redirect to login
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=login";
            </script>
          `);
        }
      
      case 'sales':
        // Sales page is accessible to anyone without authentication
        const salesTemplate = HtmlService.createTemplateFromFile('Index');
        // Pass course data
        salesTemplate.courseData = courseData; 
        // Still pass authentication info if available (for header display, etc.)
        salesTemplate.isAuthenticated = isAuthenticated;
        salesTemplate.userAccess = userAccess;
        salesTemplate.sessionToken = sessionToken;
        salesTemplate.currentPage = 'sales'; // Add current page for active state in portal switcher
        
        return salesTemplate
          .evaluate()
          .setTitle('Refund Request Form')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      case 'menu':
        // Menu page - need authentication to view
        if (!isAuthenticated) {
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=login";
            </script>
          `);
        }
        
        const menuTemplate = HtmlService.createTemplateFromFile('MenuPage');
        menuTemplate.isAuthenticated = isAuthenticated;
        menuTemplate.userAccess = userAccess;
        menuTemplate.currentPage = 'menu'; // Add current page for active state in portal switcher
        
        return menuTemplate
          .evaluate()
          .setTitle('Refund Management System')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      case 'bi':
        // Verify user has BI Portal access
        if (isAuthenticated && hasAccess(sessionToken, 'BI Portal')) {
          const template = HtmlService.createTemplateFromFile('BIPortal');
          // Pass authentication data to template
          template.isAuthenticated = true;
          template.userAccess = userAccess;
          template.sessionToken = sessionToken;
          template.currentPage = 'bi'; // Add current page for active state in portal switcher
          
          return template
            .evaluate()
            .setTitle('BI Team Portal')
            .addMetaTag('viewport', 'width=device-width, initial-scale=1')
            .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        } else {
          // If authenticated but no access, show access denied
          if (isAuthenticated) {
            return createAccessDeniedPage();
          }
          // If not authenticated, redirect to login
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=login";
            </script>
          `);
        }
      
      case 'reset':
        // Password reset page - doesn't require authentication
        // Check if token is provided
        const resetToken = e.parameter.token || '';
        if (!resetToken) {
          // Redirect to login if no token provided
          return HtmlService.createHtmlOutput(`
            <script>
              window.top.location.href = "${getScriptUrl()}?page=login";
            </script>
          `);
        }
        
        // Validate reset token
        const tokenValidation = validateResetToken(resetToken);
        console.log('Reset token validation result:', tokenValidation);
        
        const resetTemplate = HtmlService.createTemplateFromFile('ResetPassword');
        resetTemplate.tokenValid = tokenValidation.valid;
        resetTemplate.resetToken = resetToken;
        
        return resetTemplate
          .evaluate()
          .setTitle('Reset Password - Refund Management System')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      default:
        // For any unknown page, redirect to login
        return HtmlService.createHtmlOutput(`
          <script>
            window.top.location.href = "${getScriptUrl()}?page=login";
          </script>
        `);
    }
  } catch (error) {
    console.error('Error in doGet:', error);
    return HtmlService.createHtmlOutput(`
      <h2>Error: Unable to load page</h2>
      <p>Please contact the administrator with this error: ${error.toString()}</p>
      <p>Error details: ${error.stack}</p>
    `);
  }
}

function setupSpreadsheet(sheet) {
  // Define headers - must match exactly with rowData order in processForm
  const headers = [
    'Request ID',     // New column for request ID
    'Timestamp',
    'Requester',
    'Requester Email',
    'Vertical',
    'Enrollment Number',
    'Customer Name',
    'Purchase Date',           // Consolidated purchase date
    'Payment ID',
    'Purchased Amount',        // Consolidated purchased amount
    'Refund Type',
    'Mistaken Course Name',
    'Mistaken Phase Name',
    'Intended Course Name',
    'Intended Phase Name',
    'Mistaken Course Price',
    'Wrong Phone',
    'Correct Phone',
    'Sender bKash',
    'Transaction Time',
    'Transaction Time String', // Added new column for transaction time string
    'Sent Amount',
    'Refund Amount',
    'Revoke Access',
    'Access Revoke Status',    // New column for Product team
    'Access Revoke Date',      // New column for Product team
    'Access Revoke By',        // New column for Product team
    'Access Revoke Remarks',   // New column for Product team
    'Refund Reason',
    'Document Folder'          // Changed from 'Document URLs' to 'Document Folder'
  ];

  try {
    // Ensure sheet exists
    if (!sheet) {
      console.error('Sheet is undefined');
      return false;
    }

    // Check if the sheet already has data
    const lastRow = Math.max(sheet.getLastRow(), 1);
    
    // If the sheet is empty or doesn't have headers, set them up
    if (lastRow <= 1 || sheet.getRange(2, 1).getValue() === '') {
      console.log('Setting up new sheet with headers');
      
      // Set a title in cell A1
      sheet.getRange(1, 1).setValue('Refund Request Submissions');
      sheet.getRange(1, 1, 1, headers.length).merge();
      sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
      
      // Set headers
      sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      const headerRange = sheet.getRange(2, 1, 1, headers.length);
      headerRange.setBackground('#4286f4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      
      // Set column widths
      sheet.setColumnWidth(1, 180);  // Timestamp
      sheet.setColumnWidth(2, 150);  // Requester
      sheet.setColumnWidth(3, 200);  // Email
      sheet.setColumnWidth(4, 100);  // Vertical
      sheet.setColumnWidth(5, 150);  // Enrollment Number
      sheet.setColumnWidth(6, 150);  // Customer Contact
      sheet.setColumnWidth(7, 120);  // Purchase Date
      sheet.setColumnWidth(8, 150);  // Payment ID
      sheet.setColumnWidth(9, 120);  // Purchased Amount
      sheet.setColumnWidth(10, 150); // Refund Type
      sheet.setColumnWidth(11, 200); // Mistaken Course
      sheet.setColumnWidth(12, 200); // Intended Course
      sheet.setColumnWidth(13, 150); // Mistaken Price
      sheet.setColumnWidth(14, 150); // Wrong Phone
      sheet.setColumnWidth(15, 150); // Correct Phone
      sheet.setColumnWidth(16, 150); // Sender bKash
      sheet.setColumnWidth(17, 180); // Transaction Time
      sheet.setColumnWidth(18, 120); // Sent Amount
      sheet.setColumnWidth(19, 120); // Refund Amount
      sheet.setColumnWidth(20, 120); // Revoke Access
      sheet.setColumnWidth(21, 150); // Access Revoke Status
      sheet.setColumnWidth(22, 150); // Access Revoke Date
      sheet.setColumnWidth(23, 150); // Access Revoke By
      sheet.setColumnWidth(24, 250); // Access Revoke Remarks
      sheet.setColumnWidth(25, 250); // Refund Reason
      sheet.setColumnWidth(26, 300); // Document Folder
      
      // Freeze header rows
      sheet.setFrozenRows(2);
      
      // Add filters
      sheet.getRange(2, 1, 1, headers.length).createFilter();
    }
    
    // Always ensure proper formatting for any new rows that might be added
    // These apply to the entire columns regardless of existing data
    
    // Set number format for amount columns
    const amountColumns = [9, 13, 15, 18, 19];
    amountColumns.forEach(col => {
      sheet.getRange(3, col, sheet.getMaxRows() - 2).setNumberFormat('#,##0.00 ৳');
    });
    
    // Set date format for date columns
    const dateColumns = [1, 7, 13];
    dateColumns.forEach(col => {
      sheet.getRange(3, col, sheet.getMaxRows() - 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    });

    // Set column widths for new columns
    sheet.setColumnWidth(24, 150); // Access Revoke Status
    sheet.setColumnWidth(25, 150); // Access Revoke Date
    sheet.setColumnWidth(26, 150); // Access Revoke By
    sheet.setColumnWidth(27, 250); // Access Revoke Remarks

    // Set date format for Access Revoke Date column
    sheet.getRange(3, 25, sheet.getMaxRows() - 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');

    return true;
  } catch (error) {
    console.error('Error setting up spreadsheet:', error);
    console.error('Error stack:', error.stack);
    return false;
  }
}

async function handleFileUploads(formData, enrollmentNumber) {
  try {
    console.log('Starting file upload process');
    
    if (!formData.documents || formData.documents.length === 0) {
      console.error('No documents provided for upload');
      throw new Error('No documents provided');
    }

    const timestamp = new Date().getTime();
    const folderName = `Refund_${enrollmentNumber}_${timestamp}`;
    console.log('Creating folder:', folderName);
    
    const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    if (!parentFolder) {
      throw new Error(`Parent folder not found: ${DRIVE_FOLDER_ID}`);
    }
    
    const newFolder = parentFolder.createFolder(folderName);
    console.log('Folder created successfully:', newFolder.getName());
    
    // Upload all files to this folder
    for (const fileData of formData.documents) {
      console.log('Processing file:', fileData.fileName);
      const blob = Utilities.newBlob(
        Utilities.base64Decode(fileData.data),
        fileData.mimeType,
        fileData.fileName
      );
      const file = newFolder.createFile(blob);
      console.log('File uploaded successfully:', file.getName(), 'to folder:', folderName);
    }
    
    const folderUrl = newFolder.getUrl();
    console.log('Folder URL:', folderUrl);
    return folderUrl;
    
  } catch (error) {
    console.error('Error in handleFileUploads:', error);
    throw error;
  }
}

// Add function to get email recipients based on form data
function getEmailRecipients(formData) {
  console.log('Getting email recipients for:', formData);  // Add debug log
  
  const recipients = {
    to: EMAIL_CONFIG.defaultRecipients.to,
    cc: [...EMAIL_CONFIG.defaultRecipients.cc]  // Create copy of default CC list
  };

  // Add CX Team if requester is Customer
  if (formData.requester === 'Customer') {
    console.log('Adding CX Team recipients for Customer request');  // Add debug log
    recipients.cc.push(...EMAIL_CONFIG.conditionalRecipients['CX Team']);
  }
  // Add other conditional recipients based on requester
  else if (EMAIL_CONFIG.conditionalRecipients[formData.requester]) {
    console.log('Adding conditional recipients for:', formData.requester);  // Add debug log
    recipients.cc.push(...EMAIL_CONFIG.conditionalRecipients[formData.requester]);
  }

  // Add product team if access needs to be revoked
  if (formData.revokeAccess === 'Yes') {
    console.log('Adding revoke access recipients');  // Add debug log
    recipients.cc.push(...EMAIL_CONFIG.conditionalRecipients.revokeAccess);
  }

  // Remove duplicates from CC list
  recipients.cc = [...new Set(recipients.cc)];

  console.log('Final recipients:', recipients);  // Add debug log
  return recipients;
}

async function processForm(formData) {
  try {
    console.log('Starting form processing');
    
    // First consolidate the form fields
    const consolidatedData = consolidateFormFields(formData);
    console.log('Consolidated form data:', consolidatedData);
    
    // Make sure we await the file upload
    const folderUrl = await handleFileUploads(formData, consolidatedData.enrollmentNumber);
    console.log('Created folder for files, URL:', folderUrl);
    
    if (!folderUrl) {
      throw new Error('Failed to create folder for documents');
    }
    
    // Get the spreadsheet and sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    const requestId = await generateRequestId(sheet);
    console.log('Generated Request ID:', requestId);
    
    // Prepare row data - order must match headers exactly
    const timestamp = new Date();
    const rowData = [
      requestId,                                    // Request ID
      timestamp,                                    // Timestamp
      formData.requester,                          // Requester
      formData.requesterEmail,                     // Requester Email
      formData.vertical,                           // Vertical
      consolidatedData.enrollmentNumber,           // Enrollment Number - Updated to use consolidated data
      formData.customerContact,                    // Customer Name
      consolidatedData.purchaseDate ? new Date(consolidatedData.purchaseDate) : '',  // Purchase Date
      formData.paymentId,                          // Payment ID
      parseFloat(consolidatedData.purchasedAmount) || 0,  // Purchased Amount
      consolidatedData.refundType,                 // Refund Type
      consolidatedData.mistakenCourse,             // Mistaken Course Name
      consolidatedData.mistakenPhase,              // Mistaken Phase Name
      consolidatedData.intendedCourse,             // Intended Course Name
      consolidatedData.intendedPhase,              // Intended Phase Name
      parseFloat(consolidatedData.mistakenPrice) || 0,  // Mistaken Course Price
      consolidatedData.wrongPhone,                 // Wrong Phone
      consolidatedData.correctPhone,               // Correct Phone
      consolidatedData.senderBkash,                // Sender bKash
      consolidatedData.transactionTime ? new Date(consolidatedData.transactionTime) : '',  // Transaction Time
      // '',                                          // Transaction Time String - Leave empty for sheet formula
      parseFloat(consolidatedData.sentAmount) || 0,  // Sent Amount
      parseFloat(consolidatedData.refundAmount) || 0,  // Refund Amount
      consolidatedData.revokeAccess,               // Revoke Access
      consolidatedData.accessRevokeStatus,            // Access Revoke Status
      consolidatedData.accessRevokeDate ? new Date(consolidatedData.accessRevokeDate) : '',  // Access Revoke Date
      consolidatedData.accessRevokeBy || '',        // Access Revoke By
      consolidatedData.accessRevokeRemarks || '',      // Access Revoke Remarks
      consolidatedData.refundReason,               // Refund Reason
      folderUrl                                    // Document Folder
    ];
    
    // Append to sheet
    sheet.appendRow(rowData);
    console.log('Row appended successfully');
    
    // Get email recipients
    const recipients = getEmailRecipients(formData);
    console.log('Email recipients:', recipients);

    // Send notification email
    sendRefundRequestEmail(recipients, requestId, formData, consolidatedData, folderUrl, timestamp);
    
    // Send confirmation email to requester
    if (formData.requesterEmail !== 'customer@submission.com') {
      sendConfirmationEmail(formData.requesterEmail, requestId, formData, consolidatedData, folderUrl, timestamp);
    }

    console.log('Emails sent successfully');

    return { 
      success: true, 
      message: 'Request submitted successfully',
      requestId: requestId,
      folderUrl: folderUrl,
      formUrl: getScriptUrl(),  // Add form URL to response
      enrollmentNumber: consolidatedData.enrollmentNumber  // Add enrollment number to response
    };
    
  } catch (error) {
    console.error('Error in processForm:', error);
    return { 
      success: false, 
      message: error.toString() 
    };
  }
}

function consolidateFormFields(formData) {
  // For Phone Number Exchange, we want both wrongPhone and enrollmentNumber to have the same value
  const wrongPhoneValue = formData.refundType === 'Phone Number Exchange' ? formData.wrongPhone : '';
  
  // Get enrollment number based on refund type
  let enrollmentNumber = '';
  switch(formData.refundType) {
    case 'Course Exchange':
      enrollmentNumber = formData.courseExchangeEnrollmentNumber;
      break;
    case 'Phone Number Exchange':
      enrollmentNumber = formData.wrongPhone;
      break;
    case 'Partial Refund':
      enrollmentNumber = formData.partialRefundEnrollmentNumber;
      break;
    case 'Full Refund':
      enrollmentNumber = formData.fullRefundEnrollmentNumber;
      break;
    case 'Mistaken Payment':
      enrollmentNumber = formData.senderBkash;
      break;
  }
  
  return {
    // Basic info
    enrollmentNumber: enrollmentNumber,
    
    // Purchase details (consolidated from different sections)
    purchaseDate: formData.courseExchangePurchaseDate || 
                 formData.phoneExchangePurchaseDate || 
                 formData.partialRefundPurchaseDate || 
                 formData.fullRefundPurchaseDate || '',
                 
    purchasedAmount: formData.courseExchangePurchasedAmount || 
                    formData.phoneExchangePurchasedAmount || 
                    formData.partialRefundPurchasedAmount || 
                    formData.fullRefundPurchasedAmount || '',
    
    // Course details with phase names
    mistakenCourse: formData.mistakenCourse || 
                   formData.partialMistakenCourse || 
                   formData.fullRefundMistakenCourse || '',
    
    mistakenPhase: formData.mistakenPhase ||
                   formData.partialMistakenPhase ||
                   formData.fullRefundMistakenPhase || '',
                   
    intendedCourse: formData.intendedCourse || 
                   formData.partialIntendedCourse || '',
    
    intendedPhase: formData.intendedPhase ||
                   formData.partialIntendedPhase || '',
                   
    mistakenPrice: formData.partialMistakenPrice || 
                  formData.fullRefundMistakenPrice || '',
    
    // Phone exchange details
    wrongPhone: wrongPhoneValue,
    correctPhone: formData.correctPhone || '',
    
    // Mistaken payment details
    senderBkash: formData.senderBkash || '',
    transactionTime: formData.transactionTime || '',
    sentAmount: formData.sentAmount || '',
    
    // Final details
    refundType: formData.refundType,
    refundAmount: formData.refundAmount,
    revokeAccess: formData.revokeAccess,
    accessRevokeStatus: formData.accessRevokeStatus,
    accessRevokeDate: formData.accessRevokeDate,
    accessRevokeBy: formData.accessRevokeBy,
    accessRevokeRemarks: formData.accessRevokeRemarks,
    refundReason: formData.refundReason
  };
}

async function generateRequestId(sheet) {
  try {
    const currentYear = new Date().getFullYear().toString().slice(-2); // Get last 2 digits
    const lastRow = sheet.getLastRow();
    
    // If sheet is empty (only headers), start with 0001
    if (lastRow <= 2) {
      return `${currentYear}0001`;
    }
    
    // Get all request IDs from this year
    const requestIds = sheet.getRange(3, 1, lastRow - 2).getValues()
      .map(row => row[0].toString())
      .filter(id => id.startsWith(currentYear));
    
    if (requestIds.length === 0) {
      // First request of the year
      return `${currentYear}0001`;
    }
    
    // Get the highest number and increment
    const lastNumber = Math.max(...requestIds.map(id => parseInt(id.slice(-4))));
    const nextNumber = (lastNumber + 1).toString().padStart(4, '0');
    
    return `${currentYear}${nextNumber}`;
  } catch (error) {
    console.error('Error generating request ID:', error);
    throw error;
  }
}

// Add function to set up finance sheet
function setupFinanceSheet(sheet) {
  const headers = [
    'Timestamp',
    'Refund SL Number',
    'Course Enrollment Number',
    'Refunded Amount',
    'Transaction/Payment ID',
    'Refund Method',
    'Status',
    'Finance Remarks'
  ];

  try {
    // Set headers and formatting
    sheet.getRange(1, 1).setValue('Finance Refund Updates');
    sheet.getRange(1, 1, 1, headers.length).merge();
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    
    const headerRange = sheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#4286f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 180);  // Timestamp
    sheet.setColumnWidth(2, 150);  // Refund SL Number
    sheet.setColumnWidth(3, 200);  // Enrollment Number
    sheet.setColumnWidth(4, 150);  // Refunded Amount
    sheet.setColumnWidth(5, 200);  // Transaction ID
    sheet.setColumnWidth(6, 150);  // Refund Method
    sheet.setColumnWidth(7, 120);  // Status
    sheet.setColumnWidth(8, 250);  // Finance Remarks
    
    // Freeze header rows
    sheet.setFrozenRows(2);
    
    // Add filters
    sheet.getRange(2, 1, 1, headers.length).createFilter();
    
    // Set number format for amount column
    sheet.getRange(3, 4, sheet.getMaxRows() - 2).setNumberFormat('#,##0.00 ৳');
    
    return true;
  } catch (error) {
    console.error('Error setting up finance sheet:', error);
    return false;
  }
}

// Add function to process finance form submissions
function processFinanceForm(formData) {
  try {
    // Get or create finance sheet
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(FINANCE_SHEET_NAME);
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(FINANCE_SHEET_NAME);
      setupFinanceSheet(sheet);
    }
    
    // Prepare row data
    const timestamp = new Date();
    const rowData = [
      timestamp,
      formData.refundSlNumber,
      formData.enrollmentNumber,
      parseFloat(formData.refundedAmount) || 0,
      formData.transactionId,
      'Completed'  // Initial status
    ];
    
    // Append to sheet
    sheet.appendRow(rowData);
    console.log('Finance update recorded successfully');
    
    // Return success response
    return {
      success: true,
      message: 'Update submitted successfully',
      refundSlNumber: formData.refundSlNumber,
      enrollmentNumber: formData.enrollmentNumber,
      refundedAmount: formData.refundedAmount,
      transactionId: formData.transactionId
    };
    
  } catch (error) {
    console.error('Error processing finance form:', error);
    return { 
      success: false, 
      message: error.toString() 
    };
  }
}

// Function to get all refund requests
function getRefundRequests() {
  try {
    console.log('Starting getRefundRequests');
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const financeSheet = ss.getSheetByName(FINANCE_SHEET_NAME);
    const approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
    
    if (!requestSheet) {
      console.error('Request sheet not found:', SHEET_NAME);
      return { requests: [], averageSLA: 0 };
    }
    
    // Get approval statuses first
    const approvalData = approvalSheet ? approvalSheet.getDataRange().getValues() : [];
    const approvalHeaders = approvalData[1] || [];
    const approvalRequestIdCol = approvalHeaders.indexOf('Request ID');
    const approvalStatusCol = approvalHeaders.indexOf('Approval Status');
    
    // Create map of approved requests
    const approvedRequests = new Map();
    if (approvalData.length > 2) {
      for (let i = 2; i < approvalData.length; i++) {
        const row = approvalData[i];
        const requestId = row[approvalRequestIdCol];
        const status = row[approvalStatusCol];
        if (requestId) {
          approvedRequests.set(requestId, status);
        }
      }
    }
    
    // Get finance updates data for status timestamps
    const financeData = financeSheet ? financeSheet.getDataRange().getValues() : [];
    const financeHeaders = financeData[1] || [];
    const statusTimestampCol = financeHeaders.indexOf('Status Timestamp String');
    const financeRequestIdCol = financeHeaders.indexOf('Refund SL Number');
    const financeStatusCol = financeHeaders.indexOf('Status');
    
    // Create a map of latest finance status updates
    const statusUpdates = {};
    let totalCompletionTime = 0;
    let completedCount = 0;
    
    if (financeData.length > 2) {
      for (let i = 2; i < financeData.length; i++) {
        const row = financeData[i];
        const requestId = row[financeRequestIdCol];
        const status = row[financeStatusCol];
        const timestamp = row[statusTimestampCol];
        
        if (requestId && status && timestamp) {
          if (!statusUpdates[requestId] || 
              new Date(timestamp) > new Date(statusUpdates[requestId].timestamp)) {
            statusUpdates[requestId] = {
              status: status,
              timestamp: timestamp
            };
          }
        }
      }
    }
    
    // Get request data
    const headerRow = 2;
    const headers = requestSheet.getRange(headerRow, 1, 1, requestSheet.getLastColumn()).getValues()[0];
    const dataStartRow = 3;
    const lastRow = requestSheet.getLastRow();
    const lastCol = requestSheet.getLastColumn();
    
    if (lastRow < dataStartRow) {
      return { requests: [], averageSLA: 0 };
    }
    
    const dataRange = requestSheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol);
    const data = dataRange.getValues();
    
    // Map column indices (preserve existing mapping)
    const columnMap = {
      requestId: headers.indexOf('Request ID'),
      timestampString: headers.indexOf('Timestamp'),
      purchaseDateString: headers.indexOf('Purchase Date'),
      vertical: headers.indexOf('Vertical'),
      paymentId: headers.indexOf('Payment ID'),
      refundType: headers.indexOf('Refund Type'),
      refundAmount: headers.indexOf('Refund Amount'),
      refundReason: headers.indexOf('Refund Reason'),
      status: headers.indexOf('Status'),
      enrollmentNumber: headers.indexOf('Enrollment Number'),
      customerContact: headers.indexOf('Customer Name'),
      requester: headers.indexOf('Requester'),
      requesterEmail: headers.indexOf('Requester Email'),
      documentFolder: headers.indexOf('Document Folder'),
      mistakenCourse: headers.indexOf('Mistaken Course Name'),
      mistakenPhase: headers.indexOf('Mistaken Phase Name'),
      intendedCourse: headers.indexOf('Intended Course Name'),
      intendedPhase: headers.indexOf('Intended Phase Name'),
      wrongPhone: headers.indexOf('Wrong Phone'),
      correctPhone: headers.indexOf('Correct Phone'),
      senderBkash: headers.indexOf('Sender bKash'),
      transactionTimeString: headers.indexOf('Transaction Time String'),
      transactionTime: headers.indexOf('Transaction Time'),
      sentAmount: headers.indexOf('Sent Amount')
    };
    
    let totalProcessingTime = 0;
    let completedRequests = 0;
    
    // Process and filter requests
    const requests = data
      .filter(row => {
        const requestId = row[columnMap.requestId];
        const approvalStatus = approvedRequests.get(requestId);
        // Only show requests that have been approved
        return approvalStatus === REQUEST_STATUS.APPROVED;
      })
      .map(row => {
        const requestId = row[columnMap.requestId];
        const statusUpdate = statusUpdates[requestId] || {};
        const requestDate = new Date(row[columnMap.timestampString]);
        
        // Calculate processing time (preserve existing logic)
        if ((statusUpdate.status === REQUEST_STATUS.FINANCE_COMPLETED || 
             statusUpdate.status === REQUEST_STATUS.FINANCE_REJECTED) && 
            statusUpdate.timestamp) {
          const completionDate = new Date(statusUpdate.timestamp);
          if (!isNaN(requestDate.getTime()) && !isNaN(completionDate.getTime())) {
            const processingTime = completionDate - requestDate;
            totalProcessingTime += processingTime;
            completedRequests++;
          }
        }
        
        // Preserve all existing fields in the response
        return {
          requestId: requestId,
          timestampString: row[columnMap.timestampString] ? formatDateTime(row[columnMap.timestampString]) : '',
          purchaseDateString: row[columnMap.purchaseDateString] ? formatDateTime(row[columnMap.purchaseDateString]) : '',
          vertical: row[columnMap.vertical] || '',
          paymentId: row[columnMap.paymentId] || '',
          refundType: row[columnMap.refundType] || '',
          refundAmount: typeof row[columnMap.refundAmount] === 'number' ? 
            row[columnMap.refundAmount].toFixed(2) : '',
          refundReason: row[columnMap.refundReason] || '',
          status: statusUpdate.status || row[columnMap.status] || REQUEST_STATUS.PENDING,
          statusUpdateTime: statusUpdate.timestamp || '',
          enrollmentNumber: row[columnMap.enrollmentNumber] || '',
          customerContact: row[columnMap.customerContact] || '',
          requester: row[columnMap.requester] || '',
          requesterEmail: row[columnMap.requesterEmail] || '',
          documentFolder: row[columnMap.documentFolder] || '',
          mistakenCourse: row[columnMap.mistakenCourse] || '',
          mistakenPhase: row[columnMap.mistakenPhase] || '',
          intendedCourse: row[columnMap.intendedCourse] || '',
          intendedPhase: row[columnMap.intendedPhase] || '',
          wrongPhone: row[columnMap.wrongPhone] || '',
          correctPhone: row[columnMap.correctPhone] || '',
          senderBkash: row[columnMap.senderBkash] || '',
          transactionTime: row[columnMap.transactionTime] ? formatDateTime(row[columnMap.transactionTime]) : '',
          sentAmount: typeof row[columnMap.sentAmount] === 'number' ? 
            row[columnMap.sentAmount].toFixed(2) : ''
        };
      });
    
    // Calculate average SLA (preserve existing logic)
    const averageSLA = completedRequests > 0 ? 
      (totalProcessingTime / completedRequests) / (1000 * 60 * 60) : 0;
    
    console.log('Processed requests:', requests.length);
    return { 
      requests: requests,
      averageSLA: averageSLA
    };
    
  } catch (error) {
    console.error('Error in getRefundRequests:', error);
    throw new Error('Failed to load refund requests: ' + error.message);
  }
}

// Add function to get update history for a request
function getUpdateHistory(requestId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const financeSheet = ss.getSheetByName(FINANCE_SHEET_NAME);
    
    if (!financeSheet) {
      return [];
    }
    
    const data = financeSheet.getDataRange().getValues();
    // Skip title row (row 1) and get headers from row 2
    const headers = data[1];
    
    // Find column indices - using exact header names from setupFinanceSheet
    const timestampCol = 0; // First column is always Timestamp
    const requestIdCol = headers.indexOf('Refund SL Number');
    const amountCol = headers.indexOf('Refunded Amount');
    const methodCol = headers.indexOf('Refund Method');
    const statusCol = headers.indexOf('Status');
    const remarksCol = headers.indexOf('Finance Remarks');
    
    if (requestIdCol === -1) {
      console.error('Required columns not found in finance sheet');
      return [];
    }
    
    // Get all updates for this request ID
    const updates = data.slice(2) // Skip header rows
      .filter(row => String(row[requestIdCol]) === String(requestId))
        .map(row => ({
          timestamp: formatDateTime(row[timestampCol]),
          refundedAmount: row[amountCol] ? `${row[amountCol].toFixed(2)} BDT` : '',
          refundMethod: row[methodCol] || '',
          status: row[statusCol] || 'Pending',
          remarks: row[remarksCol] || ''
        }))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp)); // Sort by timestamp descending
    
    console.log(`Found ${updates.length} updates for request ${requestId}`);
    return updates;
  } catch (error) {
    console.error('Error getting update history:', error);
    console.error('Error stack:', error.stack);
    return [];
  }
}

// Helper function to format date time consistently
function formatDateTime(date) {
  if (!date) return '';
  try {
    const d = new Date(date);
    if (isNaN(d.getTime())) return '';
    
    return d.toLocaleString('en-US', {
      year: 'numeric',
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });
  } catch (error) {
    console.error('Error formatting date:', error);
    return '';
  }
}

// Update processFinanceUpdate function to handle Payment ID updates
function processFinanceUpdate(data) {
  try {
    console.log('Starting finance update with data:', data);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Update Finance Updates sheet
    let financeSheet = ss.getSheetByName(FINANCE_SHEET_NAME);
    if (!financeSheet) {
      financeSheet = ss.insertSheet(FINANCE_SHEET_NAME);
      setupFinanceSheet(financeSheet);
    }
    
    // Prepare row data for finance sheet
    const timestamp = new Date();
    const financeRowData = [
      timestamp,                    // Timestamp column
      data.requestId,              // Refund SL Number
      data.enrollmentNumber || '', // Course Enrollment Number
      parseFloat(data.refundedAmount) || 0,
      data.paymentId || '',        // Updated Payment ID
      data.refundMethod || '',
      data.status || 'Pending',
      data.remarks || ''
    ];
    
    // Append to finance sheet
    financeSheet.appendRow(financeRowData);
    console.log('Added row to finance sheet');
    
    // 2. Update main Refund Requests sheet
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const requestData = requestSheet.getDataRange().getValues();
    const headers = requestData[1];  // Get headers from row 2
    
    // Find column indices
    const requestIdCol = headers.indexOf('Request ID');
    const statusCol = headers.indexOf('Status');
    const paymentIdCol = headers.indexOf('Payment ID');
    const emailCol = headers.indexOf('Requester Email');
    const refundTypeCol = headers.indexOf('Refund Type');
    const refundAmountCol = headers.indexOf('Refund Amount');
    const requesterCol = headers.indexOf('Requester');  // Add this line
    
    if (requestIdCol === -1 || statusCol === -1 || paymentIdCol === -1) {
      throw new Error('Required columns not found in the sheet');
    }
    
    console.log('Searching for request ID:', data.requestId);
    
    // Find the row with matching Request ID
    let rowFound = false;
    let requestDetails = null;
    for (let i = 2; i < requestData.length; i++) {
      if (String(requestData[i][requestIdCol]) === String(data.requestId)) {
        console.log('Found matching row at index:', i);
        const rowNumber = i + 1;  // Convert to 1-based index
        
        // Store request details for email
        requestDetails = {
          requesterEmail: requestData[i][emailCol],
          refundType: requestData[i][refundTypeCol],
          refundAmount: requestData[i][refundAmountCol],
          paymentId: requestData[i][paymentIdCol],
          requester: requestData[i][requesterCol]  // Add this line
        };
        
        // Update status and payment ID
        requestSheet.getRange(rowNumber, statusCol + 1).setValue(data.status);
        requestSheet.getRange(rowNumber, paymentIdCol + 1).setValue(data.paymentId);
        rowFound = true;
        break;
      }
    }
    
    if (!rowFound) {
      console.error('Request ID not found:', data.requestId);
      throw new Error('Request ID not found in the main sheet');
    }
    
    // 3. Send status update emails
    const recipients = getEmailRecipients({
      requester: requestDetails.requester,  // Use the actual requester value
      revokeAccess: data.status === REQUEST_STATUS.REJECTED ? 'Yes' : 'No'
    });
    
    console.log('Email recipients:', recipients);  // Add this for debugging
    
    sendStatusUpdateEmail(data.requestId, data, requestDetails, recipients);
    
    console.log('Update completed successfully');
    return {
      success: true,
      message: 'Update processed successfully'
    };
    
  } catch (error) {
    console.error('Error processing finance update:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Function to send rejection notification
function sendRejectionEmail(requesterEmail, requestId, data) {
  try {
    if (requesterEmail && requesterEmail !== 'customer@submission.com') {
      MailApp.sendEmail({
        to: requesterEmail,
        subject: `Refund Request ${requestId} - Rejected`,
        htmlBody: `
          <div style="font-family: Arial, sans-serif; padding: 20px;">
            <h2 style="color: #dc3545;">Refund Request Rejected</h2>
            <p>Your refund request has been reviewed and cannot be processed at this time.</p>
            <div style="background-color: #f8f9fa; padding: 15px; border-radius: 4px; margin: 20px 0;">
              <p><strong>Request ID:</strong> ${requestId}</p>
              <p><strong>Remarks:</strong> ${data.remarks || 'No remarks provided'}</p>
            </div>
            <p>If you have any questions, please contact the finance team.</p>
          </div>
        `,
        noReply: true
      });
    }
  } catch (error) {
    console.error('Error sending rejection email:', error);
  }
}

// Update getAllCourseNames function to handle course-phase relationships
function getAllCourseNames() {
  try {
    console.log('Starting getAllCourseNames function');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    
    if (!settingsSheet) {
      console.error('Settings sheet not found');
      return { courses: [], coursePhaseMap: {} };
    }

    // Get all data from the sheet
    const data = settingsSheet.getDataRange().getValues();
    
    // Get headers from row 2 (index 1) instead of row 1
    const headers = data[1];  // Changed from data[0]
    
    // Find column indices
    const courseNameColIndex = headers.indexOf('All Course Names');
    const phaseNameColIndex = headers.indexOf('Phase Name');

    console.log('Column indices found:', {
      courseNameColIndex,
      phaseNameColIndex,
      headers
    });

    if (courseNameColIndex === -1 || phaseNameColIndex === -1) {
      console.error('Required columns not found in Settings sheet. Available columns:', headers);
      return { courses: [], coursePhaseMap: {} };
    }

    // Create course-phase mapping
    const coursePhaseMap = {};
    const allCourses = new Set(); // Use Set for unique course names

    // Start from row 3 (index 2) since headers are in row 2
    for (let i = 2; i < data.length; i++) {  // Changed from i = 1
      const courseName = data[i][courseNameColIndex];
      const phaseName = data[i][phaseNameColIndex];
      
      // Skip empty rows
      if (!courseName || !courseName.trim()) continue;

      // Add course to unique courses set
      allCourses.add(courseName);

      // Initialize array for course if not exists
      if (!coursePhaseMap[courseName]) {
        coursePhaseMap[courseName] = new Set(); // Use Set for unique phases
      }

      // Add phase if it exists
      if (phaseName && phaseName.trim()) {
        coursePhaseMap[courseName].add(phaseName);
      }
    }

    // Convert Sets to Arrays in coursePhaseMap
    const finalCoursePhaseMap = {};
    for (const [course, phases] of Object.entries(coursePhaseMap)) {
      finalCoursePhaseMap[course] = Array.from(phases);
    }

    const result = {
      courses: Array.from(allCourses),
      coursePhaseMap: finalCoursePhaseMap
    };

    console.log('getAllCourseNames result:', {
      courseCount: result.courses.length,
      // courses: result.courses,
      // coursePhaseMap: result.coursePhaseMap
    });
    
    return result;
  } catch (error) {
    console.error('Error in getAllCourseNames:', error);
    console.error('Stack trace:', error.stack);
    return { courses: [], coursePhaseMap: {} };
  }
}

// Add a helper function for formatting date in the required format
function formatDateForSubject(date) {
  const d = new Date(date);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${d.getDate().toString().padStart(2, '0')}-${months[d.getMonth()]}-${d.getFullYear()}`;
}

// Function to get data for product dashboard
function getProductDashboardData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const financeSheet = ss.getSheetByName(FINANCE_SHEET_NAME);
    const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    const approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
    
    // Get approval statuses
    const approvalData = approvalSheet ? approvalSheet.getDataRange().getValues() : [];
    const approvalHeaders = approvalData[1] || [];
    const approvalRequestIdCol = approvalHeaders.indexOf('Request ID');
    const approvalStatusCol = approvalHeaders.indexOf('Approval Status');
    
    // Create map of approved requests
    const approvedRequests = new Map();
    if (approvalData.length > 2) {
      for (let i = 2; i < approvalData.length; i++) {
        const row = approvalData[i];
        const requestId = row[approvalRequestIdCol];
        const status = row[approvalStatusCol];
        if (requestId && status === REQUEST_STATUS.APPROVED) {
          approvedRequests.set(requestId, true);
        }
      }
    }
    
    // Get finance completion statuses
    const financeData = financeSheet ? financeSheet.getDataRange().getValues() : [];
    const financeHeaders = financeData[1] || [];
    const financeRequestIdCol = financeHeaders.indexOf('Refund SL Number');
    const financeStatusCol = financeHeaders.indexOf('Status');
    const financeTimestampCol = financeHeaders.indexOf('Status Timestamp String');
    
    // Create map of completed finance requests
    const completedFinanceRequests = new Map();
    if (financeData.length > 2) {
      for (let i = 2; i < financeData.length; i++) {
        const row = financeData[i];
        const requestId = row[financeRequestIdCol];
        const status = row[financeStatusCol];
        const timestamp = row[financeTimestampCol];
        if (requestId && status === REQUEST_STATUS.FINANCE_COMPLETED) {
          completedFinanceRequests.set(requestId, {
            status: status,
            timestamp: timestamp
          });
        }
      }
    }

    // Get product status map - IMPROVED to get the LATEST status for each request
    const productData = productSheet ? productSheet.getDataRange().getValues() : [];
    const productHeaders = productData[0] || [];
    const productRequestIdCol = productHeaders.indexOf('Request ID');
    const productStatusCol = productHeaders.indexOf('Access Revoke Status');
    const productUserIdCol = productHeaders.indexOf('User_Id');
    const productTimestampCol = productHeaders.indexOf('Prod Timestamp String');
    
    // Create map of product statuses - Group by request ID and find latest status
    const productStatuses = new Map();
    if (productData.length > 1) {
      // Group updates by request ID
      const requestUpdates = new Map();
      
      for (let i = 1; i < productData.length; i++) {
        const row = productData[i];
        const requestId = row[productRequestIdCol];
        const status = row[productStatusCol];
        const timestamp = row[productTimestampCol] ? new Date(row[productTimestampCol]) : new Date(0);
        
        if (requestId) {
          if (!requestUpdates.has(requestId)) {
            requestUpdates.set(requestId, []);
          }
          
          requestUpdates.get(requestId).push({
            status: status || REQUEST_STATUS.PENDING,
            timestamp: timestamp
          });
        }
      }
      
      // Find latest status for each request
      requestUpdates.forEach((updates, requestId) => {
        // Sort updates by timestamp (newest first)
        updates.sort((a, b) => b.timestamp - a.timestamp);
        
        // Use the latest status
        if (updates.length > 0) {
          productStatuses.set(requestId, updates[0].status);
        } else {
          productStatuses.set(requestId, REQUEST_STATUS.PENDING);
        }
      });
    }
    
    // Get request data
    const data = requestSheet.getDataRange().getValues();
    const headers = data[1];
    const requestIdCol = headers.indexOf('Request ID');
    const revokeAccessCol = headers.indexOf('Revoke Access');
    const timestampCol = headers.indexOf('Timestamp');
    const courseCol = headers.indexOf('Mistaken Course Name');
    const studentCol = headers.indexOf('Customer Name');
    const enrollmentCol = headers.indexOf('Enrollment Number');
    
    // Initialize counters
    let pending = 0;
    let inProgress = 0;
    let completed = 0;
    let completedToday = 0;
    let totalTime = 0;
    let completedCount = 0;
    
    // Get today's date for comparison
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Process requests
    const requests = data.slice(2) // Skip header rows
      .filter(row => {
        const requestId = row[requestIdCol];
        const needsRevoke = row[revokeAccessCol] === 'Yes';
        // Only show requests that are approved, finance completed and need access revoke
        return approvedRequests.has(requestId) && completedFinanceRequests.has(requestId) && needsRevoke;
      })
      .map(row => {
        const requestId = row[requestIdCol];
        const status = productStatuses.get(requestId) || REQUEST_STATUS.PENDING;
        const timestamp = new Date(row[timestampCol]);
        
        // Count by status
        switch(status.toLowerCase()) {
          case REQUEST_STATUS.PENDING.toLowerCase():
            pending++;
            break;
          case REQUEST_STATUS.PRODUCT_IN_PROGRESS.toLowerCase():
            inProgress++;
            break;
          case REQUEST_STATUS.PRODUCT_COMPLETED.toLowerCase():
            completed++;
            // Check if completed today
            const completionDate = new Date(timestamp);
            completionDate.setHours(0, 0, 0, 0);
            if (completionDate.getTime() === today.getTime()) {
              completedToday++;
            }
            // Calculate resolution time
            const requestDate = new Date(row[timestampCol]);
            totalTime += (timestamp - requestDate) / (1000 * 60 * 60); // Convert to hours
            completedCount++;
            break;
        }
        
        // Return request object with all necessary fields
        return {
          requestId: requestId,
          courseName: row[courseCol],
          studentName: row[studentCol],
          enrollmentNumber: row[enrollmentCol],
          requestDate: formatDateTime(row[timestampCol]),
          status: status
        };
      });
    
    // Calculate average resolution time
    const averageTime = completedCount > 0 ? Math.round(totalTime / completedCount) : 0;
    
    // Get trend data (last 7 days)
    const trendData = getTrendData(data, revokeAccessCol, timestampCol);
    
    return {
      stats: {
        pending: pending,
        inProgress: inProgress,
        completedToday: completedToday,
        averageTime: `${averageTime}h`
      },
      requests: requests,
      chartData: {
        pending: pending,
        inProgress: inProgress,
        completed: completed,
        trendLabels: trendData.labels,
        trendData: trendData.data
      }
    };
  } catch (error) {
    console.error('Error in getProductDashboardData:', error);
    throw new Error('Failed to load dashboard data');
  }
}

// Helper function to get trend data
function getTrendData(data, revokeAccessCol, timestampCol) {
  const days = 7;
  const labels = [];
  const counts = new Array(days).fill(0);
  
  // Get date range
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  // Create labels
  for (let i = days - 1; i >= 0; i--) {
    const date = new Date(today);
    date.setDate(date.getDate() - i);
    labels.push(date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }));
  }
  
  // Count requests per day
  for (let i = 2; i < data.length; i++) {
    const row = data[i];
    if (row[revokeAccessCol] === 'Yes') {
      const timestamp = new Date(row[timestampCol]);
      timestamp.setHours(0, 0, 0, 0);
      const dayDiff = Math.floor((today - timestamp) / (1000 * 60 * 60 * 24));
      if (dayDiff >= 0 && dayDiff < days) {
        counts[days - 1 - dayDiff]++;
      }
    }
  }
  
  return {
    labels: labels,
    data: counts
  };
}

// Function to filter dashboard data
function filterProductDashboard(filters) {
  try {
    const data = getProductDashboardData();
    
    // Apply filters
    data.requests = data.requests.filter(request => {
      // Search filter
      if (filters.search && !Object.values(request).some(value => 
        String(value).toLowerCase().includes(filters.search.toLowerCase())
      )) {
        return false;
      }
      
      // Status filter
      if (filters.status && request.status.toLowerCase() !== filters.status.toLowerCase()) {
        return false;
      }
      
      // Course filter
      if (filters.course && request.courseName !== filters.course) {
        return false;
      }
      
      // Date filter
      if (filters.date) {
        const filterDate = new Date(filters.date);
        const requestDate = new Date(request.requestDate);
        if (filterDate.toDateString() !== requestDate.toDateString()) {
          return false;
        }
      }
      
      return true;
    });
    
    return data;
  } catch (error) {
    console.error('Error in filterProductDashboard:', error);
    throw new Error('Failed to filter dashboard data');
  }
}

// Add function to set up product tracking sheet
function setupProductSheet(sheet) {
  // Define headers consistently with how they're used in updateProductStatus
  const headers = [
    'Request ID',
    'Timestamp',
    'Mistaken Course Name',
    'Student Name',
    'Enrollment Number',
    'Access Type',          
    'Access Revoke Status', // IMPORTANT: Keep name consistent throughout code
    'User_Id',              // IMPORTANT: This column name must match exactly
    'Remarks',
    'Prod Timestamp String' // String version of timestamp for display
  ];

  try {
    console.log('Setting up Product sheet with headers:', headers);
    
    // Set title
    sheet.getRange(1, 1).setValue('Product Team Access Revoke Tracking');
    sheet.getRange(1, 1, 1, headers.length).merge();
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    
    // Set headers
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#4286f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 150);  // Request ID
    sheet.setColumnWidth(2, 180);  // Timestamp
    sheet.setColumnWidth(3, 200);  // Mistaken Course Name
    sheet.setColumnWidth(4, 180);  // Student Name
    sheet.setColumnWidth(5, 150);  // Enrollment Number
    sheet.setColumnWidth(6, 120);  // Access Type
    sheet.setColumnWidth(7, 150);  // Access Revoke Status
    sheet.setColumnWidth(8, 180);  // User_Id - increased width for email addresses
    sheet.setColumnWidth(9, 250);  // Remarks
    sheet.setColumnWidth(10, 200); // Prod Timestamp String
    
    // Freeze header rows
    sheet.setFrozenRows(2);
    
    // Add filters
    sheet.getRange(2, 1, 1, headers.length).createFilter();
    
    // Set date format for Timestamp column
    sheet.getRange(3, 2, sheet.getMaxRows() - 2).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    
    // Add data validation for status column
    const statusRange = sheet.getRange(3, 7, sheet.getMaxRows() - 2, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'In Progress', 'Completed'])
      .build();
    statusRange.setDataValidation(statusRule);
    
    // Add formula for Prod Timestamp String column
    const prodTimestampStringFormula = '=TEXT(B3, "dd/MM/yyyy HH:mm:ss")';
    sheet.getRange(3, 10).setFormula(prodTimestampStringFormula);
    
    console.log('Product sheet setup complete');
    return true;
  } catch (error) {
    console.error('Error setting up product sheet:', error);
    console.error('Error stack:', error.stack);
    return false;
  }
}

// Helper function to find column index in a case-insensitive way
function findColumnIndex(headers, name) {
  const index = headers.findIndex(header => 
    String(header).trim().toLowerCase() === name.toLowerCase()
  );
  
  // If exact match not found, try more flexible matching
  if (index === -1) {
    return headers.findIndex(header => 
      String(header).trim().toLowerCase().includes(name.toLowerCase())
    );
  }
  
  return index;
}

// Function to get request details including finance updates
function getRequestDetails(requestId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const requestSheet = ss.getSheetByName(SHEET_NAME);
  const financeSheet = ss.getSheetByName(FINANCE_SHEET_NAME);
  const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
  
  // Get request data headers and values
  const requestData = requestSheet.getDataRange().getValues();
  const requestHeaders = requestData[1];  // Second row contains headers
  
  // Find column indices for request sheet
  const requestIdCol = requestHeaders.indexOf('Request ID');
  const requesterCol = requestHeaders.indexOf('Requester');
  const requesterEmailCol = requestHeaders.indexOf('Requester Email');
  const verticalCol = requestHeaders.indexOf('Vertical');
  const enrollmentNumberCol = requestHeaders.indexOf('Enrollment Number');
  const customerNameCol = requestHeaders.indexOf('Customer Name');
  const purchaseDateCol = requestHeaders.indexOf('Purchase Date');
  const paymentIdCol = requestHeaders.indexOf('Payment ID');
  const purchasedAmountCol = requestHeaders.indexOf('Purchased Amount');
  const refundTypeCol = requestHeaders.indexOf('Refund Type');
  const mistakenCourseCol = requestHeaders.indexOf('Mistaken Course Name');
  const mistakenPhaseCol = requestHeaders.indexOf('Mistaken Phase Name');
  const intendedCourseCol = requestHeaders.indexOf('Intended Course Name');
  const intendedPhaseCol = requestHeaders.indexOf('Intended Phase Name');
  const mistakenPriceCol = requestHeaders.indexOf('Mistaken Course Price');
  const wrongPhoneCol = requestHeaders.indexOf('Wrong Phone');
  const correctPhoneCol = requestHeaders.indexOf('Correct Phone');
  const senderBkashCol = requestHeaders.indexOf('Sender bKash');
  const transactionTimeStringCol = requestHeaders.indexOf('Transaction Time String');
  const sentAmountCol = requestHeaders.indexOf('Sent Amount');
  const refundAmountCol = requestHeaders.indexOf('Refund Amount');
  const revokeAccessCol = requestHeaders.indexOf('Revoke Access');
  const refundReasonCol = requestHeaders.indexOf('Refund Reason');
  const documentFolderCol = requestHeaders.indexOf('Document Folder');
  const timestampCol = requestHeaders.indexOf('Timestamp String');
  const purchaseDateStringCol = requestHeaders.indexOf('Purchase Date String');

  // Find request row
  let requestRow = null;
  for (let i = 2; i < requestData.length; i++) {
    if (String(requestData[i][requestIdCol]) === String(requestId)) {
      requestRow = requestData[i];
      break;
    }
  }

  if (!requestRow) {
    throw new Error('Request not found');
  }

  // Get product updates
  let productUpdates = [];
  if (productSheet) {
    const productData = productSheet.getDataRange().getValues();
    const productHeaders = productData[1] || productData[0]; // Try row 2 first, then row 1
    
    // Use the findColumnIndex helper function to find columns in a case-insensitive way
    const productRequestIdCol = findColumnIndex(productHeaders, 'Request ID');
    const productStatusCol = findColumnIndex(productHeaders, 'Access Revoke Status');
    const productUserIdCol = findColumnIndex(productHeaders, 'User_Id');
    const productRemarksCol = findColumnIndex(productHeaders, 'Remarks');
    const productTimestampCol = findColumnIndex(productHeaders, 'Prod Timestamp String');

    console.log('Product headers:', productHeaders);
    console.log('Product column indices:', {
      requestId: productRequestIdCol,
      status: productStatusCol,
      user_id: productUserIdCol,
      remarks: productRemarksCol,
      timestamp: productTimestampCol
    });

    // Add validation for column indices
    if (productRequestIdCol === -1) {
      console.warn("'Request ID' column not found in Product sheet");
    }
    if (productUserIdCol === -1) {
      console.warn("'User_Id' column not found in Product sheet");
    }

    // Use a start index that skips the header row(s)
    const startIndex = productData[1] ? 2 : 1;
    
    productUpdates = productData.slice(startIndex)
      .filter(row => String(row[productRequestIdCol]) === String(requestId))
      .map(row => {
        // Safely access columns with validation
        const userId = (productUserIdCol >= 0 && row.length > productUserIdCol) ? row[productUserIdCol] : '';
        const remarks = (productRemarksCol >= 0 && row.length > productRemarksCol) ? row[productRemarksCol] : '';
        const timestamp = (productTimestampCol >= 0 && row.length > productTimestampCol) ? 
          formatDateTime(row[productTimestampCol]) : '';
        const status = (productStatusCol >= 0 && row.length > productStatusCol) ? 
          row[productStatusCol] || 'Pending' : 'Pending';
        
        return {
        type: 'Product',
          timestamp: timestamp,
          status: status,
          user_id: userId,
          remarks: remarks,
          updatedBy: userId || 'Product Team'
        };
      })
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    console.log('Product updates found:', productUpdates.length);
  }

  // Get approval updates
  let approvalUpdates = [];
  const approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
  if (approvalSheet) {
    const approvalData = approvalSheet.getDataRange().getValues();
    const approvalHeaders = approvalData[1] || approvalData[0]; // Try row 2 first, then row 1
    const approvalRequestIdCol = approvalHeaders.indexOf('Request ID');
    const approvalStatusCol = approvalHeaders.indexOf('Approval Status');
    const approvalRemarksCol = approvalHeaders.indexOf('Approver Remarks');
    const approvalTimestampCol = 0; // First column is timestamp

    console.log('Approval headers:', approvalHeaders);
    console.log('Approval column indices:', {
      requestId: approvalRequestIdCol,
      status: approvalStatusCol,
      remarks: approvalRemarksCol,
      timestamp: approvalTimestampCol
    });

    // Use a start index that skips the header row(s)
    const startIndex = approvalData[1] ? 2 : 1;
    
    approvalUpdates = approvalData.slice(startIndex)
      .filter(row => String(row[approvalRequestIdCol]) === String(requestId))
      .map(row => ({
        type: 'Approval',
        timestamp: formatDateTime(row[approvalTimestampCol]),
        status: row[approvalStatusCol] || 'Pending',
        remarks: row[approvalRemarksCol] || '',
        updatedBy: 'Refund Approver'
      }))
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    console.log('Approval updates found:', approvalUpdates.length);
  }

  // Get finance updates
  let financeUpdates = [];
  if (financeSheet) {
    const financeData = financeSheet.getDataRange().getValues();
    const financeHeaders = financeData[1];
    const financeRequestIdCol = financeHeaders.indexOf('Refund SL Number');
    const financeStatusCol = financeHeaders.indexOf('Status');
    const financeTimestampCol = 0;  // First column is timestamp
    const financeRemarksCol = financeHeaders.indexOf('Finance Remarks');

    financeUpdates = financeData.slice(2)  // Skip header rows
      .filter(row => String(row[financeRequestIdCol]) === String(requestId))
        .map(row => ({
          type: 'Finance',
          timestamp: formatDateTime(row[financeTimestampCol]),
          status: row[financeStatusCol] || 'Pending',
          remarks: row[financeRemarksCol] || '',
          updatedBy: 'Finance Team'
        }))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    console.log('Finance updates found:', financeUpdates.length);
  }

  // Combine all updates and sort by timestamp
  const allUpdates = [...approvalUpdates, ...financeUpdates, ...productUpdates]
    .filter(update => update.timestamp) // Filter out any updates with invalid timestamps
    .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

  // Log updates for debugging
  console.log('All Updates:', allUpdates);

  // Log transaction-related data for debugging
  console.log('Transaction-related data in getRequestDetails:');
  console.log('- senderBkash column index:', senderBkashCol, 'value:', requestRow[senderBkashCol]);
  console.log('- transactionTimeString column index:', transactionTimeStringCol, 'value:', requestRow[transactionTimeStringCol]);
  console.log('- sentAmount column index:', sentAmountCol, 'value:', requestRow[sentAmountCol]);

  // Construct response object
  return {
    requestId: requestRow[requestIdCol],
    requester: requestRow[requesterCol],
    requesterEmail: requestRow[requesterEmailCol],
    vertical: requestRow[verticalCol],
    enrollmentNumber: requestRow[enrollmentNumberCol],
    customerContact: requestRow[customerNameCol],
    studentName: requestRow[customerNameCol], // Add explicit studentName for consistency
    requestDate: formatDateTime(new Date(requestRow[timestampCol])),
    purchaseDate: requestRow[purchaseDateStringCol],
    paymentId: requestRow[paymentIdCol],
    purchasedAmount: requestRow[purchasedAmountCol],
    refundType: requestRow[refundTypeCol],
    mistakenCourse: requestRow[mistakenCourseCol],
    mistakenPhase: requestRow[mistakenPhaseCol],
    intendedCourse: requestRow[intendedCourseCol],
    intendedPhase: requestRow[intendedPhaseCol],
    mistakenPrice: requestRow[mistakenPriceCol],
    wrongPhone: requestRow[wrongPhoneCol],
    correctPhone: requestRow[correctPhoneCol],
    senderBkash: requestRow[senderBkashCol],
    transactionTime: requestRow[transactionTimeStringCol], // Use string column directly
    sentAmount: requestRow[sentAmountCol],
    refundAmount: requestRow[refundAmountCol],
    revokeAccess: requestRow[revokeAccessCol],
    refundReason: requestRow[refundReasonCol],
    documentFolder: requestRow[documentFolderCol],
    updates: allUpdates
  };
}

// Function to update product status\

function updateProductStatus(updateData) {
  try {
    console.log('Starting product status update with data:', updateData);
    
    // Validate required fields
    if (!updateData.requestId) {
      console.error('Missing required field: requestId');
      throw new Error('Missing required field: requestId');
    }
    
    if (!updateData.accessRevokeStatus) {
      console.error('Missing required field: accessRevokeStatus');
      throw new Error('Missing required field: accessRevokeStatus');
    }
    
    // Log critical fields for debugging
    console.log('Critical fields check:');
    console.log('- Request ID:', updateData.requestId);
    console.log('- Access Revoke Status:', updateData.accessRevokeStatus);
    console.log('- User_Id:', updateData.user_id);
    console.log('- Mistaken Course Name:', updateData.courseName);
    console.log('- Student Name:', updateData.studentName);
    console.log('- Enrollment Number:', updateData.enrollmentNumber);
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Only update the main request sheet if explicitly requested
    if (updateData.shouldUpdateMainSheet) {
      console.log('Updating main Refund Requests sheet');
      const requestSheet = ss.getSheetByName(SHEET_NAME);
      const requestData = requestSheet.getDataRange().getValues();
      const requestHeaders = requestData[1];
      
      const requestIdCol = findColumnIndex(requestHeaders, 'Request ID');
      const accessRevokeStatusCol = findColumnIndex(requestHeaders, 'Access Revoke Status');
      const accessRevokeDateCol = findColumnIndex(requestHeaders, 'Access Revoke Date');
      const accessRevokeByCol = findColumnIndex(requestHeaders, 'Access Revoke By');
      const accessRevokeRemarksCol = findColumnIndex(requestHeaders, 'Access Revoke Remarks');
      
      // Log column indices for debugging
      console.log('Main sheet column indices:', {
        requestId: requestIdCol,
        accessRevokeStatus: accessRevokeStatusCol,
        accessRevokeDate: accessRevokeDateCol,
        accessRevokeBy: accessRevokeByCol,
        accessRevokeRemarks: accessRevokeRemarksCol
      });
      
      // Find and update the request row
      let requestRowIndex = -1;
      for (let i = 2; i < requestData.length; i++) {
        if (String(requestData[i][requestIdCol]) === String(updateData.requestId)) {
          requestRowIndex = i + 1;  // Convert to 1-based index
          break;
        }
      }
    
      if (requestRowIndex === -1) {
        console.log('Request not found in main sheet, skipping main sheet update');
      } else {
        // Update request sheet with safety checks
        const timestamp = new Date();
        if (accessRevokeStatusCol >= 0) {
          requestSheet.getRange(requestRowIndex, accessRevokeStatusCol + 1).setValue(updateData.accessRevokeStatus);
        }
        if (accessRevokeDateCol >= 0) {
          requestSheet.getRange(requestRowIndex, accessRevokeDateCol + 1).setValue(timestamp);
        }
        if (accessRevokeByCol >= 0) {
          requestSheet.getRange(requestRowIndex, accessRevokeByCol + 1).setValue(updateData.user_id);
        }
        if (accessRevokeRemarksCol >= 0) {
          requestSheet.getRange(requestRowIndex, accessRevokeRemarksCol + 1).setValue(updateData.remarks);
        }
        console.log('Updated main Refund Requests sheet');
      }
    } else {
      console.log('Skipping main Refund Requests sheet update as requested');
    }
    
    // Get or create Product Access Tracking sheet
    let productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    if (!productSheet) {
      console.log('Creating new Product Access Tracking sheet');
      productSheet = ss.insertSheet(PRODUCT_SHEET_NAME);
      setupProductSheet(productSheet);
    }
    
    // Get existing data and headers
    const productData = productSheet.getDataRange().getValues();
    // Make sure we have at least header rows
    if (productData.length < 2) {
      console.log('Product sheet missing headers, setting up sheet');
      setupProductSheet(productSheet);
    }
    
    // Re-read headers after setup
    const headerRow = productSheet.getRange(2, 1, 1, productSheet.getLastColumn()).getValues()[0];
    
    // Find column indices using the case-insensitive function
    const columnIndices = {
      requestId: findColumnIndex(headerRow, 'Request ID'),
      timestamp: findColumnIndex(headerRow, 'Timestamp'),
      courseName: findColumnIndex(headerRow, 'Mistaken Course Name'),
      studentName: findColumnIndex(headerRow, 'Student Name'),
      enrollmentNumber: findColumnIndex(headerRow, 'Enrollment Number'),
      accessType: findColumnIndex(headerRow, 'Access Type'),
      accessStatus: findColumnIndex(headerRow, 'Access Revoke Status'),
      userId: findColumnIndex(headerRow, 'User_Id'),
      remarks: findColumnIndex(headerRow, 'Remarks')
    };
    
    // Log available columns and indices for debugging
    console.log('Available columns in product sheet:', headerRow);
    console.log('Product sheet column indices:', columnIndices);
    
    // Check for missing columns and log warnings
    Object.entries(columnIndices).forEach(([key, index]) => {
      if (index === -1) {
        console.warn(`'${key}' column not found in Product sheet`);
      }
    });
    
    // Create the row data array (initialize with enough elements)
    const maxIndex = Math.max(...Object.values(columnIndices).filter(i => i >= 0));
    const rowData = Array(maxIndex + 1).fill('');
    
    // Set timestamp 
    const timestamp = new Date();
    
    // Set values at correct indices with safety checks
    if (columnIndices.requestId >= 0) rowData[columnIndices.requestId] = updateData.requestId;
    if (columnIndices.timestamp >= 0) rowData[columnIndices.timestamp] = timestamp;
    if (columnIndices.courseName >= 0) rowData[columnIndices.courseName] = updateData.courseName || '';
    if (columnIndices.studentName >= 0) rowData[columnIndices.studentName] = updateData.studentName || '';
    if (columnIndices.enrollmentNumber >= 0) rowData[columnIndices.enrollmentNumber] = updateData.enrollmentNumber || '';
    if (columnIndices.accessType >= 0) rowData[columnIndices.accessType] = 'Full';  // Default to Full
    if (columnIndices.accessStatus >= 0) rowData[columnIndices.accessStatus] = updateData.accessRevokeStatus;
    if (columnIndices.userId >= 0) rowData[columnIndices.userId] = updateData.user_id || '';
    if (columnIndices.remarks >= 0) rowData[columnIndices.remarks] = updateData.remarks || '';
    
    console.log('Prepared row data:', rowData);
    
    // Always append a new row to create a log history
    productSheet.appendRow(rowData);
    console.log('Added new log entry to Product Access Tracking sheet');
    
    // Return success response
    return {
      success: true,
      message: 'Status updated successfully'
    };
    
  } catch (error) {
    console.error('Error updating product status:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Function to get data for refund approver dashboard
function getRefundApproverData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
    
    // Get request data
    const data = requestSheet.getDataRange().getValues();
    const headers = data[1];  // Headers are in row 2
    
    // Get column indices
    const requestIdCol = headers.indexOf('Request ID');
    const timestampCol = headers.indexOf('Timestamp');
    const requesterCol = headers.indexOf('Requester');
    const customerNameCol = headers.indexOf('Customer Name');
    const refundTypeCol = headers.indexOf('Refund Type');
    const refundAmountCol = headers.indexOf('Refund Amount');
    const statusCol = headers.indexOf('Status');
    
    // Get approval statuses
    const approvalData = approvalSheet ? approvalSheet.getDataRange().getValues() : [];
    const approvalHeaders = approvalData[1] || [];
    const approvalRequestIdCol = approvalHeaders.indexOf('Request ID');
    const approvalStatusCol = approvalHeaders.indexOf('Approval Status');
    const approvalTimestampCol = approvalHeaders.indexOf('Approval Timestamp String');
    
    // Create map of latest approval statuses
    const approvalStatuses = new Map();
    if (approvalData.length > 2) {
      for (let i = 2; i < approvalData.length; i++) {
        const row = approvalData[i];
        const requestId = row[approvalRequestIdCol];
        const status = row[approvalStatusCol];
        const timestamp = row[approvalTimestampCol];
        if (requestId && status) {
          if (!approvalStatuses.has(requestId) || 
              new Date(timestamp) > new Date(approvalStatuses.get(requestId).timestamp)) {
            approvalStatuses.set(requestId, {
              status: status,
              timestamp: timestamp
            });
          }
        }
      }
    }
    
    // Initialize counters
    let pending = 0;
    let approved = 0;
    let rejected = 0;
    let approvedToday = 0;
    let totalAmount = 0;
    
    // Get today's date for comparison
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    // Process requests
    const requests = [];
    const refundTypes = new Set();
    
    // Process all requests starting from row 3 (after headers)
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      const requestId = row[requestIdCol];
      const approvalInfo = approvalStatuses.get(requestId);
      const status = approvalInfo ? approvalInfo.status : REQUEST_STATUS.PENDING;
      const timestamp = approvalInfo ? new Date(approvalInfo.timestamp) : new Date(row[timestampCol]);
      const refundType = row[refundTypeCol] || '';
      
      // Add to refund types set
      if (refundType) refundTypes.add(refundType);
      
      // Count by status
      switch(status.toLowerCase()) {
        case REQUEST_STATUS.PENDING.toLowerCase():
          pending++;
          break;
        case REQUEST_STATUS.APPROVED.toLowerCase():
          approved++;
          // Check if approved today
          const approvalDate = new Date(timestamp);
          approvalDate.setHours(0, 0, 0, 0);
          if (approvalDate.getTime() === today.getTime()) {
            approvedToday++;
          }
          // Add to total amount if approved
          totalAmount += parseFloat(row[refundAmountCol]) || 0;
          break;
        case REQUEST_STATUS.REJECTED.toLowerCase():
          rejected++;
          break;
      }
      
      // Add to requests array
      requests.push({
        requestId: requestId,
        requestDate: formatDateTime(row[timestampCol]),
        requesterName: row[requesterCol],
        customerName: row[customerNameCol],
        refundType: refundType,
        refundAmount: row[refundAmountCol] ? parseFloat(row[refundAmountCol]).toFixed(2) : '0.00',
        status: status
      });
    }
    
    // Calculate average amount
    const averageAmount = approved > 0 ? totalAmount / approved : 0;
    
    return {
      stats: {
        pending: pending,
        approved: approved,
        rejected: rejected,
        approvedToday: approvedToday,
        totalAmount: totalAmount.toFixed(2),
        averageAmount: averageAmount.toFixed(2)
      },
      requests: requests,
      refundTypes: Array.from(refundTypes)
    };
  } catch (error) {
    console.error('Error in getRefundApproverData:', error);
    throw new Error('Failed to load approver dashboard data');
  }
}

// Function to set up refund approval sheet
function setupRefundApprovalSheet(sheet) {
  const headers = [
    'Timestamp',
    'Request ID',
    'Enrollment Number',
    'Approval Status',
    'Approver Remarks'
  ];

  try {
    // Set title
    sheet.getRange(1, 1).setValue('Refund Approval Updates');
    sheet.getRange(1, 1, 1, headers.length).merge();
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    
    // Set headers
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#4286f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 180);  // Timestamp
    sheet.setColumnWidth(2, 150);  // Request ID
    sheet.setColumnWidth(3, 180);  // Enrollment Number
    sheet.setColumnWidth(4, 150);  // Approval Status
    sheet.setColumnWidth(5, 300);  // Approver Remarks
    
    // Freeze header rows
    sheet.setFrozenRows(2);
    
    // Add filters
    sheet.getRange(2, 1, 1, headers.length).createFilter();
    
    return true;
  } catch (error) {
    console.error('Error setting up refund approval sheet:', error);
    return false;
  }
}

// Function to process refund approval update
function processRefundApproval(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Update main request sheet status
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const requestData = requestSheet.getDataRange().getValues();
    const requestHeaders = requestData[1];  // Headers are in row 2
    
    const requestIdCol = requestHeaders.indexOf('Request ID');
    const statusCol = requestHeaders.indexOf('Status');
    const enrollmentCol = requestHeaders.indexOf('Enrollment Number');
    
    // Find and update the request row
    let enrollmentNumber = '';
    let rowFound = false;
    for (let i = 2; i < requestData.length; i++) {
      if (String(requestData[i][requestIdCol]) === String(data.requestId)) {
        const rowNumber = i + 1;  // Convert to 1-based index
        requestSheet.getRange(rowNumber, statusCol + 1).setValue(data.approvalStatus);
        enrollmentNumber = requestData[i][enrollmentCol];
        rowFound = true;
        break;
      }
    }
    
    if (!rowFound) {
      throw new Error('Request ID not found');
    }
    
    // 2. Add entry to Refund Approval Updates sheet
    let approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
    if (!approvalSheet) {
      approvalSheet = ss.insertSheet(REFUND_APPROVAL_SHEET_NAME);
      setupRefundApprovalSheet(approvalSheet);
    }
    
    // Prepare row data
    const timestamp = new Date();
    const rowData = [
      timestamp,
      data.requestId,
      enrollmentNumber,
      data.approvalStatus,
      data.remarks || ''
    ];
    
    // Append to approval sheet
    approvalSheet.appendRow(rowData);
    
    return {
      success: true,
      message: 'Approval status updated successfully'
    };
    
  } catch (error) {
    console.error('Error processing refund approval:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Function to get approval history for a request
function getApprovalHistory(requestId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
    
    if (!approvalSheet) {
      return [];
    }
    
    const data = approvalSheet.getDataRange().getValues();
    const headers = data[1];  // Headers are in row 2
    
    const requestIdCol = headers.indexOf('Request ID');
    const timestampCol = 0;  // First column is timestamp
    const statusCol = headers.indexOf('Approval Status');
    const remarksCol = headers.indexOf('Approver Remarks');
    
    // Get all updates for this request ID
    const updates = data.slice(2)  // Skip header rows
      .filter(row => String(row[requestIdCol]) === String(requestId))
        .map(row => ({
          timestamp: formatDateTime(row[timestampCol]),
          status: row[statusCol] || '',
          remarks: row[remarksCol] || ''
        }))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));  // Sort by timestamp descending
    
    return updates;
  } catch (error) {
    console.error('Error getting approval history:', error);
    return [];
  }
}

// Helper function to include HTML templates
/**
 * Includes an HTML template file with optional data
 * @param {string} filename - The HTML template file to include
 * @param {Object} data - Optional data to pass to the template
 * @return {string} The evaluated HTML content
 */
function include(filename, data) {
  const template = HtmlService.createTemplateFromFile(filename);
  
  // Pass data to the template if provided
  if (data) {
    Object.keys(data).forEach(key => {
      template[key] = data[key];
    });
  }
  
  return template.evaluate().getContent();
}

// Create access denied page
function createAccessDeniedPage() {
  return HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
      <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.2/css/all.min.css" rel="stylesheet">
      <style>
        body {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
          background-color: #f8f9fa;
          display: flex;
          align-items: center;
          justify-content: center;
          min-height: 100vh;
          padding: 20px;
        }
        .error-container {
          max-width: 500px;
          text-align: center;
          padding: 2rem;
          background: white;
          border-radius: 12px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .error-icon {
          font-size: 4rem;
          color: #dc3545;
          margin-bottom: 1rem;
        }
      </style>
    </head>
    <body>
      <div class="error-container">
        <div class="error-icon">⚠️</div>
        <h2 class="mb-4">Access Denied</h2>
        <p class="mb-4">You do not have permission to access this page. Please contact your administrator if you believe this is an error.</p>
        <div class="d-flex justify-content-center gap-3">
          <a href="${getScriptUrl()}?page=menu" class="btn btn-primary">Go to Home</a>
          <a href="${getScriptUrl()}?page=login" class="btn btn-outline-secondary">Login Again</a>
        </div>
      </div>
    </body>
    </html>
  `)
  .setTitle('Access Denied')
  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Update or add this function to handle login authentication
function processLogin(credentials) {
  try {
    console.log('Processing login for:', credentials.email);
    const { email, password, rememberMe } = credentials;
    
    // Basic validation
    if (!email || !password) {
      return { success: false, message: 'Email and password are required' };
    }
    
    // Authenticate user
    const user = authenticateUser(email, password);
    
    if (!user) {
      console.log('Authentication failed for:', email);
      return { 
        success: false, 
        message: 'Invalid email or password',
        reason: 'authentication_failed'
      };
    }
    
    if (user.error) {
      // Only pass through specific messages for inactive accounts
      if (user.reason === 'inactive') {
        return {
          success: false,
          message: user.message,
          reason: 'inactive'
        };
      }
      // For all other authentication failures, use generic message
      return {
        success: false,
        message: 'Invalid email or password',
        reason: 'authentication_failed'
      };
    }
    
    // Generate session token
    const sessionToken = generateSessionToken(user.id, rememberMe);
    console.log('Generated session token:', sessionToken);
    
    // Get user access rights
    const userAccess = {
      id: user.id,
      name: user.name,
      email: user.email,
      sessionToken: sessionToken,
      access: user.access || []
    };
    
    // Store session in cache
    storeUserSession(sessionToken, userAccess);
    
    // Get script URL for constructing the redirect URL
    const scriptUrl = getScriptUrl();
    console.log('Script URL for redirect:', scriptUrl);
    
    // Construct the complete redirect URL
    const redirectUrl = `${scriptUrl}?page=menu&token=${sessionToken}`;
    console.log('Constructed redirect URL:', redirectUrl);
    
    // Return success with redirect URL explicitly constructed
    return {
      success: true,
      sessionToken: sessionToken,
      redirectUrl: redirectUrl,
      message: 'Login successful'
    };
  } catch (error) {
    console.error('Login error:', error);
    return { 
      success: false, 
      message: 'An error occurred during login',
      reason: 'system_error'
    };
  }
}

// Authentication system functions
// User credentials are now stored in the Credentials sheet

// Session cache
const SESSION_DURATION = 24 * 60 * 60 * 1000; // 24 hours in milliseconds
const userSessions = {};

/**
 * Authenticates a user by email and password
 * @param {string} email - User email
 * @param {string} password - User password
 * @return {Object|null} User object if authenticated, null otherwise
 */
function authenticateUser(email, password) {
  console.log('Authenticating user:', email);
  
  try {
    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    console.log('Found spreadsheet:', ss.getName());
    
    // List all sheet names for debugging
    const allSheets = ss.getSheets();
    console.log('Available sheets:', allSheets.map(sheet => sheet.getName()));
    
    // Get the credentials sheet
    const sheet = ss.getSheetByName('Credentials');
    if (!sheet) {
      console.error('Credentials sheet not found');
      return {
        error: true,
        message: 'System configuration error',
        reason: 'system_error'
      };
    }
    
    // Get all data from the credentials sheet
    const data = sheet.getDataRange().getValues();
    console.log('Found data rows:', data.length);
    
    // Extract header row to identify column indices
    const headerRow = data[0];
    console.log('Header row:', headerRow);
    
    // Find column indices
    const userNameIndex = findColumnIndex(headerRow, 'User Name');
    const emailIndex = findColumnIndex(headerRow, 'Email');
    const passwordIndex = findColumnIndex(headerRow, 'Password');
    const accessToIndex = findColumnIndex(headerRow, 'Access to');
    const statusIndex = findColumnIndex(headerRow, 'Credential Status');
    const roleIndex = findColumnIndex(headerRow, 'Role');
    
    if (emailIndex === -1 || passwordIndex === -1 || accessToIndex === -1 || statusIndex === -1) {
      console.error('Required columns not found');
      return {
        error: true,
        message: 'System configuration error',
        reason: 'system_error'
      };
    }
    
    // First find all rows matching the email
    const emailMatches = data.slice(1).filter(row => {
      const rowEmail = emailIndex >= 0 && row.length > emailIndex ? String(row[emailIndex]).toLowerCase() : '';
      return rowEmail === email.toLowerCase();
    });
    
    if (emailMatches.length === 0) {
      // No matching email found - but we don't want to reveal this
      console.log('No matching credentials found');
      return {
        error: true,
        message: 'Invalid email or password',
        reason: 'authentication_failed'
      };
    }
    
    // Check password match for any of the rows
    const passwordMatches = emailMatches.filter(row => {
      const rowPassword = passwordIndex >= 0 && row.length > passwordIndex ? String(row[passwordIndex]) : '';
      return rowPassword === password;
    });
    
    if (passwordMatches.length === 0) {
      // Wrong password - but we don't want to reveal this
      console.log('Invalid credentials');
      return {
        error: true,
        message: 'Invalid email or password',
        reason: 'authentication_failed'
      };
    }
    
    // Get all active rows for this user
    const activeRows = passwordMatches.filter(row => {
      const rowStatus = statusIndex >= 0 && row.length > statusIndex ? String(row[statusIndex]) : '';
      return rowStatus === 'Active';
    });
    
    if (activeRows.length === 0) {
      console.log('Account inactive:', email);
      return {
        error: true,
        message: 'Account is inactive',
        reason: 'inactive'
      };
    }
    
    // Collect all access permissions and roles from active rows
    const accessPermissions = [];
    const roles = new Set();
    let userName = email; // default to email if no name found
    
    activeRows.forEach(row => {
      // Get user name (use the first one found)
      if (!userName || userName === email) {
        userName = userNameIndex >= 0 && row.length > userNameIndex ? row[userNameIndex] : email;
      }
      
      // Get access permission
      const access = accessToIndex >= 0 && row.length > accessToIndex ? String(row[accessToIndex]).trim() : '';
      if (access) {
        accessPermissions.push(access);
      }
      
      // Get role
      const role = roleIndex >= 0 && row.length > roleIndex ? String(row[roleIndex]).trim() : '';
      if (role) {
        roles.add(role);
      }
    });
    
    // Create user object with combined permissions
    return {
      id: email.toLowerCase(),
      name: userName,
      email: email.toLowerCase(),
      access: [...new Set(accessPermissions)], // Remove duplicates
      roles: Array.from(roles), // Convert Set to Array
      primaryRole: Array.from(roles)[0] || 'User' // Use first role as primary
    };
    
  } catch (error) {
    console.error('Error during authentication:', error);
    return {
      error: true,
      message: 'Authentication error',
      reason: 'system_error'
    };
  }
}

/**
 * Generates a unique session token for authenticated user
 * @param {string} userId - User ID
 * @param {boolean} rememberMe - Whether to extend session duration
 * @return {string} Session token
 */
function generateSessionToken(userId, rememberMe) {
  // Generate random token
  const token = Utilities.getUuid();
  
  // Set expiration time
  const expiresAt = new Date();
  expiresAt.setTime(expiresAt.getTime() + (rememberMe ? SESSION_DURATION * 7 : SESSION_DURATION)); // 7 days if remember me
  
  console.log('Generated session token for user:', userId, 'expires at:', expiresAt);
  return token;
}

/**
 * Stores user session information
 * @param {string} token - Session token
 * @param {Object} userAccess - User access information
 */
function storeUserSession(token, userAccess) {
  // Set expiration time (24 hours from now by default)
  const expiresAt = new Date();
  expiresAt.setTime(expiresAt.getTime() + SESSION_DURATION);
  
  // Create session object
  const sessionData = {
    userAccess: userAccess,
    expiresAt: expiresAt.getTime() // Store as milliseconds timestamp
  };
  
  // Store in Script Properties - use a prefix to identify session tokens
  const sessionKey = 'session_' + token;
  PropertiesService.getScriptProperties().setProperty(
    sessionKey, 
    JSON.stringify(sessionData)
  );
  
  console.log('Stored session for token:', token, 'expires at:', expiresAt);
  return true;
}

/**
 * Validates if a session token is valid and not expired
 * @param {string} token - Session token to validate
 * @return {boolean} True if session is valid, false otherwise
 */
function validateSession(token) {
  if (!token) return false;
  
  // Get from Script Properties
  const sessionKey = 'session_' + token;
  const sessionJson = PropertiesService.getScriptProperties().getProperty(sessionKey);
  
  if (!sessionJson) {
    console.log('Session not found for token:', token);
    return false;
  }
  
  try {
    const session = JSON.parse(sessionJson);
    const now = new Date().getTime();
    
    // Check if session has expired
    if (now > session.expiresAt) {
      console.log('Session expired for token:', token);
      // Clean up expired session
      PropertiesService.getScriptProperties().deleteProperty(sessionKey);
      return false;
    }
    
    console.log('Session valid for token:', token);
    return true;
  } catch (error) {
    console.error('Error validating session:', error);
    return false;
  }
}

/**
 * Gets user access information from session token
 * @param {string} token - Session token
 * @return {Object} User access information
 */
function getUserAccess(token) {
  if (!token) return {};
  
  // Get from Script Properties
  const sessionKey = 'session_' + token;
  const sessionJson = PropertiesService.getScriptProperties().getProperty(sessionKey);
  
  if (!sessionJson) return {};
  
  try {
    const session = JSON.parse(sessionJson);
    return session.userAccess || {};
  } catch (error) {
    console.error('Error getting user access:', error);
    return {};
  }
}

/**
 * Checks if user has access to a specific portal
 * @param {string} token - Session token
 * @param {string} portal - Portal to check access for
 * @return {boolean} True if user has access, false otherwise
 */
function hasAccess(token, portal) {
  if (!token) return false;
  
  // Get session from Script Properties
  const sessionKey = 'session_' + token;
  const sessionData = PropertiesService.getScriptProperties().getProperty(sessionKey);
  
  if (!sessionData) {
    console.log('Session not found for token:', token);
    return false;
  }
  
  try {
    const userData = JSON.parse(sessionData);
    const userAccess = userData.userAccess;
    return userAccess && userAccess.access && userAccess.access.includes(portal);
  } catch (error) {
    console.error('Error parsing session data:', error);
    return false;
  }
}

/**
 * Logs out a user by invalidating their session token
 * @param {string} token - Session token to invalidate
 * @return {Object} Success/failure response
 */
function logoutUser(token) {
  if (!token) return { success: false, message: 'Invalid session token' };
  
  // Delete from Script Properties
  const sessionKey = 'session_' + token;
  const sessionExists = PropertiesService.getScriptProperties().getProperty(sessionKey);
  
  if (sessionExists) {
    PropertiesService.getScriptProperties().deleteProperty(sessionKey);
    console.log('User logged out, session deleted for token:', token);
    return { success: true, message: 'Logged out successfully' };
  }
  
  return { success: false, message: 'Invalid session token' };
}

/**
 * Gets user data by session token - used as a direct API endpoint
 * @param {string} token - Session token
 * @return {Object} User data object or null if session is invalid
 */
function getUserDataBySessionToken(token) {
  try {
    if (!token) {
      console.error('No token provided to getUserDataBySessionToken');
      return null;
    }
    
    // Check if session is valid
    if (!validateSession(token)) {
      console.error('Invalid or expired session token in getUserDataBySessionToken');
      return null;
    }
    
    // Get user access data
    const userAccess = getUserAccess(token);
    console.log('Retrieved user data by token:', {
      id: userAccess.id || null,
      email: userAccess.email || null,
      name: userAccess.name || null
    });
    
    // Return only the necessary user data
    return {
      id: userAccess.id || null,
      email: userAccess.email || null,
      name: userAccess.name || null
    };
  } catch (error) {
    console.error('Error in getUserDataBySessionToken:', error);
    return null;
  }
}

// Function to set up BI Updates sheet
function setupBISheet(sheet) {
  const headers = [
    'Timestamp',
    'Timestamp String',
    'Request ID',
    'Course Name',
    'Status',
    'Remarks',
    'Updated By'
  ];

  try {
    // Set title
    sheet.getRange(1, 1).setValue('BI Team Updates');
    sheet.getRange(1, 1, 1, headers.length).merge();
    sheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setHorizontalAlignment('center');
    
    // Set headers
    sheet.getRange(2, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    const headerRange = sheet.getRange(2, 1, 1, headers.length);
    headerRange.setBackground('#4286f4');
    headerRange.setFontColor('white');
    headerRange.setFontWeight('bold');
    
    // Set column widths
    sheet.setColumnWidth(1, 180);  // Timestamp
    sheet.setColumnWidth(2, 180);  // Timestamp String
    sheet.setColumnWidth(3, 150);  // Request ID
    sheet.setColumnWidth(4, 200);  // Course Name
    sheet.setColumnWidth(5, 150);  // Status
    sheet.setColumnWidth(6, 300);  // Remarks
    sheet.setColumnWidth(7, 180);  // Updated By
    
    // Freeze header rows
    sheet.setFrozenRows(2);
    
    // Add filters
    sheet.getRange(2, 1, 1, headers.length).createFilter();
    
    // Add data validation for Status column
    const statusRange = sheet.getRange(3, 5, sheet.getMaxRows() - 2, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        BI_STATUS.PENDING,
        BI_STATUS.IN_PROGRESS,
        BI_STATUS.COMPLETED,
        BI_STATUS.REJECTED
      ])
      .build();
    statusRange.setDataValidation(statusRule);
    
    // Add formula for Timestamp String column
    const timestampStringFormula = '=TEXT(A3, "dd/MM/yyyy HH:mm:ss")';
    sheet.getRange(3, 2).setFormula(timestampStringFormula);
    
    console.log('BI Updates sheet setup complete');
    return true;
  } catch (error) {
    console.error('Error setting up BI updates sheet:', error);
    console.error('Error stack:', error.stack);
    return false;
  }
}

// Function to get data for BI dashboard
function getBIDashboardData() {
  try {
    console.log('Starting getBIDashboardData function');
    
    // Try to get from cache first
    const cache = CacheService.getUserCache();
    const cacheKey = 'bi_dashboard_' + Session.getActiveUser().getEmail();
    const cachedData = cache.get(cacheKey);
    
    if (cachedData) {
      console.log('Returning cached dashboard data');
      return JSON.parse(cachedData);
    }
    
    console.log('No cached data found, fetching from sheets');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      throw new Error('Could not access spreadsheet');
    }

    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    const biSheet = ss.getSheetByName(BI_SHEET_NAME);
    
    if (!requestSheet) {
      throw new Error('Main request sheet not found');
    }
    
    console.log('Sheets accessed:', {
      requestSheet: !!requestSheet,
      productSheet: !!productSheet,
      biSheet: !!biSheet
    });
    
    // Get product completion statuses - ONLY COMPLETED REQUESTS
    const productData = productSheet ? productSheet.getDataRange().getValues() : [];
    const productHeaders = productData[1] || [];
    const productRequestIdCol = findColumnIndex(productHeaders, 'Request ID');
    const productStatusCol = findColumnIndex(productHeaders, 'Access Revoke Status');
    
    if (productRequestIdCol === -1 || productStatusCol === -1) {
      console.error('Required columns not found in product sheet:', {
        foundColumns: productHeaders,
        requestIdCol: productRequestIdCol,
        statusCol: productStatusCol
      });
    }
    
    console.log('Product sheet column indices:', {
      requestId: productRequestIdCol,
      status: productStatusCol
    });
    
    // Create map of completed product requests with validation
    const completedProductRequests = new Map();
    if (productData.length > 2) {
      for (let i = 2; i < productData.length; i++) {
        const row = productData[i];
        if (!row || row.length <= Math.max(productRequestIdCol, productStatusCol)) {
          console.warn(`Skipping invalid product row at index ${i}`);
          continue;
        }
        
        const requestId = row[productRequestIdCol];
        const status = row[productStatusCol];
        
        if (requestId && status === REQUEST_STATUS.PRODUCT_COMPLETED) {
          completedProductRequests.set(requestId, true);
        }
      }
    }
    
    console.log('Found completed product requests:', completedProductRequests.size);
    
    // Get BI status updates with validation
    const biData = biSheet ? biSheet.getDataRange().getValues() : [];
    const biHeaders = biData[1] || [];
    const biRequestIdCol = findColumnIndex(biHeaders, 'Request ID');
    const biStatusCol = findColumnIndex(biHeaders, 'Status');
    const biTimestampCol = findColumnIndex(biHeaders, 'Timestamp');
    
    if (biRequestIdCol === -1 || biStatusCol === -1) {
      console.error('Required columns not found in BI sheet:', {
        foundColumns: biHeaders,
        requestIdCol: biRequestIdCol,
        statusCol: biStatusCol
      });
    }
    
    // Create map of latest BI statuses with validation
    const biStatuses = new Map();
    if (biData.length > 2) {
      for (let i = 2; i < biData.length; i++) {
        const row = biData[i];
        if (!row || row.length <= Math.max(biRequestIdCol, biStatusCol, biTimestampCol)) {
          console.warn(`Skipping invalid BI row at index ${i}`);
          continue;
        }
        
        const requestId = row[biRequestIdCol];
        const status = row[biStatusCol];
        const timestamp = row[biTimestampCol];
        
        if (requestId && status) {
          if (!biStatuses.has(requestId) || 
              new Date(timestamp) > new Date(biStatuses.get(requestId).timestamp)) {
            biStatuses.set(requestId, {
              status: status,
              timestamp: timestamp
            });
          }
        }
      }
    }
    
    console.log('Processed BI statuses:', biStatuses.size);
    
    // Initialize counters
    let pending = 0;
    let inProgress = 0;
    let completed = 0;
    let rejected = 0;
    
    // Process requests with validation
    const requests = [];
    const data = requestSheet.getDataRange().getValues();
    const headers = data[1];
    
    // Map column indices with validation
    const columnMap = {
      requestId: findColumnIndex(headers, 'Request ID'),
      timestamp: findColumnIndex(headers, 'Timestamp'),
      courseName: findColumnIndex(headers, 'Mistaken Course Name'),
      refundType: findColumnIndex(headers, 'Refund Type'),
      refundAmount: findColumnIndex(headers, 'Refund Amount')
    };
    
    // Validate required columns
    const missingColumns = Object.entries(columnMap)
      .filter(([key, value]) => value === -1)
      .map(([key]) => key);
    
    if (missingColumns.length > 0) {
      throw new Error(`Required columns not found: ${missingColumns.join(', ')}`);
    }
    
    console.log('Request sheet column indices:', columnMap);
    
    // Process each request with validation
    for (let i = 2; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length <= Math.max(...Object.values(columnMap))) {
        console.warn(`Skipping invalid request row at index ${i}`);
        continue;
      }
      
      const requestId = row[columnMap.requestId];
      
      // Only process requests that have been completed by Product team
      if (!completedProductRequests.has(requestId)) {
        continue;
      }
      
      try {
        // Get BI status or default to Pending
        const biInfo = biStatuses.get(requestId);
        const biStatus = biInfo ? biInfo.status : 'Pending';
        
        // Update counters
        switch(biStatus) {
          case BI_STATUS.PENDING: pending++; break;
          case BI_STATUS.IN_PROGRESS: inProgress++; break;
          case BI_STATUS.COMPLETED: completed++; break;
          case BI_STATUS.REJECTED: rejected++; break;
        }
        
        // Create request object with safe value access
        requests.push({
          request_id: requestId,
          timestamp: new Date(row[columnMap.timestamp] || new Date()),
          timestamp_string: formatDateTime(row[columnMap.timestamp]),
          course_name: row[columnMap.courseName] || '',
          refund_type: row[columnMap.refundType] || '',
          refund_amount: parseFloat(row[columnMap.refundAmount]) || 0,
          current_bi_status: biStatus
        });
      } catch (rowError) {
        console.error(`Error processing row ${i}:`, rowError);
      }
    }
    
    console.log('Processed requests:', {
      total: requests.length,
      statusCounts: {
        pending,
        inProgress,
        completed,
        rejected
      }
    });
    
    // Prepare and validate final result
    const result = {
      requests: requests,
      statusCounts: {
        pending,
        inProgress,
        completed,
        rejected
      }
    };
    
    // Cache results for 15 minutes
    try {
      cache.put(cacheKey, JSON.stringify(result), 900);
      console.log('Dashboard data cached successfully');
    } catch (cacheError) {
      console.warn('Failed to cache dashboard data:', cacheError);
    }
    
    console.log('Dashboard data prepared successfully');
    return result;
    
  } catch (error) {
    console.error('Error in getBIDashboardData:', error);
    console.error('Error stack:', error.stack);
    throw new Error(`Failed to load dashboard data: ${error.message}`);
  }
}

// Function to create BI update history
function createBIUpdateHistory(update, sessionToken) {
  try {
    console.log('Starting BI update with data:', update);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get user data from session token
    let userEmail = '';
    if (sessionToken) {
      const userData = getUserDataBySessionToken(sessionToken);
      if (userData && userData.email) {
        userEmail = userData.email;
        console.log('Using authenticated user email from session token:', userEmail);
      } else {
        console.warn('Could not get user email from session token, falling back to Session.getActiveUser()');
      }
    }
    
    // Fall back to Session.getActiveUser() if no email from token
    if (!userEmail) {
      userEmail = Session.getActiveUser().getEmail();
      console.log('Using email from Session.getActiveUser():', userEmail);
    }
    
    // Get or create BI Updates sheet
    let biSheet = ss.getSheetByName(BI_SHEET_NAME);
    if (!biSheet) {
      console.log('Creating new BI Updates sheet');
      biSheet = ss.insertSheet(BI_SHEET_NAME);
      setupBISheet(biSheet);
    }
    
    // For single update, wrap in array
    const updates = Array.isArray(update) ? update : [update];
    
    // Validate updates
    for (const item of updates) {
      if (!item.request_id || !item.status) {
        throw new Error('Missing required fields: request_id and status');
      }
      if (![
        BI_STATUS.PENDING,
        BI_STATUS.IN_PROGRESS,
        BI_STATUS.COMPLETED,
        BI_STATUS.REJECTED
      ].includes(item.status)) {
        throw new Error('Invalid status: ' + item.status);
      }
    }
    
    // Prepare rows for batch insert
    const timestamp = new Date();
    const newRows = updates.map(item => [
      timestamp,                    // Timestamp
      '',                          // Timestamp String (formula will handle this)
      item.request_id,             // Request ID
      item.course_name || '',      // Course Name
      item.status,                 // Status
      item.remarks || '',          // Remarks
      userEmail                    // Updated By - now using email from session token
    ]);
    
    // Get last row and append new rows
    const lastRow = Math.max(biSheet.getLastRow(), 2);
    const range = biSheet.getRange(lastRow + 1, 1, newRows.length, newRows[0].length);
    range.setValues(newRows);
    
    // Apply timestamp string formula to new rows
    const timestampStringRange = biSheet.getRange(lastRow + 1, 2, newRows.length, 1);
    const formula = '=TEXT(A' + (lastRow + 1) + ', "dd/MM/yyyy HH:mm:ss")';
    timestampStringRange.setFormula(formula);
    
    // Update main request sheet status if needed
    if (updates.length === 1 && updates[0].shouldUpdateMainSheet) {
      const requestSheet = ss.getSheetByName(SHEET_NAME);
      const requestData = requestSheet.getDataRange().getValues();
      const headers = requestData[1];
      const requestIdCol = findColumnIndex(headers, 'Request ID');
      const statusCol = findColumnIndex(headers, 'Status');
      
      // Find and update the request row
      for (let i = 2; i < requestData.length; i++) {
        if (String(requestData[i][requestIdCol]) === String(updates[0].request_id)) {
          requestSheet.getRange(i + 1, statusCol + 1).setValue(updates[0].status);
          break;
        }
      }
    }
    
    // Invalidate cache
    const cache = CacheService.getUserCache();
    const cacheKey = 'bi_dashboard_' + userEmail;
    cache.remove(cacheKey);
    
    // Send notification for each update
    for (const item of updates) {
      const requestDetails = getBIRequestDetails(item.request_id);
      sendBIUpdateNotification(item.request_id, {
        status: item.status,
        remarks: item.remarks,
        updated_by: userEmail
      }, requestDetails);
    }
    
    console.log('BI update completed successfully');
    return {
      success: true,
      message: 'Update(s) processed successfully'
    };
    
  } catch (error) {
    console.error('Error creating BI update:', error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

// Function to get detailed information about a specific BI request
function getBIRequestDetails(requestId) {
  try {
    console.log('Getting BI request details for:', requestId);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // Get sheets
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    const biSheet = ss.getSheetByName(BI_SHEET_NAME);
    const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    
    // Get request data
    const requestData = requestSheet.getDataRange().getValues();
    const requestHeaders = requestData[1];
    
    // Map request column indices
    const columnMap = {
      requestId: findColumnIndex(requestHeaders, 'Request ID'),
      timestamp: findColumnIndex(requestHeaders, 'Timestamp'),
      requester: findColumnIndex(requestHeaders, 'Requester'),
      requesterEmail: findColumnIndex(requestHeaders, 'Requester Email'),
      vertical: findColumnIndex(requestHeaders, 'Vertical'),
      courseName: findColumnIndex(requestHeaders, 'Mistaken Course Name'),
      mistakenPhase: findColumnIndex(requestHeaders, 'Mistaken Phase Name'),
      intendedCourse: findColumnIndex(requestHeaders, 'Intended Course Name'),
      intendedPhase: findColumnIndex(requestHeaders, 'Intended Phase Name'),
      refundType: findColumnIndex(requestHeaders, 'Refund Type'),
      refundAmount: findColumnIndex(requestHeaders, 'Refund Amount'),
      customerName: findColumnIndex(requestHeaders, 'Customer Name'),
      enrollmentNumber: findColumnIndex(requestHeaders, 'Enrollment Number'),
      purchaseDate: findColumnIndex(requestHeaders, 'Purchase Date'),
      paymentId: findColumnIndex(requestHeaders, 'Payment ID'),
      purchasedAmount: findColumnIndex(requestHeaders, 'Purchased Amount'),
      wrongPhone: findColumnIndex(requestHeaders, 'Wrong Phone'),
      correctPhone: findColumnIndex(requestHeaders, 'Correct Phone'),
      senderBkash: findColumnIndex(requestHeaders, 'Sender bKash'),
      transactionTime: findColumnIndex(requestHeaders, 'Transaction Time'),
      sentAmount: findColumnIndex(requestHeaders, 'Sent Amount'),
      refundReason: findColumnIndex(requestHeaders, 'Refund Reason'),
      documentFolder: findColumnIndex(requestHeaders, 'Document Folder')
    };
    
    // Find request row
    let requestRow = null;
    for (let i = 2; i < requestData.length; i++) {
      if (String(requestData[i][columnMap.requestId]) === String(requestId)) {
        requestRow = requestData[i];
        break;
      }
    }
    
    if (!requestRow) {
      throw new Error('Request not found');
    }
    
    // Get product completion status
    let productStatus = REQUEST_STATUS.PENDING;
    if (productSheet) {
      const productData = productSheet.getDataRange().getValues();
      const productHeaders = productData[1] || [];
      const productRequestIdCol = findColumnIndex(productHeaders, 'Request ID');
      const productStatusCol = findColumnIndex(productHeaders, 'Access Revoke Status');
      
      // Find latest product status
      for (let i = productData.length - 1; i > 1; i--) {
        if (String(productData[i][productRequestIdCol]) === String(requestId)) {
          productStatus = productData[i][productStatusCol] || REQUEST_STATUS.PENDING;
          break;
        }
      }
    }
    
    // Get BI update history
    let history = [];
    if (biSheet) {
      const biData = biSheet.getDataRange().getValues();
      const biHeaders = biData[1] || [];
      const biRequestIdCol = findColumnIndex(biHeaders, 'Request ID');
      const biTimestampCol = findColumnIndex(biHeaders, 'Timestamp');
      const biStatusCol = findColumnIndex(biHeaders, 'Status');
      const biRemarksCol = findColumnIndex(biHeaders, 'Remarks');
      const biUpdatedByCol = findColumnIndex(biHeaders, 'Updated By');
      
      // Get all updates for this request
      history = biData.slice(2)
        .filter(row => String(row[biRequestIdCol]) === String(requestId))
        .map(row => ({
          timestamp: formatDateTime(row[biTimestampCol]),
          status: row[biStatusCol] || BI_STATUS.PENDING,
          remarks: row[biRemarksCol] || '',
          updated_by: row[biUpdatedByCol] || ''
        }))
        .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    }
    
    // Get current status from latest history entry or default to Pending
    const currentStatus = history.length > 0 ? history[0].status : BI_STATUS.PENDING;
    
    // Return detailed request information
    return {
      request_id: requestId,
      timestamp: formatDateTime(requestRow[columnMap.timestamp]),
      requester: requestRow[columnMap.requester] || '',
      requester_email: requestRow[columnMap.requesterEmail] || '',
      vertical: requestRow[columnMap.vertical] || '',
      course_name: requestRow[columnMap.courseName] || '',
      mistaken_phase: requestRow[columnMap.mistakenPhase] || '',
      intended_course: requestRow[columnMap.intendedCourse] || '',
      intended_phase: requestRow[columnMap.intendedPhase] || '',
      refund_type: requestRow[columnMap.refundType] || '',
      refund_amount: requestRow[columnMap.refundAmount] || 0,
      customer_name: requestRow[columnMap.customerName] || '',
      enrollment_number: requestRow[columnMap.enrollmentNumber] || '',
      purchase_date: requestRow[columnMap.purchaseDate] ? formatDateTime(requestRow[columnMap.purchaseDate]) : '',
      payment_id: requestRow[columnMap.paymentId] || '',
      purchased_amount: requestRow[columnMap.purchasedAmount] || 0,
      wrong_phone: requestRow[columnMap.wrongPhone] || '',
      correct_phone: requestRow[columnMap.correctPhone] || '',
      sender_bkash: requestRow[columnMap.senderBkash] || '',
      transaction_time: requestRow[columnMap.transactionTime] ? formatDateTime(requestRow[columnMap.transactionTime]) : '',
      sent_amount: requestRow[columnMap.sentAmount] || 0,
      refund_reason: requestRow[columnMap.refundReason] || '',
      document_folder: requestRow[columnMap.documentFolder] || '',
      product_status: productStatus,
      current_bi_status: currentStatus,
      history: history
    };
    
  } catch (error) {
    console.error('Error getting BI request details:', error);
    throw error;
  }
}

// Function to process bulk BI updates
function bulkUpdateBIRequests(requests, newStatus, remarks, sessionToken) {
  try {
    console.log('Starting bulk BI update process');
    const startTime = new Date().getTime();
    
    // Validate inputs
    if (!Array.isArray(requests) || requests.length === 0) {
      throw new Error('No requests provided for update');
    }
    
    if (![
      BI_STATUS.IN_PROGRESS,
      BI_STATUS.COMPLETED,
      BI_STATUS.REJECTED
    ].includes(newStatus)) {
      throw new Error('Invalid status: ' + newStatus);
    }
    
    // Process in batches to avoid timeout
    const batchSize = 20;
    const batches = [];
    
    // Prepare batches
    for (let i = 0; i < requests.length; i += batchSize) {
      batches.push(requests.slice(i, i + batchSize));
    }
    
    // Track results
    const results = [];
    let processedCount = 0;
    
    // Process each batch
    for (const batch of batches) {
      try {
        // Prepare updates with consistent format
        const updates = batch.map(req => ({
          request_id: req.request_id,
          course_name: req.course_name || '',
          status: newStatus,
          remarks: remarks || '',
          shouldUpdateMainSheet: true
        }));
        
        // Process batch with session token
        const batchResult = createBIUpdateHistory(updates, sessionToken);
        
        if (batchResult.success) {
          processedCount += batch.length;
          results.push(...batch.map(item => ({
            request_id: item.request_id,
            success: true
          })));
        } else {
          results.push(...batch.map(item => ({
            request_id: item.request_id,
            success: false,
            error: batchResult.message
          })));
        }
        
        // Check remaining time (5.5 minutes max to be safe)
        if (new Date().getTime() - startTime > 330000) {
          console.log('Time limit approaching, returning partial results');
          return {
            success: true,
            partial: true,
            processed: processedCount,
            total: requests.length,
            results: results
          };
        }
        
      } catch (batchError) {
        console.error('Error processing batch:', batchError);
        results.push(...batch.map(item => ({
          request_id: item.request_id,
          success: false,
          error: batchError.toString()
        })));
      }
    }
    
    // Return final results
    return {
      success: results.every(r => r.success),
      partial: false,
      processed: processedCount,
      total: requests.length,
      results: results
    };
    
  } catch (error) {
    console.error('Error in bulk update process:', error);
    return {
      success: false,
      error: error.toString(),
      processed: 0,
      total: requests.length,
      results: []
    };
  }
}

// Function to filter BI dashboard data
function filterBIDashboardData(filters) {
  try {
    console.log('Filtering BI dashboard data with:', filters);
    
    // Get base data first
    const dashboardData = getBIDashboardData();
    
    // If no filters or no requests, return as is
    if (!filters || !dashboardData.requests || dashboardData.requests.length === 0) {
      return dashboardData;
    }
    
    // Initialize counters
    let pending = 0;
    let inProgress = 0;
    let completed = 0;
    let rejected = 0;
    
    // Apply filters
    const filteredRequests = dashboardData.requests.filter(request => {
      let matchesFilter = true;
      
      // Status filter
      if (filters.status && filters.status !== 'all') {
        matchesFilter = matchesFilter && request.current_bi_status === filters.status;
      }
      
      // Date range filter
      if (filters.startDate && filters.endDate) {
        const requestDate = new Date(request.timestamp);
        const startDate = new Date(filters.startDate);
        const endDate = new Date(filters.endDate);
        endDate.setHours(23, 59, 59, 999); // Include full end date
        
        matchesFilter = matchesFilter && 
          requestDate >= startDate && 
          requestDate <= endDate;
      }
      
      // Search term filter
      if (filters.searchTerm) {
        const searchTerm = filters.searchTerm.toLowerCase();
        const searchableFields = [
          request.request_id,
          request.course_name,
          request.refund_type,
          (request.refund_amount || '').toString()
        ].join(' ').toLowerCase();
        
        matchesFilter = matchesFilter && searchableFields.includes(searchTerm);
      }
      
      // Update counters if request matches filters
      if (matchesFilter) {
        switch(request.current_bi_status) {
          case BI_STATUS.PENDING: pending++; break;
          case BI_STATUS.IN_PROGRESS: inProgress++; break;
          case BI_STATUS.COMPLETED: completed++; break;
          case BI_STATUS.REJECTED: rejected++; break;
        }
      }
      
      return matchesFilter;
    });
    
    // Return filtered data with updated counts
    return {
      requests: filteredRequests,
      statusCounts: {
        pending: pending,
        inProgress: inProgress,
        completed: completed,
        rejected: rejected
      }
    };
    
  } catch (error) {
    console.error('Error filtering BI dashboard data:', error);
    return {
      requests: [],
      statusCounts: {
        pending: 0,
        inProgress: 0,
        completed: 0,
        rejected: 0
      }
    };
  }
}

// Function to get quick filter presets
function getBIFilterPresets() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const yesterday = new Date(today);
  yesterday.setDate(yesterday.getDate() - 1);
  
  const lastWeekStart = new Date(today);
  lastWeekStart.setDate(lastWeekStart.getDate() - 7);
  
  const lastMonthStart = new Date(today);
  lastMonthStart.setMonth(lastMonthStart.getMonth() - 1);
  
  return {
    today: {
      startDate: today,
      endDate: today
    },
    yesterday: {
      startDate: yesterday,
      endDate: yesterday
    },
    lastWeek: {
      startDate: lastWeekStart,
      endDate: today
    },
    lastMonth: {
      startDate: lastMonthStart,
      endDate: today
    }
  };
}

// Function to send BI update notification
function sendBIUpdateNotification(requestId, update, requestDetails) {
  try {
    const subject = `BI Update: Request ${requestId} - ${update.status}`;
    
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h2 style="color: #4286f4;">BI Team Update Notification</h2>
        <div style="background-color: #f8f9fa; padding: 15px; border-radius: 4px; margin: 20px 0;">
          <p><strong>Request ID:</strong> ${requestId}</p>
          <p><strong>Course Name:</strong> ${requestDetails.course_name}</p>
          <p><strong>Status:</strong> ${update.status}</p>
          <p><strong>Remarks:</strong> ${update.remarks || 'No remarks provided'}</p>
          <p><strong>Updated By:</strong> ${update.updated_by}</p>
          <p><strong>Update Time:</strong> ${formatDateTime(new Date())}</p>
        </div>
        <p>For more details, please visit the <a href="${getScriptUrl()}?page=bi">BI Portal</a>.</p>
      </div>
    `;
    
    // Get email recipients
    const recipients = {
      to: EMAIL_CONFIG.defaultRecipients.to,
      cc: [
        ...EMAIL_CONFIG.defaultRecipients.cc,
        ...EMAIL_CONFIG.conditionalRecipients['BI Team']
      ]
    };
    
    // Send email
    MailApp.sendEmail({
      to: recipients.to,
      cc: recipients.cc.join(','),
      subject: subject,
      htmlBody: htmlBody,
      noReply: true
    });
    
    console.log('BI update notification sent successfully');
    return true;
  } catch (error) {
    console.error('Error sending BI update notification:', error);
    return false;
  }
}

// Function to get the complete request journey from all teams
function getCompleteRequestJourney(requestId) {
  try {
    console.log('Getting complete request journey for:', requestId);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let allUpdates = [];
    
    // 1. Get initial form submission
    const requestSheet = ss.getSheetByName(SHEET_NAME);
    if (requestSheet) {
      const requestData = requestSheet.getDataRange().getValues();
      const requestHeaders = requestData[1];
      const requestIdCol = findColumnIndex(requestHeaders, 'Request ID');
      const timestampCol = 0; // First column is timestamp
      const requesterCol = findColumnIndex(requestHeaders, 'Requester');
      
      // Find the submission
      for (let i = 2; i < requestData.length; i++) {
        if (String(requestData[i][requestIdCol]) === String(requestId)) {
          allUpdates.push({
            timestamp: requestData[i][timestampCol],
            status: 'Submitted',
            remarks: 'Request submitted',
            updated_by: requestData[i][requesterCol] || 'Sales Team',
            team: 'Sales'
          });
          break;
        }
      }
    }
    
    // 2. Get approver updates
    const approvalSheet = ss.getSheetByName(REFUND_APPROVAL_SHEET_NAME);
    if (approvalSheet) {
      const approvalData = approvalSheet.getDataRange().getValues();
      const approvalHeaders = approvalData[1];
      const approvalRequestIdCol = findColumnIndex(approvalHeaders, 'Request ID');
      const approvalTimestampCol = 0; // First column is timestamp
      const approvalStatusCol = findColumnIndex(approvalHeaders, 'Approval Status');
      const approvalRemarksCol = findColumnIndex(approvalHeaders, 'Approver Remarks');
      const approvalUpdatedByCol = findColumnIndex(approvalHeaders, 'Updated By');
      
      // Get all approval updates
      for (let i = 2; i < approvalData.length; i++) {
        if (String(approvalData[i][approvalRequestIdCol]) === String(requestId)) {
          allUpdates.push({
            timestamp: approvalData[i][approvalTimestampCol],
            status: approvalData[i][approvalStatusCol] || 'Pending Approval',
            remarks: approvalData[i][approvalRemarksCol] || '',
            updated_by: approvalData[i][approvalUpdatedByCol] || 'Approver',
            team: 'Approver'
          });
        }
      }
    }
    
    // 3. Get finance updates
    const financeSheet = ss.getSheetByName(FINANCE_SHEET_NAME);
    if (financeSheet) {
      const financeData = financeSheet.getDataRange().getValues();
      const financeHeaders = financeData[1];
      const financeRequestIdCol = findColumnIndex(financeHeaders, 'Refund SL Number');
      const financeTimestampCol = 0; // First column is timestamp
      const financeStatusCol = findColumnIndex(financeHeaders, 'Status');
      const financeRemarksCol = findColumnIndex(financeHeaders, 'Finance Remarks');
      const financeUpdatedByCol = findColumnIndex(financeHeaders, 'Updated By');
      
      // Get all finance updates
      for (let i = 2; i < financeData.length; i++) {
        if (String(financeData[i][financeRequestIdCol]) === String(requestId)) {
          allUpdates.push({
            timestamp: financeData[i][financeTimestampCol],
            status: financeData[i][financeStatusCol] || 'Pending Finance',
            remarks: financeData[i][financeRemarksCol] || '',
            updated_by: financeData[i][financeUpdatedByCol] || 'Finance Team',
            team: 'Finance'
          });
        }
      }
    }
    
    // 4. Get product updates
    const productSheet = ss.getSheetByName(PRODUCT_SHEET_NAME);
    if (productSheet) {
      const productData = productSheet.getDataRange().getValues();
      const productHeaders = productData[1];
      const productRequestIdCol = findColumnIndex(productHeaders, 'Request ID');
      const productTimestampCol = 0; // First column is timestamp
      const productStatusCol = findColumnIndex(productHeaders, 'Access Revoke Status');
      const productRemarksCol = findColumnIndex(productHeaders, 'Remarks');
      const productUpdatedByCol = findColumnIndex(productHeaders, 'Updated By');
      
      // Get all product updates
      for (let i = 2; i < productData.length; i++) {
        if (String(productData[i][productRequestIdCol]) === String(requestId)) {
          allUpdates.push({
            timestamp: productData[i][productTimestampCol],
            status: productData[i][productStatusCol] || 'Pending Product',
            remarks: productData[i][productRemarksCol] || '',
            updated_by: productData[i][productUpdatedByCol] || 'Product Team',
            team: 'Product'
          });
        }
      }
    }
    
    // 5. Get BI updates
    const biSheet = ss.getSheetByName(BI_SHEET_NAME);
    if (biSheet) {
      const biData = biSheet.getDataRange().getValues();
      const biHeaders = biData[1];
      const biRequestIdCol = findColumnIndex(biHeaders, 'Request ID');
      const biTimestampCol = 0; // First column is timestamp
      const biStatusCol = findColumnIndex(biHeaders, 'Status');
      const biRemarksCol = findColumnIndex(biHeaders, 'Remarks');
      const biUpdatedByCol = findColumnIndex(biHeaders, 'Updated By');
      
      // Get all BI updates
      for (let i = 2; i < biData.length; i++) {
        if (String(biData[i][biRequestIdCol]) === String(requestId)) {
          allUpdates.push({
            timestamp: biData[i][biTimestampCol],
            status: biData[i][biStatusCol] || 'Pending BI',
            remarks: biData[i][biRemarksCol] || '',
            updated_by: biData[i][biUpdatedByCol] || 'BI Team',
            team: 'BI'
          });
        }
      }
    }
    
    // Sort all updates by timestamp (newest first)
    allUpdates.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    // Format timestamps
    allUpdates = allUpdates.map(update => ({
      ...update,
      timestamp: formatDateTime(update.timestamp)
    }));
    
    console.log(`Found ${allUpdates.length} updates for request ${requestId} across all teams`);
    return allUpdates;
  } catch (error) {
    console.error('Error getting complete request journey:', error);
    console.error('Error stack:', error.stack);
    return [];
  }
}

/**
 * Changes a user's password
 * 
 * @param {string} sessionToken - User's session token
 * @param {string} currentPassword - User's current password
 * @param {string} newPassword - User's new password
 * @return {Object} Response with success status and message
 */
function changePassword(sessionToken, currentPassword, newPassword) {
  try {
    // Validate session token
    const userData = getUserDataBySessionToken(sessionToken);
    if (!userData || !userData.email) {
      return { success: false, message: 'Invalid session' };
    }
    
    // Get user's email from session data
    const userEmail = userData.email;
    
    // Open the credentials sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Credentials');
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const headerRow = values[0];
    
    // Find column indices
    const emailIndex = findColumnIndex(headerRow, 'Email');
    const passwordIndex = findColumnIndex(headerRow, 'Password');
    const lastPasswordIndex = findColumnIndex(headerRow, 'Last Password');
    
    if (emailIndex === -1 || passwordIndex === -1) {
      return { success: false, message: 'Sheet structure error' };
    }
    
    // Find all rows for this user
    let userRows = [];
    let currentPasswordVerified = false;
    
    for (let i = 1; i < values.length; i++) {
      if (values[i][emailIndex] === userEmail) {
        userRows.push(i + 1); // +1 because sheet rows are 1-indexed
        
        // Verify current password (only needs to match one row)
        if (values[i][passwordIndex] === currentPassword) {
          currentPasswordVerified = true;
        }
      }
    }
    
    if (!currentPasswordVerified) {
      return { success: false, message: 'Current password is incorrect' };
    }
    
    if (userRows.length === 0) {
      return { success: false, message: 'User not found' };
    }
    
    // Update password for all rows with this email
    for (const rowIndex of userRows) {
      // Store the current password as the last password if the column exists
      if (lastPasswordIndex !== -1) {
        const currentPwd = sheet.getRange(rowIndex, passwordIndex + 1).getValue();
        sheet.getRange(rowIndex, lastPasswordIndex + 1).setValue(currentPwd);
      }
      
      // Set the new password
      sheet.getRange(rowIndex, passwordIndex + 1).setValue(newPassword);
    }
    
    return { 
      success: true, 
      message: 'Password updated successfully.' 
    };
    
  } catch (error) {
    console.error('Error in changePassword:', error);
    return { success: false, message: 'An error occurred: ' + error.message };
  }
}

/**
 * Handles password reset requests
 * @param {string} email - User's email address
 * @return {Object} Response object with success status and message
 */
function requestPasswordReset(email) {
  try {
    if (!email) {
      return { success: false, message: 'Email is required' };
    }
    
    // Get the Credentials sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentialsSheet = ss.getSheetByName('Credentials');
    
    if (!credentialsSheet) {
      console.error('Credentials sheet not found');
      return { success: false, message: 'System error: Credentials sheet not found' };
    }
    
    // Find the user by email
    const data = credentialsSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    
    if (emailCol === -1) {
      console.error('Email column not found in Credentials sheet');
      return { success: false, message: 'System error: Email column not found' };
    }
    
    let userFound = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol].toString().toLowerCase() === email.toLowerCase()) {
        userFound = true;
        break;
      }
    }
    
    if (!userFound) {
      // For security reasons, don't reveal if the email exists or not
      return { 
        success: true, 
        message: 'If your email is registered, you will receive password reset instructions shortly.' 
      };
    }
    
    // Generate a reset token
    const resetToken = Utilities.getUuid();
    const expiresAt = new Date();
    expiresAt.setTime(expiresAt.getTime() + (24 * 60 * 60 * 1000)); // 24 hours from now
    
    // Store the reset token in Script Properties
    const resetKey = 'reset_' + resetToken;
    const resetData = {
      email: email,
      expiresAt: expiresAt.getTime(),
      used: false
    };
    
    PropertiesService.getScriptProperties().setProperty(resetKey, JSON.stringify(resetData));
    
    // Get the script URL
    const scriptUrl = getScriptUrl();
    const resetUrl = `${scriptUrl}?page=reset&token=${resetToken}`;
    
    // Send email with reset link
    const subject = 'Password Reset - Refund Management System';
    const body = `
      <p>Hello,</p>
      <p>We received a request to reset your password for the Refund Management System.</p>
      <p>To reset your password, please click on the link below:</p>
      <p><a href="${resetUrl}">${resetUrl}</a></p>
      <p>This link will expire in 24 hours.</p>
      <p>If you did not request a password reset, please ignore this email.</p>
      <p>Thank you,<br>Refund Management System</p>
    `;
    
    GmailApp.sendEmail(
      email,
      subject,
      "Please use an HTML-compatible email client to view this message.",
      { htmlBody: body }
    );
    
    console.log('Password reset email sent to:', email);
    
    return { 
      success: true, 
      message: 'If your email is registered, you will receive password reset instructions shortly.' 
    };
  } catch (error) {
    console.error('Error in requestPasswordReset:', error);
    return { success: false, message: 'An error occurred while processing your request.' };
  }
}

/**
 * Validates a password reset token
 * @param {string} token - Reset token to validate
 * @return {Object} Object containing validity status and email if valid
 */
function validateResetToken(token) {
  try {
    if (!token) {
      return { valid: false };
    }
    
    // Get the reset token data from Script Properties
    const resetKey = 'reset_' + token;
    const resetDataJson = PropertiesService.getScriptProperties().getProperty(resetKey);
    
    if (!resetDataJson) {
      console.log('Reset token not found:', token);
      return { valid: false };
    }
    
    const resetData = JSON.parse(resetDataJson);
    const now = new Date().getTime();
    
    // Check if token has expired or already been used
    if (now > resetData.expiresAt || resetData.used) {
      console.log('Reset token expired or already used:', token);
      return { valid: false };
    }
    
    return { 
      valid: true,
      email: resetData.email
    };
  } catch (error) {
    console.error('Error validating reset token:', error);
    return { valid: false };
  }
}

/**
 * Resets a user's password using a valid reset token
 * @param {string} token - Reset token
 * @param {string} newPassword - New password
 * @return {Object} Response object with success status and message
 */
function resetPassword(token, newPassword) {
  try {
    if (!token || !newPassword) {
      return { success: false, message: 'Invalid request' };
    }
    
    // Validate the reset token
    const tokenData = validateResetToken(token);
    if (!tokenData.valid) {
      return { success: false, message: 'Invalid or expired reset token' };
    }
    
    const email = tokenData.email;
    
    // Get the Credentials sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const credentialsSheet = ss.getSheetByName('Credentials');
    
    if (!credentialsSheet) {
      console.error('Credentials sheet not found');
      return { success: false, message: 'System error: Credentials sheet not found' };
    }
    
    // Find all rows for this user
    const data = credentialsSheet.getDataRange().getValues();
    const headers = data[0];
    const emailCol = headers.indexOf('Email');
    const passwordCol = headers.indexOf('Password');
    const lastPasswordCol = headers.indexOf('Last Password');
    
    if (emailCol === -1 || passwordCol === -1) {
      console.error('Required columns not found in Credentials sheet');
      return { success: false, message: 'System error: Required columns not found' };
    }
    
    let rowsUpdated = 0;
    
    // Update all rows for this user
    for (let i = 1; i < data.length; i++) {
      if (data[i][emailCol].toString().toLowerCase() === email.toLowerCase()) {
        // Store the current password as the last password if the column exists
        if (lastPasswordCol !== -1) {
          credentialsSheet.getRange(i + 1, lastPasswordCol + 1).setValue(data[i][passwordCol]);
        }
        
        // Update the password
        credentialsSheet.getRange(i + 1, passwordCol + 1).setValue(newPassword);
        rowsUpdated++;
      }
    }
    
    if (rowsUpdated === 0) {
      console.error('User not found:', email);
      return { success: false, message: 'User not found' };
    }
    
    // Mark the reset token as used
    const resetKey = 'reset_' + token;
    const resetDataJson = PropertiesService.getScriptProperties().getProperty(resetKey);
    if (resetDataJson) {
      const resetData = JSON.parse(resetDataJson);
      resetData.used = true;
      PropertiesService.getScriptProperties().setProperty(resetKey, JSON.stringify(resetData));
    }
    
    console.log('Password reset successful for:', email);
    
    return { 
      success: true, 
      message: 'Your password has been reset successfully' 
    };
  } catch (error) {
    console.error('Error in resetPassword:', error);
    return { success: false, message: 'An error occurred while resetting your password' };
  }
}

