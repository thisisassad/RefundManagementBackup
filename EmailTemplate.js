// Email template utilities and functions
function generateRefundRequestEmailBody(requestId, formData, consolidatedData, folderUrl) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f5f5f5;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <!-- Header -->
        <div style="text-align: center; margin-bottom: 30px;">
          <h1 style="color: #2196F3; margin: 0; font-size: 24px;">New Refund Request</h1>
          <p style="color: #666; margin: 10px 0 0;">Request ID: ${requestId}</p>
        </div>

        <!-- Main Content -->
        <div style="margin-bottom: 30px;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td colspan="2" style="padding: 15px; background-color: #f8f9fa; border-radius: 4px;">
                <h2 style="margin: 0; color: #333; font-size: 18px;">Request Details</h2>
              </td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee; width: 40%;"><strong>Requester</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee;">${formData.requester}</td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee;"><strong>Vertical</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee;">${formData.vertical}</td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee;"><strong>Enrollment Number</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee;">${consolidatedData.enrollmentNumber}</td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee;"><strong>Payment ID</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee;">${formData.paymentId}</td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee;"><strong>Refund Type</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee;">${consolidatedData.refundType}</td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee;"><strong>Amount</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee; color: #e53935;">${consolidatedData.refundAmount} BDT</td>
            </tr>
            <tr>
              <td style="padding: 12px; border-bottom: 1px solid #eee;"><strong>Revoke Access</strong></td>
              <td style="padding: 12px; border-bottom: 1px solid #eee;">
                <span style="
                  padding: 4px 8px;
                  border-radius: 4px;
                  font-size: 12px;
                  background-color: ${formData.revokeAccess === 'Yes' ? '#ffebee' : '#e8f5e9'};
                  color: ${formData.revokeAccess === 'Yes' ? '#c62828' : '#2e7d32'};
                ">${formData.revokeAccess}</span>
              </td>
            </tr>
          </table>
        </div>

        <!-- Reason Section -->
        <div style="margin-bottom: 30px; background-color: #f8f9fa; padding: 15px; border-radius: 4px;">
          <h2 style="margin: 0 0 15px; color: #333; font-size: 18px;">Reason for Refund</h2>
          <p style="margin: 0; color: #555; line-height: 1.5;">${formData.refundReason}</p>
        </div>

        <!-- Documents Section -->
        <div style="margin-bottom: 30px;">
          <h2 style="margin: 0 0 15px; color: #333; font-size: 18px;">Supporting Documents</h2>
          <a href="${folderUrl}" 
             style="display: inline-block; 
                    padding: 10px 20px; 
                    background-color: #2196F3; 
                    color: white; 
                    text-decoration: none; 
                    border-radius: 4px;
                    font-weight: bold;">
            View Documents
          </a>
        </div>

        <!-- Footer -->
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center;">
          <p style="color: #666; font-size: 14px; margin: 0;">Please review and process this request according to our refund policy.</p>
          <p style="color: #999; font-size: 12px; margin: 10px 0 0;">This is an automated email. Please do not reply.</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

function generateConfirmationEmailBody(requestId, formData, consolidatedData, folderUrl) {
  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin: 0; padding: 0; font-family: Arial, sans-serif; background-color: #f5f5f5;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">
        <!-- Header with Success Icon -->
        <div style="text-align: center; margin-bottom: 30px;">
          <div style="width: 60px; height: 60px; margin: 0 auto 15px; background-color: #4CAF50; border-radius: 50%; display: flex; align-items: center; justify-content: center;">
            <span style="color: white; font-size: 30px;">✓</span>
          </div>
          <h1 style="color: #4CAF50; margin: 0; font-size: 24px;">Request Received</h1>
          <p style="color: #666; margin: 10px 0 0;">Your refund request has been successfully submitted</p>
        </div>

        <!-- Request Details -->
        <div style="background-color: #f8f9fa; padding: 20px; border-radius: 4px; margin-bottom: 30px;">
          <table style="width: 100%;">
            <tr>
              <td style="padding: 8px 0;"><strong>Request ID:</strong></td>
              <td style="padding: 8px 0;">${requestId}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0;"><strong>Payment ID:</strong></td>
              <td style="padding: 8px 0;">${formData.paymentId}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0;"><strong>Refund Type:</strong></td>
              <td style="padding: 8px 0;">${consolidatedData.refundType}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0;"><strong>Amount:</strong></td>
              <td style="padding: 8px 0; color: #e53935;">${consolidatedData.refundAmount} BDT</td>
            </tr>
          </table>
        </div>

        <!-- Documents Link -->
        <div style="text-align: center; margin-bottom: 30px;">
          <p style="margin-bottom: 15px; color: #666;">Your submitted documents can be accessed here:</p>
          <a href="${folderUrl}" 
             style="display: inline-block; 
                    padding: 10px 20px; 
                    background-color: #2196F3; 
                    color: white; 
                    text-decoration: none; 
                    border-radius: 4px;
                    font-weight: bold;">
            View Documents
          </a>
        </div>

        <!-- Footer -->
        <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #eee; text-align: center;">
          <p style="color: #666; font-size: 14px; margin: 0;">The finance team will process this request upon approval.</p>
          <p style="color: #666; font-size: 14px; margin: 10px 0;">Please keep this email for your reference.</p>
          <p style="color: #999; font-size: 12px; margin: 20px 0 0;">Request ID: ${requestId}</p>
        </div>
      </div>
    </body>
    </html>
  `;
}

function generateRejectionEmailBody(requestId, data) {
  return `
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2 style="color: #dc3545;">Refund Request Rejected</h2>
      <p>Your refund request has been reviewed and cannot be processed at this time.</p>
      <div style="background-color: #f8f9fa; padding: 15px; border-radius: 4px; margin: 20px 0;">
        <p><strong>Request ID:</strong> ${requestId}</p>
        <p><strong>Remarks:</strong> ${data.remarks || 'No remarks provided'}</p>
      </div>
      <p>If you have any questions, please contact the finance team.</p>
    </div>
  `;
}

function sendRejectionEmail(email, requestId, data) {
  if (email && email !== 'customer@submission.com') {
    MailApp.sendEmail({
      to: email,
      subject: `Refund Request ${requestId} - Rejected`,
      htmlBody: generateRejectionEmailBody(requestId, data),
      noReply: true
    });
  }
}

// Add new status update email functions
function generateStatusUpdateEmailBody(requestId, data, requestDetails) {
  const statusColors = {
    'Pending': { bg: '#fff3cd', text: '#856404', icon: '⏳' },
    'In Progress': { bg: '#cce5ff', text: '#004085', icon: '🔄' },
    'Completed': { bg: '#d4edda', text: '#155724', icon: '✅' },
    'Reject': { bg: '#f8d7da', text: '#721c24', icon: '❌' }
  };

  const statusStyle = statusColors[data.status] || statusColors['Pending'];

  return `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
    </head>
    <body style="margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; line-height: 1.5; background-color: #f8fafc;">
      <div style="max-width: 600px; margin: 40px auto; background-color: #ffffff; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); overflow: hidden;">
        <!-- Status Banner -->
        <div style="background-color: ${statusStyle.bg}; padding: 20px; text-align: center;">
          <div style="font-size: 32px; margin-bottom: 10px;">${statusStyle.icon}</div>
          <h1 style="margin: 0; color: ${statusStyle.text}; font-size: 24px; font-weight: 600;">
            Status Updated: ${data.status}
          </h1>
        </div>

        <!-- Main Content -->
        <div style="padding: 32px;">
          <!-- Request Details -->
          <div style="background-color: #f8fafc; border-radius: 8px; padding: 24px; margin-bottom: 24px;">
            <h2 style="margin: 0 0 16px; color: #1a202c; font-size: 18px; font-weight: 600;">Request Details</h2>
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px 0; color: #4a5568; width: 40%;">Request ID</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${requestId}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #4a5568;">Refund Type</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${requestDetails.refundType}</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #4a5568;">Amount</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${requestDetails.refundAmount} BDT</td>
              </tr>
              <tr>
                <td style="padding: 8px 0; color: #4a5568;">Payment ID</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${requestDetails.paymentId}</td>
              </tr>
            </table>
          </div>

          <!-- Update Details -->
          <div style="background-color: #f8fafc; border-radius: 8px; padding: 24px;">
            <h2 style="margin: 0 0 16px; color: #1a202c; font-size: 18px; font-weight: 600;">Update Information</h2>
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 8px 0; color: #4a5568; width: 40%;">Updated At</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${formatDateTime(new Date())}</td>
              </tr>
              ${data.refundedAmount ? `
              <tr>
                <td style="padding: 8px 0; color: #4a5568;">Refunded Amount</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${data.refundedAmount} BDT</td>
              </tr>
              ` : ''}
              ${data.refundMethod ? `
              <tr>
                <td style="padding: 8px 0; color: #4a5568;">Refund Method</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${data.refundMethod}</td>
              </tr>
              ` : ''}
              ${data.remarks ? `
              <tr>
                <td style="padding: 8px 0; color: #4a5568;">Remarks</td>
                <td style="padding: 8px 0; color: #1a202c; font-weight: 500;">${data.remarks}</td>
              </tr>
              ` : ''}
            </table>
          </div>
        </div>

        <!-- Footer -->
        <div style="padding: 24px; background-color: #f8fafc; text-align: center; border-top: 1px solid #e2e8f0;">
          <p style="margin: 0; color: #4a5568; font-size: 14px;">
            This is an automated message from the Refund Management System.
          </p>
        </div>
      </div>
    </body>
    </html>
  `;
}

function sendStatusUpdateEmail(requestId, data, requestDetails, recipients) {
  const subject = `Refund Request ${requestId} - ${data.status} | ${formatDateForSubject(new Date())}`;
  const htmlBody = generateStatusUpdateEmailBody(requestId, data, requestDetails);

  // Send to requester
  if (requestDetails.requesterEmail && requestDetails.requesterEmail !== 'customer@submission.com') {
    MailApp.sendEmail({
      to: requestDetails.requesterEmail,
      subject: subject,
      htmlBody: htmlBody,
      noReply: true
    });
  }

  // Send to finance team and conditional recipients
  MailApp.sendEmail({
    to: recipients.to,
    cc: recipients.cc.join(','),
    subject: subject,
    htmlBody: htmlBody,
    noReply: true
  });
}

// Helper function for date formatting
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

// Email sending functions
function sendRefundRequestEmail(recipients, requestId, formData, consolidatedData, folderUrl, timestamp) {
  MailApp.sendEmail({
    to: recipients.to,
    cc: recipients.cc.join(','),
    subject: `Refund Request - ${consolidatedData.refundType} - ${requestId} | ${formatDateForSubject(timestamp)}`,
    htmlBody: generateRefundRequestEmailBody(requestId, formData, consolidatedData, folderUrl),
    noReply: true
  });
}

function sendConfirmationEmail(email, requestId, formData, consolidatedData, folderUrl, timestamp) {
  MailApp.sendEmail({
    to: email,
    subject: `Refund Request - ${consolidatedData.refundType} - ${requestId} | ${formatDateForSubject(timestamp)}`,
    htmlBody: generateConfirmationEmailBody(requestId, formData, consolidatedData, folderUrl),
    noReply: true
  });
}

// Helper function for date formatting
function formatDateForSubject(date) {
  const d = new Date(date);
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  return `${d.getDate().toString().padStart(2, '0')}-${months[d.getMonth()]}-${d.getFullYear()}`;
} 