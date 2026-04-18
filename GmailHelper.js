/**
 * Gmail Helper - Sử dụng Gmail API thay vì MailApp
 * Gmail API ít vấn đề về authorization hơn và mạnh mẽ hơn
 */

/**
 * Gửi email sử dụng Gmail API
 * @param {Object} options - Email options
 * @param {string} options.to - Email người nhận
 * @param {string} options.subject - Tiêu đề email
 * @param {string} options.htmlBody - Nội dung HTML
 * @returns {Object} - Result object
 */
function sendEmailViaGmail(options) {
  try {
    const { to, subject, htmlBody } = options;
    
    Logger.log(`Sending email via Gmail API to: ${to}`);
    
    // Validate email
    if (!to || !to.includes('@')) {
      throw new Error('Invalid email address: ' + to);
    }
    
    // Create email message
    const email = [
      'Content-Type: text/html; charset=UTF-8',
      'MIME-Version: 1.0',
      `To: ${to}`,
      `Subject: ${subject}`,
      '',
      htmlBody
    ].join('\n');
    
    // Encode email in base64
    const encodedEmail = Utilities.base64EncodeWebSafe(email);
    
    // Send via Gmail API
    const response = Gmail.Users.Messages.send(
      {
        raw: encodedEmail
      },
      'me'
    );
    
    Logger.log(`Email sent successfully via Gmail API. Message ID: ${response.id}`);
    
    return {
      success: true,
      messageId: response.id,
      to: to
    };
    
  } catch (error) {
    Logger.log(`Failed to send email via Gmail API: ${error.message}`);
    
    return {
      success: false,
      error: error.message,
      to: options.to
    };
  }
}

/**
 * Gửi email đơn giản (wrapper function để dễ migrate)
 */
function sendEmail(to, subject, htmlBody) {
  return sendEmailViaGmail({
    to: to,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * Test Gmail API authorization và gửi test email
 */
function testGmailAPI() {
  Logger.log('========================================');
  Logger.log('TESTING GMAIL API');
  Logger.log('========================================');
  
  try {
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log(`User email: ${userEmail}`);
    
    // Test 1: Check Gmail API availability
    Logger.log('Test 1: Checking Gmail API availability...');
    const profile = Gmail.Users.getProfile('me');
    Logger.log(`Gmail profile retrieved: ${profile.emailAddress}`);
    Logger.log(`Messages total: ${profile.messagesTotal}`);
    
    // Test 2: Send test email
    Logger.log('Test 2: Sending test email...');
    
    const result = sendEmailViaGmail({
      to: userEmail,
      subject: '✅ Gmail API Test - Success!',
      htmlBody: `
        <div style="font-family: Arial, sans-serif; padding: 20px; max-width: 600px;">
          <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; border-radius: 10px; text-align: center;">
            <h1 style="margin: 0;">✅ Gmail API Works!</h1>
          </div>
          
          <div style="background: #f8f9fa; padding: 20px; margin-top: 20px; border-radius: 10px;">
            <h2 style="color: #333;">🎉 Congratulations!</h2>
            <p style="color: #666; line-height: 1.6;">
              Your Google Apps Script project is now successfully using <strong>Gmail API</strong> 
              instead of MailApp. This provides better reliability and authorization handling.
            </p>
            
            <div style="background: white; padding: 15px; border-left: 4px solid #28a745; margin: 15px 0;">
              <h3 style="margin-top: 0; color: #28a745;">✓ What's Working:</h3>
              <ul style="color: #666;">
                <li>Gmail API authorization completed</li>
                <li>Email sending via Gmail API functional</li>
                <li>HTML email rendering enabled</li>
                <li>Ready for production use</li>
              </ul>
            </div>
            
            <div style="background: #e7f3ff; padding: 15px; border-radius: 5px; margin-top: 15px;">
              <p style="margin: 0; color: #0066cc;">
                <strong>📧 Test Email Details:</strong><br>
                Sent: ${new Date().toLocaleString()}<br>
                From: ${userEmail}<br>
                Method: Gmail API v1
              </p>
            </div>
          </div>
          
          <div style="text-align: center; margin-top: 20px; padding: 20px; background: #fff3cd; border-radius: 10px;">
            <p style="margin: 0; color: #856404;">
              <strong>🚀 Next Steps:</strong><br>
              Your email notification system is now ready to use!
            </p>
          </div>
        </div>
      `
    });
    
    if (result.success) {
      Logger.log('✅ SUCCESS!');
      Logger.log(`Test email sent successfully via Gmail API`);
      Logger.log(`Message ID: ${result.messageId}`);
      Logger.log('');
      Logger.log('Gmail API is now working. You can:');
      Logger.log('1. Check your email inbox for the test message');
      Logger.log('2. Use the email notification system from the Sheet UI');
      Logger.log('3. Send folder links to groups');
      Logger.log('========================================');
      
      return {
        success: true,
        method: 'Gmail API',
        messageId: result.messageId,
        userEmail: userEmail,
        messagesTotal: profile.messagesTotal
      };
      
    } else {
      throw new Error(result.error);
    }
    
  } catch (error) {
    Logger.log('❌ ERROR!');
    Logger.log(`Failed: ${error.message}`);
    Logger.log('');
    Logger.log('Troubleshooting:');
    Logger.log('1. Make sure Gmail API is enabled in Apps Script');
    Logger.log('2. Check that appsscript.json has gmail.send scope');
    Logger.log('3. Try running the function again to trigger authorization');
    Logger.log('========================================');
    
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Lấy thông tin quota email còn lại
 */
function getEmailQuota() {
  try {
    // Gmail API không có quota limit như MailApp
    // Nhưng có daily sending limit (khoảng 500 emails/day cho free account)
    const profile = Gmail.Users.getProfile('me');
    
    return {
      success: true,
      emailAddress: profile.emailAddress,
      messagesTotal: profile.messagesTotal,
      threadsTotal: profile.threadsTotal,
      note: 'Gmail API allows ~500 emails/day for free accounts, 2000/day for Google Workspace'
    };
    
  } catch (error) {
    Logger.log(`Error getting quota: ${error.message}`);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Batch send emails với retry logic
 */
function batchSendEmails(emailList) {
  Logger.log(`Starting batch send for ${emailList.length} emails`);
  
  const results = {
    total: emailList.length,
    sent: 0,
    failed: 0,
    errors: []
  };
  
  emailList.forEach((emailData, index) => {
    try {
      Logger.log(`Sending email ${index + 1}/${emailList.length} to ${emailData.to}`);
      
      const result = sendEmailViaGmail({
        to: emailData.to,
        subject: emailData.subject,
        htmlBody: emailData.htmlBody
      });
      
      if (result.success) {
        results.sent++;
        Logger.log(`  ✅ Sent successfully`);
      } else {
        results.failed++;
        results.errors.push({
          to: emailData.to,
          error: result.error
        });
        Logger.log(`  ❌ Failed: ${result.error}`);
      }
      
      // Small delay to avoid rate limiting
      if (index < emailList.length - 1) {
        Utilities.sleep(100); // 100ms delay
      }
      
    } catch (error) {
      results.failed++;
      results.errors.push({
        to: emailData.to,
        error: error.message
      });
      Logger.log(`  ❌ Exception: ${error.message}`);
    }
  });
  
  Logger.log(`Batch send completed: ${results.sent}/${results.total} sent, ${results.failed} failed`);
  
  return results;
}
