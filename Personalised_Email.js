function sendEmailsFromGroupExtension() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Email Comms');
  var lastRow = sheet.getLastRow();
  var dataRange = sheet.getRange('B4:N' + lastRow);
  var data = dataRange.getValues();

  var fixedSubject = '{Subject lime or cell containing the subject}';
  var groupEmail = '{Sender's Email}';

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var emailAddress = row[0]; // Column B has the email addresses
    var sentMarker = row[8]; // Column J for the sent marker

    // Skip if email is already sent
    if (sentMarker === 'Sent') continue;

    // Handle potentially empty CC fields
    var ccList = [row[5], row[6], row[12]].filter(function(email) {
      return email; // This will filter out empty strings
    }).join(',');

    var emailBodyHtml = row[7]; // Column H has the email body in HTML format

    if (emailAddress) {
      // Send email
      GmailApp.sendEmail(emailAddress, fixedSubject, '', {
        cc: ccList,
        from: groupEmail,
        htmlBody: emailBodyHtml
      });

      // Mark as sent in column J for the same row starting from row 2
      sheet.getRange(i + 2, 10).setValue('Sent'); // Row index is i + 2, column index for J is 10
    }
  } 
} 
