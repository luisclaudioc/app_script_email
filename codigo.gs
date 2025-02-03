const EMAIL_BODY = `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
      body {
        font-family: Arial, sans-serif;
        background-color: #f9f9f9;
        margin: 0;
        padding: 0;
        color: #333;
      }
      .container {
        max-width: 600px;
        margin: 20px auto;
        background: #ffffff;
        border-radius: 8px;
        overflow: hidden;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        padding: 20px;
      }
      .header {
        text-align: center;
        font-size: 24px;
        font-weight: bold;
        color: #4CAF50;
        margin-bottom: 20px;
      }
      .content {
        font-size: 16px;
        line-height: 1.6;
      }
      .footer {
        text-align: center;
        font-size: 12px;
        color: #777;
        margin-top: 20px;
      }
    </style>
  </head>
  <body>
    <div class="container">
      <div class="header">Thank You for Your Purchase!</div>
      <div class="content">
        <p>Hi, <strong>{clientName}</strong>!</p>
        <p>Thanks for your purchase. Your total is <strong>{total} euros</strong>.</p>
        <p>Find the invoice attached to this email.</p>
        <p>Best regards,</p>
        <br>
        <p><strong>Delícias de Minas Lda</strong></p>
      </div>
      <div class="footer">
        &copy; 2024 DuForno Salgados. All rights reserved.
      </div>
    </div>
  </body>
  </html>
`; 


function enviarEmail() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales");
  var range = sheet.getRange(2, 1, (sheet.getLastRow() -1), 8);
  var values = range.getValues();

  values.forEach((order, orderIndex) => {
    if (order[6] == "Enviado" && order[7] !== "OK") {
      var email = order[2];
      var invoiceId = order[4];
      var emailData = {
        clientName: order[1],
        total: order[5],
      }
      var emailBody = fillTemplate(EMAIL_BODY, emailData);
      sheet.getRange(orderIndex + 2, 8).setValue("OK");

      var invoice = DriveApp.getFileById(invoiceId).getBlob();

      var options = {
        htmlBody: emailBody,
        name: "Delícias de Minas Lda",
        attachments: [invoice],
      }

      // Send email
      GmailApp.sendEmail(email, "Purchase Delícias de Minas Lda", "", options);
      
    }
  });
}

function fillTemplate(emailTemplate, emailData) {
  return emailTemplate.replace(/\{(\w+)\}/g, (placeholder, word) => {
    return emailData[word] || placeholder;
  });
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Actions")
    .addItem("Enviar Email", "enviarEmail")
    .addToUi(); 
}


