/**
 * Google Apps Script - Elle Belle Billing
 * Google Sheet Save + WhatsApp via WabaStore API
 */

// ============ CONFIG ============
var WHATSAPP_TOKEN = 'K3UzL1xUoIEedE1qeTMR9XNBqXWqkpjtzCbCXKMCYBgIHeKvQG2Jr8pvoazV3onHFbn77rPgj6JEMrzSLc5hrdq6qfOFBNcPsshvlsaJOI5HqREFTSAoGYctmxzRTr68';
var WHATSAPP_URL = 'https://crmapi.wabastore.com/api/meta/v19.0/487693087754528/messages';
var TEMPLATE_NAME = 'erer_biling';
// =================================

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');

    // Headers
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Date',
        'Bill No',
        'Customer Name',
        'Phone',
        'Total Amount',
        'Payment Mode',
        'Pending Balance',
        'WhatsApp Status'
      ]);
      var headerRange = sheet.getRange(1, 1, 1, 8);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#000000');
      headerRange.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // WhatsApp message
    var whatsappStatus = 'Not Sent';
    if (data.phone) {
      var phoneNumber = data.phone.toString();
      if (!phoneNumber.startsWith('91')) {
        phoneNumber = '91' + phoneNumber;
      }

      try {
        var bodyParams = [
          { type: 'text', text: data.customerName || 'Customer' },
          { type: 'text', text: data.billNo || '-' },
          { type: 'text', text: data.date || new Date().toLocaleDateString('en-IN') },
          { type: 'text', text: 'Rs.' + String(data.totalAmount || 0) },
          { type: 'text', text: data.paymentMode || 'Cash' },
          { type: 'text', text: 'Rs.' + String(data.pendingBalance || 0) },
          { type: 'text', text: phoneNumber },
          { type: 'text', text: 'Elle Belle' },
          { type: 'text', text: 'Thank you for shopping with us!' },
          { type: 'text', text: data.date || new Date().toLocaleDateString('en-IN') },
          { type: 'text', text: 'Visit again!' },
          { type: 'text', text: 'Elle Belle Billing' }
        ];

        var whatsappPayload = {
          to: phoneNumber,
          recipient_type: 'individual',
          type: 'template',
          template: {
            language: {
              policy: 'deterministic',
              code: 'el'
            },
            name: TEMPLATE_NAME,
            components: [
              {
                type: 'body',
                parameters: bodyParams
              }
            ]
          }
        };

        var options = {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'Authorization': 'Bearer ' + WHATSAPP_TOKEN
          },
          payload: JSON.stringify(whatsappPayload),
          muteHttpExceptions: true
        };

        var waResponse = UrlFetchApp.fetch(WHATSAPP_URL, options);
        var waCode = waResponse.getResponseCode();
        var waBody = waResponse.getContentText();

        if (waCode >= 200 && waCode < 300) {
          whatsappStatus = 'Sent: ' + waBody.substring(0, 200);
        } else {
          whatsappStatus = 'Failed: ' + waCode + ' - ' + waBody.substring(0, 200);
        }
      } catch (waErr) {
        whatsappStatus = 'Error: ' + waErr.toString().substring(0, 150);
      }
    }

    // Sheet mein save
    sheet.appendRow([
      data.date || new Date().toLocaleDateString('en-IN'),
      data.billNo || '',
      data.customerName || '',
      data.phone || '',
      data.totalAmount || 0,
      data.paymentMode || '',
      data.pendingBalance || 0,
      whatsappStatus
    ]);

    return ContentService.createTextOutput(
      JSON.stringify({ status: 'success', whatsapp: whatsappStatus })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var phone = e.parameter.phone;
    if (!phone) {
      return ContentService.createTextOutput(
        JSON.stringify({ orders: [] })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');
    var data = sheet.getDataRange().getValues();
    var orders = [];
    var customerName = '';

    for (var i = 1; i < data.length; i++) {
      if (data[i][3].toString() === phone.toString()) {
        customerName = data[i][2];
        orders.push({
          date: data[i][0],
          billNo: data[i][1],
          totalAmount: data[i][4],
          paymentMode: data[i][5],
        });
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        customerName: customerName,
        orders: orders.reverse()
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ orders: [], error: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
