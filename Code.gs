/**
 * Google Apps Script - Elle Belle Billing
 * Google Sheet Save + WhatsApp via WabaStore API
 */

// ============ CONFIG ============
var WHATSAPP_TOKEN = 'K3UzL1xUoIEedE1qeTMR9XNBqXWqkpjtzCbCXKMCYBgIHeKvQG2Jr8pvoazV3onHFbn77rPgj6JEMrzSLc5hrdq6qfOFBNcPsshvlsaJOI5HqREFTSAoGYctmxzRTr68';
var WHATSAPP_URL = 'https://crmapi.wabastore.com/api/meta/v19.0/487693087754528/messages';
var TEMPLATE_NAME = 'erer_biling';
// =================================

// ============ CUSTOM MENU ============
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('WhatsApp')
    .addItem('Send Message to Selected Row', 'sendWhatsAppFromSheet')
    .addItem('Send Message to All Unsent Rows', 'sendWhatsAppToAllUnsent')
    .addToUi();
}

/**
 * Selected row se WhatsApp message bhejta hai
 * Sheet columns: A=Date, B=Bill No, C=Customer Name, D=Mobile No., E=Total Amount, F=Payment Mode, G=Pending Balance, H=Delivery Date, I=WhatsApp Status
 */
function sendWhatsAppFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');
  var ui = SpreadsheetApp.getUi();
  var activeRow = sheet.getActiveRange().getRow();

  if (activeRow <= 1) {
    ui.alert('Please select a data row (not the header).');
    return;
  }

  var rowData = sheet.getRange(activeRow, 1, 1, 9).getValues()[0];
  var phone = rowData[3].toString().trim();

  if (!phone) {
    ui.alert('This row has no phone number.');
    return;
  }

  var result = sendTemplateMessage({
    date: rowData[0],
    billNo: rowData[1],
    customerName: rowData[2],
    phone: phone,
    totalAmount: rowData[4],
    paymentMode: rowData[5],
    pendingBalance: rowData[6],
    deliveryDate: rowData[7]
  });

  // Update WhatsApp Status column (I)
  sheet.getRange(activeRow, 9).setValue(result);
  ui.alert('Result: ' + result);
}

/**
 * Sabhi unsent rows ko WhatsApp message bhejta hai
 */
function sendWhatsAppToAllUnsent() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DATA');
  var ui = SpreadsheetApp.getUi();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    ui.alert('No data rows found.');
    return;
  }

  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();
  var sentCount = 0;
  var failCount = 0;

  for (var i = 0; i < data.length; i++) {
    var status = data[i][8].toString();
    var phone = data[i][3].toString().trim();

    // Skip if already sent or no phone
    if (!phone || status.indexOf('Sent:') === 0) continue;

    var result = sendTemplateMessage({
      date: data[i][0],
      billNo: data[i][1],
      customerName: data[i][2],
      phone: phone,
      totalAmount: data[i][4],
      paymentMode: data[i][5],
      pendingBalance: data[i][6],
      deliveryDate: data[i][7]
    });

    sheet.getRange(i + 2, 9).setValue(result);

    if (result.indexOf('Sent:') === 0) {
      sentCount++;
    } else {
      failCount++;
    }

    // Rate limit - avoid API throttling
    Utilities.sleep(1000);
  }

  ui.alert('Done! Sent: ' + sentCount + ', Failed: ' + failCount);
}

/**
 * Common function - template message bhejne ke liye
 * Template: erer_biling
 * {{1}} = Customer Name
 * {{2}} = Bill No
 * {{3}} = Total Bill Amt
 * {{4}} = Cash Paid
 * {{5}} = Card Amt
 * {{6}} = Digital Amt (UPI/Online)
 * {{7}} = Balance Amt (Pending)
 * {{8}} = Delivery Date
 */
function sendTemplateMessage(data) {
  var phoneNumber = data.phone.toString();
  if (!phoneNumber.startsWith('91')) {
    phoneNumber = '91' + phoneNumber;
  }

  var totalAmt = String(data.totalAmount || 0);
  var paymentMode = (data.paymentMode || 'Cash').toString();
  var cashPaid = '0';
  var cardAmt = '0';
  var digitalAmt = '0';

  if (paymentMode === 'Cash') {
    cashPaid = totalAmt;
  } else if (paymentMode === 'Card') {
    cardAmt = totalAmt;
  } else if (paymentMode === 'UPI' || paymentMode === 'Online') {
    digitalAmt = totalAmt;
  }

  var deliveryDate = data.deliveryDate || '';
  if (deliveryDate instanceof Date) {
    deliveryDate = deliveryDate.toLocaleDateString('en-IN');
  }

  try {
    var bodyParams = [
      { type: 'text', text: data.customerName || 'Customer' },
      { type: 'text', text: 'Bill No: ' + String(data.billNo || '-') },
      { type: 'text', text: 'Total Bill Amt: Rs.' + totalAmt },
      { type: 'text', text: 'Cash Paid: Rs.' + cashPaid },
      { type: 'text', text: 'Card Amt: Rs.' + cardAmt },
      { type: 'text', text: 'Digital Amt: Rs.' + digitalAmt },
      { type: 'text', text: 'Balance Amt: Rs.' + String(data.pendingBalance || 0) },
      { type: 'text', text: 'Delivery Date: ' + (deliveryDate || 'N/A') },
      { type: 'text', text: 'Thanks For Shopping, See you Soon!' },
      { type: 'text', text: 'Save our contact for latest offers - tinyurl.com/ElleBelleDigitalCard' },
      { type: 'text', text: 'Elle Belle - she is beautiful' },
      { type: 'text', text: 'Mumbai, Ghatkopar (E) - 9892944944 / 8958957575' }
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
      return 'Sent: ' + waBody.substring(0, 200);
    } else {
      return 'Failed: ' + waCode + ' - ' + waBody.substring(0, 200);
    }
  } catch (waErr) {
    return 'Error: ' + waErr.toString().substring(0, 150);
  }
}

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
        'Mobile No.',
        'Total Amount',
        'Payment Mode',
        'Pending Balance',
        'Delivery Date',
        'WhatsApp Status',
        'Message',
        'Digital Card Link',
        'Brand',
        'Contact'
      ]);
      var headerRange = sheet.getRange(1, 1, 1, 13);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#000000');
      headerRange.setFontColor('#FFFFFF');
      sheet.setFrozenRows(1);
    }

    // WhatsApp message
    var whatsappStatus = 'Not Sent';
    if (data.phone) {
      whatsappStatus = sendTemplateMessage(data);
    }

    // Sheet mein save
    var dateVal = data.date || new Date().toLocaleDateString('en-IN');
    sheet.appendRow([
      'Date - ' + dateVal,
      'Bill No. - ' + (data.billNo || ''),
      data.customerName || '',
      'Mobile No. - ' + (data.phone || ''),
      'Total Amount - Rs.' + String(data.totalAmount || 0),
      'Payment Mode - ' + (data.paymentMode || ''),
      'Pending Balance - Rs.' + String(data.pendingBalance || 0),
      'Delivery Date - ' + (data.deliveryDate || ''),
      whatsappStatus,
      'Thanks For Shopping, See you Soon!',
      'https://tinyurl.com/ElleBelleDigitalCard',
      '"ऐल बेल" Elle Belle she is beautiful "એલ બેલ"',
      '📍Mumbai, Ghatkopar (E) 📞 9892944944 / 8958957575'
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
