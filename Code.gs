// This is a script to inject row data from a google sheet to a template in google doc.
// It is triggered by google form submission.
// It also gathered some environmental parameters from another sheet "dev" in the google sheet.
//
// This script has originally made by automagictv, https://gist.github.com/automagictv/48bc3dd1bc785601422e80b2de98359e
// Author: Shayan Ghiaseddin
// Author URI: sghiaseddin.com
//

const TEMPLATE_FILE_ID = 'YOUR_TEMPLATE_FILE_ID_HERE';
const DESTINATION_FOLDER_ID = 'YOUR_DESTINATION_FOLDER_ID_HERE';
const CURRENCY_SIGN = '$';

// Converts a float to a string value in the desired currency format
function toCurrency(num) {
    var fmt = Number(num).toFixed(2);
    return `${CURRENCY_SIGN}${fmt}`;
}

// Format datetimes to: YYYY-MM-DD
function toDateFmt(dt_string) {
  var millis = Date.parse(dt_string);
  var date = new Date(millis);
  var year = date.getFullYear();
  var month = ("0" + (date.getMonth() + 1)).slice(-2);
  var day = ("0" + date.getDate()).slice(-2);

  // Return the date in YYYY-mm-dd format
  return `${year}-${month}-${day}`;
}


// Parse and extract the data submitted through the form.
function parseFormData(values, header) {
    // Set temporary variables to hold prices and data.
    var subtotal = 0;
    var tax = 0;
    var total = 0;
    var response_data = {};
    var quantity = 1;
    var item_total = [];

    // Get the receipt number from the second sheet
    var sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("dev");
    var current_receipt_number = parseInt(sheet2.getRange("B2").getValue(), 10); // For reciept auto numbering
    var payment_methods = sheet2.getRange("B3").getValue(); // To include payment methods under the table

    // Iterate through all of our response data and add the keys (headers)
    // and values (data) to the response dictionary object.
    for (var i = 0; i < values.length; i++) {
      // Extract the key and value
      var key = header[i];
      var value = values[i];

      // If we have a quantity, store it for the next iteration
      if (key.toLowerCase().includes("quantity")) {
        quantity = value;

      // If we have a price, add it to the running subtotal and format it to the
      // desired currency.
      } else if (key.toLowerCase().includes("price")) {
        item_total.push( value * quantity);
        subtotal += value * quantity;
        if ( value != '' ) {
          value = toCurrency(value);
        }

      // Format dates
      } else if (key.toLowerCase().includes("date")) {
        value = toDateFmt(value);

      // Invoice related fields
      } else if (key.toLowerCase().includes("document title")) {
        if ( value.toLowerCase() == 'invoice' ) { // invoices must have date and payment methods
          response_data["_receipt_date_title"] = 'Invoice date';
          response_data["_receipt_date"] = toDateFmt(response_data["Timestamp"]);
          response_data["_payment_methods"] = payment_methods;
        } else { // receipt
          response_data["_receipt_date_title"] = '';
          response_data["_receipt_date"] = '';
          response_data["_payment_methods"] = '';
        }
      }

      // Add the key/value data pair to the response dictionary.
      response_data[key] = value;
    }

    // Once all data is added, we'll adjust the items total and subtotal
    for (var i = 0; i < item_total.length; i++) {
      if ( item_total[i] != '' || item_total[i] != 0 ) {
        response_data["_item_total_" + (i + 1).toString()] = toCurrency(item_total[i]);
      } else {
        response_data["_item_total_" + (i + 1).toString()] = '';
      }
    }
    
    // Setting total and tax and number
    response_data["_total"] = toCurrency(subtotal * (1 + response_data["Tax"] / 100 ));
    response_data["Tax"] = toCurrency(subtotal * response_data["Tax"] / 100);
    response_data["_subtotal"] = toCurrency(subtotal);
    response_data["_receipt_number"] = current_receipt_number + 1; // Increment receipt number

    // Update the receipt number in the sheet for next use
    sheet2.getRange("B2").setValue(current_receipt_number + 1);

    return response_data;
}

// Helper function to inject data into the template
function populateTemplate(document, response_data) {

    // Get the document header and body (which contains the text we'll be replacing).
    var document_header = document.getHeader();
    var document_body = document.getBody();

    // Replace variables in the header
    for (var key in response_data) {
      var match_text = `{{${key}}}`;
      var value = response_data[key];

      // Replace our template with the final values
      document_header.replaceText(match_text, value);
      document_body.replaceText(match_text, value);
    }

}


// Function to populate the template form
function createDocFromForm() {

  // Get active sheet and tab of our response data spreadsheet.
  var sheet = SpreadsheetApp.getActiveSheet();
  var last_row = sheet.getLastRow() - 1;

  // Get the data from the spreadsheet.
  var range = sheet.getDataRange();
 
  // Identify the most recent entry and save the data in a variable.
  var data = range.getValues()[last_row];
  
  // Extract the headers of the response data to automate string replacement in our template.
  var headers = range.getValues()[0];

  // Parse the form data.
  var response_data = parseFormData(data, headers);

  // Retreive the template file and destination folder.
  var template_file = DriveApp.getFileById(TEMPLATE_FILE_ID);
  var target_folder = DriveApp.getFolderById(DESTINATION_FOLDER_ID);

  // Copy the template file so we can populate it with our data.
  // The name of the file will be the company name and the invoice number in the format: DATE_COMPANY_NUMBER
  var filename = `${response_data["_receipt_number"]}_${response_data["Full Name"]}_${toDateFmt(response_data["Timestamp"])}`;
  var document_copy = template_file.makeCopy(filename, target_folder);

  // Open the copy.
  var document = DocumentApp.openById(document_copy.getId());

  // Populate the template with our form responses and save the file.
  populateTemplate(document, response_data);
  // document.saveAndClose();
}

// Test and debug
// createDocFromForm()