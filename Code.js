function sendEmailWithHtmlTable() {
  // Get the active spreadsheet and the relevant sheet 'Daily Email Report'
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('your email sheet name');
  
  // Get the data range starting at A1 (with the header as row 1, up to row 25, you can modify this range)
  var dataRange = sheet.getRange('A1:L25');
  var data = dataRange.getValues();
  
  // Get the styling information for the entire range
  var backgrounds = dataRange.getBackgrounds();
  // Removed fontColors since the font color is uniform
  var fontSizes = dataRange.getFontSizes();
  
  // Define headerRows as an array of sheet row numbers that should use header styling.
  var headerRows = [2, 14, 15];
  
  // Define which columns require special formatting (zero-based column indices)
  var currencyColumns = [3, 5, 7, 10];
  var percentageColumns = [8, 11];
  
  // Create HTML table with a title and adjust the heading style
  var html = '<h2 style="font-size:15px; margin-bottom:15px;">your Report</h2>';
  html += '<table style="width: 100%; border-collapse: collapse;">';
  
  // Build the table header from row 1 in the sheet (data index 0) and include all columns except the last one.
  html += '<thead><tr>';
  for (var j = 0; j < data[0].length - 1; j++) {
    html += '<th style="border: 1px solid #000; background-color: ' + backgrounds[0][j] + 
            '; font-size: ' + fontSizes[0][j] + 
            'px; padding: 8px;">' + data[0][j] + '</th>';
  }
  html += '</tr></thead><tbody>';
  
  // Loop through the remaining rows (starting from index 1 corresponds to sheet row 2)
  for (var i = 1; i < data.length; i++) {
    html += '<tr>';
    
    if (headerRows.indexOf(i + 1) !== -1) {
      // For header rows (sheet row 2, 14, 15) apply specific styling.
      // For row 13 (i.e., sheet row 14), we need to reduce our total columns by 1.
      var maxCols = (i === 13) ? data[i].length - 1 : data[i].length;
      for (var j = 0; j < maxCols; j++) {
        // Skip the third column (index 2) except in row 13
        if (j === 2 && i !== 13) {
          continue;
        }
        html += '<td style="border: 1px solid #000; background-color: ' + backgrounds[i][j] +
                '; font-size: ' + fontSizes[i][j] +
                'px; padding: 8px; font-weight: bold; text-align: center;">' + data[i][j] + '</td>';
      }
    } else if (i === 12 || i === 13) {
      // For rows 13 and 14 in the sheet, output raw cell data without formatting.
      // Again, reduce total columns by 1 for row 13.
      var maxCols = (i === 13) ? data[i].length - 1 : data[i].length;
      for (var j = 0; j < maxCols; j++) {
        // Skip the third column unless it's row 14 (i.e., when i === 13)
        if (j === 2 && i !== 13) {
          continue;
        }
        var cellValue = data[i][j];
        if (cellValue === '' || cellValue === null) {
          cellValue = '&nbsp;';
        }
        html += '<td>' + cellValue + '</td>';
      }
    } else {
      // For all other rows, apply currency and percentage formatting along with the original styling.
      for (var j = 0; j < data[i].length; j++) {
        // Skip the third column unless it's row 14 (i.e., when i === 13)
        if (j === 2 && i !== 13) {
          continue;
        }
        var cellValue = data[i][j];
        var formattedValue = cellValue;
        
        // Apply currency and percentage formatting based on the column.
        if (currencyColumns.includes(j)) {
          formattedValue = '$' + Math.round(cellValue);
        } else if (percentageColumns.includes(j)) {
          formattedValue = Math.round(cellValue * 100) + '%';
        }
        
        html += '<td style="border: 1px solid #000; background-color: ' + backgrounds[i][j] +
                '; font-size: ' + fontSizes[i][j] + 'px; padding: 8px;">' + formattedValue + '</td>';
      }
    }
    
    html += '</tr>';
  }
  
  html += '</tbody></table>';
  
  /* Display the HTML in a modal dialog for preview (or you can uncomment the email sending part below)
  var htmlOutput = HtmlService.createHtmlOutput(html)
                              .setWidth(800)
                              .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Campaign Report Preview');
  
 */
// Uncomment this section to send the email
  MailApp.sendEmail({
    to: 'xyz@gmail.com',
    subject: 'your email report subject',
    body: 'Your email client does not support HTML emails.',
    htmlBody: html
  });
  
}
