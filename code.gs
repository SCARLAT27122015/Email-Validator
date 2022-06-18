const emailColName = 'Email';

function validateEmail(email)  {
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(email);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom functions')
      .addItem('Evaluate Emails', 'evaluateEmail')
      .addToUi();
}

function evaluateEmail() {
  SpreadsheetApp.getUi(); // Or DocumentApp or FormApp.
  //.alert('You clicked the first menu item!');
  const app = SpreadsheetApp;
  const ss = app.getActiveSpreadsheet();
  const shSource = ss.getActiveSheet();
  const data = shSource.getDataRange().getValues();
  const emailPos = data[0].indexOf(emailColName);
  let validated = data.map((row, index) => {
    
    if (index === 0){
      return ['Email Validation'];
    }else {
      let email = validateEmail(row[emailPos]);
      return [email];
    }
    
  });

  shSource.getRange(1, shSource.getLastColumn() + 1, validated.length, validated[0].length).setValues(validated);


}
