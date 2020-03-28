function myFunction() {
  function testSendEmail() {
  var Sheet = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet1 = Sheet.getSheetByName("Sheet2");
  Sheet1.activate();

  var MailTo = "ryohei.kond@gmail.com";
  var MailFrom = "konpei.to98@gmail.com";
  var NameFrom = "Leader Report";
  var Title = "リーダー報告";

  var Value_B2 = parseInt(Sheet1.getRange("B2").getValue()*100);
  var Value_B3 = parseInt(Sheet1.getRange("B3").getValue()*100);
  var Value_B4 = parseInt(Sheet1.getRange("B4").getValue()*100);

  var Value_B7 = parseInt(Sheet1.getRange("B7").getValue()*100);
  var Value_B8 = parseInt(Sheet1.getRange("B8").getValue()*100);
  var Value_B9 = parseInt(Sheet1.getRange("B9").getValue()*100);

  var Value_B12 = parseInt(Sheet1.getRange("B12").getValue()*100);
  var Value_B13 = parseInt(Sheet1.getRange("B13").getValue()*100);
  var Value_B14 = parseInt(Sheet1.getRange("B14").getValue()*100);

  var Value_B17 = parseInt(Sheet1.getRange("B17").getValue()*100);
  var Value_B18 = parseInt(Sheet1.getRange("B18").getValue()*100);
  var Value_B19 = parseInt(Sheet1.getRange("B19").getValue()*100);

  var html =
"<!DOCTYPE html>
<html lang="ja">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width, initial-scale=1.0"/>
  <meta http-quiv="X-UA-Compatible" content="ie=edge"/>
  <title>Document</title>
</head>
<body>

</body>
</html>"



  GmailApp.sendEmail(
    MailTo,
    Title,
    Text,
    {
      from: MailFrom,
      name: NameFrom,
      htmlBody: html;
    }
  );
}

}
