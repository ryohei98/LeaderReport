function sendEmail(obj,members) {

  var MailTo = "ryohei.kond@gmail.com";
  var MailFrom = "konpei.to98@gmail.com";
  var NameFrom = "Leader Report";
  var Title = "リーダー報告";

  var html = HtmlService.createTemplateFromFile("mailTemplate");



  GmailApp.sendEmail(
    MailTo,
    Title,
    "htmlメールを有効にしてください",
    {
      from: MailFrom,
      name: NameFrom,
      htmlBody: html;
    }
  );
}
