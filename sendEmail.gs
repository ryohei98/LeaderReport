function sendEmail(obj,members) {

  var MailTo = "ryohei.kond@gmail.com";
  var MailFrom = "konpei.to98@gmail.com";
  var NameFrom = "Leader Report";
  var Title = "リーダー報告";

  var html =  "<!DOCTYPE html>
                <html lang='ja'>

                <head>
                  <style>

                    .main {
                      text-align: center;
                    }

                    .content-wrapper {
                      text-align: left;
                    }
                  </style>
                  <meta charset='UTF-8' />
                  <meta name='viewport' content='width=device-width, initial-scale=1.0' />
                  <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
                </head>

                <body>
                  <div class='container main'>
                    <h1 class='py-5'>リーダー報告書</h1>
                    <p>もし上手く見えなければ、ウインドウを最大化してください</p>
                    <div class='content-wrapper'>
                      <table class='table table-hover mb-5'>
                        <tbody>
                          <tr>
                            <th scope='row'>会場名</th>
                            <td>"+ obj["PlaceName"] +"</td>
                          </tr>
                          <tr>
                            <th scope='row'>リーダー名</th>
                            <td>" + obj["LeaderName"] + "</td>
                          </tr>
                          <tr>
                            <th scope='row'>開催日</th>
                            <td>" + obj["Date"] + "</td>
                          </tr>
                        </tbody>
                      </table>
                      <h1>一日目</h1>
                      <h2>個人記録</h2>
                      <table class='table table-hover'>
                        <thead>
                          <tr>
                            <th scope='col'>メンバー名</th>
                            <th scope='col'>早退遅刻</th>
                            <th scope='col'>時間</th>
                            <th scope='col'>目標入会率</th>
                            <th scope='col'>個人入会率(実績)</th>
                            <th scope='col'>接客数(実績)</th>
                            <th scope='col'>入会数(実績)</th>
                            <th scope='col'>自宅入会数(実績)</th>
                            <th scope='col'>休憩予定</th>
                            <th scope='col'>休憩結果</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            ";

for(var member in members){
  var thisMember = members[member];
  html += "<th scope='row'>" + thisMember["MemberName1"] +"</th>
          ";
  html += "<td>" + thisMember["Attendance1"] +"</td>
          ";
  html += "<td>" + thisMember["TimeLate1"]+"~"+thiMember["TimeEarly1"] +"</td>
          ";
  html += "<td>" + thisMember["TargetRate1"] +"</td>
          ";
  html += "<td>" + thisMember["NumServed1"] +"</td>
          ";
  html += "<td>" + thisMember["NumApplying1"] +"</td>
          ";
  html += "<td>" + thisMember["NumApplyingHome1"] +"</td>
          ";
  html += "<td>" + thisMember["BreakTimePlanned1"] +"</td>
          ";
  html += "<td>" + thisMember["BreakTimeActual1"] +"</td>
          </tr>";
}

html += "     </tbody>
            </table>
            <h2>時間帯別の記録</h2>
            <table class='table table-hover'>
            <thead>
              <tr>
                <th scope='col'></th>
                <th scope='col'>11:30</th>
                <th scope='col'>13:00</th>
                <th scope='col'>15:00</th>
                <th scope='col'>17:00</th>
                <th scope='col'>19:00</th>
                <th scope='col'>それ以降</th>
              </tr>
            </thead>
            <tbody>
              <tr>
              ";

for ( var num in [0,1,2,3,4,5]){
  
}




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

}
