function sendEmail(obj,members) {

  var MailTo = "ryohei.kond@gmail.com";
  var MailFrom = "konpei.to98@gmail.com";
  var NameFrom = "Leader Report";
  var Title = "リーダー報告";

  var html =  "<!DOCTYPE html>
                <html lang='ja'>

                <head>
                  <style>
                    .header{
                      background-color:#2288ff;
                      color: #fff;
                      margin:0;
                      margin-bottom: 10px;
                      padding: 10px 0 10px 0;
                      width:100%;
                      text-align: center;
                    }

                    .chapter-head{
                      background-color:#2288ff;
                      color: #fff;
                      margin:0;
                      margin-bottom: 10px;
                      padding: 10px 30px;
                      width:100%;
                      text-align:left;
                      border-radius: 5px;
                    }

                    .main {
                      text-align: center;
                    }

                    .content-wrapper {
                      text-align: left;
                    }

                    .chapter-wrapper {
                      text-align: left;
                      margin-bottom:20px;
                      border: 1px solid gray;
                      border-radius: 5px;
                    }
                  </style>
                  <meta charset='UTF-8' />
                  <meta name='viewport' content='width=device-width, initial-scale=1.0' />
                  <link rel='stylesheet' href='https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css' integrity='sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T' crossorigin='anonymous'>
                </head>

                <body>
                  <div class = 'header'>
                    <h1 class ='header'>リーダー報告書</h1>
                  </div>
                  <div class='container main'>
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
                      <div class='chapter-wrapper'>
                        <h2 class='chapter-head'>個人記録</h2>
                        <table class='table table-hover'>
                          <thead>
                            <tr>
                              <th scope='col'>メンバー名</th>
                              <th scope='col'>早退遅刻</th>
                              <th scope='col'>時間</th>
                              <th scope='col'>目標入会率</th>
                              <th scope='col'>入会率</th>
                              <th scope='col'>接客数</th>
                              <th scope='col'>入会数</th>
                              <th scope='col'>入会数</th>
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

html += "               </tbody>
                      </table>
                    </div>
                    <div class ='chapter-wrapper'>
                      <h2 class ='chapter-head'>時間帯別の記録</h2>
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
                            <th scope='col'>合計</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            <th scope='row'>一日目</th>
              ";

  var numServedArray = new Array(6).fill(0);
  for ( var num in [0,1,2,3,4,5]){
    var ID = "NumServedT" + String(num +1) + "D1";
    html += "<td>" + obj[ID] + "</td>
    ";
  }
  html += "<td>" + obj["NumServed1"] + "</td>";

  html += "</tr>
          <tr>
          "
  var nunApplyingArray = new Array(6).fill(0);
  for ( var num in [0,1,2,3,4,5]){
    var ID = "NumApplyingT" + String(num +1) + "D1";
    html += "<td>" + obj[ID] + "</td>
    ";
  }

  html += "<td>" + obj["NumApplying1"] + "</td>";
  html += "</tr>
          <tr>
          "

  var rateArray = new Array(6).fill(0);
  for( var num in [0,1,2,3,4,5]){
    var rate = String(Math.round(numApplyingArray[num] / numServedArray[num])) + "%";
    html += "<td>" + rate + "</td>
    ";
  }

  html += "</tr>
        </tbody>
        </table>
      </div>
      <div class ='chapter-wrapper'>
        <h2 class ='chapter-head'>来場者属性</h2>
        <table class='table table-hover'>
          <thead>
            <tr>
              <th scope='col'></th>
              <th scope='col'>チャレタ</th>
              <th scope='col'>公文</th>
              <th scope='col'>紙の進研ゼミ</th>
              <th scope='col'>宿題のみ</th>
              <th scope='col'>塾</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              "

  html += "<th scope='row'>受講歴のある来場者</th>
                      ";
  html += "<td>" + Obj["NumChareta1"] +"</td>
                              ";
  html += "<td>" + Obj["NumKumon1"] +"</td>
                              ";
  html += "<td>" + Obj["NumShinken1"] +"</td>
                              ";
  html += "<td>" + Obj["NumOnly1"] +"</td>
                              ";
  html += "<td>" + Obj["NumCram1"] +"</td>
                                    </tr>
                                    </tbody>
                                    </table>
                                    </div>";

  if(obj["SecondDay"]){
    html += "<h1>二日目</h1>
              <div class='chapter-wrapper'>
                <h2 class='chapter-head'>個人記録</h2>
                <table class='table table-hover'>
                  <thead>
                    <tr>
                      <th scope='col'>メンバー名</th>
                      <th scope='col'>早退遅刻</th>
                      <th scope='col'>時間</th>
                      <th scope='col'>目標入会率</th>
                      <th scope='col'>入会率</th>
                      <th scope='col'>接客数</th>
                      <th scope='col'>入会数</th>
                      <th scope='col'>入会数</th>
                      <th scope='col'>休憩予定</th>
                      <th scope='col'>休憩結果</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                    ";

for(var member in members){
var thisMember = members[member];
html += "<th scope='row'>" + thisMember["MemberName2"] +"</th>
";
html += "<td>" + thisMember["Attendance2"] +"</td>
";
html += "<td>" + thisMember["TimeLate2"]+"~"+thiMember["TimeEarly1"] +"</td>
";
html += "<td>" + thisMember["TargetRate2"] +"</td>
";
html += "<td>" + thisMember["NumServed2"] +"</td>
";
html += "<td>" + thisMember["NumApplying2"] +"</td>
";
html += "<td>" + thisMember["NumApplyingHome2"] +"</td>
";
html += "<td>" + thisMember["BreakTimePlanned2"] +"</td>
";
html += "<td>" + thisMember["BreakTimeActual2"] +"</td>
</tr>";
}

html += "               </tbody>
    </table>
  </div>
  <div class ='chapter-wrapper'>
    <h2 class ='chapter-head'>時間帯別の記録</h2>
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
          <th scope='col'>合計</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <th scope='row'>二日目</th>
";

var numServedArray = new Array(6).fill(0);
for ( var num in [0,1,2,3,4,5]){
var ID = "NumServedT" + String(num +1) + "D2";
html += "<td>" + obj[ID] + "</td>
";
}
html += "<td>" + obj["NumServed1"] + "</td>";

html += "</tr>
<tr>
"
var nunApplyingArray = new Array(6).fill(0);
for ( var num in [0,1,2,3,4,5]){
var ID = "NumApplyingT" + String(num +1) + "D2";
html += "<td>" + obj[ID] + "</td>
";
}

html += "<td>" + obj["NumApplying1"] + "</td>";
html += "</tr>
<tr>
"

var rateArray = new Array(6).fill(0);
for( var num in [0,1,2,3,4,5]){
var rate = String(Math.round(numApplyingArray[num] / numServedArray[num])) + "%";
html += "<td>" + rate + "</td>
";
}

html += "</tr>
        </tbody>
        </table>
        </div>
        <div class ='chapter-wrapper'>
        <h2 class ='chapter-head'>来場者属性</h2>
        <table class='table table-hover'>
        <thead>
        <tr>
        <th scope='col'></th>
        <th scope='col'>チャレタ</th>
        <th scope='col'>公文</th>
        <th scope='col'>紙の進研ゼミ</th>
        <th scope='col'>宿題のみ</th>
        <th scope='col'>塾</th>
        </tr>
        </thead>
        <tbody>
        <tr>
        "

  html += "<th scope='row'>受講歴のある来場者</th>
            ";
  html += "<td>" + Obj["NumChareta1"] +"</td>
            ";
  html += "<td>" + Obj["NumKumon1"] +"</td>
            ";
  html += "<td>" + Obj["NumShinken1"] +"</td>
            ";
  html += "<td>" + Obj["NumOnly1"] +"</td>
            ";
  html += "<td>" + Obj["NumCram1"] +"</td>
                  </tr>
                  </tbody>
                  </table>
                  </div>";



  }

  html += "</div>
        </div>
        </body>

        </html>";


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
