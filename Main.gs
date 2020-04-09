function Main(){

  // アクティブなスプレッドシートを取得
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 各シートの呼び出しをして変数に格納

  var SheetList = ss.getSheetByName("List");
  var SheetResult = ss.getSheetByName("RecentResult");
  var SheetService = ss.getSheetByName("Service");
  var SheetSalesGood = ss.getSheetByName("SalesGood");
  var SheetSalesBad = ss.getSheetByName("SalesBad");
  var SheetPlace = ss.getSheetByName("PlaceSpecific");
  var SheetIndivi = ss.getSheetByName("Individual");

  // Listのシートだけよく使うので取得する
  var latestArray = createLatest(SheetList)

  // 扱いやすいようにobj化する
    var listObj = makeObj(array);

  // 以下で各valueを取得
  // 会場名(質問についているIDをもとに取得してくる）
  var latestPlace = getValueByID(listObj, placeName);
  obj[placeName] = latestPlace;
  // リーダー名
  var latestLeaderName = getValueByID(listObj, leaderName);
  obj[leaderName] = latestLeaderName;

  // Logger.log("latestLeaderName:%s",latestLeaderName);
  // 開催日
  var latestDate = getValueByID(listObj, date);
  obj[date] = latestDate;

  // 教材サービス
  var latestService = getValueByID(listObj, serviceShort);
  obj[serviceShort] = latestService;

  // 営業・良かった点
  var latestSalesGood = getValueByID(listObj, salesGood);
  obj[salesGood] = latestSalesGood;

  // 営業悪かった・不明点
  var latestSalesBad = getValueByID(listObj, salesBad);
  obj[salesBad] = latestSalesBad;

  // 店舗固有情報共有
  var latestPlaceInfo = getValueByID(listObj, placeInformation);
  obj[placeInformation] = latestPlaceInfo;

  // スゴイきみ
  var latestSugoi = getValueByID(listObj, sugoi);
  obj[sugoi] = latestSugoi;

  // 気付き
  var latestNotice = getValueByID(listObj, notice);
  obj[notice] = latestNotice;

  // 以下で直近の回答を変数に格納
  // 二日目の勤務有無をBoolean型で代入しておく
  var secondDay = isSecondDay(latestArray);
  obj["secondDay"] = secondDay;

  // 一日目の来場者を合計して格納
  var latestNumServed1 = sumNums(latestArray,numServed1);
  obj[numServed1] = latestNumServed1;

  // 二日目の来場者
  var latestNumServed2 = sumNums(latestArray,numServed2);
  obj[numServed2] = latestNumServed2;

  // 各日の来場者を合計
  var latestNumServedSum = latestNumServed1 + latestNumServed2;
  obj[numServedSum] = latestNumServedSum;

  // 一日目の入会者を合計して格納
  var latestNumApplying1 = sumNums(latestArray,numApplying1);
  obj[numApplying1] = latestNumApplying1;

  // Logger.log("latestNumApplying1:" + latestNumApplying1);

  // 二日目の入会者
  var latestNumApplying2 = sumNums(latestArray,numApplying2);
  obj[numApplying2] = latestNumApplying2;

  // Logger.log("latestNumApplying2:" + latestNumApplying2);

  // 各日の入会者を合計して格納
  var latestNumApplyingSum = latestNumApplying1 + latestNumApplying2;
  obj[numApplyingSum] = latestNumApplyingSum;

  // 入会者・来場者から入会率を導出
  var latestRateSum = String(Math.round(( latestNumApplyingSum / latestNumServedSum ) * 100)) + "%";
  obj["RateSum"] = latestRateSum;
  var latestRate1 = String(Math.round(( latestNumApplying1 / latestNumServed1 ) * 100)) + "%";
  obj["Rate1"] = latestRate1;
  if( secondDay ){
    var latestRate2 = String(Math.round(( latestNumApplying2 / latestNumServed2 ) * 100)) + "%";
  }else{
    var latestRate2 = "0%";
  }
  obj["Rate2"] = latestRate2;
  // Logger.log("latestRateSum:" + latestRateSum + "%\n" + "latestRate1:" + latestRate1 + "%\n" + "latestRate2:" + latestRate2 + "%");

  // IndividualシートのsetMembers()のためにObjをつくっておく
    var keyObj = {"メンバー名":"MemberName",
                    "対応人数":"NumServedSum",
                    "一日目目標入会率":"TargetRate1",
                    "二日目目標入会率":"TargetRate2",
                    "一日目入会率":"ApplyingRate1",
                    "二日目入会率":"ApplyingRate2",
                    "合計入会率":"RateSum",
                    "日付": latestDate
                  };


  // メンバー関係はmembersという二次元配列に格納する
  var members = createMembers(latestArray, memberName1, memberName2);


// シートの更新系ははじめに二次配列に格納してからの方が処理が早いのでそれで対応したい
// まずは完成優先でのちのちの課題か？
// イメージとしては格納用の配列を用意して毎回リターンする
// 最終的にはその配列を二次元配列化してsetValues()で一発格納
// GASのAPI使わない分こっちのが早いと思われる

  // RecentResultシートの更新を行う
  // まずは店舗の列でgoDownRow()を実行
  var placeNameCol = getColByID(SheetResult,"店舗");
  // Logger.log("placeNameCol:%s",placeNameCol);

  // 当該店舗の行以外を下にずらす
  goDownRowKey(SheetResult, latestPlace, placeNameCol);

  // 一つ一つValueをセットしていく
  getNSet(SheetResult, "日付", latestDate);
  getNSet(SheetResult, "店舗", latestPlace);
  getNSet(SheetResult, "リーダー", latestLeaderName);
  getNSet(SheetResult, "来場", latestNumServedSum);
  getNSet(SheetResult, "入会者数", latestNumApplyingSum);
  getNSet(SheetResult, "入会率", latestRateSum);

  // Serviceシートについて
  // 全部の行を下におろすのとナンバリングとを行う
  goDownNNum(SheetService, "番号");

  // 各Valueをセットしていく
  getNSet(SheetService, "日付", latestDate);
  getNSet(SheetService, "店舗", latestPlace);
  getNSet(SheetService, "報告", latestService);

  // SalesGoodシートについて
  // 全部の行を下におろすのとナンバリングとを行う
  goDownNNum(SheetSalesGood, "番号");

  // 各Valueをセットしていく
  getNSet(SheetSalesGood, "日付", latestDate);
  getNSet(SheetSalesGood, "店舗", latestPlace);
  getNSet(SheetSalesGood, "報告", latestSalesGood);

  // SalesBadシートについて
  // 全部の行を下におろしてナンバリング
  goDownNNum(SheetSalesBad, "番号");

  // 各Valueをセットしていく
  getNSet(SheetSalesBad, "日付", latestDate);
  getNSet(SheetSalesBad, "店舗", latestPlace);
  getNSet(SheetSalesBad, "報告", latestSalesBad);

  // PlaceSpecificシートについて
  // 上と同じなので以下略
  goDownNNum(SheetPlace, "番号");
  getNSet(SheetPlace, "日付", latestDate);
  getNSet(SheetPlace, "店舗", latestPlace);
  getNSet(SheetPlace, "報告", latestPlaceInfo);

  // Individualシート
  var members = createMembers(latestArray,memberName1, memberName2);
  members = adjustMembers(members, numServed1,numServed2,numApplying1,numApplying2);
  members =  setMembers(SheetIndivi, members, keyObj,"メンバー名",memberName1, memberName2,"日付");

  sendEmail(obj, members);

}
