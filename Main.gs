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
  var latestArray = createLatest(SheetList)

  // 以下で各valueを取得
  // 会場名(質問についているIDをもとに取得してくる）
  var latestPlace = getValueByID(latestArray, "PlaceName");

  // リーダー名
  var latestLeaderName = getValueByID(latestArray, "LeaderName");
  // Logger.log("latestLeaderName:%s",latestLeaderName);

  // 開催日
  var latestDate = getValueByID(latestArray, "Date");


  // 教材サービス
  var latestService = getValueByID(latestArray, "ServiceShort");

  // 営業・良かった点
  var latestSalesGood = getValueByID(latestArray, "SalesGood");

  // 営業悪かった・不明点
  var latestSalesBad = getValueByID(latestArray, "SalesBad");

  // 店舗固有情報共有
  var latestPlaceInfo = getValueByID(latestArray, "PlaceInformation");

  // スゴイきみ
  var latestSugoi = getValueByID(latestArray, "Sugoi");

  // 気付き
  var latestNotice = getValueByID(latestArray, "Notice");

  // メンバー名
  var memberNames1 = getValueByID(latestArray, "MemberName1");
  var memberNames2 = getValueByID(latestArray, "MemberName2");

  // 遅刻早退
  var lateCols1 = getValueByID( latestArray, "Attendance1");
  var lateCols2 = getValueByID( latestArray, "Attendance2");

  // 目標入会率
  var rateGoalCols1 = getValueByID( latestArray, "TargetRate1");
  var rateGoalCols2 = getValueByID( latestArray, "TargetRate2");

  // 以下で直近の回答を変数に格納
  // 二日目の勤務有無をBoolean型で代入しておく
  var secondDay = isSecondDay(latestArray);

  // 一日目の来場者を合計して格納
  var latestNumServed1 = sumNums(latestArray,"NumServed1");
  // Logger.log("latestNumServed1:" + latestNumServed1);

  // 二日目の来場者
  var latestNumServed2 = sumNums(latestArray,"NumServed2");

  // Logger.log("latestNumServed2:" + latestNumServed2);

  // 各日の来場者を合計
  var latestNumServedSum = latestNumServed1 + latestNumServed2;

  // 一日目の入会者を合計して格納
  var latestNumApplying1 = sumNums(latestArray,"NumApplying1");

  // Logger.log("latestNumApplying1:" + latestNumApplying1);

  // 二日目の入会者
  var latestNumApplying2 = sumNums(latestArray,"NumApplying2");

  // Logger.log("latestNumApplying2:" + latestNumApplying2);

  // 各日の入会者を合計して格納
  var latestNumApplyingSum = latestNumApplying1 + latestNumApplying2;
  // Logger.log("latestNumApplyingSum:" + latestNumApplyingSum)

  // 入会者・来場者から入会率を導出
  var latestRateSum = String(Math.round(( latestNumApplyingSum / latestNumServedSum ) * 100)) + "%";
  var latestRate1 = String(Math.round(( latestNumApplying1 / latestNumServed1 ) * 100)) + "%";
  if( secondDay ){
    var latestRate2 = String(Math.round(( latestNumApplying2 / latestNumServed2 ) * 100)) + "%";
  }else{
    var latestRate2 = "0%";
  }
  // Logger.log("latestRateSum:" + latestRateSum + "%\n" + "latestRate1:" + latestRate1 + "%\n" + "latestRate2:" + latestRate2 + "%");

  // メンバー関係はmembersという二次元配列に格納する
  var members = createMembers(latestArray, "MemberName1", "MemberName2");




// シートの更新系ははじめに二次配列に格納してからの方が処理が早いのでそれで対応したい
// まずは完成優先でのちのちの課題か？
// イメージとしては格納用の配列を用意して毎回リターンする
// 最終的にはその配列を二次元配列化してsetValues()で一発格納
// GASのAPI使わない分こっちのが早いと思われる

  // RecentResultシートの更新を行う
  // まずは店舗の列でgoDownRow()を実行
  var placeNameCol = getColByIDwoCol(SheetResult,"店舗");
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
  var members = createMembers(latestArray,"MemberName1", "MemberName2");
  members = adjustMembers(members, "NumServed1","NumServed2","NumApplying1","NumApplying2");
  var keyObj = {"メンバー名":"MemberName",
                  "対応人数":"NumServedSum",
                  "目標入会率":"TargetRate",
                  "一日目入会率":"ApplyingRate1",
                  "二日目入会率":"ApplyingRate2",
                  "合計入会率":"RateSum",
                  "日付": latestDate
                };
  setMembers(SheetIndivi, members, keyObj,"メンバー名","MemberName1","MemberName2","日付");


}
