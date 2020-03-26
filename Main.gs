function Main(){
//  この6行は触らないでください
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var SheetList = ss.getSheetByName("List");
  var listLastRow = SheetList.getLastRow();
  var listLastCol = SheetList.getLastColumn();
  SheetList.activate();
  var latestArray = createLatest(SheetList, listLastRow, listLastCol);

//  変更があった際には以下をいじってください
//  ""のあるところは“”内のみを変更してください
//  '=' の左側の文字列は絶対にいじらないでください。
//  以下で各valueを取得
//  会場名(質問についているIDをもとに取得してくる）
  var thisPlace = getValueByID(SheetList, latestArray, "PlaceName", listLastRow, listLastCol);

//  リーダー名
  var thisLeaderName = getValueByID(SheetList, latestArray, "LeaderName", listLastRow, listLastCol);
  Logger.log("thisLeaderName:%s",thisLeaderName);

//  開催日
  var thisDate = getValueByID(SheetList, latestArray, "Date", listLastRow, listLastCol);

//  来場者/ここは列を取得
  var numServeds1 = getColByID(SheetList, "NumServed1", listLastCol);
  var numServeds2 = getColByID(SheetList, "NumServed2", listLastCol);


//  入会者/ここは列を取得
  var numApplyings1 = getColByID(SheetList, "NumApplying1", listLastCol);
  var numApplyings2 = getColByID(SheetList, "NumApplying2", listLastCol);

//  教材サービス
  var thisService = getValueByID(SheetList, latestArray, "ServiceShort", listLastRow, listLastCol);

//  営業・良かった点
  var thisSalesGood = getValueByID(SheetList, latestArray, "SalesGood", listLastRow, listLastCol);

//  営業悪かった・不明点
  var thisSalesBad = getValueByID(SheetList, latestArray, "SalesBad", listLastRow, listLastCol);

//  店舗固有情報共有
  var thisPlaceInfo = getValueByID(SheetList, latestArray, "PlaceInformation", listLastRow, listLastCol);

//  スゴイきみ
  var thisSugoi = getValueByID(SheetList, latestArray, "Sugoi", listLastRow, listLastCol);

//  気付き
  var thisNotice = getValueByID(SheetList, latestArray, "Notice", listLastRow, listLastCol);

//  メンバー名
  var memberNames1 = getValueByID(SheetList, latestArray, "MemberName1", listLastRow, listLastCol);
  var memberNames2 = getValueByID(SheetList, latestArray, "MemberName2", listLastRow, listLastCol);

//  遅刻早退
  var lateCols1 = getValueByID(SheetList, latestArray, "Attendance1", listLastRow, listLastCol);
  var lateCols2 = getValueByID(SheetList, latestArray, "Attendance2", listLastRow, listLastCol);

//  目標入会率
  var rateGoalCols1 = getValueByID(SheetList, latestArray, "TargetRate1", listLastRow, listLastCol);
  var rateGoalCols2 = getValueByID(SheetList, latestArray, "TargetRate2", listLastRow, listLastCol);

//以上がいじれるところ、以下はGASのコードに触らない方は変更厳禁です

//※※※※※　以下、変更厳禁  ※※※※※※


//  各シートの呼び出しをして変数に格納
  var SheetResult = ss.getSheetByName("RecentResult");
  var SheetService = ss.getSheetByName("Service");
  var SheetSalesGood = ss.getSheetByName("SalesGood");
  var SheetSalesBad = ss.getSheetByName("SalesBad");
  var SheetPlace = ss.getSheetByName("PlaceSpecific");
  var SheetIndivi = ss.getSheetByName("Individual");

//  以下で直近の回答を変数に格納

//  二日目の勤務有無をBoolean型で代入しておく
    var secondDay = isSecondDay(SheetList);

//  一日目の来場者を合計して格納
  var thisNumServed1 = sumNums(SheetList,numServeds1, listLastRow);
//  Logger.log("thisNumServed1:" + thisNumServed1);

//  二日目の来場者
  var thisNumServed2 = sumNums(SheetList,numServeds2, listLastRow);

//  Logger.log("thisNumServed2:" + thisNumServed2);

//  各日の来場者を合計
  var thisNumServedSum = thisNumServed1 + thisNumServed2;

//  一日目の入会者を合計して格納
  var thisNumApplying1 = sumNums(SheetList,numApplyings1, listLastRow);

//  Logger.log("thisNumApplying1:" + thisNumApplying1);

//  二日目の入会者
  var thisNumApplying2 = sumNums(SheetList,numApplyings2, listLastRow);

//  Logger.log("thisNumApplying2:" + thisNumApplying2);

//  各日の入会者を合計して格納
  var thisNumApplyingSum = thisNumApplying1 + thisNumApplying2;
//  Logger.log("thisNumApplyingSum:" + thisNumApplyingSum)

//  入会者・来場者から入会率を導出
  var thisRateSum = String(Math.round(( thisNumApplyingSum / thisNumServedSum ) * 100)) + "%";
  var thisRate1 = String(Math.round(( thisNumApplying1 / thisNumServed1 ) * 100)) + "%";
  if( secondDay ){
    var thisRate2 = String(Math.round(( thisNumApplying2 / thisNumServed2 ) * 100)) + "%";
  }else{
    var thisRate2 = "0%";
  }

  Logger.log("thisRateSum:" + thisRateSum + "%\n" + "thisRate1:" + thisRate1 + "%\n" + "thisRate2:" + thisRate2 + "%");

//  RecentResultシートの更新を行う
    SheetResult.activate();
//  まずは店舗の列でgoDownRow()を実行
  var placeNameCol = getColByIDwoCol(SheetResult,"店舗");
  Logger.log("placeNameCol:%s",placeNameCol);

//  当該店舗の行以外を下にずらす
  goDownRowKey(SheetResult, thisPlace, placeNameCol);

// 一つ一つValueをセットしていく/ほんとは配列に格納してからセットの方がいいけど、保守性のために・・・。
  getNSet(SheetResult, "日付", thisDate);
  getNSet(SheetResult, "店舗", thisPlace);
  getNSet(SheetResult, "リーダー", thisLeaderName);
  getNSet(SheetResult, "来場", thisNumServedSum);
  getNSet(SheetResult, "入会者数", thisNumApplyingSum);
  getNSet(SheetResult, "入会率", thisRateSum);

//  Serviceシートについて
  SheetService.activate();
//  全部の行を下におろすのとナンバリングとを行う
  goDownNNum(SheetService, "番号");

//  各Valueをセットしていく
  getNSet(SheetService, "日付", thisDate);
  getNSet(SheetService, "店舗", thisPlace);
  getNSet(SheetService, "報告", thisService);

//  SalesGoodシートについて
  SheetSalesGood.activate();
//  全部の行を下におろすのとナンバリングとを行う
  goDownNNum(SheetSalesGood, "番号");

//  各Valueをセットしていく
  getNSet(SheetSalesGood, "日付", thisDate);
  getNSet(SheetSalesGood, "店舗", thisPlace);
  getNSet(SheetSalesGood, "報告", thisSalesGood);

//  SalesBadシートについて
  SheetSalesBad.activate();
//  全部の行を下におろしてナンバリング
  goDownNNum(SheetSalesBad, "番号");

//  各Valueをセットしていく
  getNSet(SheetSalesBad, "日付", thisDate);
  getNSet(SheetSalesBad, "店舗", thisPlace);
  getNSet(SheetSalesBad, "報告", thisSalesBad);

//  PlaceSpecificシートについて
  SheetPlace.activate();
//  上と同じなので以下略
  goDownNNum(SheetPlace, "番号");
  getNSet(SheetPlace, "日付", thisDate);
  getNSet(SheetPlace, "店舗", thisPlace);
  getNSet(SheetPlace, "報告", thisPlaceInfo);

//  Individualシートのためにクラスを作成する

}
