//一番初めに最新の行をすべて一次配列に格納して二次配列で返す関数
function createLatest(sheet) {

  //最終行列を取得しておく
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  //見出し行の二次配列を取得
  var items = sheet.getRange(1, 1, 1, lastCol).getValues();

  //最新の行のvalueの入った二次配列を取得
  var rows = sheet.getRange(lastRow, 1, 1, lastCol).getValues();

  //return用の配列を宣言
  var array = Array(1);

  //代入
  array[0] = items[0]
  array[1] = rows[0]

  return array;
}


//見出し以下のすべての行を下に一行ずらす関数

function goDownRow(sheet) {
  //最終行の取得
  var lastRow = sheet.getLastRow();
  //一番初めの時のための条件分岐
  if (lastRow > 3) {
    //行を挿入する
    sheet.insertRowAfter(3);
  }
}


//一致するワードをもとにそれ以外の行を下に持っていく関数

function goDownRowKey(sheet, keyword, col) {
  //範囲を決定して変数宣言
  var modifiedLastRow = sheet.getLastRow() - 1;
  var lastCol = sheet.getLastColumn();

  // 範囲を決定しておく
  var findRange = sheet.getRange(2, col, modifiedLastRow, lastCol);

  //テキストファインダーを作成
  var tf = findRange.createTextFinder(String(keyword));

  //最終列を取得しておく
  var lastCol = sheet.getLastColumn();

  //検索をかける
  var ranges = tf.findAll();

  //検索に一致するものがあったかどうかで条件分岐
  if (ranges.length >= 1) {

    //行番号を取得してその行を上に持ってくる
    for (var range in ranges) {
      //Rangeがかぶるとバグるので条件分岐して回避
      if (ranges[range].getRow() > 4) {

        //下の行を移動させる
        sheet.moveRows(ranges[range], 4);
      }
    }
  } else {
    goDownRow(sheet);
  }
}

//一致するワードをもとにそれ以外の行を下に持っていく関数のキーワード2つバージョン

function goDownRowKey2(sheet, keyword1, keyword2, col) {
  //範囲を決定して変数宣言
  var modifiedLastRow = sheet.getLastRow() - 1;
  var lastCol = sheet.getLastColumn();

  // 範囲を決定しておく
  Logger.log("2,%s,%s,%s", col, modifiedLastRow, lastCol)
  var findRange = sheet.getRange(2, col, modifiedLastRow, lastCol);

  // textFinderのために正規表現を使った文字列を作っておく
  var regexp = "(.*" + keyword1 + ".*)|(.*" + keyword2 + ".*)";

  //テキストファインダーを作成
  var tf = findRange.createTextFinder(regexp).useRegularExpression(true);

  //最終列を取得しておく
  var lastCol = sheet.getLastColumn();

  //検索をかける
  var ranges = tf.findAll();

  //検索に一致するものがあったかどうかで条件分岐
  if (ranges.length >= 1) {

    //行番号を取得してその行を上に持ってくる
    for (var range in ranges) {
      //Rangeがかぶるとバグるので条件分岐して回避
      if (ranges[range].getRow() > 4) {

        //下の行を移動させる
        sheet.moveRows(ranges[range], 4);
      }
    }
  } else {
    goDownRow(sheet);
  }
}


//引数にどの列に番号が入っているかを決めるキーワードを取り行を一行さげるとともにナンバリングを行う
function goDownNNum(sheet, keyword) {
  // 最後の行列を取得する
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn()

  //すでに番号が存在しているとき
  if (lastRow > 3) {

    //渡されたキーワードでテキストファインダーを作成
    var tf = sheet.getRange(3, 1, 1, lastCol).createTextFinder(keyword);

    //探す
    var ranges = tf.findAll();

    //その行列を取得
    var numRow = ranges[0].getRow();
    var numCol = ranges[0].getColumn();

    //最新のナンバリングを取得
    var latestNum = sheet.getRange(numRow + 1, numCol).getValue();

    //一を足して変数に入れておく
    var newNum = latestNum + 1;

    //全体を一行下げる
    goDownRow(sheet);

    //番号を入れておく
    sheet.getRange(numRow + 1, numCol).setValue(newNum);
  } else {
    //上と同じ処理なので省略
    var tf = sheet.getRange(3, 1, 1, lastCol).createTextFinder(keyword);

    var ranges = tf.findAll();

    var numRow = ranges[0].getRow();
    var numCol = ranges[0].getColumn();

    //    はじめの番号として１を入れておく
    sheet.getRange(numRow + 1, numCol).setValue("1");
  }
}

//検索したいキーワードをもとにListSheetの上から一行を検索してRangeをとり、Valueを適当な位置にセットする関数

function getNSet(sheet, keyword, value) {

  //渡されたシートの最後の行・列番号を取得
  var lastCol = sheet.getLastColumn();

  //検索範囲の決定
  var findRange = sheet.getRange(3, 1, 1, lastCol);

  //textFinderを作成
  var ranges = findRange.createTextFinder(String(keyword)).findAll();


  //要素の数で条件分岐（要素が一つの時配列だと扱いにくいから）
  if (ranges.length == 1) {

    //列番号を取得して変数に代入
    var rangeRow = ranges[0].getRow();
    var rangeCol = ranges[0].getColumn();

    //valueをセットするRangeを取得
    var valueRange = sheet.getRange(rangeRow + 1, rangeCol);

    //セット
    valueRange.setValue(value);

  } else {

    //仮に配列だった時の予防線としてfor文で回しとく
    for (var range in ranges) {

      var rangeRow = ranges[range].getRow();
      var rangeCol = ranges[range].getColumn();

      var valueRange = sheet.getRange(rangeRow, rangeCol);
      valueRange.setValue(value);
    }
  }
}


//検索したいキーワードをもとにarray[0]を検索して行番号を格納した配列を返す関数

function getColByID(array, keyword) {

  //検索範囲の決定
  var searchArray = array[0];

  //return用の配列を作成
  var indexes = new Array(1).fill(0);

  //keywordを正規表現化する
  var regexp = new RegExp(".*" + keyword + ".*");

  //forで回して一つ一つ確認
  for (var element in searchArray) {

    //Logger.log("%s は %s　を含むか: %s\n", searchArray[element], keyword, regexp.test(searchArray[element]));

    //正規表現化したもので一致があるか条件分岐
    if (regexp.test(searchArray[element])) {

      //一致有れば配列の最後に追加
      indexes.push(parseInt(element));
    }
  }
  indexes.shift();
  return indexes
}



// 引数にsheetをとって該当の行番号を返す関数（SHeetList以外用)
function getColByIDwoCol(sheet, keyword) {

  //最終列の取得
  var lastCol = sheet.getLastColumn();

  //検索範囲の決定
  var findRange = sheet.getRange(3, 1, 1, lastCol);

  //textFinderを作成
  var ranges = findRange.createTextFinder(String(keyword)).findAll();

  //戻り値のための配列を作成
  var cols = Array(1);

  //列番号をそれぞれ取得して配列に格納
  for (var range in ranges) {
    cols[range] = ranges[range].getColumn();
  }

  Logger.log("getColByIDwoCol:%s", cols);

  return cols
}


//二日目の勤務の有無についてboolean型の変数を戻す関数

function isSecondDay(array) {

  //入力されたものを代入しておく
  var values = getValueByID(array, "二日目に勤務");

  //リターンする変数を作っておく
  var secondDay = false;

  //for文を回してひとつひとつはいが含まれていないか判定
  for (var value in values) {
    if (values[value].match(/.*はい.*/)) {
      secondDay = true;
    }
  }

  //
  //  ここではい、いいえがぶれていたらそれを判定して
  //  メールを送るという関数を入れるとなおよくなるだろうな
  //
  return secondDay;
}

// 引数に最新の回答の保持された二次元配列と質問のIDをとってIDから回答を返す関数
function getValueByID(array, keyword) {
  //列番号の配列を取得し変数に格納
  var cols = getColByID(array, keyword);

  //Return用の配列を作っておく
  var returnArray = Array(1);

  //二次元配列を一次元に
  var searchArray = array[1]

  //列番号の配列をもとに最下行のValueを取ってきてarrayに格納
  for (var col in cols) {

    //arrayのインデックスと行番号は一ずれるので修正
    var index = cols[col];
    //引数に取った最新のものの格納された配列から当該の者だけをとって返す
    returnArray[col] = searchArray[index];
  }

  //returnArrayの要素が一つなら配列でないほうが扱いやすいので条件分岐して代入する

  if (returnArray.length == 1) {
    return returnArray[0]
  } else {
    return returnArray
  }
}


//入場者来場者を足す関数,引数に 回答を記録したシートからとってきた配列とシートをとる
function sumNums(array, keyword) {
  //int型で変数を宣言
  var sum = parseInt(0);

  //入っている列を取得
  var indexes = getColByID(array, keyword);

  //引数に渡された列の配列から空でないセルの数字を合計していく
  for (var nn1 in indexes) {
    //空白のセルの判定
    if (array[1][indexes[nn1]] != "") {
      //int型にキャストして空白でなかったセルを足す
      sum += parseInt(array[1][indexes[nn1]]);
    }
  }
  return sum
}

//メンバー関係を格納した配列を作って返す関数
//この関数内で、一日目と二日目のメンバーが同じ場合に引っ付ける作業も行う
function createMembers(array, memberName1, memberName2) {

  //一日目と二日目のメンバーをobjectの配列の形で変数に取る
  var members1 = getMembers(array, memberName1);
  Logger.log("members1の出力を終わりました");
  var members2 = getMembers(array, memberName2);

  //members1の空白の奴を削除する
  for (var member1 in members1) {

    // 一つ目のobjectを代入しとく
    var obj = members1[member1];

    // メンバーネームが空白の時その配列を削除しておく
    if (!obj[memberName1]) {

      // 削除
      members1.splice(member1, 1);

    }
  }

  //名前の一致を探すためのfor
  for (var member1 in members1) {
    //二日あるかのためのboolean型の変数宣言
    var twoDays = false;
    //一致を探すためのfor
    for (var member2 in members2) {

      //objを変数に代入する
      var member1Obj = members1[member1];
      var member2Obj = members2[member2];

      //名前が一致するかの条件分岐
      if (member1Obj[memberName1] == member2Obj[memberName2] &&
        member1Obj[memberName1]) {

        //二日あるかどうかをtrueに
        twoDays = true;

        // 一日目と二日目に同じメンバーがいることを判定
        if (members1[member1][memberName1] == members2[member2][memberName2]) {
          // いたら削除しとく
          members2.splice(member2, 1);
        }

        //objを代入するためのforを回す
        for (var key in member2Obj) {
          //メンバーの名前だけ条件分岐で回避
          if (key != memberName2) {
            //一日目の方に引っ付けてしまう
            member1Obj[key] = member2Obj[key];
          }
          //二日あったよのやつ
          member1Obj["TwoDays"] = twoDays;
        }
      }
    }
  }
  // 最後にかぶってなかった二日目の人を代入する
  for (var member2 in members2) {

    // 空白判定
    if (members2[member2].memberName2) {

      // 追加
      members1.push(members2[member2]);
    }
  }
  return members1
}

// 各日のメンバーの各valueをObject型に代入して返す関数
function getMembers(array, keyword) {
  //keywordを使ってarrayのindexを取得
  var cols = getColByID(array, keyword);

  //各メンバーで入力項目はかわらないので項目数を取ってくる
  var difference = cols[1] - cols[0];

  //リターンするための配列を作成
  var members = new Array(1);

  for (var col in cols) {
    //各メンバーにobjを作った方が勝手がよさそうなので...
    var obj = new Object();
    //なんでcols[col+1]を取ってこないかというと、二日目の勤務の確認が入っているから
    var nextCol = cols[col] + difference;

    if (array[1][cols[col]]) {

      //cols[col]とnextColの間にある整数を取ってくる
      for (var nn2 = cols[col]; nn2 < nextCol; nn2 += 1) {
        //現在見ている質問を代入する
        var currentItem = array[0][nn2];
        //現在見ている答えを代入する
        var currentValue = array[1][nn2];
        //IDを取得するために [ か ] で文字列を分割
        var currentItemSplit = currentItem.split(/\[|\]/);
        //取得したIDをプロパティにする
        var currentProperty = currentItemSplit[1];
        //jsonにID：valueの形で
        obj[currentProperty] = currentValue;
      }
      // できた奴を代入する
      members[col] = obj;
    }
  }
  return members
}

// SheetIndiviにValueをsetするためにmembersに必要な値を追加する関数
function adjustMembers(membersArray, numServed1, numServed2, numApplying1, numApplying2) {
  var members = membersArray;
  for (var member in members) {
    // 変数にとっておく
    var thisMember = members[member];

    // それぞれの数字を取得する
    var numberServed1 = avoidBlank(thisMember, numServed1);
    var numberServed2 = avoidBlank(thisMember, numServed2);
    var numberApplying1 = avoidBlank(thisMember, numApplying1);
    var numberApplying2 = avoidBlank(thisMember, numApplying2);

    // 合計した数字を取得しとく
    var numberServedSum = numberServed1 + numberServed2;
    var numberApplyingSum = numberApplying1 + numberApplying2;

    // 一日目と二日目、合計の入会率を求める
    thisMember["NumServedSum"] = numberServedSum;
    thisMember["NumApplyingSum"] = numberApplyingSum;
    thisMember["RateSum"] = String(Math.round((numberApplyingSum / numberServedSum) * 100)) + "%";

    // 0では割れないので0を回避する
    if (numberApplying1 == 0) {
      // 分母にゼロが来るときは０を代入しておく
      thisMember["ApplyingRate1"] = "0%";
    } else {
      // それ以外の時は正当な入会率を求めておく
      thisMember["ApplyingRate1"] = String(Math.round((numberApplying1 / numberServed1) * 100)) + "%";
    }
    if (numberApplying2 == 0) {
      thisMember["ApplyingRate2"] = "0%";
    } else {
      thisMember["ApplyingRate2"] = String(Math.round((numberApplying2 / numberServed2) * 100)) + "%";
    }
    // return用の配列に戻す
    members[member] = thisMember;
  }
  return members
}

// 与えられたものが空白かを判定し、空白ならば0空白でなければ、valueをintにして返す関数
function avoidBlank(thisMember, keyword) {
  // return 用の変数を用意、初期値として0を設定
  var intNum = 0;
  // 空白でなければ
  if (thisMember[keyword]) {
    // intにキャストした値を代入する
    intNum = parseInt(thisMember[keyword]);
  }
  return intNum
}

// メンバーをセットする関数
// 引数には（セットしたいシート、メンバーのObjが格納された配列、
// どの行をもとに行を下げるかのキーワード、一日目のメンバー名のID、二日目のメンバー名のID）をとる
function setMembers(sheet, members, keyObj, keyword, memberName1, memberName2, dateKey) {
  // ここではメンバー名をもとに行をおろしたいので、メンバー名の行を取得する
  var col = getColByIDwoCol(sheet, keyword);

  // 配列の中の各メンバーを取り出す
  for (var member in members) {
    // 取り出したメンバーを代入しておく
    var thisMember = members[member];
    // 取り出したメンバーのメンバー名をもとに行を下げる
    goDownRowKey2(sheet, thisMember[memberName1], thisMember[memberName2], col);

    // メンバーの各Valueをとりだすためのfor
    for (var key in keyObj) {
      // メンバーのObjのkeyと引数のキーワードが一致するかを条件分岐
      if (key == keyword) {

        // その中でも一日目のメンバー名の方に名前があるかを分岐
        if (thisMember[memberName1]) {
          // あればその値をセット
          getNSet(sheet, key, thisMember[memberName1]);
          thisMember["MemberName"] = thisMember[memberName1];

          // 一日目になくて、二日目のみ名前がある場合
        } else if (thisMember[memberName2]) {
          // 二日目のメンバー名をセット
          getNSet(sheet, key, thisMember[memberName2]);
          thisMember["MemberName"] = thisMember[memberName2];
        }

        // 日付だけはメンバーのobjにないので別で代入
      } else if (key == dateKey) {
        // Mainの方で渡しているlatestDateをセット
        getNSet(sheet, dateKey, keyObj[key]);

        //引数のキーワードと一致せず、かつそのValueが空白でないときに
      } else if (thisMember[keyObj[key]]) {
        // メンバーの各Valueを代入していく
        getNSet(sheet, key, thisMember[keyObj[key]]);
      }
    }

    members[member] = thisMember;
  }
  return members
}

function makeObjForEmail(array){
  var obj = new Object(1);
  for (var index in array[0]){
    var question = array[0][index];
    var questionSplit = question.split(/\[|\]/);
    var answer = array[1][index];
    var id = questionSplit[1];
    obj[id] = answer;
  }
  return obj
}
