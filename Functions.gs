//一番初めに最新の行をすべて一次配列に格納して一次配列で返す関数
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


//すべての行を下に一行ずらす関数

function goDownRow(sheet) {
  //最終行の取得
  var lastRow = sheet.getLastRow();
  //一番初めの時のための条件分岐
  if (lastRow > 0) {
    //行を挿入する
    sheet.insertRowAfter(3);
  }
}


//一致するワードをもとにそれ以外の行を下に持っていく関数

function goDownRowKey(sheet, keyword, col) {
  //範囲を決定して変数宣言
  var modifiedLastRow = sheet.getLastRow() - 1;
  var lastCol = sheet.getLastColumn();
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

function goDownNNum(sheet, keyword) {

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

    //最新の番号を取得
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


//検索したいキーワードをもとにListSheetの上から一行を検索して行番号を格納した配列を返す関数

function getColByID(array, keyword) {

  //検索範囲の決定
  var searchArray = array[0];

  //return用の配列を作成
  var indexes = [];

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
  return indexes
}




function getColByIDwoCol(sheet, keyword) {

  //最終列の取得
  var lastCol = sheet.getLastColumn();

  //検索範囲の決定
  var findRange = sheet.getRange(1, 1, 1, lastCol);

  //textFinderを作成
  var ranges = findRange.createTextFinder(String(keyword)).findAll();

  //戻り値のための配列を作成
  var cols = Array(1);

  //列番号をそれぞれ取得して配列に格納
  for (var range in ranges) {
    cols[range] = ranges[range].getColumn();
  }

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


//入場者来場者を足す関数,引数に　シート、列番号の配列、最後列のintをとる
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

  //一日目と二日目のメンバーをjsonの配列の形で変数に取る
  var members1 = getMembers(array, memberName1);
  var members2 = getMembers(array, memberName2);

  //members1の空白の奴を削除する
  for (var member1 in members1) {
    var obj = members1[member1];
    if (obj[memberName1] == "") {
      members1.splice(member1, 1);
    }
  }

  Logger.log("members1:%s", members1);

  //名前の一致を探すためのfor
  for (var member1 in members1) {
    //二日あるかのためのboolean型の変数宣言
    var twoDays = false;
    //一致を探すためのfor
    for (var member2 in members2) {

      //jsonを変数に代入する
      var member1Obj = members1[member1];
      var member2Obj = members2[member2];


      //名前が一致するかの条件分岐
      if (member1Obj.MemberName1 == member2Obj.MemberName2) {


        //二日あるかどうかをtrueに
        twoDays = true;

        //jsonを代入するためのforを回す
        for (var key in member2Obj) {
          //メンバーの名前だけ条件分岐で回避
          if (key != "MemberName2") {
            //一日目の方に引っ付けてしまう
            member1Obj[key] = member2Obj[key];
          }
          //二日あったよのやつ
          member1Obj["TwoDays"] = twoDays;
        }
        //空白になってるやつを回避する
      } else if (members2[member2].MemberName != "") {
        //members1の後ろにjsonを付け足しとく
        //pushするとspliceした奴ずれるからどうやって対処するか考えないと....
        members1.push(members2[member2]);
      }
    }
  }
  return members1
}


function getMembers(array, keyword) {
  //keywordを使ってarrayのindexを取得
  var cols = getColByID(array, keyword);

  //各メンバーで入力項目はかわらないので項目数を取ってくる
  var difference = cols[1] - cols[0];

  //リターンするための配列を作成
  var members = new Array(1);

  for (var col in cols) {
    //各メンバーにobjを作った方が勝手がよさそうなので...
    var obj = {
      number: col,
    };
    //なんでcols[col+1]を取ってこないかというと、二日目の勤務の確認が入っているから
    var nextCol = cols[col] + difference;

    //cols[col]とnextColの間にある整数を取ってくる
    for (var nn2 = cols[col]; nn2 < nextCol; nn2 += 1) {

      //現在見ている質問を代入する
      var currentItem = array[0][nn2];
      //現在見ている答えを代入する
      var currentValue = array[1][nn2];
      //IDを取得するために　[　か　]　で文字列を分割
      var currentItemSplit = currentItem.split(/\[|\]/);
      //取得したIDをプロパティにする
      var currentProperty = currentItemSplit[1];
      //jsonにID：valueの形で
      obj[currentProperty] = currentValue;
    }
    //リターン用の配列にjsonを代入
    members[col] = obj
    Logger.log(obj);
  }
  Logger.log(members);
  return members
}
