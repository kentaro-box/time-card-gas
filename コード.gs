function doGet() {

  var output = HtmlService.createTemplateFromFile('index').evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  return output;
}

//------------- ユーザーの判定 ---------------//
function checkUser(data) {
  data = JSON.parse(data);

  var spreadsheet = SpreadsheetApp.openById('1b3KeY91XbHCdp9TPbeZYYtfFnS3DDxJKc6R-sZAhgPs');
  var sheet = spreadsheet.getSheetByName('Info');
  var lastRow = sheet.getDataRange().getLastRow(); //対象となるシートの最終行を取得

  var names = [];

  // nama一覧を取得
  for (var i = 1; i <= lastRow; i++) {
    names.push(sheet.getRange(i, 1).getValue());
  }

  // 入力された名前と一致するか
  if (names.indexOf(data[0]) == -1) {
    data = "名前の登録がありません";
    return data;
  }

  // 一致する名前のセルのRowを取得
  var checkRow = names.indexOf(data[0]) + 1;
  // 取得したRowのパスワードを取得
  var pass = sheet.getRange(checkRow, 2).getValue();
  // 同じく時給を取得
  var hourly_wage = sheet.getRange(checkRow, 3).getValue();
  data.push(hourly_wage);

  // 入力したパスワードとスプレッドシートのパスワードの照合
  if (data[1] == pass) {
    // 合致
    return JSON.stringify(data);
  } else {
    // 不一致
    return data = "パスワードが一致しません";
  }
}


//--------------- 出勤時間のセット ------------------//

function timeInRecord(userData) {

  var spreadsheet = getSpredSheet(userData);
  userData = JSON.parse(userData);

  var [now, year, month, day, weekName, hour, min, sec, s, today] = getToday();

  // 月替りに新しいシート作成
  if (spreadsheet.getSheetByName(year + "/" + month + "_" + userData[0]) != null) {
    var sheet = spreadsheet.getSheetByName(year + "/" + month + "_" + userData[0]);
    var lastRow = sheet.getDataRange().getLastRow();
  } else {
    spreadsheet.insertSheet(year + "/" + month + "_" + userData[0]);
    sheet = spreadsheet.getSheetByName(year + "/" + month + "_" + userData[0]);

    sheet.getRange(1, 1).setValue("日付");
    sheet.getRange(1, 2).setValue("曜日");
    sheet.getRange(1, 3).setValue("出社時間");
    sheet.getRange(1, 4).setValue("退社時間");
    sheet.getRange(1, 5).setValue("休憩開始時間");
    sheet.getRange(1, 6).setValue("休憩終了時間");
    sheet.getRange(1, 7).setValue("日中勤務時間");
    sheet.getRange(1, 8).setValue("深夜帯勤務時間");
    sheet.getRange(1, 8).setValue("深夜日付跨ぐ");
    sheet.getRange(1, 10).setValue("UserAgent");
  }


  // 祝日の取得
  const id = 'ja.japanese#holiday@group.v.calendar.google.com';
  const cal = CalendarApp.getCalendarById(id);
  const events = cal.getEventsForDay(new Date(year, month, day));
//  const events = cal.getEventsForDay(new Date('2020/07/23'));

  var foriday;
  //祝日がある日
  if (events.length > 0) {
    foriday = true;
  } else {
    foriday = false;
  }

  // 既に出社しているか判定
  if (sheet.getRange(lastRow, 1).getValue() != "日付") {
    var date = sheet.getRange(lastRow, 1).getValue();
    if (date.getTime() == today.getTime()) {
      var alert = "既に出社しています";
      return alert;
    }
  }

  // 祝日・日曜・土曜ならマーカー
  if (foriday || weekName == "日" || weekName == "土") {
    writeData(sheet);
    var range = sheet.getRange(lastRow + 1, 1, 1, 2);
    range.setBackground("#fbd3d0");
  } else {
    writeData(sheet);
  }

  // スプレッドシートにデータ書き込み
  function writeData(sheet) {
    sheet.getRange(lastRow + 1, 1).setNumberFormat("yyyy/mm/dd");
    sheet.getRange(lastRow + 1, 1).setValue(today);
    sheet.getRange(lastRow + 1, 2).setValue(weekName);
    sheet.getRange(lastRow + 1, 3).setNumberFormat("H:mm");
    sheet.getRange(lastRow + 1, 3).setValue(now);
  }

  return "出社登録完了";
}


// --------------- 休憩開始時間 ------------------- //

function startRestTime(userData) {

  var spreadsheet = getSpredSheet(userData);
  userData = JSON.parse(userData);

  var [now, year, month, day, weekName, hour, min, sec, s, today] = getToday();
  var sheet = spreadsheet.getSheetByName(year + "/" + month + "_" + userData[0]);
  var lastRow = sheet.getLastRow(); //対象となるシートの最終行を取得

  var check = sheet.getRange(lastRow, 5).getValue();

  if (check == "休憩開始時間") {
    return "まずは出社してください";
  } else if (check != "") {
    return "休憩開始時間は登録済みです";
  }

  sheet.getRange(lastRow, 5).setNumberFormat('H:mm')
  sheet.getRange(lastRow, 5).setValue(now);
  return "休憩開始時間を登録しました";
}

function endRestTime(userData) {
  var spreadsheet = getSpredSheet(userData);
  userData = JSON.parse(userData);

  var [now, year, month, day, weekName, hour, min, sec, s, today] = getToday();
  var sheet = spreadsheet.getSheetByName(year + "/" + month + "_" + userData[0]);
  var lastRow = sheet.getLastRow(); //対象となるシートの最終行を取得

  var check = sheet.getRange(lastRow, 6).getValue();

  if (check == "休憩終了時間") {
    return "まずは出社してください";
  } else if (check != "") {
    return "休憩終了時間は登録済みです";
  }

  sheet.getRange(lastRow, 6).setNumberFormat('H:mm')
  sheet.getRange(lastRow, 6).setValue(now);
  return "休憩終了時間を登録しました";
}

// --------------- 退勤時間 ------------------- //

// * morning_legal_time: 働き出した日の早朝帯nの区切り時間
// * midnight_legal_time: 働き出した日の深夜帯の区切り時間
// * 

function timeOutRecord(userData) {
  var spreadsheet = getSpredSheet(userData);
  userData = JSON.parse(userData);

  var [now, year, month, day, weekName, hour, min, sec, s, today] = getToday();
  // 法定時間外判別用

  var sheet = spreadsheet.getSheetByName(year + "/" + month + "_" + userData[0]);
  var lastRow = sheet.getLastRow(); //対象となるシートの最終行を取得
  
  var work_in_date = sheet.getRange(lastRow, 1).getValue();
  work_in_date = new Date(work_in_date);
  work_in_date = work_in_date.toLocaleDateString()
  work_in_date = work_in_date.split('/');
  
  var work_date_year = Number(work_in_date[0]);
  var work_date_month = Number(work_in_date[1]);
  var work_date_day = Number(work_in_date[2]);

  var morning_legal_time = new Date(work_date_year, work_date_month, work_date_day, 05, 00, 000);
  var midnight_legal_time = new Date(work_date_year, work_date_month, work_date_day, 22, 00, 000);
  var next_day = new Date(work_date_year, work_date_month, work_date_day + 1, 00, 00, 000);

  var check = sheet.getRange(lastRow, 4).getValue();

  if (check == "退社時間") {
    return "まずは出社してください";
  } else if (check != "") {
    return "すでに退社済み,または出社しておりません";
  } else if (sheet.getRange(lastRow, 1).getValue() == "") {
    return "まずは出社してください";
  }


  // 退社時間登録
  sheet.getRange(lastRow, 4).setNumberFormat('H:mm')
  sheet.getRange(lastRow, 4).setValue(now);
  Utilities.sleep(500);


  // 出勤時間
  var start_time = sheet.getRange(lastRow, 3).getValue();
  // 退社時間
  var end_time = sheet.getRange(lastRow, 4).getValue();
  // 各時間をミリ秒に
  var get_time_start = start_time.getTime();
  var get_time_end = end_time.getTime();

  // 休憩していれば労働時間から休憩時間を引く
  if (sheet.getRange(lastRow, 5).getValue() != "") {

    var restStartTime = sheet.getRange(lastRow, 5).getValue();
    var restEndTime = sheet.getRange(lastRow, 6).getValue();

    // 法定時間外の休憩時間
    // 早朝時間帯の休憩
    if (morning_legal_time.getTime() > restStartTime.getTime()) {
      var morning_out_legal_time = morning_legal_time.getTime() - restStartTime.getTime();
      // それ以外の休憩
      var morning_in_legal_time = restStartTime.getTime() - morning_out_legal_time;
    }
    console.log(restEndTime.getTime() > midnight_legal_time.getTime());
    console.log(restEndTime.getTime());
    console.log(midnight_legal_time.getTime());
    console.log(restEndTime);
    console.log(midnight_legal_time);
    // 深夜時間帯の休憩
    if (restEndTime.getTime() > midnight_legal_time.getTime()) {
      var midnight_out_legal_time = restEndTime.getTime() - midnight_legal_time.getTime();
      // それ以外の休憩
      var midnight_in_legal_time = restEndTime.getTime() - midnight_out_legal_time;
      // 深夜日付跨ぐ休憩
      if (restEndTime.getTime() > next_day.getTime()) {
        var next_day_out_legal_rest_time = restEndTime.getTime() - next_day.getTime();
        midnight_out_legal_time = midnight_out_legal_time - next_day_out_legal_rest_time;
      }
    }
    // 休憩時間全部
    var rest_time = restEndTime.getTime() - restStartTime.getTime();

    rest_time = computeDuration(rest_time);
    sheet.getRange(lastRow, 7).setNumberFormat('H:mm')
    sheet.getRange(lastRow, 7).setValue(rest_time);


  }

  // 休憩していなければ
  // 法定時間内か法定時間外か、法定時間外であれば日付跨いでないか
  if (get_time_start < morning_legal_time.getTime() && midnight_legal_time.getTime() < get_time_end) {
    var work_time = get_time_end - get_time_start;
    var work_out_leagl_time = (morning_legal_time.getTime() - get_time_start) + (get_time_end - midnight_legal_time.getTime());
//    work_time = work_time - work_out_leagl_time;


    if (get_time_end > next_day.getTime()) {
      var next_day_out_leagl_time = get_time_end - next_day.getTime();
      work_out_leagl_time = work_out_leagl_time - next_day_out_leagl_time;
//      work_time - work_out_leagl_time;
    }

  } else if (get_time_start < morning_legal_time.getTime() && midnight_legal_time.getTime() < get_time_end) {
    var work_time = get_time_end - get_time_start;
    var work_out_leagl_time = (morning_legal_time.getTime() - get_time_start) - (get_time_end - midnight_legal_time.getTime());
    work_time -work_out_leagl_time;
  }  else if (get_time_start < morning_legal_time.getTime()) {
    var work_time = get_time_end - get_time_start;
    var work_out_leagl_time = (morning_legal_time.getTime() - get_time_start);
//    work_time - work_out_leagl_time;
  } else if (midnight_legal_time.getTime() < get_time_end) {
    var work_time = get_time_end - get_time_start;
    work_out_leagl_time = work_out_leagl_time - (get_time_end - midnight_legal_time.getTime());
//    work_time - work_out_leagl_time;

    if (get_time_end > next_day.getTime()) {
      var next_day_out_leagl_time = get_time_end - next_day.getTime();
      work_out_leagl_time = work_out_leagl_time - next_day_out_leagl_time;
//      work_time - work_out_leagl_time;
    }
  }

  console.log(typeof midnight_out_legal_time);
  // 法定外で休憩
  if (typeof morning_out_legal_time != "undefined") {
    work_out_leagl_time = work_out_leagl_time - morning_out_legal_time;
//    work_time - morning_out_legal_time;
  }

  if (typeof midnight_out_legal_time != "undefined") {
    work_out_leagl_time = work_out_leagl_time - midnight_out_legal_time;
    console.log("aaaa");
//    work_time - midnight_out_legal_time;
  }

  if (typeof next_day_out_legal_rest_time != "undefined") {
    work_out_leagl_time = next_day_out_leagl_time - next_day_out_legal_rest_time;
//    work_time - next_day_out_legal_rest_time;
  }

  
 
  if (typeof work_out_leagl_time == 'undefined') {
    sheet.getRange(lastRow, 9).setNumberFormat('H:mm')
    sheet.getRange(lastRow, 9).setValue(0);
  } else {
    sheet.getRange(lastRow, 9).setNumberFormat('H:mm')
    work_time - work_out_leagl_time;
    work_out_leagl_time = computeDuration(work_out_leagl_time);
    sheet.getRange(lastRow, 9).setValue(work_out_leagl_time);

    if (typeof next_day_out_leagl_time == 'undefined') {

      sheet.getRange(lastRow, 10).setNumberFormat('H:mm')
      sheet.getRange(lastRow, 10).setValue(0);
    } else {


      // 祝日の取得
      const id = 'ja.japanese#holiday@group.v.calendar.google.com';
      const cal = CalendarApp.getCalendarById(id);
      const events = cal.getEventsForDay(new Date(year, month, day));
//      const events = cal.getEventsForDay(new Date('2020/07/23'));

      var foriday;
      //祝日がある日
      if (events.length > 0) {
        foriday = true;
      } else {
        foriday = false;
      }
      
      var weekName = new Date(year, month, day).getDay();

      // 祝日・日曜・土曜ならマーカー
      if (foriday == false || weekName != 0 || weekName != 6) {
        sheet.getRange(lastRow, 10).setBackground("#fffff");
      } else {
        sheet.getRange(lastRow, 10).setBackground("#e8f0fe");
      }
      sheet.getRange(lastRow, 10).setNumberFormat('H:mm')
      next_day_out_leagl_time = computeDuration(next_day_out_leagl_time);
      sheet.getRange(lastRow, 10).setValue(next_day_out_leagl_time);
    } 
    
     
     work_time = computeDuration(work_time);
     sheet.getRange(lastRow, 8).setNumberFormat('H:mm')
     sheet.getRange(lastRow, 8).setValue(work_time);

  }



  // ミリ秒変換
  function computeDuration(ms) {
    var h = String(Math.floor(ms / 3600000) + 100).substring(1);
    var m = String(Math.floor((ms - h * 3600000) / 60000) + 100).substring(1);
    return h + ':' + m;
  }


  // 時間の計算

  var has_color = [];
  var no_color = [];
  var now_lastRow = sheet.getLastRow();

  for (var i = 2; i <= now_lastRow; i++) {
    var check_background = sheet.getRange(i, 1).getBackground();

    if (check_background == "#ffffff") {
      no_color.push(i);
    } else {
      has_color.push(i);
    }

  }
  // 通常時間集計
  if (no_color.length > 0) {
    sheet.getRange(1, 13).setValue("通常時間合計");
    var sum;
    for (var j = 0; j < no_color.length; j++) {
      if (j == 0) {
        if (no_color.length == 1) {
          sum = "=SUM(H" + no_color[j] + ")";
        }
        sum = "=SUM(H" + no_color[j];
      } else if (j == no_color.length - 1) {
        sum = sum + ",H" + no_color[j] + ")";

      } else {
        sum = sum + ",H" + no_color[j];
      }

    }
    sheet.getRange(2, 13).setNumberFormat('[h]:mm')
    sheet.getRange(2, 13).setFormula(sum);
  }
  
  // 通常時間割増
  if (has_color.length > 0) {
    sheet.getRange(1, 14).setValue("休日割増時間合計");
    var has_color_sum;
    for (var h = 0; h < has_color.length; h++) {
      if (h == 0) {
        if (has_color.length == 1) {
          has_color_sum = "=SUM(H" + has_color[h] + ")";
        }
        has_color_sum = "=SUM(H" + has_color[h];
      } else if (h == has_color.length - 1) {
        has_color_sum = has_color_sum + ",H" + has_color[h] + ")";

      } else {
        has_color_sum = has_color_sum + ",H" + has_color[h];
      }

    }
    sheet.getRange(2, 14).setNumberFormat('[h]:mm')
    sheet.getRange(2, 14).setFormula(has_color_sum);
  }


  if (typeof morning_out_legal_time != "undefined" && midnight_out_legal_time != "undefined") {
    return "長時間お疲れ様でした！\nお気をつけておかえりくださいね！";
  }


  // 深夜勤務時間であれば深夜時間切り分け
  return "お疲れ様でした！\nお気をつけておかえりくださいね！";
}









// ------------ スプレッドシートの取得 ------------------ //

function getSpredSheet(data) {

  data = JSON.parse(data);
  switch (data[0]) {
    case "山田":
      var spreadsheet = SpreadsheetApp.openById('1loV6Ob1WLxTl5Yn8qJeagBSaNG7B5u4B8xcYdtqCLlw');
      break;
    case "市川":
      var spreadsheet = SpreadsheetApp.openById('1Z7o6PUaFyIpL95MVQEQCMLywGfcTfUJr5f8UNvLWBwQ');
      break;
  }
  return spreadsheet;
}

// ------------- 日付の取得 ----------------------- //

function getToday() {
  var now = new Date();
  var year = now.getFullYear();
  var month = now.getMonth() + 1; //１を足すこと
  var day = now.getDate();
  var week = now.getDay(); //曜日(0～6=日～土)
  var weeks = ["日", "月", "火", "水", "木", "金", "土"];
  var weekName = weeks[week];
  var hour = now.getHours();
  var min = now.getMinutes();
  var sec = now.getSeconds();
  now = new Date(year, month - 1, day, hour, min, 000);
  var today = new Date(year, month - 1, day, 00, 00, 000);

  //曜日の選択肢
  var youbi = new Array("日", "月", "火", "水", "木", "金", "土");
  //出力用
  var s = year + "/" + month + "/" + day + "(" + weekName + ")   " + hour + ":" + min;
  return [now, year, month, day, weekName, hour, min, sec, s, today];
}


