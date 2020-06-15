function doGet() {

  var output = HtmlService.createTemplateFromFile('index').evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
  return output;
}

//------------- ユーザーの判定 ---------------//
function checkUser(data) {
  data = JSON.parse(data);



  //スクリプトプロパティの値を取得
  var prop = PropertiesService.getScriptProperties();
  var res = prop.getProperty("判定用シートID");

  var spreadsheet = SpreadsheetApp.openById(res);
  var sheet = spreadsheet.getSheetByName('EmployeeInfo');
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

  var prop = PropertiesService.getScriptProperties();
  var res = prop.getProperty("判定用シートID");

  var timesCardSpreadsheet = SpreadsheetApp.openById(res);
  var TimeCardsheet = timesCardSpreadsheet.getSheetByName('CompanyRegulations');
  var ceiling = TimeCardsheet.getRange(2, 1).getValue();

  console.log(ceiling);
  console.log(typeof ceiling);

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
    sheet.getRange(1, 7).setValue("休憩時間合計");
    sheet.getRange(1, 8).setValue("日中勤務時間");
    sheet.getRange(1, 9).setValue("深夜帯勤務時間");
    sheet.getRange(1, 10).setValue("深夜日付跨ぐ");
    sheet.getRange(1, 11).setValue("UserAgent");
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

  // 丸める時間取得
  //  var prop = PropertiesService.getScriptProperties();
  //  var res = prop.getProperty("判定用シートID");
  //
  //  var spreadsheet = SpreadsheetApp.openById(res);
  //  var TimeCardsheet = spreadsheet.getSheetByName('CompanyRegulations');
  //  var ceiling = TimeCardsheet.getRange(2, 1).getValue();
  //  
  //  console.log(ceiling);
  //  console.log(typeof ceiling);
  //  
  //  consoe.log(2222);

  // 日時取得
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

  var morning_legal_time = new Date(work_date_year, work_date_month - 1, work_date_day, 05, 00, 000);
  console.log(morning_legal_time);
  console.log(morning_legal_time.getTime());
  var midnight_legal_time = new Date(work_date_year, work_date_month - 1, work_date_day, 22, 00, 000);
  console.log(midnight_legal_time);
  var next_day = new Date(work_date_year, work_date_month - 1, work_date_day + 1, 00, 00, 000);
  console.log(next_day);

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
  console.log(start_time);
  // 退社時間
  var end_time = sheet.getRange(lastRow, 4).getValue();
  // 各時間をミリ秒に
  var get_time_start = start_time.getTime();
  console.log("get_time_start" + get_time_start);
  var get_time_end = end_time.getTime();

  // 休憩していれば労働時間から休憩時間を引く
  if (sheet.getRange(lastRow, 5).getValue() != "") {

    var restStartTime = sheet.getRange(lastRow, 5).getValue();
    var restEndTime = sheet.getRange(lastRow, 6).getValue();

    // 法定時間外の休憩時間
    // 早朝時間帯の休憩
    if (morning_legal_time.getTime() > restStartTime.getTime()) {
      console.log(1 - 1);
      var morning_out_legal_time = morning_legal_time.getTime() - restStartTime.getTime();
      // それ以外の休憩
      var morning_in_legal_time = restStartTime.getTime() - morning_out_legal_time;
    }

    // 深夜時間帯の休憩
    if (restEndTime.getTime() > midnight_legal_time.getTime()) {
      console.log(1 - 2);
      var midnight_out_legal_time = restEndTime.getTime() - midnight_legal_time.getTime();
      // それ以外の休憩
      var midnight_in_legal_time = restEndTime.getTime() - midnight_out_legal_time;
      // 深夜日付跨ぐ休憩
      if (restEndTime.getTime() > next_day.getTime()) {
        console.log(1 - 3);
        var next_day_out_legal_rest_time = restEndTime.getTime() - next_day.getTime();
        midnight_out_legal_time = midnight_out_legal_time - next_day_out_legal_rest_time;
      }
    }
    // 休憩時間全部
    var rest_time = restEndTime.getTime() - restStartTime.getTime();
    console.log(1 - 4);
    rest_time = computeDuration(rest_time);
    sheet.getRange(lastRow, 7).setNumberFormat('H:mm')
    sheet.getRange(lastRow, 7).setValue(rest_time);
  }

  // 休憩していなければ
  // 法定時間内か法定時間外か、法定時間外であれば日付跨いでないか
  if (get_time_start < morning_legal_time.getTime() && midnight_legal_time.getTime() < get_time_end) {
    console.log(1);
    var work_time = get_time_end - get_time_start;
    var work_out_leagl_time = (morning_legal_time.getTime() - get_time_start) + (get_time_end - midnight_legal_time.getTime());
    work_time = work_time - work_out_leagl_time;


    if (get_time_end > next_day.getTime()) {
      console.log(2);
      var work_time = get_time_end - get_time_start;
      var next_day_out_leagl_time = get_time_end - next_day.getTime();
      work_out_leagl_time = work_out_leagl_time - next_day_out_leagl_time;
      work_time = work_time - work_out_leagl_time;
    }

  } else if (get_time_start < morning_legal_time.getTime()) {
    console.log(4);
    var work_time = get_time_end - get_time_start;
    var work_out_leagl_time = morning_legal_time.getTime() - get_time_start;
    work_time = work_time - work_out_leagl_time;
  } else if (midnight_legal_time.getTime() < get_time_end) {
    console.log(5);
    var work_time = get_time_end - get_time_start;
    var work_out_leagl_time = get_time_end - midnight_legal_time.getTime();
    console.log("work_out_leagl_time" + work_out_leagl_time);
    work_time = work_time - work_out_leagl_time;

    if (get_time_end > next_day.getTime()) {
      console.log(6);
      var next_day_out_leagl_time = get_time_end - next_day.getTime();
      work_out_leagl_time = work_out_leagl_time - next_day_out_leagl_time;
      work_time = work_time - work_out_leagl_time;
    }
  } else {
    var work_time = get_time_end - get_time_start;
    console.log(work_time);
  }

  // 法定外で休憩
  if (typeof morning_out_legal_time != "undefined") {
    work_out_leagl_time = work_out_leagl_time - morning_out_legal_time;
    work_time - morning_out_legal_time;
  }

  if (typeof midnight_out_legal_time != "undefined") {
    work_out_leagl_time = work_out_leagl_time - midnight_out_legal_time;
    console.log("aaaa");
    work_time - midnight_out_legal_time;
  }

  if (typeof next_day_out_legal_rest_time != "undefined") {
    work_out_leagl_time = next_day_out_leagl_time - next_day_out_legal_rest_time;
    work_time - next_day_out_legal_rest_time;
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
      if (next_day == foriday || next_day.getDay() == 0 || next_day.getDay() == 6) {
        sheet.getRange(lastRow, 10).setBackground("#e8f0fe");
      } else {
        sheet.getRange(lastRow, 10).setBackground("#fffff");
      }
      sheet.getRange(lastRow, 10).setNumberFormat('H:mm')

      next_day_out_leagl_time = computeDuration(next_day_out_leagl_time);
      sheet.getRange(lastRow, 10).setValue(next_day_out_leagl_time);
    }
  }

  work_time = computeDuration(work_time);

  sheet.getRange(lastRow, 8).setNumberFormat('H:mm')
  //     sheet.getRange(lastRow, 8).strformula("=CEILING("+ work_time + ",0:15)");
  sheet.getRange(lastRow, 8).setValue(work_time);

  // ミリ秒変換
  function computeDuration(ms) {
    var h = String(Math.floor(ms / 3600000) + 100).substring(1);
    var m = String(Math.floor((ms - h * 3600000) / 60000) + 100).substring(1);
    return h + ':' + m;
  }


  writeTime(sheet);
  
  sumTime(sheet, spreadsheet, userData);


  if (typeof morning_out_legal_time != "undefined" && midnight_out_legal_time != "undefined") {
    return "長時間お疲れ様でした！\nお気をつけておかえりくださいね！";
  }




  // 深夜勤務時間であれば深夜時間切り分け
  return "お疲れ様でした！\nお気をつけておかえりくださいね！";
}

// ------------ スプレッドシートから計算 ------------------ //

function onOpen() {

  //メニュー配列
  var myMenu = [
    { name: "メール送信", functionName: "sendMail" },
    { name: "配信リスト更新", functionName: "inportContacts2" }
  ];

  SpreadsheetApp.getActiveSpreadsheet().addMenu("メール", myMenu); //メニューを追加

}

// ------------ スプレッドシートの取得 ------------------ //

function getSpredSheet(data) {

  data = JSON.parse(data);

  var prop = PropertiesService.getScriptProperties();
  var resYamada = prop.getProperty("山田シートID");
  var resIchikawa = prop.getProperty("市川シートID");

  switch (data[0]) {
    case "山田":
      var spreadsheet = SpreadsheetApp.openById(resYamada);
      break;
    case "市川":
      var spreadsheet = SpreadsheetApp.openById(resIchikawa);
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


// ------------ 合計時間記述 ------------------ //

function writeTime(sheet) {

  // 時間の計算
  // シートの色付き色なしの行を取得
  var has_color = [];
  var no_color = [];
  var now_lastRow = sheet.getLastRow();

  var midnight_has_color = [];
  var midnight_no_color = [];


  for (var i = 2; i <= now_lastRow; i++) {
    var check_background = sheet.getRange(i, 1).getBackground();

    if (check_background == "#ffffff") {
      no_color.push(i);
    } else {
      has_color.push(i);
    }

    var check_midnight_background = sheet.getRange(i, 10).getBackground();
    if (check_midnight_background == "#e8f0fe") {
      midnight_has_color.push(i);
    } else {
      midnight_no_color.push(i);
    }
  }


  // 通常時間集計
  if (no_color.length > 0) {
    sheet.getRange(1, 13).setValue("通常時間合計");
    sheet.getRange(1, 15).setValue("通常深夜");
    var sum;
    var out_time_sum;
    for (var j = 0; j < no_color.length; j++) {
      if (j == 0) {
        if (no_color.length == 1) {
          sum = "=SUM(H" + no_color[j] + ")";
          out_time_sum = "=SUM(I" + no_color[j] + ")";
        } else {
          sum = "=SUM(H" + no_color[j];
          out_time_sum = "=SUM(I" + no_color[j];
        }

      } else if (j == no_color.length - 1) {
        sum = sum + ",H" + no_color[j] + ")";
        out_time_sum = out_time_sum + ",I" + no_color[j] + ")";
      } else {
        sum = sum + ",H" + no_color[j];
        out_time_sum = out_time_sum + ",I" + no_color[j];
      }
    }
    sheet.getRange(2, 13).setNumberFormat('[h]:mm')
    sheet.getRange(2, 13).setFormula(sum);
    sheet.getRange(2, 15).setNumberFormat('[h]:mm')
    sheet.getRange(2, 15).setFormula(out_time_sum);

  }

  // 通常時間割増
  if (has_color.length > 0) {
    sheet.getRange(1, 14).setValue("休日シフト時間");
    sheet.getRange(1, 16).setValue("休日シフト深夜時間");
    var has_color_sum;
    var has_color_out_time_sum;
    for (var h = 0; h < has_color.length; h++) {
      if (h == 0) {
        if (has_color.length == 1) {
          console.log("あ");
          has_color_sum = "=SUM(H" + has_color_sum[h] + ")";
          has_color_out_time_sum = "=SUM(I" + has_color[h] + ")";
        } else {
          has_color_sum = "=SUM(H" + has_color[h];
          has_color_out_time_sum = "=SUM(I" + has_color[h];
        }


      } else if (h == has_color.length - 1) {
        console.log("う");
        has_color_sum = has_color_sum + ",H" + has_color[h] + ")";
        has_color_out_time_sum = has_color_out_time_sum + ",I" + has_color[h] + ")";
      } else {
        console.log("え");
        has_color_sum = has_color_sum + ",H" + has_color[h];
        has_color_out_time_sum = has_color_out_time_sum + ",I" + has_color[h] + ")";
      }
    }
    // 休日時間割増
    sheet.getRange(2, 14).setNumberFormat('[h]:mm')
    sheet.getRange(2, 14).setFormula(has_color_sum);
    sheet.getRange(2, 16).setNumberFormat('[h]:mm')
    sheet.getRange(2, 16).setFormula(has_color_out_time_sum);

  }
  
  if (midnight_no_color.length > 0) {
    sheet.getRange(1, 17).setValue("深夜翌日またぎ 平日");
    var midnight_no_color_sum;

    for (var k = 0; k < midnight_no_color.length; k++) {
      if (k == 0) {
        if (midnight_no_color.length == 1) {
          console.log(11);
          midnight_no_color_sum = "=SUM(J" + midnight_no_color[k] + ")";
        } else {
          console.log(12);
          midnight_no_color_sum = "=SUM(J" + midnight_no_color[k];
        }

      } else if (k == midnight_no_color.length - 1) {
        console.log(13);
        midnight_no_color_sum = midnight_no_color_sum + ",J" + midnight_no_color[k] + ")";

      } else {
        console.log(14);
        midnight_no_color_sum = midnight_no_color_sum + ",J" + midnight_no_color[k];
      }
    }
    // 休日時間割増
    sheet.getRange(2, 17).setNumberFormat('[h]:mm')
    sheet.getRange(2, 17).setFormula(midnight_no_color_sum);
  }


  if (midnight_has_color.length > 0) {
    sheet.getRange(1, 18).setValue("休日深夜翌日またぎ 割増");
    var midnight_has_color_sum;

    for (var k = 0; k < midnight_has_color.length; k++) {
      if (k == 0) {
        if (midnight_has_color.length == 1) {
          console.log(11);
          midnight_has_color_sum = "=SUM(J" + midnight_has_color[k] + ")";
        } else {
          console.log(12);
          midnight_has_color_sum = "=SUM(J" + midnight_has_color[k];
        }

      } else if (k == midnight_has_color.length - 1) {
        console.log(13);
        midnight_has_color_sum = midnight_has_color_sum + ",J" + midnight_has_color[k] + ")";

      } else {
        console.log(14);
        midnight_has_color_sum = midnight_has_color_sum + ",J" + midnight_has_color[k];
      }
    }
    // 休日時間割増
    sheet.getRange(2, 18).setNumberFormat('[h]:mm')
    sheet.getRange(2, 18).setFormula(midnight_has_color_sum);
  }

}

function sumTime(sheet, spreadsheet, userData) {
  
  
  // 丸める時間取得
  //  var prop = PropertiesService.getScriptProperties();
  //  var res = prop.getProperty("判定用シートID");
  //
  //  var spreadsheet = SpreadsheetApp.openById(res);
  //  var TimeCardsheet = spreadsheet.getSheetByName('CompanyRegulations');
  //  var ceiling = TimeCardsheet.getRange(2, 1).getValue();
  //  
  //  console.log(ceiling);
  //  console.log(typeof ceiling);
  //  
  //  consoe.log(2222);

  // 日時取得
  var [now, year, month, day, weekName, hour, min, sec, s, today] = getToday();
  // 法定時間外判別用

  var lastRow = sheet.getLastRow(); //対象となるシートの最終行を取得
  
  var multiply_col13 = sheet.getRange(2, 13).getValue();
  var multiply_col14 = sheet.getRange(2, 14).getValue();
  var multiply_col15 = sheet.getRange(2, 15).getValue();
  var multiply_col16 = sheet.getRange(2, 16).getValue();
  var multiply_col17 = sheet.getRange(2, 17).getValue();
  var multiply_col18 = sheet.getRange(2, 18).getValue();
  
 
  var multiply_1 = Number(new Date(multiply_col13).getHours()) + Number(new Date(multiply_col13).getMinutes()/60);
  var multiply_125 = (Number(new Date(multiply_col17).getHours()) + Number(new Date(multiply_col17).getMinutes()/60)) + (Number(new Date(multiply_col15).getHours()) + Number(new Date(multiply_col15).getMinutes()/60));
  var multiply_135 = Number(new Date(multiply_col14).getHours()) + Number(new Date(multiply_col14).getMinutes()/60);
  var multiply_160 = (Number(new Date(multiply_col16).getHours()) + Number(new Date(multiply_col16).getMinutes()/60)) + (Number(new Date(multiply_col18).getHours()) + Number(new Date(multiply_col18).getMinutes()/60));
  
  sheet.getRange(14, 1).setValue(multiply_1);
  sheet.getRange(14, 2).setValue(multiply_125);
  sheet.getRange(14, 3).setValue(multiply_135);
  sheet.getRange(14, 4).setValue(multiply_160);
  
  
}

