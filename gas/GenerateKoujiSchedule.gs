/**
 * 入居者一覧（Googleスプレッドシート）から住戸内工事日程表を生成するGoogle Apps Script。
 *
 * 使い方:
 *   1. 入居者一覧のスプレッドシートを開く
 *   2. 拡張機能 > Apps Script を開き、このファイルの内容を貼り付ける
 *   3. スプレッドシートを再読み込みすると「工程表」メニューが追加される
 *   4. 「工程表」 > 「工程表を作成」を実行する
 *      （初回はGoogleの権限承認が必要です）
 *   5. 「工程表」という名前のシートが作成・更新される
 *
 * 入居者一覧シートのフォーマット（各シート共通）:
 *   1行目: タイトル（A1: 物件名）
 *   2行目: 見出し（部屋番号, 氏名, 電話, 携帯, 日程, 備考）
 *   3行目以降: データ
 *     A列: 部屋番号
 *     B列: 氏名（「空室」の場合は空室として表示）
 *     E列: 日程（例: 6/19（金）13:00 / 3月28日(土)9：00）
 *     F列: 備考（「工期外希望」「カメラ付」「受話器」に対応）
 *
 * オンライン回答フォーム（紙のアンケート回収の代替）:
 *   「工程表」>「回答フォームを作成する」を実行すると、入居者一覧の部屋番号を
 *   選択肢にしたGoogleフォームが作成され、このスプレッドシートに接続されます。
 *   住民がフォームに回答すると、自動的に該当する部屋番号のE列・F列に反映されます。
 *   詳しい手順は gas/README.md を参照してください。
 */

const TIME_SLOTS = [9, 10, 11, 13, 14, 15, 16, 17];
const AM_SLOTS = [9, 10, 11];
const PM_SLOTS = [13, 14, 15, 16, 17];
const TOTAL_COLS = 1 + AM_SLOTS.length + PM_SLOTS.length; // 工事予定日 + 8時間帯
const SCHEDULE_SHEET_NAME = "工程表";

const WEEKDAYS = ["月", "火", "水", "木", "金", "土", "日"];

const YELLOW = "#FFFF00";
const ORANGE = "#FFC000";
const GRAY = "#D9D9D9";

const DATE_RE = /(\d{1,2})[/月](\d{1,2})日?[（(](.)[）)]\s*(\d{1,2})[:：](\d{2})/;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("工程表")
    .addItem("工程表を作成", "generateKoujiSchedule")
    .addSeparator()
    .addItem("回答フォームを作成する", "createOrUpdateKoujiForm")
    .addItem("フォーム回答を再取り込み", "resyncFormResponses")
    .addToUi();
}

function parseSchedule(text) {
  if (!text) return null;
  const m = DATE_RE.exec(String(text));
  if (!m) return null;
  return {
    month: parseInt(m[1], 10),
    day: parseInt(m[2], 10),
    weekday: m[3],
    hour: parseInt(m[4], 10),
    minute: parseInt(m[5], 10),
  };
}

function optionMarker(remark) {
  if (!remark) return "";
  let suffix = "";
  if (remark.indexOf("カメラ") !== -1) suffix += "A";
  if (remark.indexOf("受話器") !== -1) suffix += "B";
  return suffix;
}

function formatRoom(room) {
  if (typeof room === "number") {
    return Number.isInteger(room) ? String(room) : String(room);
  }
  return String(room).trim();
}

function commonAreaDate(month, day, weekday) {
  const d = new Date(2001, month - 1, day);
  d.setDate(d.getDate() - 1);
  const prevWeekday = WEEKDAYS[(WEEKDAYS.indexOf(weekday) - 1 + 7) % 7];
  return { month: d.getMonth() + 1, day: d.getDate(), weekday: prevWeekday };
}

function collectRoomData(ss) {
  let buildingName = "";
  const dates = {}; // "m-d" -> {weekday, slots: {hour: [[room, minute, option], ...]}}
  const outOfPeriod = [];
  const vacant = [];
  const notSubmitted = [];

  ss.getSheets().forEach((sheet) => {
    if (sheet.getName() === SCHEDULE_SHEET_NAME) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    if (!buildingName) {
      buildingName = String(sheet.getRange(1, 1).getValue() || "");
    }

    const values = sheet.getRange(3, 1, lastRow - 2, 6).getValues();
    values.forEach((row) => {
      const [roomRaw, name, , , schedText, remark] = row;
      if (roomRaw === "" || roomRaw === null) return;

      const room = formatRoom(roomRaw);
      const remarkStr = String(remark || "");

      if (String(name || "") === "空室") {
        vacant.push(room);
        return;
      }
      if (remarkStr.indexOf("工期外") !== -1) {
        outOfPeriod.push({ room, schedule: parseSchedule(schedText) });
        return;
      }

      const sched = parseSchedule(schedText);
      if (!sched) {
        notSubmitted.push(room);
        return;
      }

      const key = sched.month + "-" + sched.day;
      if (!dates[key]) dates[key] = { weekday: sched.weekday, slots: {} };
      if (!dates[key].slots[sched.hour]) dates[key].slots[sched.hour] = [];
      dates[key].slots[sched.hour].push([room, sched.minute, optionMarker(remarkStr)]);
    });
  });

  return { buildingName, dates, outOfPeriod, vacant, notSubmitted };
}

function generateKoujiSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const { buildingName, dates, outOfPeriod, vacant, notSubmitted } = collectRoomData(ss);

  let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);

  // タイトル
  sheet.getRange(1, 1, 1, TOTAL_COLS).merge()
    .setValue(`${buildingName}　様　住戸内工事日程表`)
    .setFontSize(16).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true);
  sheet.setRowHeight(1, 36);

  // 注記
  sheet.getRange(2, 1, 1, TOTAL_COLS).merge()
    .setValue("※表中の部屋数は工事可能な部屋数を示します。")
    .setHorizontalAlignment("left");

  const headerRow1 = 3;
  const headerRow2 = 4;

  // 工事予定日
  sheet.getRange(headerRow1, 1, 2, 1).merge()
    .setValue("工事予定日")
    .setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 午前
  sheet.getRange(headerRow1, 2, 1, AM_SLOTS.length).merge()
    .setValue("午前（9時～12時）")
    .setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 午後
  const pmStart = 2 + AM_SLOTS.length;
  sheet.getRange(headerRow1, pmStart, 1, PM_SLOTS.length).merge()
    .setValue("午後（13時～18時）")
    .setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // 時間帯ラベル
  TIME_SLOTS.forEach((h, i) => {
    sheet.getRange(headerRow2, 2 + i)
      .setValue(`${h}時頃`)
      .setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
  });

  sheet.getRange(headerRow1, 1, 2, TOTAL_COLS)
    .setBorder(true, true, true, true, true, true);

  // データ行
  let row = headerRow2 + 1;

  const sortedKeys = Object.keys(dates).sort((a, b) => {
    const [am, ad] = a.split("-").map(Number);
    const [bm, bd] = b.split("-").map(Number);
    return am * 100 + ad - (bm * 100 + bd);
  });

  // 初日の前日を共用部工事日として追加
  if (sortedKeys.length > 0) {
    const [fm, fd] = sortedKeys[0].split("-").map(Number);
    const fwd = dates[sortedKeys[0]].weekday;
    const ca = commonAreaDate(fm, fd, fwd);

    sheet.getRange(row, 1)
      .setValue(`${ca.month}月${ca.day}日（${ca.weekday}）`)
      .setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    sheet.getRange(row, 2, 1, TOTAL_COLS - 1).merge()
      .setValue("共　用　部　工　事\n（お部屋の工事は出来ません）")
      .setFontWeight("bold").setFontSize(12)
      .setHorizontalAlignment("center").setVerticalAlignment("middle")
      .setWrap(true)
      .setBackground(GRAY);

    sheet.getRange(row, 1, 1, TOTAL_COLS)
      .setBorder(true, true, true, true, true, true)
      .setBackground(GRAY);
    sheet.setRowHeight(row, 50);
    row += 1;
  }

  // 各日付の行
  sortedKeys.forEach((key) => {
    const [month, day] = key.split("-").map(Number);
    const { weekday, slots } = dates[key];

    let nRows = 1;
    TIME_SLOTS.forEach((h) => {
      if (slots[h] && slots[h].length > nRows) nRows = slots[h].length;
    });

    sheet.getRange(row, 1, nRows, 1).merge()
      .setValue(`${month}月${day}日（${weekday}）`)
      .setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    TIME_SLOTS.forEach((h, i) => {
      const col = 2 + i;
      const entries = (slots[h] || []).slice().sort((a, b) => a[1] - b[1]);
      for (let sub = 0; sub < nRows; sub++) {
        const cell = sheet.getRange(row + sub, col);
        cell.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
        if (sub < entries.length) {
          const [room, minute, option] = entries[sub];
          const minuteStr = minute < 10 ? "0" + minute : String(minute);
          cell.setValue(`${h}:${minuteStr}\n${room}${option}`);
          cell.setFontWeight("bold");
          cell.setBackground(option ? ORANGE : YELLOW);
        }
      }
    });

    sheet.getRange(row, 1, nRows, TOTAL_COLS)
      .setBorder(true, true, true, true, true, true);

    row += nRows;
  });

  // 列幅・行高
  sheet.setColumnWidth(1, 110);
  for (let i = 0; i < TIME_SLOTS.length; i++) {
    sheet.setColumnWidth(2 + i, 70);
  }
  for (let r = headerRow2 + 1; r < row; r++) {
    sheet.setRowHeight(r, 45);
  }

  // オプション凡例
  const hasOption = sortedKeys.some((key) =>
    Object.values(dates[key].slots).some((entries) => entries.some(([, , opt]) => opt))
  );
  if (hasOption) {
    sheet.getRange(row, 1)
      .setValue("A：カメラ付き　B：受話器付きのお部屋です")
      .setBackground(ORANGE)
      .setFontWeight("bold");
    row += 1;
  }

  // 工期外希望
  row += 1;
  outOfPeriod.forEach(({ room, schedule }) => {
    let text = `${room} 工期外希望`;
    if (schedule) {
      const minuteStr = schedule.minute < 10 ? "0" + schedule.minute : String(schedule.minute);
      text += `（${schedule.month}/${schedule.day}${schedule.weekday}${schedule.hour}:${minuteStr}）`;
    }
    sheet.getRange(row, 1)
      .setValue(text)
      .setFontColor("#0000FF")
      .setFontWeight("bold");
    row += 1;
  });

  // 未提出
  if (notSubmitted.length > 0) {
    sheet.getRange(row, 1)
      .setValue(`未提出　${notSubmitted.join(", ")}`)
      .setFontColor("#FF0000")
      .setFontWeight("bold");
    row += 1;
  }

  // 空室
  if (vacant.length > 0) {
    sheet.getRange(row, 1)
      .setValue(`空室　${vacant.join(", ")}`)
      .setFontColor("#FF0000")
      .setFontWeight("bold");
    row += 1;
  }

  // 印刷設定（A4横・1ページに収める）はSpreadsheetのUIから
  // ファイル > 印刷 > 用紙サイズ:A4, 向き:横, 拡大縮小:1ページに収める
  // で設定してください（Apps Scriptからは設定できません）。

  sheet.setFrozenRows(headerRow2);
  SpreadsheetApp.getActiveSpreadsheet().toast("工程表を作成しました。");
}

/**
 * オンライン回答フォーム（紙アンケート・ポスト回収の代替）
 * ------------------------------------------------------------
 * 入居者一覧の部屋番号を選択肢にしたGoogleフォームを作成し、このスプレッドシートに
 * 接続する。住民の回答は自動でE列（日程）・F列（備考）に反映される。
 */

const FORM_Q_ROOM = "部屋番号";
const FORM_Q_DATE = "工事希望日";
const FORM_Q_TIME = "工事希望時間";
const FORM_Q_REMARK = "オプション・ご要望";
const FORM_ID_PROP = "KOUJI_FORM_ID";

function buildTimeOptions() {
  const options = [];
  TIME_SLOTS.forEach((h) => {
    options.push(`${h}:00`);
    options.push(`${h}:30`);
  });
  return options;
}

function getRoomList(ss) {
  const rooms = [];
  ss.getSheets().forEach((sheet) => {
    if (sheet.getName() === SCHEDULE_SHEET_NAME) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const values = sheet.getRange(3, 1, lastRow - 2, 2).getValues();
    values.forEach(([roomRaw, name]) => {
      if (roomRaw === "" || roomRaw === null) return;
      if (String(name || "") === "空室") return;
      rooms.push(formatRoom(roomRaw));
    });
  });
  return rooms;
}

function createOrUpdateKoujiForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buildingName = String(ss.getSheets()[0].getRange(1, 1).getValue() || "工事");
  const rooms = getRoomList(ss);
  if (rooms.length === 0) {
    ui.alert("入居者一覧に部屋番号が見つかりません。先に部屋番号・氏名を入力してください。");
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  const existingFormId = props.getProperty(FORM_ID_PROP);
  let form = null;
  if (existingFormId) {
    try {
      form = FormApp.openById(existingFormId);
    } catch (err) {
      form = null;
    }
  }

  const isNew = !form;
  if (isNew) {
    form = FormApp.create(`${buildingName} 工事日程アンケート`);
    props.setProperty(FORM_ID_PROP, form.getId());
    form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
    form.setCollectEmail(false);
    form.setDescription(
      "工事の希望日時をご回答ください。同じ部屋番号で再度回答すると、内容は最新の回答で上書きされます。"
    );
    ScriptApp.newTrigger("onKoujiFormSubmit").forForm(form).onFormSubmit().create();
  } else {
    // 部屋番号リストが変わっている場合に備え、質問を作り直す
    form.getItems().forEach((item) => form.deleteItem(item));
  }

  form.addListItem().setTitle(FORM_Q_ROOM).setChoiceValues(rooms).setRequired(true);
  form.addDateItem().setTitle(FORM_Q_DATE).setIncludesYear(true).setRequired(true);
  form.addListItem().setTitle(FORM_Q_TIME).setChoiceValues(buildTimeOptions()).setRequired(true);
  form.addCheckboxItem()
    .setTitle(FORM_Q_REMARK)
    .setChoiceValues([
      "インターホンカメラ付きオプションを希望",
      "インターホン受話器増設を希望",
      "指定期間内に都合がつかない（工期外希望）",
    ])
    .setRequired(false);

  ui.alert(
    "回答フォームを作成しました",
    `フォームURL:\n${form.getPublishedUrl()}\n\n` +
      "Googleフォームの編集画面右上の「送信」ボタンからQRコードを表示できます。\n" +
      "掲示物や配布物にQRコードを貼り、住民にスマホから回答してもらってください。\n\n" +
      "住民の回答は自動的に入居者一覧の該当部屋番号（日程・備考欄）に反映されます。",
    ui.ButtonSet.OK
  );
}

function weekdayJP(year, month, day) {
  const dow = new Date(year, month - 1, day).getDay();
  return WEEKDAYS[(dow + 6) % 7];
}

function findRoomRow(ss, room) {
  let found = null;
  ss.getSheets().some((sheet) => {
    if (sheet.getName() === SCHEDULE_SHEET_NAME) return false;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return false;
    const values = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (formatRoom(values[i][0]) === room) {
        found = { sheet, row: i + 3 };
        return true;
      }
    }
    return false;
  });
  return found;
}

function applyFormResponse(ss, itemResponses) {
  const answers = {};
  itemResponses.forEach((ir) => {
    answers[ir.getItem().getTitle()] = ir.getResponse();
  });

  const room = String(answers[FORM_Q_ROOM] || "").trim();
  if (!room) return;

  const dateStr = String(answers[FORM_Q_DATE] || "");
  const [y, m, d] = dateStr.split("-").map(Number);
  const timeStr = String(answers[FORM_Q_TIME] || "");
  const [hour, minute] = timeStr.split(":");
  if (!y || !m || !d || !hour) return;

  const weekday = weekdayJP(y, m, d);
  const schedText = `${m}/${d}（${weekday}）${hour}:${minute}`;

  const remarkAnswer = answers[FORM_Q_REMARK];
  const remarks = Array.isArray(remarkAnswer) ? remarkAnswer : remarkAnswer ? [remarkAnswer] : [];
  const remarkText = remarks.join("、");

  const target = findRoomRow(ss, room);
  if (!target) {
    Logger.log(`部屋番号「${room}」が入居者一覧に見つかりませんでした。`);
    return;
  }
  target.sheet.getRange(target.row, 5).setValue(schedText);
  target.sheet.getRange(target.row, 6).setValue(remarkText);
}

function onKoujiFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyFormResponse(ss, e.response.getItemResponses());
}

function resyncFormResponses() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getDocumentProperties();
  const formId = props.getProperty(FORM_ID_PROP);
  if (!formId) {
    ui.alert("先に「回答フォームを作成する」を実行してください。");
    return;
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const form = FormApp.openById(formId);
  const responses = form.getResponses();
  responses.forEach((response) => applyFormResponse(ss, response.getItemResponses()));
  ui.alert(`${responses.length}件の回答を取り込みました。`);
}
