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
 *   2行目: 見出し（部屋番号, 氏名, 電話, 携帯, 日程, 備考, 確認コード, 第1希望, 第2希望, 第3希望）
 *   3行目以降: データ
 *     A列: 部屋番号
 *     B列: 氏名（「空室」の場合は空室として表示）
 *     E列: 日程（工程表の作成に使われる「確定日程」。既定では第1希望が入る。
 *          第1希望どおりに施工できない場合は、H〜J列を見ながらここを
 *          手動で書き換えてから「工程表を作成」を実行する）
 *     F列: 備考（「工期外希望」「カメラ付」「受話器」に対応）
 *     H・I・J列: 第1〜第3希望（回答フォームで集めた希望日時。参考情報として保存される）
 *
 * オンライン回答フォーム（紙のアンケート回収の代替）:
 *   「工程表」>「回答フォームを作成する」を実行すると、入居者一覧の部屋番号を
 *   選択肢にしたGoogleフォームが作成され、このスプレッドシートに接続されます。
 *   住民は氏名・電話番号・第1〜第3希望日時をフォーム上で入力し、回答すると
 *   自動的に該当する部屋番号のB列（氏名）・C列（電話）・E列（確定日程＝第1希望）・
 *   F列（備考）・H〜J列（第1〜第3希望）に反映されます。
 *
 *   「工程表」>「各住戸QRコードを作成する」を実行すると、住戸ごとに専用の
 *   QRコード（部屋番号＋確認コードを埋め込んだ回答リンク）を印刷用スライドとして
 *   生成します。他の部屋のQRコードを使って回答することはできません
 *   （確認コードが一致しない回答はスプレッドシートに反映されず、
 *   「フォーム取込エラー」シートに記録されます）。
 *
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
    .addItem("各住戸QRコードを作成する", "createPerRoomQrSlips")
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

// 工程表シート・フォーム取込エラーシートは部屋データとして扱わない
function isRoomDataSheet(sheet) {
  const name = sheet.getName();
  return name !== SCHEDULE_SHEET_NAME && name !== ERROR_SHEET_NAME;
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
    if (!isRoomDataSheet(sheet)) return;
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
const FORM_Q_CODE = "確認コード";
const FORM_Q_NAME = "氏名";
const FORM_Q_TEL = "電話番号";
const FORM_Q_DATE1 = "第1希望日";
const FORM_Q_TIME1 = "第1希望時間";
const FORM_Q_DATE2 = "第2希望日";
const FORM_Q_TIME2 = "第2希望時間";
const FORM_Q_DATE3 = "第3希望日";
const FORM_Q_TIME3 = "第3希望時間";
const FORM_Q_REMARK = "オプション・ご要望";
const FORM_ID_PROP = "KOUJI_FORM_ID";
const ERROR_SHEET_NAME = "フォーム取込エラー";
const CODE_COL = 7; // G列: 住戸ごとの確認コード（QRコードに埋め込む合言葉）
const PREF_COLS = [8, 9, 10]; // H・I・J列: 第1〜第3希望
const PREF_FIELDS = [
  [FORM_Q_DATE1, FORM_Q_TIME1, true],
  [FORM_Q_DATE2, FORM_Q_TIME2, false],
  [FORM_Q_DATE3, FORM_Q_TIME3, false],
];
// 旧バージョン（単一希望のみ）で使っていた質問。フォーム更新時に見つかれば削除する。
const LEGACY_FORM_TITLES = ["工事希望日", "工事希望時間"];

function buildTimeOptions() {
  const options = [];
  TIME_SLOTS.forEach((h) => {
    options.push(`${h}:00`);
    options.push(`${h}:30`);
  });
  return options;
}

function generatePassword() {
  return String(Math.floor(1000 + Math.random() * 9000)); // 4桁の数字
}

// 各部屋（空室を除く）にG列の確認コードが無ければ発行し、H〜J列（第1〜第3希望）の
// 見出しを整える。既存の確認コード・希望日時は変更しない。
function ensureSheetLayout(ss) {
  ss.getSheets().forEach((sheet) => {
    if (!isRoomDataSheet(sheet)) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    if (!sheet.getRange(2, CODE_COL).getValue()) {
      sheet.getRange(2, CODE_COL).setValue("確認コード");
    }
    ["第1希望", "第2希望", "第3希望"].forEach((label, i) => {
      const col = PREF_COLS[i];
      if (!sheet.getRange(2, col).getValue()) {
        sheet.getRange(2, col).setValue(label);
      }
    });

    const range = sheet.getRange(3, 1, lastRow - 2, CODE_COL);
    const values = range.getValues();
    let changed = false;
    values.forEach((row) => {
      const room = row[0];
      const name = row[1];
      if (room === "" || room === null) return;
      if (String(name || "") === "空室") return;
      if (!row[CODE_COL - 1]) {
        row[CODE_COL - 1] = generatePassword();
        changed = true;
      }
    });
    if (changed) range.setValues(values);
  });
}

function getRoomList(ss) {
  const rooms = [];
  ss.getSheets().forEach((sheet) => {
    if (!isRoomDataSheet(sheet)) return;
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

// 部屋番号と確認コードの一覧（各住戸QRコードの生成に使う）
function getRoomEntries(ss) {
  const entries = [];
  ss.getSheets().forEach((sheet) => {
    if (!isRoomDataSheet(sheet)) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const values = sheet.getRange(3, 1, lastRow - 2, CODE_COL).getValues();
    values.forEach((row) => {
      const roomRaw = row[0];
      const name = row[1];
      if (roomRaw === "" || roomRaw === null) return;
      if (String(name || "") === "空室") return;
      entries.push({ room: formatRoom(roomRaw), code: String(row[CODE_COL - 1] || "") });
    });
  });
  return entries;
}

function findItemByTitle(form, title) {
  return form.getItems().find((item) => item.getTitle() === title) || null;
}

function logFormError(ss, message) {
  let sheet = ss.getSheetByName(ERROR_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(ERROR_SHEET_NAME);
    sheet.appendRow(["日時", "内容"]);
  }
  sheet.appendRow([new Date(), message]);
}

function createOrUpdateKoujiForm() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buildingName = String(ss.getSheets()[0].getRange(1, 1).getValue() || "工事");

  ensureSheetLayout(ss);
  const rooms = getRoomList(ss);
  if (rooms.length === 0) {
    ui.alert("入居者一覧に部屋番号が見つかりません。先に部屋番号を入力してください（空室はB列に「空室」と入力）。");
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
      "工事の希望日時を第1希望〜第3希望までご回答ください（第2・第3希望は任意です）。" +
        "同じ部屋番号で再度回答すると、内容は最新の回答で上書きされます。"
    );
    ScriptApp.newTrigger("onKoujiFormSubmit").forForm(form).onFormSubmit().create();
  }

  // 既存の質問はできる限り作り直さず、内容だけ更新する
  // （質問を削除・再作成するとIDが変わり、配布済みの住戸別QRコードが無効になるため）
  const roomItem = findItemByTitle(form, FORM_Q_ROOM);
  if (roomItem) {
    roomItem.asListItem().setChoiceValues(rooms);
  } else {
    form.addListItem().setTitle(FORM_Q_ROOM).setChoiceValues(rooms).setRequired(true);
  }

  if (!findItemByTitle(form, FORM_Q_CODE)) {
    form.addTextItem()
      .setTitle(FORM_Q_CODE)
      .setHelpText("QRコードから自動的に入力されます。ご自身での入力は不要です。")
      .setRequired(true);
  }

  if (!findItemByTitle(form, FORM_Q_NAME)) {
    form.addTextItem().setTitle(FORM_Q_NAME).setRequired(true);
  }

  if (!findItemByTitle(form, FORM_Q_TEL)) {
    form.addTextItem()
      .setTitle(FORM_Q_TEL)
      .setHelpText("工事当日にご連絡が取れる番号をご記入ください。")
      .setRequired(true);
  }

  PREF_FIELDS.forEach(([dateTitle, timeTitle, required]) => {
    if (!findItemByTitle(form, dateTitle)) {
      form.addDateItem().setTitle(dateTitle).setIncludesYear(true).setRequired(required);
    }
    if (!findItemByTitle(form, timeTitle)) {
      form.addListItem().setTitle(timeTitle).setChoiceValues(buildTimeOptions()).setRequired(required);
    }
  });

  // 旧バージョンの単一希望の質問が残っていれば削除する
  LEGACY_FORM_TITLES.forEach((title) => {
    const legacyItem = findItemByTitle(form, title);
    if (legacyItem) form.deleteItem(legacyItem);
  });

  if (!findItemByTitle(form, FORM_Q_REMARK)) {
    form.addCheckboxItem()
      .setTitle(FORM_Q_REMARK)
      .setChoiceValues([
        "インターホンカメラ付きオプションを希望",
        "インターホン受話器増設を希望",
        "指定期間内に都合がつかない（工期外希望）",
      ])
      .setRequired(false);
  }

  // 質問の表示順を整える（IDは変わらないため、配布済みQRコードには影響しない）
  [
    FORM_Q_ROOM,
    FORM_Q_CODE,
    FORM_Q_NAME,
    FORM_Q_TEL,
    FORM_Q_DATE1,
    FORM_Q_TIME1,
    FORM_Q_DATE2,
    FORM_Q_TIME2,
    FORM_Q_DATE3,
    FORM_Q_TIME3,
    FORM_Q_REMARK,
  ].forEach((title, index) => {
    const item = findItemByTitle(form, title);
    if (item) form.moveItem(item.getIndex(), index);
  });

  ui.alert(
    isNew ? "回答フォームを作成しました" : "回答フォームを更新しました",
    `フォームURL:\n${form.getPublishedUrl()}\n\n` +
      "続けて「工程表」>「各住戸QRコードを作成する」を実行すると、" +
      "住戸ごとに専用のQRコード（他の部屋の回答には使えません）を印刷用スライドとして生成できます。",
    ui.ButtonSet.OK
  );
}

function weekdayJP(year, month, day) {
  const dow = new Date(year, month - 1, day).getDay();
  return WEEKDAYS[(dow + 6) % 7];
}

// フォームの日付・時間の回答を "m/d（weekday）h:mm" 形式にまとめる。
// どちらか未回答なら null（第2・第3希望は任意のため、未回答もあり得る）。
function buildSchedText(answers, dateTitle, timeTitle) {
  const dateStr = String(answers[dateTitle] || "");
  const [y, m, d] = dateStr.split("-").map(Number);
  const timeStr = String(answers[timeTitle] || "");
  const [hour, minute] = timeStr.split(":");
  if (!y || !m || !d || !hour) return null;
  const weekday = weekdayJP(y, m, d);
  return `${m}/${d}（${weekday}）${hour}:${minute}`;
}

function findRoomRow(ss, room) {
  let found = null;
  ss.getSheets().some((sheet) => {
    if (!isRoomDataSheet(sheet)) return false;
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

  const target = findRoomRow(ss, room);
  if (!target) {
    logFormError(ss, `部屋番号「${room}」が入居者一覧に見つかりませんでした。`);
    return;
  }

  // 住戸別QRコードに埋め込まれた確認コードと一致するか照合する。
  // 一致しない場合は「別の部屋のQRコードで回答された」可能性があるため反映しない。
  const expectedCode = String(target.sheet.getRange(target.row, CODE_COL).getValue() || "").trim();
  const submittedCode = String(answers[FORM_Q_CODE] || "").trim();
  if (expectedCode && submittedCode !== expectedCode) {
    logFormError(ss, `部屋番号「${room}」の確認コードが一致しませんでした（別の部屋のQRコードの可能性）。`);
    return;
  }

  const preferences = PREF_FIELDS.map(([dateTitle, timeTitle]) => buildSchedText(answers, dateTitle, timeTitle));
  if (!preferences[0]) return; // 第1希望は必須のため通常ここには来ない

  const remarkAnswer = answers[FORM_Q_REMARK];
  const remarks = Array.isArray(remarkAnswer) ? remarkAnswer : remarkAnswer ? [remarkAnswer] : [];
  const remarkText = remarks.join("、");

  const name = String(answers[FORM_Q_NAME] || "").trim();
  const tel = String(answers[FORM_Q_TEL] || "").trim();
  if (name) target.sheet.getRange(target.row, 2).setValue(name);
  if (tel) target.sheet.getRange(target.row, 3).setValue(tel);

  // E列（確定日程）は既定で第1希望を採用する。第1希望どおりに施工できない場合は、
  // H〜J列（第1〜第3希望）を見ながら管理者がE列を手動で書き換えてから工程表を作成する。
  target.sheet.getRange(target.row, 5).setValue(preferences[0]);
  target.sheet.getRange(target.row, 6).setValue(remarkText);
  PREF_COLS.forEach((col, i) => {
    target.sheet.getRange(target.row, col).setValue(preferences[i] || "");
  });
}

function onKoujiFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyFormResponse(ss, e.response.getItemResponses());
}

// 住戸ごとに「部屋番号＋確認コード」を埋め込んだ回答用URLをQRコード化し、
// 1住戸1ページの印刷用Googleスライドとして生成する。
// 他の部屋のQRコードを使って回答しても、確認コードが一致しないため反映されない。
function createPerRoomQrSlips() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const formId = props.getProperty(FORM_ID_PROP);
  if (!formId) {
    ui.alert("先に「回答フォームを作成する」を実行してください。");
    return;
  }

  ensureSheetLayout(ss);
  const entries = getRoomEntries(ss);
  if (entries.length === 0) {
    ui.alert("部屋番号が見つかりません。");
    return;
  }

  const form = FormApp.openById(formId);
  const roomItem = findItemByTitle(form, FORM_Q_ROOM);
  const codeItem = findItemByTitle(form, FORM_Q_CODE);
  if (!roomItem || !codeItem) {
    ui.alert("フォームの質問が見つかりません。先に「回答フォームを作成する」を実行してください。");
    return;
  }

  const buildingName = String(ss.getSheets()[0].getRange(1, 1).getValue() || "工事");
  const presentation = SlidesApp.create(`${buildingName} 工事アンケートQRコード`);
  const placeholderSlide = presentation.getSlides()[0];
  let failedCount = 0;

  entries.forEach(({ room, code }) => {
    const formResponse = form.createResponse();
    formResponse.withItemResponse(roomItem.asListItem().createResponse(room));
    formResponse.withItemResponse(codeItem.asTextItem().createResponse(code));
    const url = formResponse.toPrefilledUrl();

    const slide = presentation.appendSlide(SlidesApp.PredefinedLayout.BLANK);

    slide.insertTextBox(`${room} 号室`, 40, 30, 400, 50)
      .getText().getTextStyle().setFontSize(28).setBold(true);
    slide.insertTextBox(
      "スマホのカメラでQRコードを読み取り、工事希望日時をご回答ください。",
      40, 90, 550, 40
    ).getText().getTextStyle().setFontSize(14);

    const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=" + encodeURIComponent(url);
    let inserted = false;
    try {
      const resp = UrlFetchApp.fetch(qrUrl, { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        slide.insertImage(resp.getBlob(), 150, 150, 250, 250);
        inserted = true;
      }
    } catch (err) {
      inserted = false;
    }
    if (!inserted) {
      failedCount += 1;
      slide.insertTextBox(
        `（QR画像の取得に失敗しました。下記URLを直接ご案内ください）\n${url}`,
        40, 150, 550, 120
      ).getText().getTextStyle().setFontSize(10);
    }

    slide.insertTextBox(
      `確認コード: ${code}（QRコードに自動で含まれています。手入力は不要です）`,
      40, 420, 550, 30
    ).getText().getTextStyle().setFontSize(10);
  });

  placeholderSlide.remove();

  ui.alert(
    "各住戸専用QRコードを作成しました",
    `スライドURL:\n${presentation.getUrl()}\n\n` +
      "1ページ＝1住戸のQRコードです。印刷してポストや玄関先など、各住戸ごとに配布・掲示してください。\n" +
      "QRコードには部屋番号と確認コードが埋め込まれているため、他の部屋の回答として使うことはできません。" +
      (failedCount > 0
        ? `\n\n※${failedCount}件はQR画像の自動取得に失敗しました。該当ページのURLを別のQR作成サイトなどでご利用ください。`
        : ""),
    ui.ButtonSet.OK
  );
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
