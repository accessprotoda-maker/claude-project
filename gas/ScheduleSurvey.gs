/**
 * 住戸ごとのQRコード・パスワードを発行し、スマホから工事希望日程（第1〜第3希望）を
 * 入力してもらい、自動で時間枠に振り分けるGoogle Apps Scriptプロトタイプ。
 *
 * 前提:
 *   ・1つの時間枠（例: 6/19 9時頃）に工事できるのは1住戸まで（容量=1）。
 *   ・各住戸は第1〜第3希望（日付＋時間帯）を入力する。
 *   ・自動振り分けは「提出が早い住戸」を優先し、第1→第2→第3希望の順に空き枠を探す。
 *     3つとも埋まっていた住戸は「振分不可」として一覧化する。
 *
 * 導入手順:
 *   1. スプレッドシートに「住戸一覧」シートを作成し、A列に部屋番号・B列に氏名を入力する
 *      （1行目は見出し行、2行目以降がデータ）
 *   2. 拡張機能 > Apps Script を開き、このファイルの内容を貼り付けて保存する
 *   3. 「デプロイ」>「新しいデプロイ」>「ウェブアプリ」として公開する
 *      （アクセスできるユーザー: 全員 にしておく）
 *   4. スプレッドシートを再読み込みすると「工事日程アンケート」メニューが追加される
 *   5. 「工事日程アンケート」>「QR・パスワードを発行」を実行する
 *      （住戸一覧シートに パスワード・回答用URL・QRコード が追加される）
 *   6. QRコード・パスワードを工事案内に印刷して各住戸に配布する
 *   7. 住人がQRを読み取り → パスワードを入力 → 第1〜第3希望を入力・送信すると
 *      「回答」シートに記録される
 *   8. 「工事日程アンケート」>「自動振り分けを実行」で「工程表（自動振分）」シートが作成される
 *      振分不可の住戸は赤字で一覧表示される
 */

const SURVEY_TARGET_SHEET = "住戸一覧";
const SURVEY_RESPONSE_SHEET = "回答";
const SURVEY_RESULT_SHEET = "工程表（自動振分）";

const SURVEY_TIME_SLOTS = [9, 10, 11, 13, 14, 15, 16, 17];
const SURVEY_WEEKDAYS = ["日", "月", "火", "水", "木", "金", "土"];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("工事日程アンケート")
    .addItem("QR・パスワードを発行", "setupSurveyLinks")
    .addItem("自動振り分けを実行", "allocateSurveySchedule")
    .addToUi();
}

// ---------------------------------------------------------------------------
// QR・パスワード発行
// ---------------------------------------------------------------------------

function generatePassword_() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // 紛らわしい文字を除外
  let s = "";
  for (let i = 0; i < 6; i++) {
    s += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return s;
}

function generateToken_() {
  return Utilities.getUuid().replace(/-/g, "").slice(0, 16);
}

function setupSurveyLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SURVEY_TARGET_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`「${SURVEY_TARGET_SHEET}」シートが見つかりません。A列に部屋番号、B列に氏名を入力したシートを用意してください。`);
    return;
  }

  const header = ["部屋番号", "氏名", "パスワード", "トークン", "回答用URL", "QRコード"];
  const headerRange = sheet.getRange(1, 1, 1, header.length);
  if (sheet.getRange(1, 1).getValue() !== header[0]) {
    headerRange.setValues([header]).setFontWeight("bold");
  }

  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    SpreadsheetApp.getUi().alert("ウェブアプリとしてデプロイしてから実行してください（デプロイ > 新しいデプロイ > ウェブアプリ）。");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const data = sheet.getRange(2, 1, numRows, 6).getValues();

  for (let i = 0; i < data.length; i++) {
    const room = data[i][0];
    if (room === "" || room === null) continue;

    let password = data[i][2];
    let token = data[i][3];
    if (!password) password = generatePassword_();
    if (!token) token = generateToken_();

    const url = `${webAppUrl}?t=${token}`;
    const qr = `=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=160x160&data="&ENCODEURL("${url}"))`;

    data[i][2] = password;
    data[i][3] = token;
    data[i][4] = url;
  }

  sheet.getRange(2, 1, numRows, 5).setValues(data);

  // QRコード列（数式）はまとめて設定し直す
  for (let i = 0; i < data.length; i++) {
    const room = data[i][0];
    if (room === "" || room === null) continue;
    const url = data[i][4];
    sheet.getRange(2 + i, 6).setFormula(
      `=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=160x160&data="&ENCODEURL("${url}"))`
    );
  }

  sheet.setColumnWidths(1, 5, 140);
  sheet.setColumnWidth(6, 170);
  for (let i = 0; i < numRows; i++) sheet.setRowHeight(2 + i, 170);

  SpreadsheetApp.getActiveSpreadsheet().toast("QR・パスワードを発行しました。");
}

function lookupRoomByToken_(token) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SURVEY_TARGET_SHEET);
  if (!sheet) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  for (const row of data) {
    const [room, name, password, rowToken] = row;
    if (rowToken && String(rowToken) === token) {
      return { room: String(room), name: String(name), password: String(password) };
    }
  }
  return null;
}

// ---------------------------------------------------------------------------
// 入力フォーム（ウェブアプリ）
// ---------------------------------------------------------------------------

function doGet(e) {
  const token = e && e.parameter && e.parameter.t;
  const entry = token ? lookupRoomByToken_(token) : null;

  if (!entry) {
    return HtmlService.createHtmlOutput("<p>無効なURLです。お手元の案内に記載のQRコードから再度アクセスしてください。</p>");
  }

  const template = HtmlService.createTemplateFromFile("SurveyForm");
  template.token = token;
  template.room = entry.room;
  template.timeSlots = SURVEY_TIME_SLOTS;
  return template.evaluate().setTitle("工事日程アンケート").addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function verifyAndSaveSurvey(payload) {
  const entry = lookupRoomByToken_(payload.token);
  if (!entry) {
    return { ok: false, message: "無効なURLです。" };
  }
  if (String(payload.password || "").trim() !== entry.password) {
    return { ok: false, message: "パスワードが違います。" };
  }

  const choices = [payload.choice1, payload.choice2, payload.choice3];
  for (const c of choices) {
    if (!c || !c.date || !c.hour) {
      return { ok: false, message: "第1〜第3希望をすべて入力してください。" };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SURVEY_RESPONSE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SURVEY_RESPONSE_SHEET);
    sheet.appendRow(["タイムスタンプ", "部屋番号", "氏名",
      "第1希望日", "第1希望時間", "第2希望日", "第2希望時間", "第3希望日", "第3希望時間"]);
    sheet.getRange(1, 1, 1, 9).setFontWeight("bold");
  }

  sheet.appendRow([
    new Date(), entry.room, entry.name,
    choices[0].date, choices[0].hour,
    choices[1].date, choices[1].hour,
    choices[2].date, choices[2].hour,
  ]);

  return { ok: true, message: "回答を受け付けました。ありがとうございました。" };
}

// ---------------------------------------------------------------------------
// 自動振り分け
// ---------------------------------------------------------------------------

function latestResponsesByRoom_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SURVEY_RESPONSE_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const byRoom = {};
  data.forEach((row) => {
    const [timestamp, room, name, d1, h1, d2, h2, d3, h3] = row;
    if (!room) return;
    const key = String(room);
    const entry = {
      timestamp,
      room: key,
      name: String(name),
      choices: [
        { date: formatDate_(d1), hour: Number(h1) },
        { date: formatDate_(d2), hour: Number(h2) },
        { date: formatDate_(d3), hour: Number(h3) },
      ],
    };
    // 同じ住戸から複数回回答があった場合は最新（タイムスタンプが新しい方）を採用
    if (!byRoom[key] || entry.timestamp > byRoom[key].timestamp) {
      byRoom[key] = entry;
    }
  });

  // 提出が早い住戸を優先して割り当てる
  return Object.values(byRoom).sort((a, b) => a.timestamp - b.timestamp);
}

function formatDate_(value) {
  if (value instanceof Date) {
    const y = value.getFullYear();
    const m = value.getMonth() + 1;
    const d = value.getDate();
    return `${y}-${pad2_(m)}-${pad2_(d)}`;
  }
  return String(value);
}

function pad2_(n) {
  return n < 10 ? `0${n}` : String(n);
}

function allocateSurveySchedule() {
  const responses = latestResponsesByRoom_();
  if (responses.length === 0) {
    SpreadsheetApp.getUi().alert("回答がまだありません。");
    return;
  }

  const takenSlots = new Set(); // "YYYY-MM-DD|H"
  const assigned = []; // {room, name, date, hour, choiceRank}
  const unassigned = []; // {room, name, choices}

  responses.forEach((r) => {
    let placed = false;
    for (let rank = 0; rank < r.choices.length; rank++) {
      const c = r.choices[rank];
      if (!c.date || !c.hour) continue;
      const key = `${c.date}|${c.hour}`;
      if (!takenSlots.has(key)) {
        takenSlots.add(key);
        assigned.push({ room: r.room, name: r.name, date: c.date, hour: c.hour, choiceRank: rank + 1 });
        placed = true;
        break;
      }
    }
    if (!placed) {
      unassigned.push(r);
    }
  });

  writeAllocationResult_(assigned, unassigned);
}

function writeAllocationResult_(assigned, unassigned) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SURVEY_RESULT_SHEET);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(SURVEY_RESULT_SHEET);

  const totalCols = 1 + SURVEY_TIME_SLOTS.length;

  sheet.getRange(1, 1, 1, totalCols).merge()
    .setValue("住戸内工事日程表（自動振り分け結果）")
    .setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center");

  const headerRow = 2;
  sheet.getRange(headerRow, 1).setValue("日付").setFontWeight("bold");
  SURVEY_TIME_SLOTS.forEach((h, i) => {
    sheet.getRange(headerRow, 2 + i).setValue(`${h}時頃`).setFontWeight("bold").setHorizontalAlignment("center");
  });
  sheet.getRange(headerRow, 1, 1, totalCols).setBorder(true, true, true, true, true, true);

  // 日付ごとにグループ化
  const byDate = {};
  assigned.forEach((a) => {
    if (!byDate[a.date]) byDate[a.date] = {};
    byDate[a.date][a.hour] = a;
  });
  const sortedDates = Object.keys(byDate).sort();

  let row = headerRow + 1;
  sortedDates.forEach((date) => {
    const weekday = SURVEY_WEEKDAYS[new Date(date + "T00:00:00").getDay()];
    sheet.getRange(row, 1).setValue(`${date}（${weekday}）`).setFontWeight("bold");
    SURVEY_TIME_SLOTS.forEach((h, i) => {
      const a = byDate[date][h];
      const cell = sheet.getRange(row, 2 + i);
      cell.setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
      if (a) {
        cell.setValue(`${a.room}\n${a.name}` + (a.choiceRank > 1 ? `\n(第${a.choiceRank}希望)` : ""));
        cell.setBackground(a.choiceRank === 1 ? "#FFFF00" : "#FFC000");
        cell.setFontWeight("bold");
      }
    });
    sheet.getRange(row, 1, 1, totalCols).setBorder(true, true, true, true, true, true);
    row += 1;
  });

  sheet.setColumnWidth(1, 130);
  for (let i = 0; i < SURVEY_TIME_SLOTS.length; i++) sheet.setColumnWidth(2 + i, 110);
  for (let r = headerRow + 1; r < row; r++) sheet.setRowHeight(r, 50);

  // 振り分け不可の住戸
  row += 1;
  if (unassigned.length > 0) {
    sheet.getRange(row, 1).setValue("振分不可（要手動調整）").setFontColor("#FF0000").setFontWeight("bold");
    row += 1;
    unassigned.forEach((u) => {
      const wishes = u.choices.map((c, i) => `第${i + 1}希望: ${c.date} ${c.hour}時頃`).join(" / ");
      sheet.getRange(row, 1, 1, totalCols).merge()
        .setValue(`${u.room}（${u.name}） — ${wishes}`)
        .setFontColor("#FF0000");
      row += 1;
    });
  } else {
    sheet.getRange(row, 1).setValue("振分不可の住戸はありません。").setFontColor("#008000").setFontWeight("bold");
    row += 1;
  }

  sheet.setFrozenRows(headerRow);
  SpreadsheetApp.getActiveSpreadsheet().toast("自動振り分けが完了しました。");
}
