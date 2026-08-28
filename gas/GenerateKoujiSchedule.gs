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
 *   2行目: 見出し（部屋番号, 氏名, 電話, 携帯, 日程, 備考, 確認コード, 第1希望, 第2希望, 第3希望, 所有者メール）
 *   3行目以降: データ
 *   （A1が空欄ならセルのメモに「物件名を入力してください」、2行目のA〜F列が空欄なら
 *    見出し文字列とメモを自動で設定する。いずれかのメニューを一度実行すれば反映される。
 *    既に値が入っているセルは上書きしない）
 *     A列: 部屋番号
 *     B列: 氏名（「空室」の場合は空室として表示）
 *     E列: 日程（工程表の作成に使われる「確定日程」。既定では第1希望が入る。
 *          第1希望どおりに施工できない場合は、H〜J列を見ながらここを
 *          手動で書き換えてから「工程表を作成」を実行する）
 *     F列: 備考（「工期外希望」「カメラ付」「受話器」に対応）
 *     H・I・J列: 第1〜第3希望（回答フォームで集めた希望日時。参考情報として保存される）
 *     K列: 所有者メール（任意。住戸に居住していない所有者〈賃貸オーナー等〉宛に
 *          回答リンクをメール送付したい場合のみ入力する）
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
 *   「工程表」>「住戸QRコード一覧表を作成する（社内用）」を実行すると、
 *   全住戸の部屋番号・確認コード・QRコードを1枚の一覧表（スプレッドシート）に
 *   まとめます。配布状況の管理・照合用の社内資料であり、全住戸の確認コードが
 *   1か所にまとまるため住民には配布しないこと。
 *
 * 「設定」シート（無ければ「回答フォームを作成する」実行時に自動作成される）:
 *   回答期限・問い合わせ先・担当者通知メールを1か所で管理する。ここに入力した値は、
 *   フォームの説明文・所有者宛メール・未回答レポートに自動で反映される。
 *   回答期限を過ぎると、日次トリガーによりフォームの回答受付が自動的に停止する
 *   （「未回答状況レポートを表示」で、期限までの残り日数・未回答の部屋も確認できる）。
 *
 * 日程の重複チェック:
 *   複数の部屋が同じ日時（E列）を確定日程として選んでいる場合、住民の回答が
 *   反映されるたびに自動でチェックされ、該当するE列のセルがピンク色で強調表示される
 *   （メモに重複している部屋番号が記載される）。「設定」シートに担当者通知メールを
 *   入力しておくと、重複が発生するたびにそのアドレス宛にメールで通知される。
 *   「工程表」>「日程重複チェックを表示」でいつでも一覧確認・再チェックできる。
 *
 * マザースプレッドシートへのデータ集約（複数建物を横断して閲覧したい場合）:
 *   「設定」シートに、あらかじめ用意した集約用の空のGoogleスプレッドシートのURLを
 *   入力しておくと、住民の回答が反映されるたび（および「★初期セットアップを一括実行」
 *   実行時）に、この建物の部屋データ（確認コードを除く）がそのスプレッドシートの
 *   「集約データ（マザーシート）」シートに自動で書き込まれる。複数の建物で同じ
 *   マザースプレッドシートのURLを設定すれば、1か所で全建物の状況を横断的に閲覧できる。
 *   「工程表」>「マザーデータへ同期する」でいつでも手動同期もできる。
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
    .addItem("★初期セットアップを一括実行", "runInitialSetup")
    .addSeparator()
    .addItem("工程表を作成", "generateKoujiSchedule")
    .addSeparator()
    .addItem("回答フォームを作成する", "createOrUpdateKoujiForm")
    .addItem("各住戸QRコードを作成する", "createPerRoomQrSlips")
    .addItem("住戸QRコード一覧表を作成する（社内用）", "createRoomQrList")
    .addItem("所有者へ回答リンクをメールで送る", "emailAbsenteeOwners")
    .addItem("フォーム回答を再取り込み", "resyncFormResponses")
    .addItem("日程重複チェックを表示", "reportScheduleConflicts")
    .addItem("マザーデータへ同期する", "syncToMotherSheetManual")
    .addSeparator()
    .addItem("未回答状況レポートを表示", "reportUnanswered")
    .addToUi();
}

// 新しい建物で使い始める際のショートカット。
// 「回答フォームを作成する」→「各住戸QRコードを作成する」→
// 「住戸QRコード一覧表を作成する（社内用）」を順番にまとめて実行する。
// 各ステップの完了ダイアログはそのまま表示される（OKを押すと次のステップに進む）。
function runInitialSetup() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rooms = getRoomList(ss);
  if (rooms.length === 0) {
    ui.alert(
      "入居者一覧に部屋番号が見つかりません。先に入居者一覧のA列に部屋番号を入力してから" +
        "実行してください（空室はB列に「空室」と入力）。"
    );
    return;
  }
  createOrUpdateKoujiForm();
  createPerRoomQrSlips();
  createRoomQrList();
  syncToMotherSheet(ss);
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

// 工程表・フォーム取込エラー・QR一覧（社内用）・設定シートは部屋データとして扱わない
function isRoomDataSheet(sheet) {
  const name = sheet.getName();
  return (
    name !== SCHEDULE_SHEET_NAME &&
    name !== ERROR_SHEET_NAME &&
    name !== QR_LIST_SHEET_NAME &&
    name !== SETTINGS_SHEET_NAME
  );
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
  checkAndFlagConflicts(ss);
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

// 操作方法が分からない場合の問い合わせ先（既定値）。
// 実際に使われる値は「設定」シートで上書きできる（getSettings参照）。
const CONTACT_PHONE_LABEL = "施工会社";
const CONTACT_PHONE_NUMBER = "052-269-9100";

function telUriFor(number) {
  return "tel:" + String(number).replace(/[^\d+]/g, "");
}

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
const OWNER_EMAIL_COL = 11; // K列: 住戸に居住していない所有者（賃貸オーナー等）のメールアドレス（任意）
const SETTINGS_SHEET_NAME = "設定";
const CLOSE_FORM_TRIGGER_HANDLER = "closeFormIfDeadlinePassed";
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

    // 入力ガイド：「1行目=物件名、2行目=見出し、3行目以降=部屋データ」という決まった
    // レイアウトを、記入する場所そのものに明記する（行数の少ない新規シートでも設定されるよう、
    // 下の行数チェックより前に置く）。
    if (!sheet.getRange(1, 1).getValue() && !sheet.getRange(1, 1).getNote()) {
      sheet.getRange(1, 1).setNote(
        "ここに物件名（マンション名）を入力してください。\n" +
          "1行目=物件名、2行目=見出し、3行目以降が部屋データです。"
      );
    }
    const roomHeaderWasBlank = !sheet.getRange(2, 1).getValue();
    ["部屋番号", "氏名", "電話", "携帯", "日程", "備考"].forEach((label, i) => {
      const col = i + 1;
      if (!sheet.getRange(2, col).getValue()) {
        sheet.getRange(2, col).setValue(label);
      }
    });
    if (roomHeaderWasBlank && !sheet.getRange(2, 1).getNote()) {
      sheet.getRange(2, 1).setNote("この行は見出し行です。部屋番号のデータは3行目から入力してください。");
    }

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
    if (!sheet.getRange(2, OWNER_EMAIL_COL).getValue()) {
      sheet.getRange(2, OWNER_EMAIL_COL).setValue("所有者メール（任意）");
    }

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

// 回答期限・問い合わせ先を1か所で管理する「設定」シート。無ければ既定値で作成する。
// 既存の設定シート（項目が3行しかない古いバージョン）にも、無い項目だけ追加する。
function ensureSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const isNew = !sheet;
  if (isNew) {
    sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 2)
      .setValues([["設定項目", "値"]])
      .setFontWeight("bold")
      .setBackground(GRAY);
    sheet.setColumnWidth(1, 260);
    sheet.setColumnWidth(2, 220);
  }

  const rows = [
    ["回答期限（例: 2026/08/10）※空欄可", ""],
    ["問い合わせ先（表示名）", CONTACT_PHONE_LABEL],
    ["問い合わせ先電話番号", CONTACT_PHONE_NUMBER],
    ["担当者通知メール（任意・日程重複時に通知）", ""],
    ["マザースプレッドシートのURL（任意・データ集約用）", ""],
    ["工事の種別（例: インターホン・自動火災報知設備工事）", ""],
  ];
  rows.forEach(([label, defaultValue], i) => {
    const row = 2 + i;
    if (!sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(label);
    }
    if (isNew) {
      sheet.getRange(row, 2).setValue(defaultValue);
    }
  });

  return sheet;
}

// GoogleスプレッドシートのURLからスプレッドシートIDを取り出す。
// 既にIDだけが入力されている場合はそのまま返す。
function extractSpreadsheetId(urlOrId) {
  const str = String(urlOrId || "").trim();
  const m = /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/.exec(str);
  return m ? m[1] : str;
}

// 設定シートの値を読み取る。回答期限は未入力ならnull。
function getSettings(ss) {
  const sheet = ensureSettingsSheet(ss);
  const values = sheet.getRange(2, 1, 6, 2).getValues();
  const deadlineRaw = values[0][1];
  const contactLabel = String(values[1][1] || CONTACT_PHONE_LABEL).trim();
  const contactNumber = String(values[2][1] || CONTACT_PHONE_NUMBER).trim();
  const notifyEmail = String(values[3][1] || "").trim();
  const motherSheetId = extractSpreadsheetId(values[4][1]);
  const workType = String(values[5][1] || "").trim();

  let deadline = null;
  if (Object.prototype.toString.call(deadlineRaw) === "[object Date]") {
    deadline = deadlineRaw;
  } else if (deadlineRaw) {
    const parsed = new Date(String(deadlineRaw));
    if (!isNaN(parsed.getTime())) deadline = parsed;
  }

  return { deadline, contactLabel, contactNumber, notifyEmail, motherSheetId, workType };
}

const MOTHER_SHEET_NAME = "集約データ（マザーシート）";
const MOTHER_HEADERS = [
  "物件名", "工事の種別", "部屋番号", "氏名", "電話", "携帯", "確定日程", "備考",
  "第1希望", "第2希望", "第3希望", "最終更新日時",
];

// マザースプレッドシート側に集約データシートが無ければ、見出し付きで作成する。
// 見出し行は毎回最新のMOTHER_HEADERSで上書きする（列を追加した場合も、既存の
// マザーシートを作り直さずに済むようにするため。データ行には影響しない）。
function ensureMotherSheet(motherSs) {
  let sheet = motherSs.getSheetByName(MOTHER_SHEET_NAME);
  if (!sheet) {
    sheet = motherSs.insertSheet(MOTHER_SHEET_NAME);
    sheet.setFrozenRows(1);
  }
  sheet.getRange(1, 1, 1, MOTHER_HEADERS.length)
    .setValues([MOTHER_HEADERS])
    .setFontWeight("bold")
    .setBackground(GRAY);
  return sheet;
}

// この建物（入居者一覧）の現在のデータを、「設定」シートで指定されたマザー
// スプレッドシートに反映する。同じ物件名の既存行はいったん削除してから書き直す
// （洗い替え方式。追記だと再同期のたびに行が際限なく重複するため）。
// 確認コード（G列）は複数建物分が1か所に集まるリスクがあるため一切書き込まない。
// マザーが未設定・開けない場合は何もせずfalseを返す（呼び出し元の処理は止めない）。
function syncToMotherSheet(ss) {
  const settings = getSettings(ss);
  if (!settings.motherSheetId) return false;

  let motherSs;
  try {
    motherSs = SpreadsheetApp.openById(settings.motherSheetId);
  } catch (err) {
    logFormError(ss, `マザースプレッドシートを開けませんでした（${err}）。URL・アクセス権をご確認ください。`);
    return false;
  }

  const buildingName = String(ss.getSheets()[0].getRange(1, 1).getValue() || "");
  const sheet = ensureMotherSheet(motherSs);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const names = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = names.length - 1; i >= 0; i--) {
      if (String(names[i][0]) === buildingName) {
        sheet.deleteRow(i + 2);
      }
    }
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  const rowsToAppend = [];
  ss.getSheets().forEach((s) => {
    if (!isRoomDataSheet(s)) return;
    const last = s.getLastRow();
    if (last < 3) return;
    const values = s.getRange(3, 1, last - 2, PREF_COLS[2]).getValues(); // A〜J列（G列=確認コードは使わない）
    values.forEach((row) => {
      const roomRaw = row[0];
      if (roomRaw === "" || roomRaw === null) return;
      rowsToAppend.push([
        buildingName,
        settings.workType,
        formatRoom(roomRaw),
        row[1], // 氏名
        row[2], // 電話
        row[3], // 携帯
        row[4], // 確定日程
        row[5], // 備考
        row[PREF_COLS[0] - 1], // 第1希望
        row[PREF_COLS[1] - 1], // 第2希望
        row[PREF_COLS[2] - 1], // 第3希望
        now,
      ]);
    });
  });

  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, MOTHER_HEADERS.length).setValues(rowsToAppend);
  }
  return true;
}

// メニュー「マザーデータへ同期する」から実行する。
function syncToMotherSheetManual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  if (!settings.motherSheetId) {
    ui.alert(
      "マザースプレッドシートが未設定です。\n" +
        "「設定」シートの「マザースプレッドシートのURL（任意）」にURLを入力してから、もう一度実行してください。"
    );
    return;
  }
  const ok = syncToMotherSheet(ss);
  if (ok) {
    ui.alert("マザーデータへの同期が完了しました。");
  } else {
    ui.alert(
      "マザースプレッドシートへの同期に失敗しました。「フォーム取込エラー」シートに詳細を記録しました。" +
        "URL・アクセス権をご確認ください。"
    );
  }
}

const CONFLICT_NOTE_PREFIX = "⚠ 日程が重複しています：";
const CONFLICT_BACKGROUND = "#FF9999";

// 入居者一覧のE列（確定日程）を全シート横断でチェックし、同じ日時が複数の部屋に
// 設定されていれば、該当セルをピンク色で強調表示・メモ追記する。
// 前回のチェックで付けた強調は、このチェック時点の最新状態に基づいて一旦クリアしてから
// 付け直すため、解消済みの重複は自動的に元の表示に戻る。
// 戻り値は重複グループの配列（{schedText, entries:[{room, sheet, row}]}）。
function checkAndFlagConflicts(ss) {
  const bySchedule = {};
  const cells = [];

  ss.getSheets().forEach((sheet) => {
    if (!isRoomDataSheet(sheet)) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;
    const values = sheet.getRange(3, 1, lastRow - 2, 5).getValues();
    values.forEach((row, i) => {
      const roomRaw = row[0];
      const name = row[1];
      if (roomRaw === "" || roomRaw === null) return;
      if (String(name || "") === "空室") return;
      const rowNum = i + 3;
      cells.push({ sheet, row: rowNum });

      const schedText = String(row[4] || "").trim();
      if (!schedText) return;
      const room = formatRoom(roomRaw);
      if (!bySchedule[schedText]) bySchedule[schedText] = [];
      bySchedule[schedText].push({ room, sheet, row: rowNum });
    });
  });

  cells.forEach(({ sheet, row }) => {
    const cell = sheet.getRange(row, 5);
    if (String(cell.getNote() || "").indexOf(CONFLICT_NOTE_PREFIX) === 0) {
      cell.setNote("");
    }
    cell.setBackground(null);
  });

  const conflicts = Object.keys(bySchedule)
    .map((schedText) => ({ schedText, entries: bySchedule[schedText] }))
    .filter((c) => c.entries.length > 1);

  conflicts.forEach((c) => {
    const roomList = c.entries.map((e) => e.room).join("、");
    c.entries.forEach(({ sheet, row }) => {
      const cell = sheet.getRange(row, 5);
      cell.setBackground(CONFLICT_BACKGROUND);
      cell.setNote(`${CONFLICT_NOTE_PREFIX}${roomList}（${c.schedText}）`);
    });
  });

  return conflicts;
}

// メニュー「日程重複チェックを表示」から実行する。
function reportScheduleConflicts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const conflicts = checkAndFlagConflicts(ss);

  if (conflicts.length === 0) {
    ui.alert("日程の重複チェック", "重複している日程はありません。", ui.ButtonSet.OK);
    return;
  }

  const message = conflicts
    .map((c) => `${c.schedText}：${c.entries.map((e) => e.room).join("、")}`)
    .join("\n");
  ui.alert(
    "日程の重複チェック",
    "以下の日時が複数の部屋で重複しています（入居者一覧のE列がピンク色で表示されます）。\n" +
      "「工程表を作成」の前に、いずれかの部屋のE列を第2・第3希望などに書き換えてください。\n\n" +
      message,
    ui.ButtonSet.OK
  );
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
    ScriptApp.newTrigger("onKoujiFormSubmit").forForm(form).onFormSubmit().create();
  }

  // 「設定」シートの回答期限を過ぎたら自動でフォームを締め切るための日次トリガーを
  // （まだ無ければ）用意する。判定・締切処理自体は closeFormIfDeadlinePassed が行う。
  const hasCloseTrigger = ScriptApp.getProjectTriggers().some(
    (t) => t.getHandlerFunction() === CLOSE_FORM_TRIGGER_HANDLER
  );
  if (!hasCloseTrigger) {
    ScriptApp.newTrigger(CLOSE_FORM_TRIGGER_HANDLER).timeBased().everyDays(1).atHour(0).create();
  }

  // タイトル・説明文は毎回更新する（入居者一覧のA1（物件名）を後から修正した場合や、
  // 問い合わせ先・回答期限の変更を、既存フォームにも反映するため）
  form.setTitle(`${buildingName} 工事日程アンケート`);
  const settings = getSettings(ss);
  const contactText = `${settings.contactLabel}：${settings.contactNumber}`;
  const deadlineText = settings.deadline
    ? `回答期限：${Utilities.formatDate(settings.deadline, Session.getScriptTimeZone(), "yyyy年M月d日")}まで\n\n`
    : "";
  form.setDescription(
    "工事の希望日時を第1希望〜第3希望までご回答ください（第2・第3希望は任意です）。\n" +
      "同じ部屋番号で再度回答すると、内容は最新の回答で上書きされます。\n\n" +
      deadlineText +
      `スマートフォンの操作でご不明な点がございましたら、${contactText}までお電話ください。` +
      "ご本人以外（ご家族など）が代わりにご回答いただいても構いません。"
  );

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

// 日次トリガーから呼ばれる。「設定」シートの回答期限を過ぎていたら、
// フォームの新規回答受付を自動的に停止する。UIを持たないため通知はしない。
function closeFormIfDeadlinePassed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  if (!settings.deadline) return;

  const props = PropertiesService.getDocumentProperties();
  const formId = props.getProperty(FORM_ID_PROP);
  if (!formId) return;

  // 締切日の終わり（23:59:59）までは受け付ける
  const deadlineEnd = new Date(
    settings.deadline.getFullYear(),
    settings.deadline.getMonth(),
    settings.deadline.getDate() + 1
  );
  if (new Date() < deadlineEnd) return;

  const form = FormApp.openById(formId);
  if (form.isAcceptingResponses()) {
    form.setAcceptingResponses(false);
    form.setCustomClosedFormMessage(
      "回答期限を過ぎたため、このフォームでの受付は終了しました。" +
        `ご不明な点は${settings.contactLabel}：${settings.contactNumber}までお問い合わせください。`
    );
  }
}

function weekdayJP(year, month, day) {
  const dow = new Date(year, month - 1, day).getDay();
  return WEEKDAYS[(dow + 6) % 7];
}

// フォームの「日付」質問の回答から年月日を取り出す。
// DateItemの回答はDateオブジェクトで返る場合と "yyyy-MM-dd" 文字列で返る場合の
// 両方があるため、どちらでも解釈できるようにする。
function parseDateAnswer(dateAnswer) {
  if (!dateAnswer) return null;
  if (Object.prototype.toString.call(dateAnswer) === "[object Date]") {
    return { y: dateAnswer.getFullYear(), m: dateAnswer.getMonth() + 1, d: dateAnswer.getDate() };
  }
  const m = /(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(dateAnswer));
  if (!m) return null;
  return { y: Number(m[1]), m: Number(m[2]), d: Number(m[3]) };
}

// フォームの日付・時間の回答を "m/d（weekday）h:mm" 形式にまとめる。
// どちらか未回答なら null（第2・第3希望は任意のため、未回答もあり得る）。
function buildSchedText(answers, dateTitle, timeTitle) {
  const date = parseDateAnswer(answers[dateTitle]);
  const timeStr = String(answers[timeTitle] || "");
  const [hour, minute] = timeStr.split(":");
  if (!date || !hour) return null;
  const weekday = weekdayJP(date.y, date.m, date.d);
  return `${date.m}/${date.d}（${weekday}）${hour}:${minute}`;
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

  // 他の部屋と確定日程（E列）が重複していないか確認し、重複していればセルを
  // 強調表示した上で、「設定」シートに担当者通知メールが設定されていれば通知する。
  const conflicts = checkAndFlagConflicts(ss);
  const myConflict = conflicts.find((c) => c.entries.some((entry) => entry.room === room));
  if (myConflict) {
    const settings = getSettings(ss);
    if (settings.notifyEmail) {
      try {
        MailApp.sendEmail({
          to: settings.notifyEmail,
          subject: `【日程重複】${myConflict.schedText} に複数の部屋が重複しています`,
          body:
            `以下の部屋が同じ日時（${myConflict.schedText}）で回答しています。\n` +
            `該当部屋：${myConflict.entries.map((e) => e.room).join("、")}\n\n` +
            "入居者一覧のE列（該当セルはピンク色で表示されます）を確認し、" +
            "いずれかの部屋の日程を第2・第3希望などに調整してください。",
        });
      } catch (err) {
        logFormError(ss, `日程重複の通知メール送信に失敗しました（${err}）。`);
      }
    }
  }

  // 「設定」シートにマザースプレッドシートが指定されていれば、最新の状態を反映する。
  syncToMotherSheet(ss);
}

function onKoujiFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  applyFormResponse(ss, e.response.getItemResponses());
}

// QRコード画像を取得する。1つ目のサービスが失敗した場合は2つ目を試す
// （どちらも失敗した場合はnullを返す）。
function fetchQrImageBlob(dataUrl) {
  const encoded = encodeURIComponent(dataUrl);
  const providers = [
    `https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=${encoded}`,
    `https://quickchart.io/qr?size=400&text=${encoded}`,
  ];
  for (let i = 0; i < providers.length; i++) {
    try {
      const resp = UrlFetchApp.fetch(providers[i], { muteHttpExceptions: true });
      if (resp.getResponseCode() === 200) {
        return resp.getBlob();
      }
    } catch (err) {
      // 次の候補を試す
    }
  }
  return null;
}

// 部屋番号タイトルを縦書き風に変換する（例: "101" -> "1\n0\n1\n号\n室"）。
// Slides APIには本物の縦書き（文字を正立させたまま上から下に流し込む設定）が無いため、
// 1文字ずつ改行して縦長のテキストボックスに配置することで近い見た目にする。
function verticalRoomTitle(room) {
  return Array.from(`${room}号室`).join("\n");
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

  const settings = getSettings(ss);
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

    const titleShape = slide.insertTextBox(verticalRoomTitle(room), 615, 20, 70, 260);
    titleShape.getText().getTextStyle().setFontSize(28).setBold(true);
    titleShape.getText().getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.CENTER);
    slide.insertTextBox(
      "📱 スマホでぴっと読み取るだけ！QRコードをかざして、工事の希望日時をご回答ください。\n" +
        `📞 お電話でのご案内をご希望の方は、${settings.contactLabel}（${settings.contactNumber}）までご連絡ください。`,
      40, 90, 550, 70
    ).getText().getTextStyle().setFontSize(14);

    const blob = fetchQrImageBlob(url);
    let inserted = false;
    if (blob) {
      slide.insertImage(blob, 150, 180, 250, 250);
      inserted = true;
    }
    if (!inserted) {
      failedCount += 1;
      slide.insertTextBox(
        `（QR画像の取得に失敗しました。下記URLを直接ご案内ください）\n${url}`,
        40, 180, 550, 120
      ).getText().getTextStyle().setFontSize(10);
    }

    slide.insertTextBox(
      `確認コード: ${code}（QRコードに自動で含まれています。手入力は不要です）`,
      40, 450, 550, 30
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

const QR_LIST_SHEET_NAME = "QR一覧（社内用）";

// 全住戸の「部屋番号・確認コード・QRコード」を1枚の一覧表（スプレッドシート）にまとめる。
// 住民への配布物ではなく、社内で配布状況を管理するための一覧。
// 全住戸の確認コードが1か所に載るため、住民には配布しないこと。
function createRoomQrList() {
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

  let sheet = ss.getSheetByName(QR_LIST_SHEET_NAME);
  if (sheet) ss.deleteSheet(sheet);
  sheet = ss.insertSheet(QR_LIST_SHEET_NAME);

  const totalCols = 5;
  sheet.getRange(1, 1, 1, totalCols).merge()
    .setValue("住戸別QRコード・確認コード一覧（社内管理用）")
    .setFontWeight("bold").setFontSize(14)
    .setHorizontalAlignment("center");
  sheet.getRange(2, 1, 1, totalCols).merge()
    .setValue(
      "※全住戸の確認コードが含まれます。住民への配布物には使用せず、社内で厳重に管理してください。" +
        "「回答用URL」は、お電話で回答を代行入力する際などにクリックしてお使いください。"
    )
    .setFontColor("#FF0000").setWrap(true);

  const headerRow = 3;
  ["部屋番号", "確認コード", "QRコード", "回答用URL", "配布チェック"].forEach((label, i) => {
    sheet.getRange(headerRow, i + 1)
      .setValue(label)
      .setFontWeight("bold")
      .setBackground(GRAY)
      .setHorizontalAlignment("center");
  });

  entries.forEach(({ room, code }, i) => {
    const row = headerRow + 1 + i;
    const formResponse = form.createResponse();
    formResponse.withItemResponse(roomItem.asListItem().createResponse(room));
    formResponse.withItemResponse(codeItem.asTextItem().createResponse(code));
    const url = formResponse.toPrefilledUrl();
    const qrUrl = "https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=" + encodeURIComponent(url);

    sheet.getRange(row, 1).setValue(room).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(row, 2).setValue(code).setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.getRange(row, 3).setFormula(`=IMAGE("${qrUrl}")`);
    sheet.getRange(row, 4).setFormula(`=HYPERLINK("${url}","${room}号室の回答フォームを開く")`);
    sheet.getRange(row, 5).insertCheckboxes();
    sheet.setRowHeight(row, 90);
  });

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 220);
  sheet.setColumnWidth(5, 100);
  sheet.setFrozenRows(headerRow);

  ui.alert(
    "住戸QRコード一覧表を作成しました",
    `「${QR_LIST_SHEET_NAME}」シートに、全${entries.length}住戸のQRコード・確認コード・回答用URLを一覧化しました。\n\n` +
      "この一覧表は全住戸の確認コードが1か所にまとまっているため、住民への配布物には使わず、" +
      "社内での配布状況の管理・照合用としてご利用ください。\n" +
      "電話で回答内容を伺い代行入力する場合は、「回答用URL」列のリンクから該当住戸のフォームを開いて入力できます。\n" +
      "住民配布用には「各住戸QRコードを作成する」で作成される、1住戸1ページのスライドをお使いください。",
    ui.ButtonSet.OK
  );
}

// 住戸に居住していない所有者（賃貸オーナー等）向けに、入居者一覧のK列に登録された
// メールアドレス宛てに、その住戸専用の回答リンクを送信する。
// QRコードを掲示物として見る機会がない所有者向けの代替配布手段。
function emailAbsenteeOwners() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();
  const formId = props.getProperty(FORM_ID_PROP);
  if (!formId) {
    ui.alert("先に「回答フォームを作成する」を実行してください。");
    return;
  }

  ensureSheetLayout(ss);

  const form = FormApp.openById(formId);
  const roomItem = findItemByTitle(form, FORM_Q_ROOM);
  const codeItem = findItemByTitle(form, FORM_Q_CODE);
  if (!roomItem || !codeItem) {
    ui.alert("フォームの質問が見つかりません。先に「回答フォームを作成する」を実行してください。");
    return;
  }

  const buildingName = String(ss.getSheets()[0].getRange(1, 1).getValue() || "工事");
  const settings = getSettings(ss);
  const deadlineLine = settings.deadline
    ? `回答期限：${Utilities.formatDate(settings.deadline, Session.getScriptTimeZone(), "yyyy年M月d日")}まで\n\n`
    : "";
  let sentCount = 0;
  const failed = [];

  ss.getSheets().forEach((sheet) => {
    if (!isRoomDataSheet(sheet)) return;
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return;

    const values = sheet.getRange(3, 1, lastRow - 2, OWNER_EMAIL_COL).getValues();
    values.forEach((row) => {
      const roomRaw = row[0];
      const name = row[1];
      const code = row[CODE_COL - 1];
      const ownerEmail = String(row[OWNER_EMAIL_COL - 1] || "").trim();
      if (roomRaw === "" || roomRaw === null) return;
      if (String(name || "") === "空室") return;
      if (!ownerEmail) return;

      const room = formatRoom(roomRaw);
      const formResponse = form.createResponse();
      formResponse.withItemResponse(roomItem.asListItem().createResponse(room));
      formResponse.withItemResponse(codeItem.asTextItem().createResponse(String(code || "")));
      const url = formResponse.toPrefilledUrl();

      const subject = `【${buildingName}】工事日程アンケートのお願い（${room}号室）`;
      // htmlMailClient向け: 電話番号をタップでそのまま発信できるようtel:リンクにする
      const plainBody =
        `${room}号室の所有者様\n\n` +
        "平素より大変お世話になっております。インターホン・自動火災報知設備の取替工事にあたり、" +
        "工事希望日時のご回答をお願いしております。\n\n" +
        deadlineLine +
        "以下のリンクより、工事希望日時（第1〜第3希望。第1希望のみ必須）をご回答ください。\n" +
        `${url}\n\n` +
        `※このリンクは${room}号室専用です。他の住戸の回答にはご利用いただけません。\n` +
        "※実際にお住まいの方がいらっしゃる場合は、そちらの方にご回答いただいても構いません。\n\n" +
        `ご不明な点がございましたら、${settings.contactLabel}：${settings.contactNumber}までお問い合わせください。`;
      const htmlBody =
        `<p>${room}号室の所有者様</p>` +
        "<p>平素より大変お世話になっております。インターホン・自動火災報知設備の取替工事にあたり、" +
        "工事希望日時のご回答をお願いしております。</p>" +
        (settings.deadline
          ? `<p>回答期限：${Utilities.formatDate(settings.deadline, Session.getScriptTimeZone(), "yyyy年M月d日")}まで</p>`
          : "") +
        `<p><a href="${url}">こちらのリンクより工事希望日時（第1〜第3希望。第1希望のみ必須）をご回答ください</a></p>` +
        `<p>※このリンクは${room}号室専用です。他の住戸の回答にはご利用いただけません。<br>` +
        "※実際にお住まいの方がいらっしゃる場合は、そちらの方にご回答いただいても構いません。</p>" +
        `<p>ご不明な点がございましたら、${settings.contactLabel}` +
        `<a href="${telUriFor(settings.contactNumber)}">${settings.contactNumber}</a>までお問い合わせください。</p>`;

      try {
        MailApp.sendEmail({ to: ownerEmail, subject, body: plainBody, htmlBody });
        sentCount += 1;
      } catch (err) {
        failed.push(`${room}（${ownerEmail}）`);
      }
    });
  });

  if (sentCount === 0 && failed.length === 0) {
    ui.alert(
      "送信対象がありません",
      `入居者一覧のK列（${OWNER_EMAIL_COL}列目・見出し「所有者メール（任意）」）に、` +
        "住戸に居住していない所有者のメールアドレスを入力してから実行してください。",
      ui.ButtonSet.OK
    );
    return;
  }

  ui.alert(
    "所有者へのメール送信が完了しました",
    `${sentCount}件送信しました。` +
      (failed.length > 0 ? `\n\n送信に失敗した住戸: ${failed.join("、")}` : ""),
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

// 「設定」シートの回答期限を基準に、未回答の部屋数・一覧を表示する。
function reportUnanswered() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  const { notSubmitted } = collectRoomData(ss);

  let deadlineText = "回答期限：未設定（「設定」シートに入力すると表示されます）";
  if (settings.deadline) {
    const deadlineStr = Utilities.formatDate(settings.deadline, Session.getScriptTimeZone(), "yyyy年M月d日");
    const today = new Date();
    const todayMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const deadlineMidnight = new Date(
      settings.deadline.getFullYear(),
      settings.deadline.getMonth(),
      settings.deadline.getDate()
    );
    const diffDays = Math.round((deadlineMidnight - todayMidnight) / (1000 * 60 * 60 * 24));
    if (diffDays > 0) {
      deadlineText = `回答期限：${deadlineStr}（あと${diffDays}日）`;
    } else if (diffDays === 0) {
      deadlineText = `回答期限：${deadlineStr}（本日締切）`;
    } else {
      deadlineText = `回答期限：${deadlineStr}（${-diffDays}日超過）`;
    }
  }

  const message =
    `${deadlineText}\n\n` +
    `未回答：${notSubmitted.length}件\n` +
    (notSubmitted.length > 0 ? notSubmitted.join("、") : "（すべての部屋が回答済みです）");

  ui.alert("未回答状況レポート", message, ui.ButtonSet.OK);
}
