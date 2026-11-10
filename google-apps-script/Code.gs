/**
 * 奇美醫院 SEL 學習小棧｜12月聖誕感恩活動 — 留言後端（Google Apps Script）
 *
 * 用途：把「願意分享」的感謝留言存進這份 Google 試算表，讓所有人都看得到。
 * 設定步驟請看同資料夾的「設定教學.md」。
 *
 * 管理方式：
 *  - 試算表「感恩留言」分頁中，取消勾選「顯示」欄 → 該留言立刻從大家的聖誕樹上隱藏
 *  - 想要「先審核、再公開」：把下方 NEED_APPROVAL 改成 true，並重新部署（新版本）
 *  - 「溫暖數」欄（H 欄）是大家按「溫暖 +1」的累計，可直接查看
 */

const SHEET_NAME = '感恩留言';
const NEED_APPROVAL = false;   // true = 新留言預設不顯示，需管理者勾選「顯示」後才公開
const MAX_TEXT = 200;
const MAX_NAME = 20;
const KINDS = ['person', 'event', 'thing'];
const RATE_LIMIT_PER_MIN = 30; // 簡易防洗版：全站每分鐘最多新增幾則

// 欄位順序：A id | B 時間 | C 類別 | D 內容 | E 署名 | F 顯示 | G token雜湊 | H 溫暖數
const COL = { ID: 0, TIME: 1, KIND: 2, TEXT: 3, NAME: 4, SHOW: 5, HASH: 6, LIKES: 7 };
const LIKE_LIMIT_PER_MIN = 600; // 簡易防灌：全站每分鐘最多幾次「溫暖 +1」

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', '時間', '類別(person人/event事/thing物)', '內容', '署名', '顯示', 'token雜湊(勿改)', '溫暖數']);
    sh.setFrozenRows(1);
    sh.getRange('A1:H1').setFontWeight('bold').setBackground('#FEE8D5');
    sh.setColumnWidth(4, 420);
    sh.hideColumns(7);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function hash_(s) {
  return Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8)
  );
}

// 清理文字：去掉控制字元、限制長度，並避免被試算表當成公式執行
function clean_(s, max) {
  let v = String(s == null ? '' : s).replace(/[\u0000-\u0009\u000B-\u001F\u007F]/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(v)) v = "'" + v;
  return v;
}

function isShown_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

/** GET：讀取所有「顯示」的留言 */
function doGet() {
  const values = getSheet_().getDataRange().getValues().slice(1);
  const items = values
    .filter(function (r) { return r[COL.ID] && isShown_(r[COL.SHOW]); })
    .map(function (r) {
      return {
        id: String(r[COL.ID]),
        ts: new Date(r[COL.TIME]).getTime() || 0,
        k: KINDS.indexOf(String(r[COL.KIND])) > -1 ? String(r[COL.KIND]) : 'person',
        t: String(r[COL.TEXT]),
        n: String(r[COL.NAME] || ''),
        l: Math.max(0, Number(r[COL.LIKES]) || 0)
      };
    });
  return json_({ ok: true, items: items });
}

/** POST：新增 / 刪除自己的留言 */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return json_({ ok: false, error: 'busy' });
  try {
    const sh = getSheet_();

    if (d.action === 'add') {
      const id = clean_(d.id, 40).replace(/[^0-9a-z]/gi, '');
      const token = String(d.token || '');
      const k = KINDS.indexOf(d.k) > -1 ? d.k : 'person';
      const t = clean_(d.t, MAX_TEXT);
      const n = clean_(d.n, MAX_NAME);
      if (!id || !t || token.length < 16) return json_({ ok: false, error: 'invalid' });

      const cache = CacheService.getScriptCache();
      const count = Number(cache.get('rate') || 0);
      if (count >= RATE_LIMIT_PER_MIN) return json_({ ok: false, error: 'rate_limited' });
      cache.put('rate', String(count + 1), 60);

      const now = new Date();
      const visible = !NEED_APPROVAL;
      sh.appendRow([id, now, k, t, n, visible, hash_(token), 0]);
      sh.getRange(sh.getLastRow(), COL.SHOW + 1).insertCheckboxes();
      if (visible) sh.getRange(sh.getLastRow(), COL.SHOW + 1).check();

      return json_({ ok: true, visible: visible, item: { id: id, ts: now.getTime(), k: k, t: t.replace(/^'/, ''), n: n.replace(/^'/, '') } });
    }

    if (d.action === 'like') {
      const delta = d.delta === -1 ? -1 : 1;          // 一次只能 +1 或 -1
      const cache = CacheService.getScriptCache();
      const count = Number(cache.get('likerate') || 0);
      if (count >= LIKE_LIMIT_PER_MIN) return json_({ ok: false, error: 'rate_limited' });
      cache.put('likerate', String(count + 1), 60);
      const values = sh.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][COL.ID]) === String(d.id)) {
          if (!isShown_(values[i][COL.SHOW])) return json_({ ok: false, error: 'not_found' });
          const l = Math.max(0, (Number(values[i][COL.LIKES]) || 0) + delta);
          sh.getRange(i + 1, COL.LIKES + 1).setValue(l);
          return json_({ ok: true, l: l });
        }
      }
      return json_({ ok: false, error: 'not_found' });
    }

    if (d.action === 'delete') {
      const values = sh.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][COL.ID]) === String(d.id)) {
          if (values[i][COL.HASH] !== hash_(d.token || '')) return json_({ ok: false, error: 'forbidden' });
          sh.deleteRow(i + 1);
          return json_({ ok: true });
        }
      }
      return json_({ ok: false, error: 'not_found' });
    }

    return json_({ ok: false, error: 'unknown_action' });
  } finally {
    lock.releaseLock();
  }
}
