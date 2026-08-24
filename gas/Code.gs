/**
 * 秘密客評鑑表 後端（Google Apps Script）
 * 功能：接收表單 JSON → 附件存入 Drive「秘密客證據」資料夾（每筆一個子資料夾）→ Sheet 寫一列含各附件連結。
 * 部署方式見同資料夾 部署說明.md。
 */

// 綁定在回收試算表上的 Apps Script 免填；獨立腳本才要填試算表 ID
const SHEET_ID = "";
const SHEET_NAME = "回收";          // 分頁名稱，不存在會自動建立
const ROOT_FOLDER_NAME = "秘密客證據"; // Drive 根資料夾，不存在會自動建立

const HEADER = ["時間戳","版本","來源","店別","日期","時段","人數","場次","服務員",
  "迎賓(秒)","入座→茶水(分)","茶水→點餐(分)","點餐→首道(分)","首道→最後道(分)","結帳(分)",
  "加權總分","等級","食安違規","上菜紀錄",
  "證據資料夾","發票連結","出餐照連結","環境照連結","錄音連結",
  "客觀摘要","文字摘要","評分原始JSON","食安原始JSON"];

function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const links = saveFiles_(d);
  appendRow_(d, links);
  return ContentService.createTextOutput("ok");
}

function saveFiles_(d) {
  const out = { folder: "", "發票": [], "出餐照": [], "環境照": [], "錄音": [] };
  if (!d.files || !d.files.length) return out;
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const root = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  const stamp = Utilities.formatDate(new Date(), "Asia/Taipei", "yyyyMMdd_HHmmss");
  const store = (d.meta && d.meta.m_store) ? d.meta.m_store : "未填店";
  const sub = root.createFolder(store + "_" + stamp);
  d.files.forEach(function (f, i) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(f.data),
        f.mime || "application/octet-stream",
        f.kind + "_" + (i + 1) + "_" + (f.name || "file"));
      const file = sub.createFile(blob);
      if (out[f.kind]) out[f.kind].push(file.getUrl());
      else out[f.kind] = [file.getUrl()];
    } catch (err) { /* 單檔失敗不擋整筆回收 */ }
  });
  out.folder = sub.getUrl();
  return out;
}

function appendRow_(d, links) {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(HEADER);
  const m = d.meta || {};
  const dishes = (d.dishes || []).map(function (x) {
    return "第" + x.n + "道 " + Math.floor(x.sec / 60) + "分" + (x.sec % 60) + "秒 " + (x.name || "");
  }).join("\n");
  sh.appendRow([
    new Date(), d.version || "", m.m_src || "", m.m_store || "", m.m_date || "", m.m_time || "",
    m.m_party || "", m.m_sess || "", m.m_server || "",
    m.t_greet || "", m.t_tea || "", m.t_order || "", m.t_food || "", m.t_serve || "", m.t_pay || "",
    d.total != null ? d.total : "", d.grade || "", d.veto ? "是" : "否", dishes,
    links.folder,
    (links["發票"] || []).join("\n"), (links["出餐照"] || []).join("\n"),
    (links["環境照"] || []).join("\n"), (links["錄音"] || []).join("\n"),
    m.m_note || "", d.summary || "",
    JSON.stringify(d.scores || {}), JSON.stringify(d.A || {})
  ]);
}
