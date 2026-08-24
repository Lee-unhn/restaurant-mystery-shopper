// 秘密客 · Google Sheet 收集器 v2（附件上傳版）
// 沿用原部署網址：部署 → 管理部署作業 → 編輯 → 新版本
const SHEET_ID = '貼你的SheetID';  // 部署時填入，勿 commit 真實 ID
const MAIN = 'submissions_v2';   // 新分頁起頭，舊 submissions 不動
const RAW = 'raw';
const ROOT_FOLDER_NAME = '秘密客證據';   // Drive 根資料夾，不存在自動建立

const HEADER = [
  '送出時間','版本','秘客來源','店別','日期','時段','人數組成','場次','服務員',
  '迎賓秒','入座→茶水分','茶水→點餐分','點餐→首道分','首道→最後道分','結帳分',
  '食安違規數','食安違規題號',
  'B 接待服務(25%)','C 出餐品質(20%)','D 五感環境(15%)','E VIP(15%·VIP場)','F 款待精神(15%)','G 轉換引導(10%)',
  '加權總分','等級',
  '上菜紀錄','證據件數','證據資料夾','發票連結','出餐照連結','環境照連結','錄音連結',
  '客觀事實摘要','完整摘要'
];
const WT = {B:25, C:20, D:15, E:15, F:15, G:10};

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const main = ss.getSheetByName(MAIN) || ss.insertSheet(MAIN);
    const raw = ss.getSheetByName(RAW) || ss.insertSheet(RAW);

    if (main.getLastRow() === 0) main.appendRow(HEADER);
    if (raw.getLastRow() === 0) raw.appendRow(['ts','店別','raw_json']);

    const m = data.meta || {};
    const A = data.A || {};
    const scores = data.scores || {};

    // ── 附件 → Drive ──
    const links = saveFiles_(data, m);

    const violIds = Object.keys(A).filter(k => A[k] === 'viol');
    const conAvg = (letter) => {
      const vals = Object.keys(scores)
        .filter(k => k.toUpperCase().startsWith(letter))
        .map(k => scores[k])
        .filter(v => typeof v === 'number');
      return vals.length ? vals.reduce((s,x)=>s+x, 0) / vals.length : null;
    };

    const avgs = {};
    ['B','C','D','E','F','G'].forEach(k => avgs[k] = conAvg(k));

    const isVIP = m.m_sess === 'VIP場';   // 前端已改名（舊值 VIP情境場）
    let sum = 0, wsum = 0;
    Object.keys(WT).forEach(k => {
      if (k === 'E' && !isVIP) return;
      if (avgs[k] != null) { sum += avgs[k] * WT[k]; wsum += WT[k]; }
    });
    const score = wsum ? sum / wsum : null;
    const veto = violIds.length > 0;

    let totalCell = '';
    let grade = '';
    if (veto) { totalCell = '一票否決'; grade = '不合格'; }
    else if (score != null) {
      totalCell = score.toFixed(2);
      grade = score >= 4.5 ? '標竿' : score >= 4 ? '達標' : score >= 3.5 ? '待改善' : '不合格';
    }

    const dishes = (data.dishes || []).map(x =>
      '第' + x.n + '道 ' + Math.floor(x.sec/60) + '分' + (x.sec%60) + '秒 ' + (x.name || '')
    ).join('\n');
    const ec = data.evidence || {};
    const eviCnt = Object.keys(ec).filter(k => ec[k]).map(k => k + '×' + ec[k]).join('、');

    main.appendRow([
      data.ts, data.version || '', m.m_src, m.m_store, m.m_date, m.m_time, m.m_party, m.m_sess, m.m_server,
      m.t_greet, m.t_tea, m.t_order, m.t_food, m.t_serve, m.t_pay,
      violIds.length, violIds.join(','),
      avgs.B?.toFixed(2) || '', avgs.C?.toFixed(2) || '', avgs.D?.toFixed(2) || '',
      isVIP ? (avgs.E?.toFixed(2) || '') : 'N/A',
      avgs.F?.toFixed(2) || '', avgs.G?.toFixed(2) || '',
      totalCell, grade,
      dishes, eviCnt, links.folder,
      links['發票'].join('\n'), links['出餐照'].join('\n'), links['環境照'].join('\n'), links['錄音'].join('\n'),
      m.m_note, data.summary
    ]);

    // raw 不能塞 base64 附件（單格 5 萬字上限會炸），剝掉 files 只留件數
    const rawData = Object.assign({}, data);
    delete rawData.files;
    rawData.files_saved = links.folder ? Object.keys(links).filter(k => k !== 'folder')
      .map(k => k + ':' + links[k].length).join(' ') : '無附件';
    raw.appendRow([data.ts, m.m_store, JSON.stringify(rawData)]);

    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function saveFiles_(data, m) {
  const out = { folder: '', '發票': [], '出餐照': [], '環境照': [], '錄音': [] };
  if (!data.files || !data.files.length) return out;
  const it = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  const root = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
  const stamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd_HHmmss');
  const sub = root.createFolder((m.m_store || '未填店') + '_' + stamp);
  data.files.forEach(function (f, i) {
    try {
      const blob = Utilities.newBlob(
        Utilities.base64Decode(f.data),
        f.mime || 'application/octet-stream',
        f.kind + '_' + (i + 1) + '_' + (f.name || 'file'));
      const file = sub.createFile(blob);
      if (out[f.kind]) out[f.kind].push(file.getUrl());
      else out[f.kind] = [file.getUrl()];
    } catch (err) { /* 單檔失敗不擋整筆回收 */ }
  });
  out.folder = sub.getUrl();
  return out;
}

function doGet() {
  return ContentService.createTextOutput('OK · POST only');
}
