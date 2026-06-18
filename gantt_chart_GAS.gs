/**
 * ════════════════════════════════════════════════════════════
 *  甘特圖管理系統 — Google Apps Script 後端
 *  版本：2.0（分享頁面與編輯器視覺一致）
 * ════════════════════════════════════════════════════════════
 */

// ── 必須修改的設定 ──────────────────────────────────────────
const CONFIG = {
  SHEET_ID: '1xizWmZ8KGJMqYJpcnkIJ-10WMIyE7-rHodlwh-UYCw8',
  ROOT_FOLDER_ID: '1PceXpD4lT1Xqy_vtZTqELghY4qvoecx-',
  DEPLOY_URL: 'https://script.google.com/macros/s/AKfycbzhEfQW7CkeaALVQk-g2lbGly3zf-6tcsy_tGzDseJOetv7TUdWL2uAaWIi52DZFBgI/exec',
};
// ──────────────────────────────────────────────────────────

function corsHeaders() {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
  };
}

function doOptions() {
  return ContentService.createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const shareId = e.parameter.share || '';

    if (shareId) {
      return handleShareView(shareId);
    }

    if (e.parameter.data) {
      try {
        const bytes = Utilities.base64Decode(e.parameter.data);
        const decoded = Utilities.newBlob(bytes).getDataAsString('UTF-8');
        const body = JSON.parse(decoded);
        switch (body.action) {
          case 'saveProjects':  return respond(handleSaveProjects(body));
          case 'createShare':   return respond(handleCreateShare(body));
          default:              return respond({ ok: false, error: 'Unknown action in data' });
        }
      } catch(decodeErr) {
        return respond({ ok: false, error: '資料解碼失敗: ' + decodeErr.message });
      }
    }

    switch (action) {
      case 'login':         return respond(handleLogin(e.parameter));
      case 'loadProjects':  return respond(handleLoadProjects(e.parameter));
      case 'getShare':      return respond(handleGetShare(e.parameter));
      default:              return respond({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    let body;
    try {
      body = JSON.parse(e.postData.contents);
    } catch(parseErr) {
      return respond({ ok: false, error: '無法解析請求內容: ' + parseErr.message });
    }
    const action = body.action || '';
    switch (action) {
      case 'saveProjects':  return respond(handleSaveProjects(body));
      case 'createShare':   return respond(handleCreateShare(body));
      default:              return respond({ ok: false, error: 'Unknown POST action: ' + action });
    }
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════
//  LOGIN
// ════════════════════════════════════════════════════════════
function handleLogin(params) {
  const { username, password } = params;
  if (!username || !password) return { ok: false, error: '請輸入帳號和密碼' };

  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName('users');
  if (!sheet) return { ok: false, error: '找不到帳號資料表（確認 Sheet 名稱為 users）' };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const [uname, pass, displayName, color, folderId] = data[i];
    if (uname === username && pass === password) {
      const userFolderId = ensureUserFolder(username, folderId, sheet, i + 1);
      return {
        ok: true,
        user: { username, displayName: displayName || username, color: color || '#4F6BED', folderId: userFolderId }
      };
    }
  }
  return { ok: false, error: '帳號或密碼錯誤' };
}

function ensureUserFolder(username, existingFolderId, sheet, rowIndex) {
  if (existingFolderId) {
    try {
      DriveApp.getFolderById(existingFolderId);
      return existingFolderId;
    } catch (e) {}
  }
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const newFolder = root.createFolder('user_' + username);
  const newId = newFolder.getId();
  sheet.getRange(rowIndex, 5).setValue(newId);
  return newId;
}

// ════════════════════════════════════════════════════════════
//  LOAD PROJECTS
// ════════════════════════════════════════════════════════════
function handleLoadProjects(params) {
  const { username, folderId } = params;
  if (!username || !folderId) return { ok: false, error: '缺少參數' };

  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFilesByType(MimeType.PLAIN_TEXT);
  const projects = [];

  while (files.hasNext()) {
    const file = files.next();
    if (!file.getName().endsWith('.gantt.json')) continue;
    try {
      const content = file.getBlob().getDataAsString('UTF-8');
      const proj = JSON.parse(content);
      proj._fileId = file.getId();
      projects.push(proj);
    } catch (e) {}
  }

  return { ok: true, projects };
}

// ════════════════════════════════════════════════════════════
//  SAVE PROJECTS
// ════════════════════════════════════════════════════════════
function handleSaveProjects(body) {
  const { username, folderId, projects } = body;
  if (!username || !folderId || !projects) return { ok: false, error: '缺少參數' };

  const folder = DriveApp.getFolderById(folderId);
  let saved = 0;

  projects.forEach(proj => {
    const filename = 'proj_' + proj.id + '.gantt.json';
    const content = JSON.stringify(proj, null, 2);

    if (proj._fileId) {
      try {
        const file = DriveApp.getFileById(proj._fileId);
        file.setContent(content);
        saved++;
        return;
      } catch (e) {}
    }

    const existingFiles = folder.getFilesByName(filename);
    if (existingFiles.hasNext()) {
      existingFiles.next().setContent(content);
    } else {
      folder.createFile(filename, content, MimeType.PLAIN_TEXT);
    }
    saved++;
  });

  return { ok: true, saved };
}

// ════════════════════════════════════════════════════════════
//  CREATE SHARE LINK
// ════════════════════════════════════════════════════════════
function handleCreateShare(body) {
  const { username, folderId, projectId, projectData, shareType, taskDetailHiddenIds, statsCollapsed } = body;
  if (!projectId || !projectData) return { ok: false, error: '缺少專案資料' };

  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  let sharesFolder;
  const existing = root.getFoldersByName('_shares');
  sharesFolder = existing.hasNext() ? existing.next() : root.createFolder('_shares');

  const shareId = 'share_' + projectId + '_' + Date.now();
  const shareData = {
    shareId,
    shareType: shareType || 'interactive',
    createdAt: new Date().toISOString(),
    createdBy: username || 'guest',
    project: projectData,
    taskDetailHiddenIds: taskDetailHiddenIds || [],
    statsCollapsed: statsCollapsed || false,
  };

  sharesFolder.createFile(shareId + '.json', JSON.stringify(shareData), MimeType.PLAIN_TEXT);

  const shareUrl = CONFIG.DEPLOY_URL + '?share=' + shareId;
  return { ok: true, shareId, shareUrl };
}

// ════════════════════════════════════════════════════════════
//  GET SHARE DATA
// ════════════════════════════════════════════════════════════
function handleGetShare(params) {
  const { shareId } = params;
  if (!shareId) return { ok: false, error: '缺少 shareId' };

  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const existing = root.getFoldersByName('_shares');
  if (!existing.hasNext()) return { ok: false, error: '找不到分享資料夾' };

  const sharesFolder = existing.next();
  const files = sharesFolder.getFilesByName(shareId + '.json');
  if (!files.hasNext()) return { ok: false, error: '找不到此分享連結，可能已過期' };

  const content = files.next().getBlob().getDataAsString('UTF-8');
  return { ok: true, data: JSON.parse(content) };
}

// ════════════════════════════════════════════════════════════
//  SHARE VIEW — 回傳完整的唯讀 HTML 頁面
// ════════════════════════════════════════════════════════════
function handleShareView(shareId) {
  let shareData;
  try {
    const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
    const sharesFolder = root.getFoldersByName('_shares').next();
    const file = sharesFolder.getFilesByName(shareId + '.json').next();
    shareData = JSON.parse(file.getBlob().getDataAsString('UTF-8'));
  } catch (e) {
    return HtmlService.createHtmlOutput('<h2 style="font-family:sans-serif;padding:40px;color:#DC2626">找不到此分享連結，可能已過期或被刪除。</h2>');
  }

  const proj = shareData.project;
  const createdAt = shareData.createdAt ? shareData.createdAt.slice(0, 10) : '';
  const createdBy = shareData.createdBy || '';
  const taskDetailHiddenIds = shareData.taskDetailHiddenIds || [];
  const statsCollapsed = shareData.statsCollapsed || false;

  const html = buildShareHtml(proj, createdAt, createdBy, shareId, taskDetailHiddenIds, statsCollapsed);
  return HtmlService.createHtmlOutput(html)
    .setTitle(proj.name + ' — 甘特圖分享')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ════════════════════════════════════════════════════════════
//  BUILD SHARE HTML — 與編輯器視覺一致的唯讀頁面
// ════════════════════════════════════════════════════════════
function buildShareHtml(proj, createdAt, createdBy, shareId, taskDetailHiddenIds, statsCollapsed) {
  const projJson = JSON.stringify(proj);
  const detailHiddenJson = JSON.stringify(taskDetailHiddenIds || []);
  const createdByJson = JSON.stringify(createdBy || '');
  const statsCollapsedVal = (statsCollapsed ? 'true' : 'false');

  const css = `
:root{
  --acc:#4F6BED;--acc-d:#3451C7;--acc-l:rgba(79,107,237,.1);
  --done:#059669;--done-l:rgba(5,150,105,.12);
  --del:#DC2626;--del-l:rgba(220,38,38,.1);
  --pend:#6B7280;--pend-l:rgba(107,114,128,.1);
  --bg:#F0F2F8;--surf:#fff;--bdr:rgba(0,0,0,.08);
  --tx:#111827;--tx2:#6B7280;--tx3:#9CA3AF;
  --r:8px;--rl:12px;--bar-h:26px;
}
*{box-sizing:border-box;margin:0;padding:0}
html,body{font-family:'DM Sans',sans-serif;background:var(--bg);color:var(--tx);font-size:14px;height:100%}
body{display:flex;flex-direction:column;overflow:hidden}

/* ── Topbar ── */
.topbar{background:#1E2235;padding:0;display:flex;flex-direction:column;flex-shrink:0}
.topbar-row1{display:flex;align-items:center;justify-content:space-between;padding:9px 20px;gap:12px;border-bottom:1px solid rgba(255,255,255,.08);flex-wrap:wrap}
.topbar-row2{display:flex;align-items:center;padding:7px 20px;gap:6px;flex-wrap:wrap;background:rgba(255,255,255,.04);border-bottom:1px solid rgba(255,255,255,.06)}
.tl-left{display:flex;align-items:center;gap:10px}
.pdot{width:10px;height:10px;border-radius:50%;flex-shrink:0}
.proj-title{font-size:16px;font-weight:600;color:#fff;cursor:default}
.proj-badge{font-size:11px;padding:3px 8px;border-radius:20px;background:rgba(79,107,237,.3);color:#93C5FD;font-weight:600;border:1px solid rgba(79,107,237,.4)}
.tl-right{display:flex;align-items:center;gap:7px}
.btn{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;font-size:12px;font-family:inherit;font-weight:500;border:1px solid rgba(255,255,255,.15);border-radius:var(--r);background:rgba(255,255,255,.08);color:rgba(255,255,255,.8);cursor:pointer;white-space:nowrap;height:30px;transition:all .12s}
.btn:hover{background:rgba(255,255,255,.15);color:#fff}
.btn.g{background:#10B981;color:#fff;border-color:#10B981}.btn.g:hover{background:#059669}
.btn.o{background:#D97706;color:#fff;border-color:#D97706}.btn.o:hover{background:#B45309}
.pmeta{font-size:11px;color:rgba(255,255,255,.4);margin-left:4px}

/* ── Content area ── */
.main{flex:1;display:flex;flex-direction:column;overflow:hidden;padding:14px 20px 0}

/* ── Stats panel ── */
.stats-wrap{margin-bottom:10px;flex-shrink:0}
.stats-toggle-bar{display:flex;align-items:center;gap:6px;cursor:pointer;user-select:none;padding:4px 0 4px 2px;color:var(--tx2);font-size:11px;font-weight:600}
.stats-toggle-bar:hover{color:var(--tx)}
.stats-toggle-bar .sti{font-size:10px;transition:transform .2s;display:inline-block}
.stats-toggle-bar.closed .sti{transform:rotate(-90deg)}
.stats-body{overflow:hidden;transition:max-height .25s ease,opacity .2s ease;max-height:200px;opacity:1}
.stats-body.closed{max-height:0;opacity:0}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:8px}
.stat{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--rl);padding:10px 12px}
.sl{font-size:10px;color:var(--tx2);text-transform:uppercase;letter-spacing:.5px;margin-bottom:3px}
.sv{font-size:20px;font-weight:600}

/* ── Gantt container ── */
.gc{background:var(--surf);border:1px solid var(--bdr);border-radius:var(--rl);overflow:hidden;display:flex;flex-direction:column;flex:1}
.gantt-head-row{display:flex;flex-shrink:0;border-bottom:2px solid rgba(0,0,0,.12)}
.gantt-head-left{flex-shrink:0;overflow:hidden;border-right:2px solid rgba(0,0,0,.18);background:#F3F4F6}
.gantt-head-right{flex:1;overflow:hidden;background:#F3F4F6}
.gantt-body-row{display:flex;flex:1;overflow:hidden}
.gantt-body-left{flex-shrink:0;overflow:hidden;border-right:2px solid rgba(0,0,0,.18);background:var(--surf)}
.gantt-body-right{flex:1;overflow:auto}
.gantt-table-head-left,.gantt-table-body-left{border-collapse:collapse;width:100%}
.gantt-table-head-right,.gantt-table-body-right{border-collapse:collapse}
.gantt-table-head-left th,.gantt-table-head-right th{background:#F3F4F6;font-size:10px;font-weight:600;padding:5px 8px;border-bottom:1px solid rgba(0,0,0,.08);white-space:nowrap;text-transform:uppercase;letter-spacing:.4px}
.gantt-table-head-right th.dc{text-align:center;padding:4px 1px;min-width:28px;width:28px;font-size:9px;font-weight:500}
.lth{text-align:left;padding:9px 12px;background:#F3F4F6;vertical-align:middle}
.mo-0{background:#EEF2FF;color:#3730A3}.mo-1{background:#F0FDF4;color:#166534}.mo-2{background:#FEF9C3;color:#854D0E}.mo-3{background:#FFF1F2;color:#9F1239}.mo-4{background:#ECFEFF;color:#155E75}.mo-5{background:#FFF7ED;color:#9A3412}.mo-6{background:#F5F3FF;color:#5B21B6}.mo-7{background:#FAFAFA;color:#374151}.mo-8{background:#F0FDFA;color:#115E59}.mo-9{background:#FDF2F8;color:#831843}.mo-10{background:#FFFBEB;color:#78350F}.mo-11{background:#EFF6FF;color:#1E40AF}

/* ── Table body rows ── */
.gantt-table-body-left tbody tr,.gantt-table-body-right tbody tr{height:62px}
.gantt-table-body-left tbody tr.compact,.gantt-table-body-right tbody tr.compact{height:34px}
.gantt-table-body-left tbody tr.compact .bw,.gantt-table-body-right tbody tr.compact .bw{height:34px}
.gantt-table-body-left tbody td,.gantt-table-body-right tbody td{border-bottom:1px solid rgba(0,0,0,.05);vertical-align:middle}
.gantt-table-body-left tbody tr:last-child td,.gantt-table-body-right tbody tr:last-child td{border-bottom:none}

/* ── Task name cell ── */
.nc{padding:7px 12px;background:#fff;vertical-align:middle}
.nc.critical-task{border-left:3px solid #DC2626}
.gantt-table-body-left tbody tr:hover .nc{background:#F0F4FF}
.gantt-table-body-right tbody tr:hover td{background:rgba(79,107,237,.025)}
.ni{display:flex;align-items:flex-start;gap:7px}
.nt{flex:1;min-width:0}
.nn{font-size:13px;font-weight:500;line-height:1.4;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden;word-break:break-all}
.nm{font-size:11px;color:var(--tx2);margin-top:3px;display:flex;align-items:center;gap:5px;flex-wrap:wrap}
.nm.nm-hidden{display:none}
.day-chip{display:inline-flex;align-items:center;background:#F3F4F6;color:var(--tx2);font-size:10px;font-weight:600;padding:1px 6px;border-radius:10px;white-space:nowrap;font-family:'DM Mono',monospace}
.badge{font-size:10px;padding:2px 7px;border-radius:10px;font-weight:600;flex-shrink:0}
.bd{background:var(--done-l);color:#065F46}.bo{background:var(--acc-l);color:var(--acc-d)}
.bp{background:var(--pend-l);color:#374151}.bl{background:var(--del-l);color:#991B1B}
.detail-tog{display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;font-size:9px;color:var(--tx3);cursor:pointer;flex-shrink:0;opacity:.6;border-radius:3px;border:none;background:transparent;font-family:inherit}
.detail-tog:hover{opacity:1;background:rgba(0,0,0,.06)}
.group-toggle{cursor:pointer;font-size:12px;color:var(--acc);margin-right:4px;user-select:none}

/* ── Bar ── */
.bc{padding:2px 0;position:relative;min-width:28px;width:28px}
.bw{position:relative;height:62px;display:flex;align-items:center}
.bar{height:var(--bar-h);border-radius:5px;position:absolute;display:flex;align-items:center;padding:0 8px;font-size:11px;font-weight:600;white-space:nowrap;overflow:hidden;color:#fff;box-shadow:0 1px 3px rgba(0,0,0,.14)}
.bar.critical-bar{box-shadow:0 0 0 2px #DC2626,0 1px 3px rgba(0,0,0,.14)!important}
.bar.milestone-bar{clip-path:polygon(50% 0%,100% 50%,50% 100%,0% 50%);border-radius:0;background:#7C3AED!important;width:26px!important;height:26px!important}
.bp2{position:absolute;left:0;top:0;height:100%;background:rgba(255,255,255,.22);border-radius:5px 0 0 5px;pointer-events:none}
.bl2{position:relative;z-index:1;display:flex;align-items:center;gap:5px}
.bd2{font-size:9.5px;opacity:.75;font-family:'DM Mono',monospace;font-weight:400}
.tl{position:absolute;top:0;left:50%;width:2px;height:100%;opacity:.3;pointer-events:none;transform:translateX(-50%)}

/* ── Legend ── */
.leg{display:flex;align-items:center;gap:12px;flex-wrap:wrap;padding:8px 12px;border-top:1px solid var(--bdr);background:#FAFAFA;font-size:11px;color:var(--tx2);flex-shrink:0}
.li{display:flex;align-items:center;gap:4px}
.ld{width:9px;height:9px;border-radius:3px}
.ln{margin-left:auto;font-size:10px;color:var(--tx3)}

/* ── Export overlay ── */
#exportOverlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.6);z-index:500;align-items:center;justify-content:center}
#exportOverlay.open{display:flex}
.export-box{background:#fff;border-radius:12px;padding:28px 36px;text-align:center;min-width:260px}
.export-spinner{width:36px;height:36px;border:3px solid #e5e7eb;border-top-color:#4F6BED;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 14px}
@keyframes spin{to{transform:rotate(360deg)}}`;

  const js = `
var PROJ = ${projJson};
var DETAIL_HIDDEN_INIT = ${detailHiddenJson};
var STATS_COLLAPSED_INIT = ${statsCollapsedVal};
var taskDetailHidden = new Set(DETAIL_HIDDEN_INIT);
var allDetailCollapsed = DETAIL_HIDDEN_INIT.length > 0;
var statsCollapsed = STATS_COLLAPSED_INIT;

var today = new Date(); today.setHours(0,0,0,0);
var BASE_DATE = new Date(2000,0,1);
var COL_W = 28;
var SL = {done:'已完成',ongoing:'進行中',pending:'待開始',delayed:'延遲'};
var SC = {done:'bd',ongoing:'bo',pending:'bp',delayed:'bl'};
var MN = ['1月','2月','3月','4月','5月','6月','7月','8月','9月','10月','11月','12月'];
var MO_BG = ['#EEF2FF','#F0FDF4','#FEF9C3','#FFF1F2','#ECFEFF','#FFF7ED','#F5F3FF','#FAFAFA','#F0FDFA','#FDF2F8','#FFFBEB','#EFF6FF'];
var MO_COL = ['#3730A3','#166534','#854D0E','#9F1239','#155E75','#9A3412','#5B21B6','#374151','#115E59','#831843','#78350F','#1E40AF'];

function addD(d,n){var r=new Date(d);r.setDate(r.getDate()+n);return r;}
function parseD(s){if(!s)return today;var p=s.replace(/\\//g,'-').split('-');return new Date(+p[0],+p[1]-1,+p[2]);}
function fmt(s){var d=parseD(s);return(d.getMonth()+1)+'/'+(d.getDate());}
function ds(d){return d.getFullYear()+'-'+(d.getMonth()+1<10?'0':'')+(d.getMonth()+1)+'-'+(d.getDate()<10?'0':'')+d.getDate();}
function dayCount(s,e){return Math.round((parseD(e)-parseD(s))/86400000)+1;}
function taskBarBg(t){if(t.color)return t.color;var m={done:'#059669',ongoing:'#4F6BED',pending:'#9CA3AF',delayed:'#DC2626'};return m[t.status]||'#9CA3AF';}
function getMoClass(m){return 'mo-'+m;}
function getMoStyle(m){return 'background:'+MO_BG[m]+';color:'+MO_COL[m];}
function getDepIds(deps){if(!deps||!deps.length)return[];return deps.map(function(d){return typeof d==='object'?d.id:d;}).filter(Boolean);}

function getVisibleTasks(tasks){
  var collapsed = {};
  tasks.forEach(function(t){if(t.type==='group'&&t.collapsed)collapsed[t.id]=true;});
  return tasks.filter(function(t){return !t.parentId||!collapsed[t.parentId];});
}

function getTaskLevel(task){return task.parentId?1:0;}

function computeCPM(tasks){
  var result={};
  if(!tasks||!tasks.length)return result;
  tasks.forEach(function(t){
    var es=Math.round((parseD(t.start)-BASE_DATE)/86400000);
    var ef=Math.round((parseD(t.end)-BASE_DATE)/86400000);
    var dur=ef-es+1;
    result[t.id]={es:es,ef:ef,dur:dur,ls:ef,lf:ef,float:0,critical:false};
  });
  var sorted=tasks.slice().reverse();
  sorted.forEach(function(t){
    var r=result[t.id];
    var successors=tasks.filter(function(x){return getDepIds(x.deps||[]).indexOf(t.id)>=0;});
    if(successors.length===0){r.lf=r.ef;}
    else{r.lf=Math.min.apply(null,successors.map(function(s){return result[s.id].ls;}));}
    r.ls=r.lf-r.dur+1;
    r.float=r.ls-r.es;
    r.critical=r.float<=0;
  });
  return result;
}

// ── Stats panel ──
function toggleStatsPanel(){
  statsCollapsed=!statsCollapsed;
  var body=document.getElementById('statsBody');
  var bar=document.getElementById('statsToggleBar');
  var icon=document.getElementById('statsToggleIcon');
  var label=document.getElementById('statsToggleLabel');
  if(statsCollapsed){body.classList.add('closed');bar.classList.add('closed');label.textContent='統計資訊（已收合）';}
  else{body.classList.remove('closed');bar.classList.remove('closed');label.textContent='統計資訊';}
}

// ── Task detail collapse ──
function toggleAllTaskDetail(){
  allDetailCollapsed=!allDetailCollapsed;
  var tasks=PROJ.tasks||[];
  if(allDetailCollapsed){tasks.forEach(function(t){taskDetailHidden.add(t.id);});}
  else{taskDetailHidden.clear();}
  var btn=document.getElementById('btnCollapseDetail');
  if(btn)btn.textContent=allDetailCollapsed?'☰ 展開詳情':'☰ 收合詳情';
  render();
}

function toggleTaskDetail(id,evt){
  if(evt){evt.stopPropagation();}
  if(taskDetailHidden.has(id)){taskDetailHidden.delete(id);}
  else{taskDetailHidden.add(id);}
  render();
}

function toggleGroupCollapse(id){
  var t=PROJ.tasks.find(function(x){return x.id===id;});
  if(!t)return;
  t.collapsed=!t.collapsed;
  render();
}

// ── Row hover sync ──
function rowHover(id,on){
  var rows=document.querySelectorAll('[data-id="'+id+'"]');
  rows.forEach(function(r){
    r.querySelectorAll('td').forEach(function(td){
      td.style.background=on?(td.classList.contains('nc')?'#F0F4FF':'rgba(79,107,237,.025)'):'';
    });
  });
}

// ── Scroll sync ──
function syncScroll(el){
  document.getElementById('gHR').scrollLeft=el.scrollLeft;
  document.getElementById('gBL').scrollTop=el.scrollTop;
}
function syncScrollLeft(el){
  document.getElementById('gBR').scrollTop=el.scrollTop;
}

// ── Main render ──
function render(){
  var allTasks=PROJ.tasks||[];
  var tasks=getVisibleTasks(allTasks);
  var cpm=computeCPM(tasks);

  if(!tasks.length){
    document.getElementById('gbL').innerHTML='<tr><td style="padding:20px;color:#9CA3AF">無任務</td></tr>';
    document.getElementById('gbR').innerHTML='';
    return;
  }

  var allS=tasks.map(function(t){return parseD(t.start);});
  var allE=tasks.map(function(t){return parseD(t.end);});
  var vs=addD(new Date(Math.min.apply(null,allS)),-2);
  var ve=addD(new Date(Math.max.apply(null,allE)),2);
  var days=Math.round((ve-vs)/86400000)+1;

  // Month groups
  var mg=[];
  for(var i=0;i<days;i++){
    var d=addD(vs,i),m=d.getMonth();
    if(!mg.length||mg[mg.length-1].m!==m)mg.push({m:m,count:1});
    else mg[mg.length-1].count++;
  }

  // Left header
  document.getElementById('ghL').innerHTML='<tr><th class="lth" rowspan="2">任務</th></tr><tr></tr>';

  // Right header
  var rH='<tr>';
  mg.forEach(function(g){
    rH+='<th colspan="'+g.count+'" style="text-align:center;font-size:11px;font-weight:700;padding:5px 4px;border-bottom:1px solid rgba(0,0,0,.08);border-right:2px solid rgba(0,0,0,.2);'+getMoStyle(g.m)+'">'+MN[g.m]+'</th>';
  });
  rH+='</tr><tr>';
  for(var i=0;i<days;i++){
    var d=addD(vs,i);
    var isT=d.getTime()===today.getTime(),isWk=d.getDay()===0||d.getDay()===6;
    var isLast=addD(d,1).getDate()===1;
    var ts2=isT?'background:rgba(79,107,237,.18);color:#3451C7;font-weight:700':(isWk?'background:rgba(0,0,0,.04);color:#9CA3AF':getMoStyle(d.getMonth()));
    var brd=isLast?';border-right:2px solid rgba(0,0,0,.2)':'';
    rH+='<th class="dc" style="'+ts2+brd+'">'+d.getDate()+'</th>';
  }
  rH+='</tr>';
  document.getElementById('ghR').innerHTML=rH;

  // Sync left column width
  var leftW=document.getElementById('gHL').querySelector('table').offsetWidth;
  document.getElementById('gBL').querySelector('table').style.width=leftW+'px';

  // Body
  var lB='',rB='';
  tasks.forEach(function(t){
    var ts2=parseD(t.start),te2=parseD(t.end);
    var so=Math.round((ts2-vs)/86400000),eo=Math.round((te2-vs)/86400000);
    var cs=Math.max(0,so),ce=Math.min(days-1,eo);
    var bg=taskBarBg(t),vis=ce>=0&&cs<days;
    var dc=dayCount(t.start,t.end);
    var depIds=getDepIds(t.deps||[]);
    var hasDep=depIds.length>0;
    var meta=[t.owner,t.note].filter(Boolean).join(' · ');
    var isCritical=cpm[t.id]&&cpm[t.id].critical;
    var isGroup=t.type==='group';
    var isMilestone=t.type==='milestone';
    var level=getTaskLevel(t);
    var detailHidden=taskDetailHidden.has(t.id);
    var critClass=isCritical?' critical-task':'';
    var compactClass=detailHidden?' compact':'';
    var paddingLeft=12+(level>0?24:0);
    var childCount=isGroup?allTasks.filter(function(x){return x.parentId===t.id;}).length:0;
    var floatInfo=(!isCritical&&cpm[t.id]&&cpm[t.id].float>0)?'<span style="font-size:9px;color:#9CA3AF;font-family:monospace">浮時:'+cpm[t.id].float+'d</span>':'';
    var critBadge=isCritical?'<span class="badge" style="background:#FEE2E2;color:#991B1B;font-size:9px">關鍵</span>':'';
    var critDot=isCritical?'<span style="font-size:8px;color:#DC2626;vertical-align:middle;margin-right:2px">●</span>':'';

    lB+='<tr data-id="'+t.id+'" class="'+compactClass.trim()+'" onmouseenter="rowHover('+t.id+',true)" onmouseleave="rowHover('+t.id+',false)">';
    lB+='<td class="nc'+critClass+'" style="padding-left:'+paddingLeft+'px">';
    lB+='<div class="ni">';
    if(isGroup){lB+='<span class="group-toggle" onclick="event.stopPropagation();toggleGroupCollapse('+t.id+')" title="'+(t.collapsed?'展開':'摺疊')+'">'+(t.collapsed?'▶':'▼')+'</span>';}
    lB+='<div class="nt">';
    lB+='<div class="nn" style="color:'+(isMilestone?'#7C3AED':bg)+';font-weight:'+(isCritical||isGroup?'700':'500')+'">'+(isMilestone?'◆ ':'')+critDot+t.name+(isGroup?' ('+childCount+')':'')+'</div>';
    lB+='<div class="nm'+(detailHidden?' nm-hidden':'')+'">';
    lB+='<span class="day-chip">'+(isMilestone?'里程碑':dc+'天')+'</span>';
    if(meta)lB+='<span>'+meta+'</span>';
    lB+='</div>';
    lB+='</div>';
    lB+='<div style="display:flex;align-items:center;gap:4px;flex-shrink:0">';
    lB+=critBadge+floatInfo;
    lB+='<span class="badge '+SC[t.status]+'">'+SL[t.status]+'</span>';
    lB+='<button class="detail-tog" onclick="toggleTaskDetail('+t.id+',event)" title="'+(detailHidden?'展開詳情':'收合詳情')+'">'+(detailHidden?'▶':'▽')+'</button>';
    lB+='</div>';
    lB+='</div></td></tr>';

    rB+='<tr data-id="'+t.id+'" class="'+compactClass.trim()+'" onmouseenter="rowHover('+t.id+',true)" onmouseleave="rowHover('+t.id+',false)">';
    for(var i=0;i<days;i++){
      var d=addD(vs,i);
      var isT=d.getTime()===today.getTime(),isWk=d.getDay()===0||d.getDay()===6;
      var isLast=addD(d,1).getDate()===1;
      var bkSty=isT?'background:rgba(79,107,237,.06)':(isWk?'background:rgba(0,0,0,.02)':'');
      var brd=isLast?';border-right:2px solid rgba(0,0,0,.12)':'';
      rB+='<td class="bc" style="'+bkSty+brd+'">';
      if(isT)rB+='<div class="tl" style="background:#4F6BED"></div>';
      if(vis&&i===cs){
        var span=ce-cs+1;
        var barW=isMilestone?26:span*COL_W-4;
        var critBarClass=isCritical?' critical-bar':'';
        var msBarClass=isMilestone?' milestone-bar':'';
        var innerText='';
        if(!isMilestone&&span>=3)innerText=''+t.pct+'% <span class="bd2">'+fmt(t.start)+'–'+fmt(t.end)+'</span>';
        else if(!isMilestone&&span===2)innerText=''+t.pct+'%';
        rB+='<div class="bw"><div class="bar'+critBarClass+msBarClass+'" style="left:2px;width:'+barW+'px;background:'+bg+'">';
        rB+='<div class="bp2" style="width:'+t.pct+'%"></div>';
        rB+='<span class="bl2">'+innerText+'</span>';
        rB+='</div></div>';
      }
      rB+='</td>';
    }
    rB+='</tr>';
  });

  document.getElementById('gbL').innerHTML=lB;
  document.getElementById('gbR').innerHTML=rB;

  // Dependency arrows SVG
  var wrap=document.getElementById('gBR');
  var totalW=days*COL_W,totalH=0;
  var rowYMap={};
  tasks.forEach(function(t){
    rowYMap[t.id]=totalH;
    totalH+=taskDetailHidden.has(t.id)?34:62;
  });
  var svg=document.createElementNS('http://www.w3.org/2000/svg','svg');
  svg.setAttribute('width',totalW);svg.setAttribute('height',totalH);
  svg.style.cssText='position:absolute;top:0;left:0;pointer-events:none;z-index:4;overflow:visible';
  var mid='da'+Date.now();
  svg.innerHTML='<defs><marker id="'+mid+'" markerWidth="6" markerHeight="6" refX="5" refY="3" orient="auto"><path d="M0,0 L6,3 L0,6 Z" fill="#6B7280"/></marker></defs>';
  tasks.forEach(function(t){
    if(!t.deps||!t.deps.length)return;
    var depIds2=getDepIds(t.deps||[]);
    var toR=rowYMap[t.id];if(toR===undefined)return;
    var rh=taskDetailHidden.has(t.id)?34:62;
    var toX=Math.max(0,Math.round((parseD(t.start)-vs)/86400000))*COL_W;
    var toY=toR+rh/2;
    depIds2.forEach(function(did){
      var dep=tasks.find(function(x){return x.id===did;});if(!dep)return;
      var fr=rowYMap[did];if(fr===undefined)return;
      var drh=taskDetailHidden.has(did)?34:62;
      var fromX=(Math.min(days-1,Math.round((parseD(dep.end)-vs)/86400000))+1)*COL_W;
      var fromY=fr+drh/2;
      var mx=(fromX+toX)/2;
      var path=document.createElementNS('http://www.w3.org/2000/svg','path');
      path.setAttribute('d','M'+fromX+','+fromY+' C'+mx+','+fromY+' '+mx+','+toY+' '+(toX-2)+','+toY);
      path.setAttribute('stroke','#6B7280');path.setAttribute('stroke-width','1.5');
      path.setAttribute('stroke-dasharray','5,3');path.setAttribute('fill','none');
      path.setAttribute('marker-end','url(#'+mid+')');path.setAttribute('opacity','.6');
      svg.appendChild(path);
    });
  });
  var tbl=wrap.querySelector('table');
  if(tbl){tbl.style.position='relative';var oldSvg=tbl.querySelector('svg');if(oldSvg)oldSvg.remove();tbl.appendChild(svg);}
}

// ── Stats render ──
function renderStats(){
  var tasks=PROJ.tasks||[];
  var done=tasks.filter(function(t){return t.status==='done';}).length;
  var ong=tasks.filter(function(t){return t.status==='ongoing';}).length;
  var del=tasks.filter(function(t){return t.status==='delayed';}).length;
  var avg=tasks.length?Math.round(tasks.reduce(function(a,t){return a+t.pct;},0)/tasks.length):0;
  var allS=tasks.map(function(t){return parseD(t.start);});
  var allE=tasks.map(function(t){return parseD(t.end);});
  var spanDays=tasks.length?Math.round((new Date(Math.max.apply(null,allE))-new Date(Math.min.apply(null,allS)))/86400000)+1:0;
  document.getElementById('statsGrid').innerHTML=
    '<div class="stat"><div class="sl">總任務</div><div class="sv" style="color:#4F6BED">'+tasks.length+'</div></div>'+
    '<div class="stat"><div class="sl">已完成</div><div class="sv" style="color:#059669">'+done+'</div></div>'+
    '<div class="stat"><div class="sl">進行中</div><div class="sv" style="color:#4F6BED">'+ong+'</div></div>'+
    '<div class="stat"><div class="sl">延遲中</div><div class="sv" style="color:#DC2626">'+del+'</div></div>'+
    '<div class="stat"><div class="sl">整體完成度</div><div class="sv" style="color:#D97706">'+avg+'%</div></div>'+
    '<div class="stat"><div class="sl">專案總天數</div><div class="sv" style="color:#6B7280">'+spanDays+'天</div></div>';
}

// ── Canvas export (matches editor drawGanttCanvas) ──
function drawGanttCanvas(){
  var allTasks=PROJ.tasks||[];
  var tasks=getVisibleTasks(allTasks);
  if(!tasks.length){var c=document.createElement('canvas');c.width=100;c.height=100;return c;}
  var cpm=computeCPM(tasks);
  var ROW_FULL=62,ROW_COMPACT=34;
  function rh(t){return taskDetailHidden.has(t.id)?ROW_COMPACT:ROW_FULL;}
  var taskYMap={};var cumY=0;
  tasks.forEach(function(t){taskYMap[t.id]=cumY;cumY+=rh(t);});
  var totalTaskH=cumY;
  var allS=tasks.map(function(t){return parseD(t.start);});
  var allE=tasks.map(function(t){return parseD(t.end);});
  var minD=new Date(Math.min.apply(null,allS)),maxD=new Date(Math.max.apply(null,allE));
  var vs=addD(minD,-2),ve=addD(maxD,2);
  var days=Math.round((ve-vs)/86400000)+1;
  var DPR=2,NAME_W=240,COL_W=30,FOOTER_H=44,TITLE_H=56,HEAD_H=52,PAD=24;
  var W=NAME_W+days*COL_W+PAD*2,H=TITLE_H+HEAD_H+totalTaskH+FOOTER_H;
  var canvas=document.createElement('canvas');
  canvas.width=W*DPR;canvas.height=H*DPR;
  var cx=canvas.getContext('2d');cx.scale(DPR,DPR);
  function X(i){return PAD+NAME_W+i*COL_W;}
  function Y(id){return TITLE_H+HEAD_H+taskYMap[id];}
  var todayOff=Math.round((today-vs)/86400000);
  var today_s=ds(today);
  var minS2=tasks.reduce(function(a,t){return a<t.start?a:t.start;},tasks[0].start);
  var maxE2=tasks.reduce(function(a,t){return a>t.end?a:t.end;},tasks[0].end);
  cx.fillStyle='#fff';cx.fillRect(0,0,W,H);
  cx.fillStyle='#1E2235';cx.fillRect(0,0,W,TITLE_H);
  cx.fillStyle='#fff';cx.font='600 18px "DM Sans",sans-serif';cx.textBaseline='middle';
  cx.fillText(PROJ.name,PAD+40,TITLE_H/2-6);
  cx.font='400 11px "DM Sans",sans-serif';cx.fillStyle='rgba(255,255,255,.55)';
  cx.fillText(tasks.length+' 個任務 | '+minS2+' 至 '+maxE2+' | 生成日期：'+today_s,PAD+40,TITLE_H/2+10);
  cx.fillStyle=PROJ.color||'#4F6BED';cx.beginPath();cx.arc(PAD+12,TITLE_H/2,8,0,Math.PI*2);cx.fill();
  cx.fillStyle='#F3F4F6';cx.fillRect(0,TITLE_H,W,HEAD_H);
  cx.fillStyle='#6B7280';cx.font='600 10px "DM Sans",sans-serif';cx.textBaseline='middle';
  cx.fillText('任務',PAD+12,TITLE_H+HEAD_H/2);
  var prevMonth=-1;
  for(var i=0;i<days;i++){
    var d2=addD(vs,i);var isT2=d2.getTime()===today.getTime();var isWk2=d2.getDay()===0||d2.getDay()===6;
    var x2=X(i);
    if(isT2){cx.fillStyle='rgba(79,107,237,.15)';cx.fillRect(x2,TITLE_H,COL_W,HEAD_H);}
    else if(isWk2){cx.fillStyle='rgba(0,0,0,.025)';cx.fillRect(x2,TITLE_H,COL_W,HEAD_H);}
    var m2=d2.getMonth();var lbl='';
    if(m2!==prevMonth){lbl=(m2+1)+'月';prevMonth=m2;}
    else if(d2.getDay()===1||i===0)lbl=''+d2.getDate();
    if(lbl){cx.fillStyle=isT2?'#4F6BED':'#6B7280';cx.font=(isT2?600:500)+' 9px "DM Mono",monospace';cx.textAlign='center';cx.textBaseline='middle';cx.fillText(lbl,x2+COL_W/2,TITLE_H+HEAD_H/2);}
  }
  cx.textAlign='left';
  cx.strokeStyle='#E5E7EB';cx.lineWidth=1;cx.beginPath();cx.moveTo(0,TITLE_H+HEAD_H);cx.lineTo(W,TITLE_H+HEAD_H);cx.stroke();
  var BC={done:{bg:'#D1FAE5',tx:'#065F46'},ongoing:{bg:'#E0E7FF',tx:'#3730A3'},pending:{bg:'#F3F4F6',tx:'#6B7280'},delayed:{bg:'#FEE2E2',tx:'#991B1B'}};
  tasks.forEach(function(t,ri){
    var y=Y(t.id);var rowH=rh(t);var isCompact=taskDetailHidden.has(t.id);
    var bg=taskBarBg(t);var isCritical=cpm[t.id]&&cpm[t.id].critical;
    var isMilestone=t.type==='milestone';var isGroup=t.type==='group';
    var level=getTaskLevel(t);var indentX=level>0?20:0;
    var meta=[t.owner,t.note].filter(Boolean).join(' · ');
    if(ri%2===1){cx.fillStyle='rgba(0,0,0,.018)';cx.fillRect(0,y,W,rowH);}
    cx.strokeStyle='rgba(0,0,0,.06)';cx.lineWidth=0.5;cx.beginPath();cx.moveTo(0,y+rowH);cx.lineTo(W,y+rowH);cx.stroke();
    if(todayOff>=0&&todayOff<days){cx.fillStyle='rgba(79,107,237,.04)';cx.fillRect(X(todayOff),y,COL_W,rowH);}
    for(var i=0;i<days;i++){var d3=addD(vs,i);if(d3.getDay()===0||d3.getDay()===6){cx.fillStyle='rgba(0,0,0,.018)';cx.fillRect(X(i),y,COL_W,rowH);}}
    cx.fillStyle='#fff';cx.fillRect(0,y,NAME_W+PAD,rowH);
    cx.strokeStyle='#E5E7EB';cx.lineWidth=0.5;cx.beginPath();cx.moveTo(PAD+NAME_W,y);cx.lineTo(PAD+NAME_W,y+rowH);cx.stroke();
    if(isCritical){cx.fillStyle='#DC2626';cx.fillRect(0,y,3,rowH);}
    var bc=BC[t.status]||BC.pending;var blabel=SL[t.status];
    cx.font='600 9px "DM Sans",sans-serif';cx.textBaseline='alphabetic';
    var bw2=cx.measureText(blabel).width+14;var badgeX=PAD+NAME_W-bw2-8;var badgeMidY=y+rowH/2;
    cx.fillStyle=bc.bg;rrect(cx,badgeX,badgeMidY-8,bw2,16,8);cx.fill();
    cx.fillStyle=bc.tx;cx.textAlign='center';cx.fillText(blabel,badgeX+bw2/2,badgeMidY+4);cx.textAlign='left';
    var nameX=PAD+10+indentX;var nameMaxW=NAME_W-bw2-24-indentX;
    var nameY2=isCompact?y+rowH*0.65:y+rowH*0.38;
    if(isCritical){cx.fillStyle='#DC2626';cx.beginPath();cx.arc(nameX,nameY2-5,3.5,0,Math.PI*2);cx.fill();}
    var dotOff=isCritical?10:0;
    cx.fillStyle='#111827';cx.font=(isCritical||isGroup?'700':'500')+' 12.5px "DM Sans",sans-serif';cx.textBaseline='alphabetic';
    cx.fillText(trunc(cx,(isMilestone?'◆ ':'')+t.name,nameMaxW-dotOff),nameX+dotOff,nameY2);
    if(!isCompact){
      var metaY=y+rowH*0.70;var dc=dayCount(t.start,t.end);
      cx.font='500 9.5px "DM Mono",monospace';cx.textBaseline='alphabetic';
      var chipLabel=isMilestone?'里程碑':dc+'天';var chipTxtW=cx.measureText(chipLabel).width;
      var chipW=chipTxtW+12;
      cx.fillStyle='#E5E7EB';rrect(cx,nameX-2,metaY-11,chipW,14,5);cx.fill();
      cx.fillStyle='#4B5563';cx.fillText(chipLabel,nameX+4,metaY);
      if(meta){cx.font='400 10px "DM Sans",sans-serif';cx.fillStyle='#9CA3AF';cx.fillText(trunc(cx,meta,nameMaxW-chipW-6),nameX+chipW+6,metaY);}
    }
    var ts3=parseD(t.start),te3=parseD(t.end);
    var so=Math.round((ts3-vs)/86400000),eo=Math.round((te3-vs)/86400000);
    var cs=Math.max(0,so),ce=Math.min(days-1,eo);
    if(ce>=0&&cs<days){
      var barH=20,barY=y+(rowH-barH)/2;
      var bx3=X(cs)+2,bw3=(ce-cs+1)*COL_W-4;
      if(isMilestone){
        var mx2=bx3+bw3/2,my2=barY+barH/2,ms=10;
        cx.fillStyle=bg;cx.beginPath();cx.moveTo(mx2,my2-ms);cx.lineTo(mx2+ms,my2);cx.lineTo(mx2,my2+ms);cx.lineTo(mx2-ms,my2);cx.closePath();cx.fill();
      }else{
        cx.shadowColor='rgba(0,0,0,.15)';cx.shadowBlur=3;cx.shadowOffsetY=1;
        cx.fillStyle=bg;rrect(cx,bx3,barY,bw3,barH,5);cx.fill();
        cx.shadowColor='transparent';cx.shadowBlur=0;cx.shadowOffsetY=0;
        if(isCritical){cx.strokeStyle='#DC2626';cx.lineWidth=2;rrect(cx,bx3-1,barY-1,bw3+2,barH+2,6);cx.stroke();}
        if(t.pct>0){cx.fillStyle='rgba(255,255,255,.22)';rrect(cx,bx3,barY,bw3*t.pct/100,barH,5);cx.fill();}
        var span=ce-cs+1,showDate=span>=5;
        var barTxt=t.pct>0||span>2?t.pct+'%':'';
        if(showDate)barTxt+=(barTxt?' ':'')+fmt(t.start)+'–'+fmt(t.end);
        if(barTxt){cx.fillStyle='rgba(255,255,255,.95)';cx.font='600 10px "DM Sans",sans-serif';cx.textBaseline='middle';cx.fillText(trunc(cx,barTxt,bw3-12),bx3+8,barY+barH/2);}
      }
    }
  });
  if(todayOff>=0&&todayOff<days){
    var tx3=X(todayOff)+COL_W/2;
    cx.strokeStyle='rgba(79,107,237,.4)';cx.lineWidth=1.5;cx.setLineDash([4,3]);
    cx.beginPath();cx.moveTo(tx3,TITLE_H+HEAD_H);cx.lineTo(tx3,TITLE_H+HEAD_H+totalTaskH);cx.stroke();cx.setLineDash([]);
  }
  tasks.forEach(function(t){
    if(!t.deps||!t.deps.length)return;
    var depIds2=getDepIds(t.deps);
    var ts3=parseD(t.start);var toOff=Math.round((ts3-vs)/86400000);
    var toX=X(Math.max(0,toOff));var toY=Y(t.id)+rh(t)/2;
    depIds2.forEach(function(did){
      var dep=tasks.find(function(x){return x.id===did;});if(!dep)return;
      var depEnd=parseD(dep.end);var fromOff=Math.round((depEnd-vs)/86400000);
      var fromX=X(Math.min(days-1,fromOff))+COL_W;var fromY=Y(dep.id)+rh(dep)/2;
      var mx=(fromX+toX)/2;
      cx.strokeStyle='rgba(107,114,128,.55)';cx.lineWidth=1.2;cx.setLineDash([5,3]);
      cx.beginPath();cx.moveTo(fromX,fromY);cx.bezierCurveTo(mx,fromY,mx,toY,toX-2,toY);cx.stroke();cx.setLineDash([]);
      cx.fillStyle='rgba(107,114,128,.55)';cx.beginPath();cx.moveTo(toX,toY);cx.lineTo(toX-6,toY-4);cx.lineTo(toX-6,toY+4);cx.closePath();cx.fill();
    });
  });
  cx.strokeStyle='#E5E7EB';cx.lineWidth=1;cx.strokeRect(0.5,0.5,W-1,H-1);
  cx.beginPath();cx.moveTo(PAD+NAME_W,TITLE_H);cx.lineTo(PAD+NAME_W,TITLE_H+HEAD_H);cx.stroke();
  var footerY=TITLE_H+HEAD_H+totalTaskH;
  cx.fillStyle='#F9FAFB';cx.fillRect(0,footerY,W,FOOTER_H);
  cx.strokeStyle='#E5E7EB';cx.lineWidth=0.5;cx.beginPath();cx.moveTo(0,footerY);cx.lineTo(W,footerY);cx.stroke();
  var legends=[{color:'#059669',label:'已完成'},{color:'#4F6BED',label:'進行中'},{color:'#9CA3AF',label:'待開始'},{color:'#DC2626',label:'延遲'},{color:'#DC2626',label:'● 關鍵路徑'}];
  var lx=PAD;cx.textBaseline='middle';
  legends.forEach(function(l){
    cx.fillStyle=l.color;cx.beginPath();cx.arc(lx+5,footerY+FOOTER_H/2,5,0,Math.PI*2);cx.fill();
    cx.fillStyle='#6B7280';cx.font='400 10px "DM Sans",sans-serif';
    cx.fillText(l.label,lx+14,footerY+FOOTER_H/2);
    lx+=cx.measureText(l.label).width+32;
  });
  cx.fillStyle='#9CA3AF';cx.font='400 10px "DM Sans",sans-serif';cx.textAlign='right';
  cx.fillText('唯讀分享 · '+today_s,W-PAD,footerY+FOOTER_H/2);cx.textAlign='left';
  return canvas;
}

function rrect(ctx,x,y,w,h,r){ctx.beginPath();ctx.moveTo(x+r,y);ctx.lineTo(x+w-r,y);ctx.arcTo(x+w,y,x+w,y+r,r);ctx.lineTo(x+w,y+h-r);ctx.arcTo(x+w,y+h,x+w-r,y+h,r);ctx.lineTo(x+r,y+h);ctx.arcTo(x,y+h,x,y+h-r,r);ctx.lineTo(x,y+r);ctx.arcTo(x,y,x+r,y,r);ctx.closePath();}
function trunc(ctx,text,maxW){if(ctx.measureText(text).width<=maxW)return text;while(text.length>1&&ctx.measureText(text+'...').width>maxW)text=text.slice(0,-1);return text+'...';}

// ── PNG / PDF export ──
async function exportImg(format){
  document.getElementById('exportOverlay').classList.add('open');
  await new Promise(function(r){setTimeout(r,60);});
  try{
    var canvas=drawGanttCanvas();
    if(format==='png'){
      canvas.toBlob(function(blob){
        var url=URL.createObjectURL(blob);var a=document.createElement('a');
        a.href=url;a.download=PROJ.name+'_甘特圖.png';a.click();
        setTimeout(function(){URL.revokeObjectURL(url);},3000);
        document.getElementById('exportOverlay').classList.remove('open');
      },'image/png');
    }else{
      var jsPDF2=window.jspdf.jsPDF;
      var W2=canvas.width,H2=canvas.height,mmW=W2*0.264583,mmH=H2*0.264583;
      var pdf=new jsPDF2({orientation:W2>H2?'landscape':'portrait',unit:'mm',format:[mmW,mmH]});
      pdf.addImage(canvas.toDataURL('image/png',1.0),'PNG',0,0,mmW,mmH,undefined,'FAST');
      pdf.save(PROJ.name+'_甘特圖.pdf');
      document.getElementById('exportOverlay').classList.remove('open');
    }
  }catch(e){document.getElementById('exportOverlay').classList.remove('open');alert('匯出失敗：'+e.message);}
}

// ── Init ──
window.addEventListener('load',function(){
  // Header
  document.getElementById('projDot').style.background=PROJ.color||'#4F6BED';
  document.getElementById('projNameH').textContent=PROJ.name;
  var cb=${createdByJson};
  document.getElementById('projMeta').textContent='共 '+PROJ.tasks.length+' 個任務 · 分享於 ${createdAt}'+(cb?' · '+cb:'');
  // Stats initial collapse state
  if(STATS_COLLAPSED_INIT){
    document.getElementById('statsBody').classList.add('closed');
    document.getElementById('statsToggleBar').classList.add('closed');
    document.getElementById('statsToggleLabel').textContent='統計資訊（已收合）';
  }
  // Collapse detail button
  var btn=document.getElementById('btnCollapseDetail');
  if(btn)btn.textContent=allDetailCollapsed?'☰ 展開詳情':'☰ 收合詳情';
  renderStats();
  render();
});`;

  const html = '<!DOCTYPE html>\n'
    + '<html lang="zh-TW">\n<head>\n'
    + '<meta charset="UTF-8">\n'
    + '<meta name="viewport" content="width=device-width,initial-scale=1.0">\n'
    + '<title>' + proj.name + ' — 甘特圖分享</title>\n'
    + '<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"><\/script>\n'
    + '<style>\n@import url(\'https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=DM+Mono:wght@400;500&display=swap\');\n'
    + css + '\n</style>\n</head>\n<body>\n'

    // Export overlay
    + '<div id="exportOverlay"><div class="export-box"><div class="export-spinner"></div>'
    + '<div style="font-size:14px;font-weight:500;color:#111">正在產生圖表...</div></div></div>\n'

    // Topbar row1
    + '<div class="topbar">\n'
    + '<div class="topbar-row1">\n'
    + '  <div class="tl-left">\n'
    + '    <div class="pdot" id="projDot"></div>\n'
    + '    <div class="proj-title" id="projNameH"></div>\n'
    + '    <span class="proj-badge">🔒 唯讀模式</span>\n'
    + '    <span class="pmeta" id="projMeta"></span>\n'
    + '  </div>\n'
    + '  <div class="tl-right">\n'
    + '    <button class="btn g" onclick="exportImg(\'png\')">🖼️ PNG</button>\n'
    + '    <button class="btn o" onclick="exportImg(\'pdf\')">📄 PDF</button>\n'
    + '  </div>\n'
    + '</div>\n'
    // Topbar row2 (toolbar)
    + '<div class="topbar-row2">\n'
    + '  <button class="btn" id="btnCollapseDetail" onclick="toggleAllTaskDetail()">☰ 收合詳情</button>\n'
    + '  <span style="font-size:11px;color:rgba(255,255,255,.4);margin-left:6px">分享 ID: ' + shareId + '</span>\n'
    + '</div>\n'
    + '</div>\n'

    // Main content
    + '<div class="main">\n'

    // Stats panel
    + '  <div class="stats-wrap">\n'
    + '    <div class="stats-toggle-bar" id="statsToggleBar" onclick="toggleStatsPanel()">\n'
    + '      <span class="sti" id="statsToggleIcon">▼</span>\n'
    + '      <span id="statsToggleLabel">統計資訊</span>\n'
    + '    </div>\n'
    + '    <div class="stats-body" id="statsBody">\n'
    + '      <div class="stats" id="statsGrid"></div>\n'
    + '    </div>\n'
    + '  </div>\n'

    // Gantt chart
    + '  <div class="gc">\n'
    + '    <div class="gantt-head-row">\n'
    + '      <div class="gantt-head-left" id="gHL" style="width:220px;min-width:220px">'
    + '<table class="gantt-table-head-left"><thead id="ghL"></thead></table></div>\n'
    + '      <div class="gantt-head-right" id="gHR" style="overflow:hidden">'
    + '<table class="gantt-table-head-right"><thead id="ghR"></thead></table></div>\n'
    + '    </div>\n'
    + '    <div class="gantt-body-row">\n'
    + '      <div class="gantt-body-left" id="gBL" style="width:220px;min-width:220px;overflow:hidden" onscroll="syncScrollLeft(this)">'
    + '<table class="gantt-table-body-left"><tbody id="gbL"></tbody></table></div>\n'
    + '      <div class="gantt-body-right" id="gBR" onscroll="syncScroll(this)" style="overflow:auto">'
    + '<table class="gantt-table-body-right" style="position:relative"><tbody id="gbR"></tbody></table></div>\n'
    + '    </div>\n'
    + '    <div class="leg">\n'
    + '      <div class="li"><div class="ld" style="background:#059669"></div>已完成</div>\n'
    + '      <div class="li"><div class="ld" style="background:#4F6BED"></div>進行中</div>\n'
    + '      <div class="li"><div class="ld" style="background:#9CA3AF"></div>待開始</div>\n'
    + '      <div class="li"><div class="ld" style="background:#DC2626"></div>延遲</div>\n'
    + '      <div class="li"><div class="ld" style="background:#DC2626;border-radius:50%"></div>關鍵路徑</div>\n'
    + '      <div class="ln">唯讀分享</div>\n'
    + '    </div>\n'
    + '  </div>\n'
    + '</div>\n'
    + '<script>\n' + js + '\n<\/script>\n'
    + '</body>\n</html>';

  return html;
}
