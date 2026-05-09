// ================================================================
//  Production Timer System — Code.gs  (v4 Enhanced)
//  Google Apps Script Backend
//  รองรับข้อมูลหลักล้าน record ด้วย Batch Write + Index + Cache
//  เพิ่ม: stdMs (เวลามาตรฐาน) ใน TimerLog และ DailySummary
// ================================================================

const CONFIG = {
  SHEET_NAME:         "TimerLog",
  INDEX_SHEET_NAME:   "Index",
  SUMMARY_SHEET_NAME: "DailySummary",
  BATCH_SIZE:         500,
  MAX_ROWS_PER_SHEET: 900000,
  CACHE_TTL:          300,
  TIMEZONE:           "Asia/Bangkok",
};

// เพิ่ม StdMs และ VsStd ในหัว column
const HEADERS = [
  "Timestamp","Date","EmpID","EmpName","Department",
  "Station","StartTime","EndTime","Duration","ElapsedMs",
  "StdMs","VsStd","Remark","SheetSeq",
];

// ── doGet: serve HTML app ────────────────────────────────────────
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || "page";

  if (action === "page") {
    return HtmlService.createTemplateFromFile("index")
      .evaluate()
      .setTitle("Production Timer System")
      .addMetaTag("viewport","width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (action === "dashboard") return apiDashboard(e);
  if (action === "history")   return apiHistory(e);
  if (action === "search")    return apiSearch(e);
  if (action === "export")    return apiExport(e);

  return jsonResp({ status:"error", message:"Unknown action" });
}

// ── doPost: รับข้อมูลจาก Timer (HTTP POST) ──────────────────────
function doPost(e) {
  try {
    const raw  = JSON.parse(e.postData.contents);
    const rows = Array.isArray(raw) ? raw : [raw];

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    const now   = new Date();
    const dateStr = fmtDate(now);
    const seq   = getRolloverSeq(ss, sheet);

    const writeRows = rows.map(d => buildRow(d, now, dateStr, seq));

    batchAppend(sheet, writeRows);
    updateIndex(ss, rows, dateStr, seq);
    updateDailySummary(ss, rows, dateStr);

    CacheService.getScriptCache().remove("dash_" + dateStr);

    return jsonResp({ status:"ok", saved: writeRows.length });
  } catch(err) {
    console.error(err);
    return jsonResp({ status:"error", message: err.message });
  }
}

// ── include() สำหรับใช้ใน index.html ────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── สร้าง row array พร้อม stdMs และ vsStd ───────────────────────
function buildRow(d, now, dateStr, seq) {
  const elapsedMs = Number(d.elapsedMs) || 0;
  const stdMs     = Number(d.stdMs)     || 0;
  let vsStd = "";
  if (stdMs > 0 && elapsedMs > 0) {
    const diff = elapsedMs - stdMs;
    vsStd = diff <= 0
      ? "ปกติ (" + fmtMs(Math.abs(diff)) + " ต่ำกว่า)"
      : "เกิน +" + fmtMs(diff);
  }
  return [
    now, dateStr,
    san(d.empId), san(d.empName), san(d.dept), san(d.station),
    san(d.startTime), san(d.endTime), san(d.duration),
    elapsedMs, stdMs, vsStd, san(d.remark), seq,
  ];
}

// ── API: dashboard ───────────────────────────────────────────────
function apiDashboard(e) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = (e.parameter && e.parameter.date) || todayStr();
  const key     = "dash_" + dateStr;
  const cache   = CacheService.getScriptCache();
  const hit     = cache.get(key);
  if (hit) return jsonResp(JSON.parse(hit));

  const sumSheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  let result = { date:dateStr, count:0, avgMs:0, bestMs:0, totalMs:0, worstMs:0, overStdCount:0 };

  if (sumSheet) {
    const data = sumSheet.getDataRange().getValues();
    for (let i=1; i<data.length; i++) {
      if (String(data[i][0])===dateStr) {
        result = {
          date:         dateStr,
          count:        Number(data[i][1]) || 0,
          totalMs:      Number(data[i][2]) || 0,
          avgMs:        Number(data[i][3]) || 0,
          bestMs:       Number(data[i][4]) || 0,
          worstMs:      Number(data[i][5]) || 0,
          overStdCount: Number(data[i][6]) || 0,
        };
        break;
      }
    }
  }
  cache.put(key, JSON.stringify(result), CONFIG.CACHE_TTL);
  return jsonResp(result);
}

// ── API: history ─────────────────────────────────────────────────
function apiHistory(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const empId = (e.parameter&&e.parameter.empId) || "";
  const dept  = (e.parameter&&e.parameter.dept)  || "";
  const date  = (e.parameter&&e.parameter.date)  || "";
  const limit = Math.min(parseInt((e.parameter&&e.parameter.limit)||50), 500);

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return jsonResp({ rows:[] });

  const lastRow = sheet.getLastRow();
  if (lastRow<=1) return jsonResp({ rows:[] });

  const readCount = Math.min(limit*4, lastRow-1);
  const startRow  = Math.max(2, lastRow - readCount + 1);
  const values    = sheet.getRange(startRow, 1, lastRow-startRow+1, HEADERS.length).getValues();

  const filtered = values.reverse()
    .filter(r => {
      if (empId && r[2]!==empId) return false;
      if (dept  && r[4]!==dept)  return false;
      if (date  && String(r[1])!==date)  return false;
      return true;
    })
    .slice(0, limit)
    .map(rowToObj);

  return jsonResp({ rows:filtered });
}

// ── API: search ──────────────────────────────────────────────────
function apiSearch(e) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const dateFrom = (e.parameter&&e.parameter.dateFrom) || todayStr();
  const dateTo   = (e.parameter&&e.parameter.dateTo)   || todayStr();
  const empId    = (e.parameter&&e.parameter.empId)    || "";
  const dept     = (e.parameter&&e.parameter.dept)     || "";
  const limit    = Math.min(parseInt((e.parameter&&e.parameter.limit)||200), 1000);

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return jsonResp({ rows:[] });

  const lastRow = sheet.getLastRow();
  if (lastRow<=1) return jsonResp({ rows:[] });

  const CHUNK=5000, results=[];
  for (let row=lastRow; row>=2 && results.length<limit; row-=CHUNK) {
    const from  = Math.max(2, row-CHUNK+1);
    const chunk = sheet.getRange(from,1,row-from+1,HEADERS.length).getValues();
    chunk.reverse().forEach(r=>{
      const d=String(r[1]);
      if(d<dateFrom||d>dateTo) return;
      if(empId&&r[2]!==empId) return;
      if(dept&&r[4]!==dept)   return;
      if(results.length>=limit) return;
      results.push(rowToObj(r));
    });
  }
  return jsonResp({ rows:results, total:results.length });
}

// ── API: export CSV ──────────────────────────────────────────────
function apiExport(e) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const dateStr = (e.parameter&&e.parameter.date) || todayStr();
  const sheet   = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) return ContentService.createTextOutput("No data")
    .setMimeType(ContentService.MimeType.TEXT);

  const lastRow = sheet.getLastRow();
  const lines   = [HEADERS.join(",")];

  if (lastRow>1) {
    const values = sheet.getRange(2,1,lastRow-1,HEADERS.length).getValues();
    values.forEach(r=>{
      if (String(r[1])===dateStr)
        lines.push(r.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(","));
    });
  }
  return ContentService.createTextOutput("\uFEFF"+lines.join("\n"))
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Helpers ──────────────────────────────────────────────────────
function batchAppend(sheet, rows) {
  for (let i=0; i<rows.length; i+=CONFIG.BATCH_SIZE) {
    const chunk   = rows.slice(i, i+CONFIG.BATCH_SIZE);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow+1,1,chunk.length,chunk[0].length).setValues(chunk);
  }
}

function updateIndex(ss, rows, dateStr, seq) {
  const idx = getOrCreateSheet(ss, CONFIG.INDEX_SHEET_NAME);
  if (idx.getLastRow()===0)
    idx.appendRow(["Date","EmpID","Dept","Count","TotalMs","Updated"]);

  const groups={};
  rows.forEach(r=>{
    const k=dateStr+"|"+(r.empId||"");
    if(!groups[k]) groups[k]={date:dateStr,empId:r.empId||"",dept:r.dept||"",count:0,totalMs:0};
    groups[k].count++;
    groups[k].totalMs+=Number(r.elapsedMs)||0;
  });

  const data = idx.getDataRange().getValues();
  Object.values(groups).forEach(g=>{
    let found=false;
    for(let i=1;i<data.length;i++){
      if(String(data[i][0])===g.date && String(data[i][1])===g.empId){
        data[i][3]=Number(data[i][3])+g.count;
        data[i][4]=Number(data[i][4])+g.totalMs;
        data[i][5]=new Date();
        found=true;
        break;
      }
    }
    if(!found) data.push([g.date,g.empId,g.dept,g.count,g.totalMs,new Date()]);
  });
  idx.clearContents();
  idx.getRange(1,1,data.length,data[0].length).setValues(data);
}

function updateDailySummary(ss, rows, dateStr) {
  const sum = getOrCreateSheet(ss, CONFIG.SUMMARY_SHEET_NAME);
  // เพิ่ม OverStdCount ในหัว summary
  if (sum.getLastRow()===0)
    sum.appendRow(["Date","Count","TotalMs","AvgMs","BestMs","WorstMs","OverStdCount","Updated"]);

  const msList   = rows.map(r=>Number(r.elapsedMs)||0).filter(v=>v>0);
  const newTotal  = msList.reduce((s,v)=>s+v,0);
  const newCount  = rows.length;
  const newBest   = msList.length ? Math.min(...msList) : 0;
  const newWorst  = msList.length ? Math.max(...msList) : 0;
  // นับรอบที่เกินมาตรฐาน
  const newOver   = rows.filter(r=>{
    const em=Number(r.elapsedMs)||0, sm=Number(r.stdMs)||0;
    return sm>0 && em>sm;
  }).length;

  const data=sum.getDataRange().getValues();
  let found=false;
  for(let i=1;i<data.length;i++){
    if(String(data[i][0])===dateStr){
      const cnt=Number(data[i][1])+newCount;
      const tot=Number(data[i][2])+newTotal;
      data[i][1]=cnt;
      data[i][2]=tot;
      data[i][3]=cnt ? Math.round(tot/cnt) : 0;
      data[i][4]=data[i][4]>0 ? Math.min(Number(data[i][4]),newBest||Number(data[i][4])) : newBest;
      data[i][5]=Math.max(Number(data[i][5]),newWorst);
      data[i][6]=(Number(data[i][6])||0)+newOver;
      data[i][7]=new Date();
      found=true;
      break;
    }
  }
  if(!found) data.push([
    dateStr, newCount, newTotal,
    newCount ? Math.round(newTotal/newCount) : 0,
    newBest, newWorst, newOver, new Date()
  ]);

  sum.clearContents();
  sum.getRange(1,1,data.length,data[0].length).setValues(data);
}

function getRolloverSeq(ss, sheet) {
  if (sheet.getLastRow()<CONFIG.MAX_ROWS_PER_SHEET) return 1;
  const names=ss.getSheets().map(s=>s.getName());
  let seq=2;
  while(names.includes(CONFIG.SHEET_NAME+"_"+seq)) seq++;
  const ns=ss.insertSheet(CONFIG.SHEET_NAME+"_"+seq);
  ns.appendRow(HEADERS);
  return seq;
}

function getOrCreateSheet(ss, name) {
  let s=ss.getSheetByName(name);
  if(!s){
    s=ss.insertSheet(name);
    if(name===CONFIG.SHEET_NAME) s.appendRow(HEADERS);
  }
  return s;
}

function rowToObj(r) {
  return {
    timestamp: r[0],  date:      r[1],  empId:    r[2],
    empName:   r[3],  dept:      r[4],  station:  r[5],
    startTime: r[6],  endTime:   r[7],  duration: r[8],
    elapsedMs: r[9],  stdMs:     r[10], vsStd:    r[11],
    remark:    r[12], sheetSeq:  r[13],
  };
}

function fmtMs(ms) {
  if (!ms || ms <= 0) return "0:00";
  const m = Math.floor(ms/60000);
  const s = Math.floor((ms%60000)/1000);
  return m + ":" + String(s).padStart(2,"0");
}

function san(v){ return String(v||"").trim().substring(0,200); }
function todayStr(){ return fmtDate(new Date()); }
function fmtDate(d){ return Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyy-MM-dd"); }
function jsonResp(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Functions สำหรับ google.script.run (เรียกจาก index.html) ────

/**
 * บันทึกข้อมูลจาก Timer
 * @param {Object|Array} payload - ข้อมูล 1 รอบ หรือ Array หลายรอบ
 */
function saveTimerData(payload) {
  try {
    const rows  = Array.isArray(payload) ? payload : [payload];
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, CONFIG.SHEET_NAME);
    const now   = new Date();
    const dateStr = fmtDate(now);
    const seq   = getRolloverSeq(ss, sheet);

    const writeRows = rows.map(d => buildRow(d, now, dateStr, seq));

    batchAppend(sheet, writeRows);
    updateIndex(ss, rows, dateStr, seq);
    updateDailySummary(ss, rows, dateStr);
    CacheService.getScriptCache().remove("dash_" + dateStr);

    return { status: "ok", saved: writeRows.length };
  } catch(err) {
    console.error(err);
    return { status: "error", message: err.message };
  }
}

/**
 * ดึงข้อมูล Dashboard สำหรับวันที่กำหนด
 */
function getDashboardData(dateStr) {
  if (!dateStr) dateStr = todayStr();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const key   = "dash_" + dateStr;
  const cache = CacheService.getScriptCache();
  const hit   = cache.get(key);
  if (hit) return JSON.parse(hit);

  const sumSheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME);
  let result = { date:dateStr, count:0, avgMs:0, bestMs:0, totalMs:0, worstMs:0, overStdCount:0 };
  if (sumSheet) {
    const data = sumSheet.getDataRange().getValues();
    for (let i=1; i<data.length; i++) {
      if (String(data[i][0])===dateStr) {
        result = {
          date:         dateStr,
          count:        Number(data[i][1]) || 0,
          totalMs:      Number(data[i][2]) || 0,
          avgMs:        Number(data[i][3]) || 0,
          bestMs:       Number(data[i][4]) || 0,
          worstMs:      Number(data[i][5]) || 0,
          overStdCount: Number(data[i][6]) || 0,
        };
        break;
      }
    }
  }
  cache.put(key, JSON.stringify(result), CONFIG.CACHE_TTL);
  return result;
}

/**
 * ดึง History ล่าสุด
 */
function getHistoryData(params) {
  params = params || {};
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const limit = Math.min(parseInt(params.limit||50), 500);
  const empId = params.empId || "";
  const dept  = params.dept  || "";
  const date  = params.date  || "";

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) return { rows: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow<=1) return { rows: [] };

  const readCount = Math.min(limit*4, lastRow-1);
  const startRow  = Math.max(2, lastRow-readCount+1);
  const values    = sheet.getRange(startRow,1,lastRow-startRow+1,HEADERS.length).getValues();

  const filtered = values.reverse()
    .filter(r => {
      if (empId && r[2]!==empId) return false;
      if (dept  && r[4]!==dept)  return false;
      if (date  && String(r[1])!==date) return false;
      return true;
    })
    .slice(0, limit)
    .map(rowToObj);

  return { rows: filtered };
}

// ── Setup (รันครั้งเดียวหลัง Deploy) ─────────────────────────────
function setupSpreadsheet() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEET_NAME, CONFIG.INDEX_SHEET_NAME, CONFIG.SUMMARY_SHEET_NAME]
    .forEach(n=>getOrCreateSheet(ss,n));

  const main=ss.getSheetByName(CONFIG.SHEET_NAME);
  main.setFrozenRows(1);
  main.getRange(1,1,1,HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#0f1923")
    .setFontColor("white");

  // จัด column width
  main.setColumnWidth(1,160);  // Timestamp
  main.setColumnWidth(2,100);  // Date
  main.setColumnWidth(11,100); // StdMs
  main.setColumnWidth(12,160); // VsStd

  // Format StdMs และ ElapsedMs columns
  main.getRange("J:J").setNumberFormat("0");
  main.getRange("K:K").setNumberFormat("0");

  // ลบ trigger เดิม แล้วสร้างใหม่
  ScriptApp.getProjectTriggers()
    .filter(t=>t.getHandlerFunction()==="dailyMaintenance")
    .forEach(t=>ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("dailyMaintenance").timeBased().atHour(2).everyDays(1).create();

  SpreadsheetApp.getUi().alert(
    "✅ Setup เสร็จสิ้น! (v4 Enhanced)\n\n" +
    "ขั้นตอนถัดไป:\n" +
    "1. Deploy → New Deployment → Web App\n" +
    "2. Execute as: Me\n" +
    "3. Access: Anyone\n" +
    "4. Copy Deployment URL ไปวางใน GAS_URL ของ index.html\n\n" +
    "เพิ่มใหม่:\n" +
    "• Column StdMs — เวลามาตรฐาน (ms)\n" +
    "• Column VsStd — เปรียบเทียบมาตรฐาน\n" +
    "• DailySummary: OverStdCount — รอบที่เกินมาตรฐาน"
  );
}

function dailyMaintenance() {
  console.log("Maintenance: " + new Date().toISOString());
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) console.log("TimerLog rows: " + sheet.getLastRow());
}

// ── Test ──────────────────────────────────────────────────────────
function testSaveTimerData() {
  const result = saveTimerData([{
    empName:   "สมชาย ใจดี",
    dept:      "โซนสกปรก (Zone A)",
    station:   "เชือด",
    startTime: "08:00:00",
    endTime:   "08:00:45",
    duration:  "00:00:45",
    elapsedMs: 45000,
    stdMs:     45000,
    remark:    "ทดสอบ",
  }]);
  console.log(JSON.stringify(result));
}

function testGetDashboard() {
  const result = getDashboardData(todayStr());
  console.log(JSON.stringify(result));
}
