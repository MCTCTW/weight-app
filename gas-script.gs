/**
 * 金洲工作室 · 重量核算 — 雲端後端 (Google Apps Script)
 *
 * 設計：app 端算、GAS 只存。upsert-by-key（非純 append，避免重複列）。
 * 四個分頁：
 *   _data     批次原始 JSON（給 app 重新載入用；每批一列，key=batchId）
 *   主明細     每件商品一列（會計唯一資料源；replace-by-batch）
 *   批次彙總   每批一列（只彙總明細；key=batchId）
 *   應收客人   有補款/代運才列（replace-by-batch）
 *   修改軌跡   每次改動一列（誰/何時/把哪件的什麼從X改成Y；append-only 永不覆蓋，供追責）
 *
 * 部署：擴充功能→Apps Script→貼上→部署→新增部署→類型「網頁應用程式」
 *   執行身分＝我、誰可存取＝所有人→部署→授權→複製網址(/exec)貼回工具。
 */

function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function tab(name){ var s = ss().getSheetByName(name); if(!s) s = ss().insertSheet(name); return s; }
function json(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- 讀取 ----------
function doGet(e){
  try{
    var action = (e && e.parameter && e.parameter.action) || '';
    if(action === 'batches'){
      var sh = tab('_data');
      var vals = sh.getDataRange().getValues();
      var out = [];
      for(var i=1;i<vals.length;i++){
        if(vals[i][0]){ try{ out.push(JSON.parse(vals[i][1])); }catch(_){} }
      }
      return json({ success:true, batches: out });
    }
    if(action === 'audit'){
      var ash = tab('修改軌跡');
      var av = ash.getDataRange().getValues();
      var head = av.length ? av[0] : [];
      var rows = [];
      var start = Math.max(1, av.length - 50);   // 最近 50 筆
      for(var j=start;j<av.length;j++){
        var o = {}; for(var k=0;k<head.length;k++) o[head[k]] = av[j][k];
        rows.push(o);
      }
      return json({ success:true, count: Math.max(0, av.length-1), recent: rows });
    }
    return json({ success:true, ping:true, ts: new Date().toISOString() });
  }catch(err){ return json({ success:false, error: err.message }); }
}

// ---------- 寫入（POST，text/plain 免 CORS preflight）----------
function doPost(e){
  try{
    var p = JSON.parse(e.postData.contents);

    // 修改軌跡（append-only，永不覆蓋；save 與 delete 都會帶）
    if(p.audit && p.audit.length) appendRows(tab('修改軌跡'), p.audit);

    if(p.deleteBatch){
      removeByKey(tab('_data'), 'batchId', p.deleteBatch);
      replaceByBatch(tab('主明細'), [], p.deleteBatch);
      removeByKey(tab('批次彙總'), 'batchId', p.deleteBatch);
      replaceByBatch(tab('應收客人'), [], p.deleteBatch);
      return json({ success:true, deleted: p.deleteBatch });
    }

    if(!p.batchId) return json({ success:false, error:'missing batchId' });

    // _data：批次原始 JSON（upsert by batchId）
    upsertRow(tab('_data'), ['batchId','json'], 'batchId',
      { batchId: p.batchId, json: JSON.stringify(p.json || {}) });

    // 主明細 / 應收客人：replace-by-batch（先刪該批舊列、再寫新列，處理增刪改）
    replaceByBatch(tab('主明細'), p.detail || [], p.batchId);
    replaceByBatch(tab('應收客人'), p.ar || [], p.batchId);

    // 批次彙總：每批一列（upsert by batchId）
    if(p.summary) upsertRow(tab('批次彙總'), keysOf(p.summary), 'batchId', p.summary);

    return json({ success:true });
  }catch(err){ return json({ success:false, error: err.message }); }
}

// ---------- helpers ----------
function keysOf(obj){ return Object.keys(obj); }

// 確保表頭含 keys（缺的補在最後），回傳目前表頭陣列
function ensureHeaders(sheet, keys){
  var header = sheet.getLastRow() >= 1
    ? sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getValues()[0]
    : [];
  header = header.filter(function(h){ return h !== '' && h != null; });
  var changed = false;
  keys.forEach(function(k){ if(header.indexOf(k) < 0){ header.push(k); changed = true; } });
  if(changed || sheet.getLastRow() === 0){
    sheet.getRange(1,1,1,header.length).setValues([header]);
    sheet.getRange(1,1,1,header.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return header;
}

function rowFromObj(header, obj){
  return header.map(function(h){ return (h in obj) ? obj[h] : ''; });
}

// upsert 單列（key 欄相同就覆蓋，否則 append）
function upsertRow(sheet, keys, keyCol, obj){
  var header = ensureHeaders(sheet, keys);
  var ki = header.indexOf(keyCol);
  var last = sheet.getLastRow();
  var target = -1;
  if(last >= 2){
    var col = sheet.getRange(2, ki+1, last-1, 1).getValues();
    for(var r=0;r<col.length;r++){ if(String(col[r][0]) === String(obj[keyCol])){ target = r+2; break; } }
  }
  var row = rowFromObj(header, obj);
  if(target > 0) sheet.getRange(target, 1, 1, header.length).setValues([row]);
  else sheet.appendRow(row);
}

// 刪掉 keyCol == val 的列（單列）
function removeByKey(sheet, keyCol, val){
  var last = sheet.getLastRow(); if(last < 2) return;
  var header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var ki = header.indexOf(keyCol); if(ki < 0) return;
  var col = sheet.getRange(2, ki+1, last-1, 1).getValues();
  for(var r=col.length-1;r>=0;r--){ if(String(col[r][0]) === String(val)) sheet.deleteRow(r+2); }
}

// 純 append（不刪不覆蓋；供修改軌跡）
function appendRows(sheet, rows){
  if(!rows || !rows.length) return;
  var header = ensureHeaders(sheet, keysOf(rows[0]));
  var matrix = rows.map(function(o){ return rowFromObj(header, o); });
  sheet.getRange(sheet.getLastRow()+1, 1, matrix.length, header.length).setValues(matrix);
}

// 先刪該 batchId 所有列、再 append 新 rows（rows 為物件陣列，每列須含 batchId）
function replaceByBatch(sheet, rows, batchId){
  if(rows && rows.length) ensureHeaders(sheet, keysOf(rows[0]));
  removeByKey(sheet, 'batchId', batchId);
  if(!rows || !rows.length) return;
  var header = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var matrix = rows.map(function(o){ return rowFromObj(header, o); });
  sheet.getRange(sheet.getLastRow()+1, 1, matrix.length, header.length).setValues(matrix);
}
