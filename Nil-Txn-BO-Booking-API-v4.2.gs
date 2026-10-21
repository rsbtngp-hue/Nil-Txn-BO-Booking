/**
 * Nil Transaction BO + BO Booking Performance API — v4.2
 * Complete replacement for the current Apps Script Code.gs.
 *
 * Keeps existing sheet names, headers, and upload actions; adds:
 * - getDailySince / getBookingSince incremental reads
 * - getExclusions / saveExclusions shared across all dashboard users
 * - metadata in status for incremental-sync decisions
 *
 * After replacing the existing script, save and deploy a NEW version of the
 * existing web app. Keep the current /exec URL. No existing data rows are
 * cleared by setupSheets(); daily month replacement remains upload-driven.
 */
const MASTER_SHEET='Office Master';
const DAILY_SHEET='Daily Nil Transactions';
const DAILY_STAGING_SHEET='Daily Nil Transactions Upload Staging';
const BOOKING_SHEET='BO Booking Performance';
const BOOKING_STAGING_SHEET='BO Booking Performance Staging';
const EXCLUSIONS_SHEET='Dashboard Office Exclusions';

const MASTER_HEADERS=['office_id','office_name','office_type_code','sol_id','pli_id','valid_from','valid_to','delivery_office_flag','office_class','office_status','division_name','sub_division_name','ho_name','so_name'];
const DAILY_HEADERS=['office_id','office_name','transac_date','total_bo_offices','mail_nil_offices','mo_nil_offices','cbs_nil_offices','pli_nil_offices','stamp_nil_offices','ippb_nil_offices','uploaded_at'];
const BOOKING_KEYS=['circle_id','circle_name','region_id','region_name','division_id','division_name','sub_division_id','sub_division_name','office_id','office_name','booking_date','product_name','walk_in_customer','registered_bulk','on_postal_service','govt_service','walk_in_non_commercial','non_commercial','commercial','e_commerce','pidpi','student_mail','others','walk_in_commercial','walk_in_e_commerce','bulk_non_commercial','bulk_commercial','bulk_e_commerce','bo_rebooking','free_active_service','inter_fpo'];
const BOOKING_HEADERS=['Circle ID','Circle Name','Region ID','Region Name','Division ID','Division Name','Sub Division ID','Sub Division Name','Office ID','Office Name','Booking Date','Product Name','Walk-in Customer','Registered Bulk','On Postal Service','Govt Service','Walk-in Non-Commercial','Non-Commercial','Commercial','E-commerce','PIDPI','Student Mail','Others','Walk-in Commercial','Walk-in E-commerce','Bulk Non-Commercial','Bulk Commercial','Bulk E-commerce','BO Rebooking','Free Active Service','Inter FPO'];
const EXCLUSIONS_HEADERS=['office_id','office_name','excluded','updated_at'];

function doGet(e){
  try{
    const p=(e&&e.parameter)||{}, action=p.action||'health'; let result;
    if(action==='health'){
      const ss=SpreadsheetApp.getActiveSpreadsheet();
      result={ok:true,service:'Nil Transaction BO Detailed Report API',version:'4.3-safe-missing-daily-upload',spreadsheetId:ss.getId(),spreadsheetName:ss.getName(),sheets:[MASTER_SHEET,DAILY_SHEET,BOOKING_SHEET,EXCLUSIONS_SHEET],timestamp:new Date().toISOString()};
    }else if(action==='getMaster'||action==='downloadMaster'){
      result={ok:true,action:action,rows:readRowsAsObjects_(MASTER_SHEET,MASTER_HEADERS),count:Math.max(0,getRequiredSheet_(MASTER_SHEET).getLastRow()-1),timestamp:new Date().toISOString()};
    }else if(action==='getDaily'){
      const sh=getRequiredSheet_(DAILY_SHEET), total=Math.max(0,sh.getLastRow()-1);
      const rawOffset=Number(p.offset),rawLimit=Number(p.limit);
      const offset=Number.isFinite(rawOffset)?Math.max(0,Math.floor(rawOffset)):0;
      const limit=Number.isFinite(rawLimit)?Math.min(1000,Math.max(1,Math.floor(rawLimit))):500;
      const count=Math.max(0,Math.min(limit,total-offset)); let rows=[];
      if(count){rows=sh.getRange(offset+2,1,count,DAILY_HEADERS.length).getDisplayValues().map(r=>objectFromRow_(DAILY_HEADERS,r));}
      result={ok:true,action:action,rows:rows,count:rows.length,totalCount:total,offset:offset,limit:limit,hasMore:offset+rows.length<total,timestamp:new Date().toISOString()};
    }else if(action==='getDailySince'){
      result=getRowsSince_(DAILY_SHEET,DAILY_HEADERS,'transac_date',p.since,'getDailySince');
    }else if(action==='getBookingSince'){
      result=getRowsSince_(BOOKING_SHEET,BOOKING_HEADERS,'Booking Date',p.since,'getBookingSince',BOOKING_KEYS);
    }else if(action==='getBooking'||action==='downloadBooking'){
      const sh=getRequiredSheet_(BOOKING_SHEET), last=sh.getLastRow(); let rows=[];
      if(last>1) rows=sh.getRange(2,1,last-1,BOOKING_HEADERS.length).getDisplayValues().map(r=>objectFromRow_(BOOKING_KEYS,r));
      result={ok:true,action:action,rows:rows,count:rows.length,totalCount:rows.length,timestamp:new Date().toISOString()};
    }else if(action==='getExclusions'){
      result=getExclusions_();
    }else if(action==='status'||action==='verify'){
      result=getStatus_();
    }else{
      result={ok:false,error:'Unsupported GET action: '+action};
    }
    return jsonOutput_(result);
  }catch(err){return jsonOutput_({ok:false,error:String(err&&err.message||err)});}
}

function doPost(e){
  try{
    if(!e||!e.postData||!e.postData.contents) throw new Error('Missing POST JSON body.');
    const req=JSON.parse(e.postData.contents); let result;
    if(req.action==='beginSync') result=beginSync_(req.type,req.months);
    else if(req.action==='saveBatch') result=saveBatch_(req.type,normalizeRows_(req.rows));
    else if(req.action==='commitBooking') result=commitBooking_();
    else if(req.action==='commitDailyMissing') result=commitDailyMissing_();
    else if(req.action==='saveExclusions') result=saveExclusions_(req.excludedOfficeIds);
    else throw new Error('Unsupported action: '+String(req.action||''));
    return jsonOutput_({ok:true,result:result});
  }catch(err){return jsonOutput_({ok:false,error:String(err&&err.message||err)});}
}

function setupSheets(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss,MASTER_SHEET,MASTER_HEADERS);
  ensureSheet_(ss,DAILY_SHEET,DAILY_HEADERS);
  ensureSheet_(ss,BOOKING_SHEET,BOOKING_HEADERS);
  const staging=ensureSheet_(ss,BOOKING_STAGING_SHEET,BOOKING_HEADERS);
  const dailyStaging=ensureSheet_(ss,DAILY_STAGING_SHEET,DAILY_HEADERS);
  try{dailyStaging.hideSheet();}catch(e){}
  try{staging.hideSheet();}catch(e){}
  ensureExclusionsSheet_();
  return 'Google Sheets storage is ready, including shared dashboard exclusions.';
}
function beginSync(type,months){return beginSync_(type,months);}
function saveBatch(type,rows){return saveBatch_(type,normalizeRows_(rows));}
function commitBooking(){return commitBooking_();}
function commitDailyMissing(){return commitDailyMissing_();}

function normalizeRows_(rows){
  if(Array.isArray(rows))return rows;
  if(typeof rows==='string'){const t=rows.trim();if(!t)return[];try{rows=JSON.parse(t);}catch(e){throw new Error('rows was text but not valid JSON. Send rows as an array.');}}
  if(rows&&typeof rows==='object'){if(Array.isArray(rows.rows))return rows.rows;if(Array.isArray(rows.data))return rows.data;}
  if(rows===undefined||rows===null)throw new Error('Missing rows array. Pass rows:[...] to saveBatch.');
  throw new Error('Invalid rows payload. Expected an array, JSON array string, or object containing rows/data array.');
}

function beginSync_(type,months){
  if(type==='daily_missing'){
    const staging=ensureDailyStaging_(),last=ensureDailyStaging_().getLastRow();
    if(last>1)staging.getRange(2,1,last-1,DAILY_HEADERS.length).clearContent();
    return{sheet:staging.getName(),clearedRows:Math.max(0,last-1),scope:'staged missing daily rows; live data preserved until merge'};
  }
  if(type==='booking'){
    const staging=getRequiredSheet_(BOOKING_STAGING_SHEET),last=staging.getLastRow();
    if(last>1)staging.getRange(2,1,last-1,BOOKING_HEADERS.length).clearContent();
    return{sheet:staging.getName(),clearedRows:Math.max(0,last-1),scope:'new booking snapshot staged; live sheet unchanged until commit'};
  }
  const sh=getTargetSheet_(type),lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const last=sh.getLastRow();
    if(type==='master'){
      if(last>1)sh.getRange(2,1,last-1,sh.getLastColumn()).clearContent();
      stampSync_('master');
      return{sheet:sh.getName(),clearedRows:Math.max(0,last-1),scope:'all master rows'};
    }
    const monthSet=new Set((Array.isArray(months)?months:[]).map(m=>String(m).slice(0,7)).filter(Boolean));
    if(!monthSet.size)return{sheet:sh.getName(),clearedRows:0,scope:'no month supplied; retained existing daily data'};
    if(last<=1)return{sheet:sh.getName(),clearedRows:0,months:Array.from(monthSet)};
    const dates=sh.getRange(2,3,last-1,1).getDisplayValues(), toDelete=[];
    dates.forEach((r,i)=>{const d=normalizeSheetDate_(r[0]);if(d&&monthSet.has(d.slice(0,7)))toDelete.push(i+2);});
    const blocks=[];toDelete.forEach(rowNum=>{const b=blocks[blocks.length-1];if(b&&rowNum===b.start+b.count)b.count++;else blocks.push({start:rowNum,count:1});});
    blocks.reverse().forEach(b=>sh.deleteRows(b.start,b.count));stampSync_('daily');
    return{sheet:sh.getName(),clearedRows:toDelete.length,months:Array.from(monthSet),scope:'uploaded months only'};
  }finally{lock.releaseLock();}
}

function saveBatch_(type,rows){
  if(!Array.isArray(rows))rows=normalizeRows_(rows);
  if(!rows.length)return{saved:0};
  if(rows.length>500)throw new Error('Batch too large. Send no more than 500 rows per request.');
  const isMaster=type==='master',isBooking=type==='booking',isDailyMissing=type==='daily_missing';
  const sh=isBooking?getRequiredSheet_(BOOKING_STAGING_SHEET):(isDailyMissing?ensureDailyStaging_():getTargetSheet_(type));
  const width=isMaster?MASTER_HEADERS.length:(isBooking?BOOKING_HEADERS.length:DAILY_HEADERS.length),now=new Date();
  const values=rows.map(r=>{
    r=r||{};
    if(isMaster)return['office_id','office_name','office_type_code','sol_id','pli_id','valid_from','valid_to','delivery_office_flag','office_class','office_status','division_name','sub_division_name','ho_name','so_name'].map(k=>safeCell_(r[k]));
    if(isBooking)return BOOKING_KEYS.map(k=>safeCell_(r[k]));
    return[safeCell_(r.office_id),safeCell_(r.office_name),safeCell_(r.transac_date),safeCell_(r.total_bo_offices),safeCell_(r.mail_nil_offices),safeCell_(r.mo_nil_offices),safeCell_(r.cbs_nil_offices),safeCell_(r.pli_nil_offices),safeCell_(r.stamp_nil_offices),safeCell_(r.ippb_nil_offices),now];
  });
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const start=Math.max(2,sh.getLastRow()+1);sh.getRange(start,1,values.length,width).setValues(values);
    if(!isMaster&&!isBooking){sh.getRange(start,3,values.length,1).setNumberFormat('@');sh.getRange(start,11,values.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');}
    if(!isBooking&&!isDailyMissing)stampSync_(type);SpreadsheetApp.flush();
    return{sheet:sh.getName(),saved:values.length,startRow:start};
  }finally{lock.releaseLock();}
}

/**
 * Merge staged local Daily Nil Transactions into the live sheet without deleting
 * historical data. Existing rows are identified by office_id + transac_date.
 * Only keys not already present in the live sheet are appended.
 */
function commitDailyMissing_(){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const live=getRequiredSheet_(DAILY_SHEET),stage=ensureDailyStaging_();
    const liveCount=Math.max(0,live.getLastRow()-1),stageCount=Math.max(0,stage.getLastRow()-1);
    const existing=new Set();
    if(liveCount){
      live.getRange(2,1,liveCount,3).getDisplayValues().forEach(r=>{
        const id=String(r[0]||'').trim(),date=normalizeSheetDate_(r[2]);
        if(id&&date)existing.add(id+'|'+date);
      });
    }
    const staged=stageCount?stage.getRange(2,1,stageCount,DAILY_HEADERS.length).getValues():[];
    const append=[],seenStage=new Set();
    staged.forEach(r=>{
      const id=String(r[0]||'').trim(),date=normalizeSheetDate_(r[2]);
      if(!id||!date)return;
      const key=id+'|'+date;
      if(existing.has(key)||seenStage.has(key))return;
      seenStage.add(key);append.push(r);
    });
    if(append.length){
      const start=Math.max(2,live.getLastRow()+1);
      live.getRange(start,1,append.length,DAILY_HEADERS.length).setValues(append);
      live.getRange(start,3,append.length,1).setNumberFormat('@');
      live.getRange(start,11,append.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }
    if(stage.getLastRow()>1)stage.getRange(2,1,stage.getLastRow()-1,DAILY_HEADERS.length).clearContent();
    stampSync_('daily');SpreadsheetApp.flush();
    return{sheet:DAILY_SHEET,stagedRows:stageCount,insertedRows:append.length,skippedExisting:stageCount-append.length,timestamp:new Date().toISOString()};
  }finally{lock.releaseLock();}
}

function commitBooking_(){
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const staging=getRequiredSheet_(BOOKING_STAGING_SHEET),live=getRequiredSheet_(BOOKING_SHEET);
    const count=Math.max(0,staging.getLastRow()-1),liveLast=live.getLastRow();
    if(liveLast>1)live.getRange(2,1,liveLast-1,BOOKING_HEADERS.length).clearContent();
    if(count){const v=staging.getRange(2,1,count,BOOKING_HEADERS.length).getValues();live.getRange(2,1,v.length,BOOKING_HEADERS.length).setValues(v);}
    SpreadsheetApp.flush();PropertiesService.getScriptProperties().setProperty('lastBookingWrite',new Date().toISOString());
    return{sheet:BOOKING_SHEET,committedRows:count,timestamp:new Date().toISOString()};
  }finally{lock.releaseLock();}
}

/** Date-filtered incremental read. The latest local date is inclusive. */
function getRowsSince_(sheetName,headers,dateHeader,sinceValue,action,objectKeys){
  const since=normalizeSheetDate_(sinceValue);
  if(!/^\d{4}-\d{2}-\d{2}$/.test(since))throw new Error(action+' requires since=YYYY-MM-DD.');
  const sh=getRequiredSheet_(sheetName),last=sh.getLastRow(),total=Math.max(0,last-1);
  if(!total)return{ok:true,action:action,since:since,rows:[],count:0,totalCount:0,timestamp:new Date().toISOString()};
  const dateCol=headers.indexOf(dateHeader)+1;
  if(dateCol<1)throw new Error('Date column not configured for '+sheetName);
  const dates=sh.getRange(2,dateCol,total,1).getDisplayValues(),rowNums=[];
  dates.forEach((r,i)=>{const d=normalizeSheetDate_(r[0]);if(d&&d>=since)rowNums.push(i+2);});
  const blocks=[];rowNums.forEach(n=>{const b=blocks[blocks.length-1];if(b&&n===b.start+b.count)b.count++;else blocks.push({start:n,count:1});});
  const keys=objectKeys||headers,rows=[];
  blocks.forEach(b=>sh.getRange(b.start,1,b.count,headers.length).getDisplayValues().forEach(r=>rows.push(objectFromRow_(keys,r))));
  return{ok:true,action:action,since:since,rows:rows,count:rows.length,totalCount:total,timestamp:new Date().toISOString()};
}

function ensureExclusionsSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(EXCLUSIONS_SHEET);
  if(!sh)sh=ss.insertSheet(EXCLUSIONS_SHEET);
  if(sh.getLastRow()===0)sh.getRange(1,1,1,EXCLUSIONS_HEADERS.length).setValues([EXCLUSIONS_HEADERS]);
  const got=sh.getRange(1,1,1,EXCLUSIONS_HEADERS.length).getDisplayValues()[0];
  if(!EXCLUSIONS_HEADERS.every((h,i)=>String(got[i]||'').trim()===h))throw new Error('Header mismatch in "'+EXCLUSIONS_SHEET+'". Expected: '+EXCLUSIONS_HEADERS.join(', '));
  sh.setFrozenRows(1);sh.getRange(1,1,1,EXCLUSIONS_HEADERS.length).setFontWeight('bold').setBackground('#dbeafe').setWrap(true);
  return sh;
}
function getExclusions_(){
  const sh=ensureExclusionsSheet_(),last=sh.getLastRow(),ids=[];
  if(last>1)sh.getRange(2,1,last-1,3).getDisplayValues().forEach(r=>{const id=String(r[0]||'').trim(),v=String(r[2]||'').trim().toLowerCase();if(id&&['true','yes','1'].includes(v))ids.push(id);});
  return{ok:true,action:'getExclusions',excludedOfficeIds:Array.from(new Set(ids)),count:new Set(ids).size,timestamp:new Date().toISOString()};
}
function saveExclusions_(ids){
  if(!Array.isArray(ids))throw new Error('saveExclusions requires excludedOfficeIds as an array.');
  const unique=Array.from(new Set(ids.map(x=>String(x||'').trim()).filter(Boolean))),lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const sh=ensureExclusionsSheet_(),masterById=new Map();
    readRowsAsObjects_(MASTER_SHEET,MASTER_HEADERS).forEach(o=>{const id=String(o.office_id||'').trim();if(id)masterById.set(id,String(o.office_name||''));});
    const now=new Date(),values=unique.map(id=>[id,masterById.get(id)||'',true,now]),last=sh.getLastRow();
    if(last>1)sh.getRange(2,1,last-1,EXCLUSIONS_HEADERS.length).clearContent();
    if(values.length){sh.getRange(2,1,values.length,EXCLUSIONS_HEADERS.length).setValues(values);sh.getRange(2,4,values.length,1).setNumberFormat('yyyy-mm-dd hh:mm:ss');}
    SpreadsheetApp.flush();return{sheet:EXCLUSIONS_SHEET,excludedCount:values.length,timestamp:now.toISOString()};
  }finally{lock.releaseLock();}
}

function getStatus_(){
  const ms=getRequiredSheet_(MASTER_SHEET),ds=getRequiredSheet_(DAILY_SHEET),bs=getRequiredSheet_(BOOKING_SHEET);
  const masterRows=Math.max(0,ms.getLastRow()-1),dailyRows=Math.max(0,ds.getLastRow()-1),bookingRows=Math.max(0,bs.getLastRow()-1);
  const monthCounts={};let latestTransactionDate='',latestBookingDate='';
  if(dailyRows){ds.getRange(2,3,dailyRows,1).getDisplayValues().forEach(r=>{const d=normalizeSheetDate_(r[0]);if(!d)return;const m=d.slice(0,7);monthCounts[m]=(monthCounts[m]||0)+1;if(d>latestTransactionDate)latestTransactionDate=d;});}
  if(bookingRows){bs.getRange(2,11,bookingRows,1).getDisplayValues().forEach(r=>{const d=normalizeSheetDate_(r[0]);if(d&&d>latestBookingDate)latestBookingDate=d;});}
  const props=PropertiesService.getScriptProperties();
  return{ok:true,service:'Nil Transaction BO Detailed Report API',version:'4.3-safe-missing-daily-upload',spreadsheetName:SpreadsheetApp.getActiveSpreadsheet().getName(),masterRows:masterRows,dailyRows:dailyRows,bookingRows:bookingRows,monthCounts:monthCounts,latestTransactionDate:latestTransactionDate,latestBookingDate:latestBookingDate,lastMasterWrite:props.getProperty('lastMasterWrite')||'',lastDailyWrite:props.getProperty('lastDailyWrite')||'',lastBookingWrite:props.getProperty('lastBookingWrite')||'',timestamp:new Date().toISOString()};
}

function readRowsAsObjects_(sheetName,headers){
  const sh=getRequiredSheet_(sheetName),last=sh.getLastRow();if(last<=1)return[];
  const got=sh.getRange(1,1,1,headers.length).getDisplayValues()[0];
  if(!headers.every((h,i)=>String(got[i]||'').trim()===h))throw new Error('Header mismatch in "'+sheetName+'". Run setupSheets() and check column order.');
  return sh.getRange(2,1,last-1,headers.length).getDisplayValues().map(r=>objectFromRow_(headers,r));
}
function objectFromRow_(headers,row){const o={};headers.forEach((h,i)=>o[h]=row[i]===undefined?'':row[i]);return o;}
function normalizeSheetDate_(v){
  if(v===null||v===undefined||v==='')return'';
  if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const s=String(v).trim().split('T')[0].split(' ')[0].replace(/\//g,'-'),p=s.split('-');
  if(p.length===3&&p[0].length===4)return p[0]+'-'+p[1].padStart(2,'0')+'-'+p[2].padStart(2,'0');
  if(p.length===3&&p[2].length===4)return p[2]+'-'+p[1].padStart(2,'0')+'-'+p[0].padStart(2,'0');
  return s;
}
function getTargetSheet_(type){
  if(!['master','daily','booking'].includes(type))throw new Error('Invalid type. Expected "master", "daily", or "booking".');
  return getRequiredSheet_(type==='master'?MASTER_SHEET:type==='daily'?DAILY_SHEET:BOOKING_SHEET);
}
function ensureDailyStaging_(){const sh=ensureSheet_(SpreadsheetApp.getActiveSpreadsheet(),DAILY_STAGING_SHEET,DAILY_HEADERS);try{sh.hideSheet();}catch(e){}return sh;}
function getRequiredSheet_(name){const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);if(!sh)throw new Error('Sheet "'+name+'" not found. Run setupSheets() first.');return sh;}
function ensureSheet_(ss,name,headers){
  let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);
  if(sh.getLastRow()===0){
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }else{
    const got=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getDisplayValues()[0].map(v=>String(v||'').trim());
    const exact=headers.every((h,i)=>got[i]===h);
    if(!exact && (name===BOOKING_SHEET || name===BOOKING_STAGING_SHEET)){
      repairBookingSheetColumns_(sh,headers,got);
    }else if(!exact){
      throw new Error('Header mismatch in "'+name+'". Expected: '+headers.join(', ')+'. Actual: '+got.join(', '));
    }
  }
  sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#dbeafe').setWrap(true);return sh;
}

/**
 * Repairs a Booking sheet whose columns were reordered or whose header labels
 * differ slightly. Data is remapped by header name, not by column position.
 * Unknown columns are not included in the canonical schema; a backup is made
 * before rewriting so the original can be recovered if needed.
 */
function repairBookingSheetColumns_(sh,headers,actualHeaders){
  const lastRow=sh.getLastRow(),lastCol=Math.max(sh.getLastColumn(),actualHeaders.length);
  const normalize=h=>String(h||'').toLowerCase().replace(/[^a-z0-9]/g,'');
  const aliases={
    'circleid':'circle_id','circlename':'circle_name','regionid':'region_id','regionname':'region_name',
    'divisionid':'division_id','divisionname':'division_name','subdivisionid':'sub_division_id',
    'subdivisionname':'sub_division_name','officeid':'office_id','officename':'office_name',
    'bookingdate':'booking_date','date':'booking_date','productname':'product_name',
    'walkincustomer':'walk_in_customer','registeredbulk':'registered_bulk','onpostalservice':'on_postal_service',
    'govtservice':'govt_service','walkinnoncommercial':'walk_in_non_commercial','noncommercial':'non_commercial',
    'commercial':'commercial','ecommerce':'e_commerce','pidpi':'pidpi','studentmail':'student_mail',
    'others':'others','walkincommercial':'walk_in_commercial','walkinecommerce':'walk_in_e_commerce',
    'bulknoncommercial':'bulk_non_commercial','bulkcommercial':'bulk_commercial','bulkecommerce':'bulk_e_commerce',
    'borebooking':'bo_rebooking','freeactiveservice':'free_active_service','interfpo':'inter_fpo'
  };
  const canonicalByNorm={};headers.forEach(h=>canonicalByNorm[normalize(h)]=h);
  const sourceIndex={};
  actualHeaders.forEach((h,i)=>{
    const n=normalize(h),key=aliases[n]||canonicalByNorm[n];
    const canonicalHeader=BOOKING_KEYS.includes(key)?headers[BOOKING_KEYS.indexOf(key)]:key;
    if(canonicalHeader && sourceIndex[canonicalHeader]===undefined)sourceIndex[canonicalHeader]=i;
  });
  const values=lastRow>1?sh.getRange(2,1,lastRow-1,lastCol).getValues():[];
  const mapped=values.map(row=>headers.map(h=>sourceIndex[h]===undefined?'':row[sourceIndex[h]]));
  // Make a backup only when data exists and the header schema actually differs.
  if(lastRow>1){
    const ss=sh.getParent(),backupName=sh.getName()+' Backup '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
    const backup=ss.insertSheet(backupName);
    sh.getRange(1,1,lastRow,lastCol).copyTo(backup.getRange(1,1),{contentsOnly:false});
  }
  sh.clearContents();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  if(mapped.length)sh.getRange(2,1,mapped.length,headers.length).setValues(mapped);
  SpreadsheetApp.flush();
}
function safeCell_(v){if(v===null||v===undefined)return'';if(typeof v==='string'&&/^[=+@]/.test(v))return"'"+v;return v;}
function stampSync_(type){PropertiesService.getScriptProperties().setProperty(type==='master'?'lastMasterWrite':'lastDailyWrite',new Date().toISOString());}
function jsonOutput_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
