// Standalone Google Apps Script. Run setupArchive once, then deploy as a web app.
const HEADERS=['Reading ID','Created UTC','Updated UTC','Revision','Question','Present','Crossing','Foundation','Recent past','Crown / goal','Near future','Self','External influences','Hopes / fears','Outcome','Interpretation'];
function setupArchive(){
  const props=PropertiesService.getScriptProperties();
  const old=props.getProperty('ARCHIVE_ID');
  if(old){console.log('Private archive: '+SpreadsheetApp.openById(old).getUrl());return}
  const book=SpreadsheetApp.create('Virtual Tarot — Shared Readings');
  const sh=book.getSheets()[0];sh.setName('Readings');sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#271d38').setFontColor('#f5e2ba');sh.setFrozenRows(1);sh.setColumnWidth(1,245);sh.setColumnWidths(2,2,175);sh.setColumnWidth(4,80);sh.setColumnWidth(5,320);sh.setColumnWidths(6,10,145);sh.setColumnWidth(16,440);sh.getRange('E:P').setWrap(true).setVerticalAlignment('top').setNumberFormat('@');
  props.setProperty('ARCHIVE_ID',book.getId());console.log('Private archive: '+book.getUrl());
}
function result_(ok){return ContentService.createTextOutput(JSON.stringify({ok:ok})).setMimeType(ContentService.MimeType.JSON)}
function doGet(){return result_(true)}
function literal_(v){return /^[\s]*[=+\-@]/.test(v)?"'"+v:v}
function validate_(x){return x&&x.consent===true&&x.version===1&&typeof x.id==='string'&&/^[a-f0-9-]{36}$/.test(x.id)&&Number.isSafeInteger(x.rev)&&x.rev>0&&x.rev<=1000000&&typeof x.question==='string'&&x.question.length<=240&&typeof x.notes==='string'&&x.notes.length<=4000&&Array.isArray(x.cards)&&x.cards.length===10&&new Set(x.cards).size===10&&x.cards.every(c=>typeof c==='string'&&c.length>0&&c.length<=80)}
function doPost(e){
  let lock;
  try{
    const raw=e&&e.postData&&e.postData.contents;
    if(typeof raw!=='string'||raw.length>16000)return result_(false);
    const x=JSON.parse(raw);if(!validate_(x))return result_(false);
    const id=PropertiesService.getScriptProperties().getProperty('ARCHIVE_ID');if(!id)return result_(false);
    lock=LockService.getScriptLock();if(!lock.tryLock(8000))return result_(false);
    const sh=SpreadsheetApp.openById(id).getSheetByName('Readings');if(!sh)return result_(false);
    const end=sh.getLastRow(),found=end>1?sh.getRange(2,1,end-1,1).createTextFinder(x.id).matchEntireCell(true).findNext():null;
    const row=found?found.getRow():end+1;if(found&&Number(sh.getRange(row,4).getValue())>=x.rev)return result_(true);
    if(row>sh.getMaxRows())sh.insertRowsAfter(sh.getMaxRows(),100);
    const now=new Date().toISOString(),created=found?sh.getRange(row,2).getValue():now;
    sh.getRange(row,1,1,HEADERS.length).setValues([[x.id,created,now,x.rev,literal_(x.question),...x.cards.map(literal_),literal_(x.notes)]]);
    SpreadsheetApp.flush();return result_(true);
  }catch(err){console.error('Archive write failed');return result_(false)}finally{if(lock&&lock.hasLock())lock.releaseLock()}
}
