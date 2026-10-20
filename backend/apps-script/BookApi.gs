/** Trade Log -> Life app "Book" feed. Paste into the Trade Log sheet: Extensions > Apps Script.
 *  Read-only. Returns LIVE trades + headline Book metrics as JSON when ?token= matches. */
function onOpen(){ SpreadsheetApp.getUi().createMenu('Book app').addItem('Phone link (set up / show)','showPhoneLink').addItem('Rotate token','rotateToken').addToUi(); }
function getToken_(){ var p=PropertiesService.getScriptProperties(), t=p.getProperty('BOOK_TOKEN'); if(!t){ t=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''); p.setProperty('BOOK_TOKEN',t);} return t; }
function rotateToken(){ PropertiesService.getScriptProperties().deleteProperty('BOOK_TOKEN'); showPhoneLink(); }
function showPhoneLink(){
  var url=ScriptApp.getService().getUrl(), ui=SpreadsheetApp.getUi();
  if(!url){ ui.alert('Deploy first: Deploy > New deployment > Web app (Execute as: Me, Access: Anyone). Then run this again.'); return; }
  var link='https://oliverrising96.github.io/life-app/#book='+encodeURIComponent(url+'|'+getToken_());
  ui.showModalDialog(HtmlService.createHtmlOutput('<p>Open this on your phone once:</p><textarea style="width:100%;height:120px">'+link+'</textarea>'),'Book phone link');
}
function doGet(e){
  var out;
  if(!e||!e.parameter||e.parameter.token!==getToken_()) out={error:'unauthorised'};
  else out=bookJson_();
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function bookJson_(){
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('Trades'), v=sh.getDataRange().getValues(), h=v[0], ix={};
  h.forEach(function(n,i){ ix[n]=i; });
  function g(r,n){ var x=r[ix[n]]; return x instanceof Date ? Utilities.formatDate(x,'Europe/London','yyyy-MM-dd') : (x===''?null:x); }
  var keys=['ID','Instrument','GF ticker','Asset class','Region','Type','Direction','Ccy','Entry','Stop','Target','Mark','Price used','Mark source','Risk £','Risk % NAV','Notional % NAV','R:R','P&L £','R multiple','P&L % NAV','Days held','Opened','Notion rationale','Notes'];
  var live=v.slice(1).filter(function(r){ return r[ix['Status']]==='Live' && r[ix['Instrument']]!==''; }).map(function(r){ var o={}; keys.forEach(function(k){ o[k]=g(r,k); }); return o; });
  var b=ss.getSheetByName('Book').getRange('A2:B20').getValues(), m={};
  b.forEach(function(r){ if(r[0]) m[r[0]]=r[1]; });
  return {asOf:new Date().toISOString(), book:m, live:live};
}
