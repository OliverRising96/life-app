/** Trade Log v2 -> Life app "Model Portfolio" feed.  Paste into the Trade Log sheet: Extensions > Apps Script (replace everything).
 *  Read-only feed. Returns ALL trades + Model Portfolio metrics as JSON when ?token= matches.
 *  One-off: Model Portfolio menu > "Upgrade sheet to v2 layout" rebuilds Trades + Model Portfolio tabs in place (old tab kept as "Trades (v1)"). */
var COLS=['Risk Opened','Trade Idea','Type','L/S','Conviction','GBP at risk','Entry','Stop Loss','Take Profit','Current Level','Exit Date','P&L (£)','P&L (%)','Notes','Notion link',
          'Status','R multiple','R:R (plan)','Days held'];
var NIN=15, N=500, NAV='Settings!$B$3';
function onOpen(){ SpreadsheetApp.getUi().createMenu('Model Portfolio').addItem('Phone link (set up / show)','showPhoneLink').addItem('Rotate token','rotateToken')
  .addSeparator().addItem('Upgrade sheet to v2 layout (one-off)','upgradeToV2').addToUi(); }
function getToken_(){ var p=PropertiesService.getScriptProperties(), t=p.getProperty('BOOK_TOKEN'); if(!t){ t=Utilities.getUuid().replace(/-/g,'')+Utilities.getUuid().replace(/-/g,''); p.setProperty('BOOK_TOKEN',t);} return t; }
function rotateToken(){ PropertiesService.getScriptProperties().deleteProperty('BOOK_TOKEN'); showPhoneLink(); }
function showPhoneLink(){
  var url=ScriptApp.getService().getUrl(), ui=SpreadsheetApp.getUi();
  if(!url){ ui.alert('Deploy first: Deploy > New deployment > Web app (Execute as: Me, Access: Anyone). Then run this again.'); return; }
  var link='https://oliverrising96.github.io/life-app/#book='+encodeURIComponent(url+'|'+getToken_());
  ui.showModalDialog(HtmlService.createHtmlOutput('<p>Open this on your phone once:</p><textarea style="width:100%;height:120px">'+link+'</textarea>'),'Model Portfolio phone link');
}

/* ---------- one-off migration: v1 (32 cols, GOOGLEFINANCE) -> v2 (15 inputs + helpers) ---------- */
function upgradeToV2(){
  var ss=SpreadsheetApp.getActive(), ui=SpreadsheetApp.getUi(), old=ss.getSheetByName('Trades');
  if(old && old.getRange(1,1).getValue()==='Risk Opened'){ ui.alert('Already on v2.'); return; }
  var moved=[], skipped=0, nav=ss.getSheetByName('Settings').getRange('B3').getValue()||1e6, defR=ss.getSheetByName('Settings').getRange('B4').getValue()||0.01;
  if(old){
    var v=old.getDataRange().getValues(), h=v[0], ix={}; h.forEach(function(n,i){ix[n]=i;});
    v.slice(1).forEach(function(r){
      if(r[ix['Instrument']]==='') return;
      var st=r[ix['Status']];
      if(st!=='Live' && st!=='Closed'){ skipped++; return; }           // Idea / Cancelled stay in "Trades (v1)"
      var ty=String(r[ix['Type']]||''), rv=/Relative|Curve|Spread/.test(ty)?'RV':'Dir';
      var cur = st==='Closed' ? r[ix['Exit price']] : (r[ix['Manual mark']]!==''?r[ix['Manual mark']]:r[ix['Mark']]);
      var exitD = st==='Closed' ? (r[ix['Closed']]||new Date()) : '';
      var rp=r[ix['Risk % NAV']]===''?defR:r[ix['Risk % NAV']];
      moved.push([r[ix['Opened']],r[ix['Instrument']],rv,r[ix['Direction']]||'Long','',nav*rp,r[ix['Entry']],r[ix['Stop']],r[ix['Target']],cur,exitD,'','',
                  r[ix['Notes']],r[ix['Notion rationale']]]);
    });
    old.setName('Trades (v1)');
  }
  var t=ss.insertSheet('Trades',0); buildTrades_(t);
  if(moved.length){ moved.forEach(function(m){m[11]=null;m[12]=null;});
    var rng=t.getRange(2,1,moved.length,NIN), vals=moved.map(function(m){return m.slice();});
    // write inputs except formula cols L,M
    t.getRange(2,1,moved.length,11).setValues(vals.map(function(m){return m.slice(0,11);}));
    t.getRange(2,14,moved.length,2).setValues(vals.map(function(m){return m.slice(13,15);}));
  }
  buildBook_(ss);
  ui.alert('Upgraded. Moved '+moved.length+' Live/Closed trade(s); '+skipped+' Idea/Cancelled row(s) left in "Trades (v1)". Delete that tab once you are happy.');
}
function buildTrades_(t){
  var H=t.getRange(1,1,1,COLS.length); H.setValues([COLS]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F2937').setWrap(true).setVerticalAlignment('middle');
  t.setFrozenRows(1); t.setFrozenColumns(2); t.setRowHeight(1,34);
  if(t.getMaxRows()<N+1) t.insertRowsAfter(t.getMaxRows(),N+1-t.getMaxRows());
  if(t.getMaxColumns()>COLS.length) t.deleteColumns(COLS.length+1,t.getMaxColumns()-COLS.length);
  [100,190,60,65,80,95,80,80,80,90,95,95,75,260,220,70,75,70,70].forEach(function(w,i){t.setColumnWidth(i+1,w);});
  t.getRange(2,1,N,NIN).setBackground('#FFF7D6'); t.getRange(2,12,N,2).setBackground('#EEF2F7'); t.getRange(2,16,N,4).setBackground('#EEF2F7');
  var f={L:'=IF(OR($B2="",$F2="",$G2="",$H2="",$J2="",$G2=$H2),"",$F2*($J2-$G2)/($G2-$H2))',
         M:'=IF($L2="","",$L2/'+NAV+')',
         P:'=IF($B2="","",IF($K2="","Live","Closed"))',
         Q:'=IF(OR($B2="",$G2="",$H2="",$J2="",$G2=$H2),"",($J2-$G2)/($G2-$H2))',
         R:'=IF(OR($B2="",$G2="",$H2="",$I2="",$G2=$H2),"",($I2-$G2)/($G2-$H2))',
         S:'=IF(OR($B2="",$A2=""),"",IF($K2<>"",$K2,TODAY())-$A2)'};
  for(var c in f) t.getRange(c+'2:'+c+(N+1)).setFormula(f[c]);
  t.getRange('A2:A'+(N+1)).setNumberFormat('dd-mmm-yy'); t.getRange('K2:K'+(N+1)).setNumberFormat('dd-mmm-yy');
  t.getRange('F2:F'+(N+1)).setNumberFormat('£#,##0'); t.getRange('L2:L'+(N+1)).setNumberFormat('£#,##0;[Red]-£#,##0');
  t.getRange('M2:M'+(N+1)).setNumberFormat('0.00%;[Red]-0.00%'); t.getRange('Q2:R'+(N+1)).setNumberFormat('0.00'); t.getRange('S2:S'+(N+1)).setNumberFormat('0');
  var dv=function(a,list){ t.getRange(a).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(list,true).setAllowInvalid(false).build()); };
  dv('C2:C'+(N+1),['Dir','RV']); dv('D2:D'+(N+1),['Long','Short']); dv('E2:E'+(N+1),['1','2','3','4','5']);
  var dd=SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build(); t.getRange('A2:A'+(N+1)).setDataValidation(dd); t.getRange('K2:K'+(N+1)).setDataValidation(dd);
  t.getRange('F2:F'+(N+1)).setDataValidation(SpreadsheetApp.newDataValidation().requireFormulaSatisfied('=AND(F2>0,F2<=Settings!$B$3*Settings!$B$5)')
    .setAllowInvalid(false).setHelpText('£ at risk must be >0 and within the per-trade cap (Settings B5 x NAV).').build());
  var rules=[
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$P2="Live"').setBackground('#DCFCE7').setRanges([t.getRange('P2:P'+(N+1))]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$P2="Closed"').setBackground('#E5E7EB').setRanges([t.getRange('P2:P'+(N+1))]).build(),
    // stop on the wrong side for the L/S you picked
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($G2<>"",$H2<>"",OR(AND($D2="Long",$H2>=$G2),AND($D2="Short",$H2<=$G2)))').setBackground('#FECACA').setRanges([t.getRange('H2:H'+(N+1))]).build()];
  t.setConditionalFormatRules(rules);
  t.getRange('J1').setNote('Update daily. When you close, leave your exit level here and fill Exit Date.');
  t.getRange('C1').setNote('Dir = one leg. RV = log the spread/ratio as ONE level (e.g. 2s10s in bp, FTSE/SX5E ratio). Legs go in Notes.');
  t.getRange('L1').setNote('Auto: GBP at risk x (Current - Entry) / (Entry - Stop). Works for longs and shorts.');
}
function buildBook_(ss){
  var b=ss.getSheetByName('Model Portfolio')||ss.getSheetByName('Book'); if(b){ b.clear(); b.setName('Model Portfolio'); } else b=ss.insertSheet('Model Portfolio',1);
  var L=N+1, P='Trades!$P$2:$P$'+L, S=function(c,st){return 'SUMIFS(Trades!$'+c+'$2:$'+c+'$'+L+','+P+',"'+st+'")';};
  var items=[['Starting NAV (£)','='+NAV,'£#,##0'],['Realised P&L (£)','='+S('L','Closed'),'£#,##0;[Red]-£#,##0'],['Unrealised P&L (£)','='+S('L','Live'),'£#,##0;[Red]-£#,##0'],
   ['Current NAV (£)','=B2+B3+B4','£#,##0'],['Return on NAV','=B5/B2-1','0.00%'],['Live trades','=COUNTIF('+P+',"Live")','0'],['Closed trades','=COUNTIF('+P+',"Closed")','0'],
   ['Open risk (£)','='+S('F','Live'),'£#,##0'],['Open risk (% NAV)','=B9/B2','0.00%'],['Open risk limit (% NAV)','=Settings!B6','0.00%'],['Risk headroom (% NAV)','=B11-B10','0.00%'],
   ['Largest single-trade risk (% NAV)','=IFERROR(MAXIFS(Trades!$F$2:$F$'+L+','+P+',"Live")/B2,0)','0.00%'],['Per-trade cap (% NAV)','=Settings!B5','0.00%'],
   ['Limit status','=IF(OR(B10>B11,B13>B14),"BREACH","OK")','@'],
   ['Hit rate (closed)','=IFERROR(COUNTIFS('+P+',"Closed",Trades!$L$2:$L$'+L+',">0")/B8,"")','0%'],['Avg R (closed)','=IFERROR(AVERAGEIFS(Trades!$Q$2:$Q$'+L+','+P+',"Closed"),"")','0.00'],
   ['Total R (closed)','=SUMIFS(Trades!$Q$2:$Q$'+L+','+P+',"Closed")','0.00']];
  b.getRange('A1:B1').setValues([['Metric','Value']]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F2937');
  items.forEach(function(it,i){ b.getRange(i+2,1).setValue(it[0]); b.getRange(i+2,2).setFormula(it[1]).setNumberFormat(it[2]); });
  b.setColumnWidth(1,240); b.getRange('D1').setValue('Live trades (auto)').setFontWeight('bold');
  b.getRange(2,4,1,COLS.length).setValues([COLS]).setFontWeight('bold').setFontColor('#ffffff').setBackground('#1F2937');
  b.getRange('D3').setFormula('=IFERROR(FILTER(Trades!A2:S'+L+','+P+'="Live"),"No live trades")');
}

/* ---------- feed ---------- */
function doGet(e){
  var out; if(!e||!e.parameter||e.parameter.token!==getToken_()) out={error:'unauthorised'}; else out=bookJson_();
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function bookJson_(){
  var ss=SpreadsheetApp.getActive(), v=ss.getSheetByName('Trades').getDataRange().getValues(), h=v[0], ix={};
  h.forEach(function(n,i){ ix[n]=i; });
  if(ix['Risk Opened']===undefined) return {error:'sheet not on v2 layout: run Model Portfolio menu > Upgrade sheet to v2 layout'};
  function g(r,n){ var x=r[ix[n]]; return x instanceof Date ? Utilities.formatDate(x,'Europe/London','yyyy-MM-dd') : (x===''||x===undefined?null:x); }
  var trades=[];
  v.slice(1).forEach(function(r,i){ if(r[ix['Trade Idea']]==='') return; var o={row:i+2}; COLS.forEach(function(k){ o[k]=g(r,k); }); trades.push(o); });
  trades.sort(function(a,b){ if(a.Status!==b.Status) return a.Status==='Live'?-1:1;
    var ka=a.Status==='Live'?a['Risk Opened']:a['Exit Date'], kb=b.Status==='Live'?b['Risk Opened']:b['Exit Date']; return String(kb||'').localeCompare(String(ka||'')); });
  var m={}; (ss.getSheetByName('Model Portfolio')||ss.getSheetByName('Book')).getRange('A2:B20').getValues().forEach(function(r){ if(r[0]) m[r[0]]=r[1]; });
  return {v:2, asOf:new Date().toISOString(), book:m, trades:trades, live:trades.filter(function(t){return t.Status==='Live';})};
}
