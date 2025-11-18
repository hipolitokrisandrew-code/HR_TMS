/*******************************************************
 * KPIReportsBackend.gs — KPI readers & aggregators
 *******************************************************/

// Read rows for current SS or a provided SS
function __kpi_readTmsRowsForCurrentCompany__(filters){ const ss=getActiveSpreadsheet_(); return __kpi_readTmsRowsForSS__(ss,filters); }
function __kpi_readTmsRowsForSS__(ss, filters){
  const sh=ss.getSheetByName(CONFIG.SHEETS.TMS);
  const rows=readSheetAsObjects(sh);
  const f=filters||{};
  const startMs=f.startDate?new Date(f.startDate).getTime():null;
  const endMs=f.endDate?(new Date(f.endDate).getTime()+24*60*60*1000-1):null;
  return rows.filter(r=>{
    if(!startMs && !endMs) return true;
    const sMs=r['Start']?new Date(r['Start']).getTime():null;
    const eMs=r['End']?new Date(r['End']).getTime():null;
    if (eMs!=null){ if((startMs&&eMs<startMs)||(endMs&&eMs>endMs)) return false; return true; }
    if (sMs!=null){ if((startMs&&sMs<startMs)||(endMs&&sMs>endMs)) return false; return true; }
    return true;
  });
}

// Buckets and SLA sheet readers (moved as-is)
function __kpi_fromTmsRowsToBuckets__(rows){
  if(!rows||!rows.length) return [];
  const bucket={};
  rows.forEach(r=>{
    const service=String(r['Service']||'').trim()||'Unspecified';
    const step=String(r['Process Step']||'').trim()||'—';
    const key=service+'␟'+step;
    if(!bucket[key]) bucket[key]={service,step,closedWithin:0,closedExceed:0,openExceed:0,openWithin:0,newlyOpened:0};
    const ended=!!(r['End']&&String(r['End']).trim()!=='');
    const tat=parseFloat(String(r['TAT (mins)']||r['Total TAT (mins)']||'0').replace(/,/g,''))||0;
    const exceed=tat>60*24*2;
    if (ended){ if (exceed) bucket[key].closedExceed+=1; else bucket[key].closedWithin+=1; }
    else { if (exceed) bucket[key].openExceed+=1; else bucket[key].openWithin+=1; bucket[key].newlyOpened+=0; }
  });
  return Object.values(bucket);
}
function __kpi_fmtTrendLabel__(ymd){
  if(!ymd) return '—';
  const parts=String(ymd).split('-').map(Number);
  if(parts.length>=3 && parts.every(n=>!isNaN(n))){
    const [y,m,d]=parts; const dt=new Date(Date.UTC(y,(m||1)-1,d||1,12));
    return dt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'});
  }
  return String(ymd);
}
function __kpi_fromTmsByServiceStep__(filters){ return __kpi_fromTmsRowsToBuckets__(__kpi_readTmsRowsForCurrentCompany__(filters)); }
function __kpi_fromTmsByServiceStep_forSS__(ss,filters){ return __kpi_fromTmsRowsToBuckets__(__kpi_readTmsRowsForSS__(ss,filters)); }
function __kpi_vals__(sheet){ if(!sheet) return []; return sheet.getDataRange().getDisplayValues(); }
function __kpi_getSheet__(ss,names){
  for(const n of names){ const sh=ss.getSheetByName(n); if (sh) return sh; }
  const map=new Map(); ss.getSheets().forEach(s=>map.set(s.getName().toLowerCase(),s));
  for(const n of names){ const sh=map.get(String(n).toLowerCase()); if (sh) return sh; }
  return null;
}
function __kpi_readSlaBlockFromSummarySheets__(ss){
  const sh=__kpi_getSheet__(ss,['SLA - Service Quality','SLA Performance']); const vals=__kpi_vals__(sh); if(!vals.length) return null;
  let headerRow=-1; for(let r=0;r<Math.min(vals.length,15);r++){ const row=vals[r].map(v=>String(v||'').trim().toLowerCase()); if(row.includes('closed within sla')){ headerRow=r; break; } }
  if (headerRow===-1) return null;
  const header=vals[headerRow].map(x=>String(x||'').trim());
  const idxOf=name=>header.findIndex(h=>String(h||'').trim().toLowerCase()===String(name).toLowerCase());
  const cWithin=idxOf('Closed within SLA'); const cExceed=idxOf('Closed Exceed SLA'); const oExceed=idxOf('Open Exceed SLA'); const oWithin=idxOf('Open within SLA'); const nOpened=idxOf('Newly Opened');
  if (cWithin<0 || cExceed<0) return null;
  const num=v=>{ const s=String(v||'').replace(/,/g,'').trim(); const n=parseFloat(s); return Number.isFinite(n)?n:0; };
  const out=[]; for(let r=headerRow+1;r<vals.length;r++){ const name=String((vals[r][0]||'')).trim(); if(!name) continue; out.push({ service:name, closedWithin:num(vals[r][cWithin]), closedExceed:num(vals[r][cExceed]), openExceed:oExceed>=0?num(vals[r][oExceed]):0, openWithin:oWithin>=0?num(vals[r][oWithin]):0, newlyOpened:nOpened>=0?num(vals[r][nOpened]):0 }); }
  return out;
}
function __kpi_enrichFromTmsOnly__(fromTms){
  return (fromTms||[]).map(x=>{
    const totalClosed=(x.closedWithin||0)+(x.closedExceed||0);
    const totalAll=totalClosed+(x.openWithin||0)+(x.openExceed||0)+(x.newlyOpened||0);
    const pctClosed=totalAll>0?(totalClosed/totalAll)*100:0;
    return Object.assign({},x,{ total:totalAll, reminderNotice:0, excessCount:(x.closedExceed||0)+(x.openExceed||0), minorTerminal:0, pctClosed:+pctClosed.toFixed(2) });
  });
}
function __kpi_mergeSheetTotalsWithTms__(fromSheet,fromTms){
  const svcMap={}; fromSheet.forEach(s=>{ svcMap[s.service]=Object.assign({},s); });
  const byService={}; fromTms.forEach(r=>{ (byService[r.service]=byService[r.service]||[]).push(r); });
  const out=[];
  Object.keys(byService).forEach(service=>{
    const steps=byService[service];
    const tot=steps.reduce((a,b)=>({closedWithin:a.closedWithin+(b.closedWithin||0),closedExceed:a.closedExceed+(b.closedExceed||0),openExceed:a.openExceed+(b.openExceed||0),openWithin:a.openWithin+(b.openWithin||0),newlyOpened:a.newlyOpened+(b.newlyOpened||0)}),{closedWithin:0,closedExceed:0,openExceed:0,openWithin:0,newlyOpened:0});
    const svcTotals=svcMap[service]||tot;
    const denom=(tot.closedWithin+tot.closedExceed+tot.openWithin+tot.openExceed+tot.newlyOpened)||1;
    steps.forEach(stp=>{
      const weight=((stp.closedWithin||0)+(stp.closedExceed||0)+(stp.openWithin||0)+(stp.openExceed||0)+(stp.newlyOpened||0))/denom;
      const closedWithin=Math.round((svcTotals.closedWithin||0)*weight);
      const closedExceed=Math.round((svcTotals.closedExceed||0)*weight);
      const openExceed=Math.round((svcTotals.openExceed||0)*weight);
      const openWithin=Math.round((svcTotals.openWithin||0)*weight);
      const newlyOpened=Math.round((svcTotals.newlyOpened||0)*weight);
      const total=closedWithin+closedExceed+openExceed+openWithin+newlyOpened;
      const pctClosed=total>0?((closedWithin+closedExceed)/total)*100:0;
      out.push({ service, step:stp.step, closedWithin, closedExceed, openExceed, openWithin, newlyOpened, total, reminderNotice:0, excessCount:closedExceed+openExceed, minorTerminal:0, pctClosed:+pctClosed.toFixed(2) });
    });
  });
  return out;
}
function __kpi_buildServiceStep__(ss,filters){ const fromSheet=__kpi_readSlaBlockFromSummarySheets__(ss); const fromTms=__kpi_fromTmsByServiceStep__(filters); return (!fromSheet||!fromSheet.length)?__kpi_enrichFromTmsOnly__(fromTms):__kpi_mergeSheetTotalsWithTms__(fromSheet,fromTms); }
function __kpi_buildServiceStepForSS__(ss,filters){ const fromSheet=__kpi_readSlaBlockFromSummarySheets__(ss); const fromTms=__kpi_fromTmsByServiceStep_forSS__(ss,filters); return (!fromSheet||!fromSheet.length)?__kpi_enrichFromTmsOnly__(fromTms):__kpi_mergeSheetTotalsWithTms__(fromSheet,fromTms); }

// Public APIs (renamed)
function getKPIReportDataV2Backend(filters){
  requireAuth();
  const ss=getActiveSpreadsheet_();
  const rows=__kpi_buildServiceStep__(ss, filters||{});
  const svcAgg={};
  rows.forEach(r=>{
    const s=r.service;
    if(!svcAgg[s]) svcAgg[s]={service:s,closedWithin:0,closedExceed:0,openExceed:0,openWithin:0,newlyOpened:0,total:0,reminderNotice:0,excessCount:0,minorTerminal:0};
    const a=svcAgg[s];
    a.closedWithin+=r.closedWithin||0; a.closedExceed+=r.closedExceed||0; a.openExceed+=r.openExceed||0; a.openWithin+=r.openWithin||0; a.newlyOpened+=r.newlyOpened||0; a.total+=r.total||0; a.reminderNotice+=r.reminderNotice||0; a.excessCount+=r.excessCount||0; a.minorTerminal+=r.minorTerminal||0;
  });
  const services=Object.values(svcAgg).map(a=>{ const closed=(a.closedWithin||0)+(a.closedExceed||0); const pctClosed=a.total>0?(closed/a.total)*100:0; return Object.assign(a,{pctClosed:+pctClosed.toFixed(2)}); }).sort((x,y)=>y.total-x.total);
  const overall=services.reduce((o,a)=>{ o.closedWithin+=a.closedWithin||0; o.closedExceed+=a.closedExceed||0; o.openExceed+=a.openExceed||0; o.openWithin+=a.openWithin||0; o.newlyOpened+=a.newlyOpened||0; o.total+=a.total||0; o.reminderNotice+=a.reminderNotice||0; o.excessCount+=a.excessCount||0; o.minorTerminal+=a.minorTerminal||0; return o; },{closedWithin:0,closedExceed:0,openExceed:0,openWithin:0,newlyOpened:0,total:0,reminderNotice:0,excessCount:0,minorTerminal:0});
  const closedOverall=(overall.closedWithin||0)+(overall.closedExceed||0);
  overall.pctClosed=overall.total>0?+((closedOverall/overall.total)*100).toFixed(2):0;

  const targetPct=95, trend=[], tatBins=[];
  try{
    const f=filters||{}; const rowsAll=__kpi_readTmsRowsForCurrentCompany__(f);
    if(rowsAll&&rowsAll.length){
      const startMs=f&&f.startDate?new Date(f.startDate).getTime():null;
      const endMs=f&&f.endDate?(new Date(f.endDate).getTime()+24*60*60*1000-1):null;
      const nowMs=Date.now(); const SLA_MIN=60*24*2;
      const dayAgg={}; function addDay(key,fld){ if(!dayAgg[key]) dayAgg[key]={dateLabel:key,closedWithin:0,closedExceed:0,openWithin:0,openExceed:0,total:0}; dayAgg[key][fld]+=1; dayAgg[key].total+=1; }
      function ymd(d){ const dt=(d instanceof Date)?d:new Date(d); if(isNaN(dt))return null; const y=dt.getFullYear(); const m=(dt.getMonth()+1).toString().padStart(2,'0'); const da=dt.getDate().toString().padStart(2,'0'); return y+'-'+m+'-'+da; }
      const bins=[{label:'≤ 1d',min:0,max:60*24},{label:'1–2d',min:60*24,max:60*48},{label:'2–3d',min:60*48,max:60*72},{label:'3–5d',min:60*72,max:60*120},{label:'5–7d',min:60*120,max:60*168},{label:'> 7d',min:60*168,max:Infinity}], binCounts=bins.map(b=>({label:b.label,count:0}));
      rowsAll.forEach(r=>{
        const start=r['Start'], end=r['End']; const sMs=start?new Date(start).getTime():null; const eMs=end?new Date(end).getTime():null;
        let tat=0; const tatCell=r['Total TAT (mins)'], curTatCell=r['TAT (mins)'];
        if(end) tat=parseFloat(String(tatCell||'0').replace(/,/g,''))||0;
        else if(curTatCell!=null&&curTatCell!=='') tat=parseFloat(String(curTatCell).replace(/,/g,''))||0;
        else if(sMs) tat=Math.max(0, Math.round((nowMs-sMs)/60000));
        if(!isNaN(tat)){ for(let i=0;i<bins.length;i++){ if(tat>=bins[i].min && tat<bins[i].max){ binCounts[i].count+=1; break; } } }
        const within=tat<=SLA_MIN;
        if(eMs){ if((startMs==null||eMs>=startMs)&&(endMs==null||eMs<=endMs)){ addDay(ymd(eMs), within?'closedWithin':'closedExceed'); } }
        else if(sMs){ if((startMs==null||sMs>=startMs)&&(endMs==null||sMs<=endMs)){ addDay(ymd(sMs), within?'openWithin':'openExceed'); } }
      });
      Object.keys(dayAgg).sort().forEach(k=>{
        const d=dayAgg[k]; const closed=(d.closedWithin||0)+(d.closedExceed||0); const pctClosed=d.total>0?(closed/d.total)*100:0;
        trend.push({date:d.dateLabel, dateISO:d.dateLabel, dateLabel:d.dateLabel, displayLabel:__kpi_fmtTrendLabel__(d.dateLabel), closedWithin:d.closedWithin, closedExceed:d.closedExceed, openWithin:d.openWithin, openExceed:d.openExceed, total:d.total, pctClosed:+pctClosed.toFixed(2), targetPct});
      });
      bins.forEach((b,i)=>tatBins.push({label:b.label,count:binCounts[i].count}));
    }
  }catch(err){ Logger.log('KPI trend/TAT calc error: '+err); }

  return { targetPct, overall, services, rows, trend, tatBins };
}

function getKPIReportDataV3Backend(filters){
  requireAuth();
  const f=Object.assign({},filters||{}); const co=String(f.company||'').toUpperCase(); const wantAll=!co||co==='ALL';
  const bucketsBySS=(ss)=>__kpi_buildServiceStepForSS__(ss,f);
  const ssMap={ ONWARD:getOnwardSs(), ITAM:getItamSs(), IREAL:getIrealSs(), LATTE:getLatteSs() };
  const targetPct=95; let rows=[];
  if (wantAll){ rows=[].concat(bucketsBySS(ssMap.ONWARD), bucketsBySS(ssMap.ITAM), bucketsBySS(ssMap.IREAL), bucketsBySS(ssMap.LATTE)); }
  else { const ss=ssMap[co]||getActiveSpreadsheet_(); rows=bucketsBySS(ss); }

  const svcAgg={};
  rows.forEach(r=>{
    const s=r.service;
    if(!svcAgg[s]) svcAgg[s]={service:s,closedWithin:0,closedExceed:0,openExceed:0,openWithin:0,newlyOpened:0,total:0,reminderNotice:0,excessCount:0,minorTerminal:0};
    const a=svcAgg[s];
    a.closedWithin+=r.closedWithin||0; a.closedExceed+=r.closedExceed||0; a.openExceed+=r.openExceed||0; a.openWithin+=r.openWithin||0; a.newlyOpened+=r.newlyOpened||0; a.total+=r.total||0; a.reminderNotice+=r.reminderNotice||0; a.excessCount+=r.excessCount||0; a.minorTerminal+=r.minorTerminal||0;
  });
  const services=Object.values(svcAgg).map(a=>{ const closed=(a.closedWithin||0)+(a.closedExceed||0); const pctClosed=a.total>0?(closed/a.total)*100:0; return Object.assign(a,{pctClosed:+pctClosed.toFixed(2)}); }).sort((x,y)=>y.total-x.total);
  const overall=services.reduce((o,a)=>{ o.closedWithin+=a.closedWithin||0; o.closedExceed+=a.closedExceed||0; o.openExceed+=a.openExceed||0; o.openWithin+=a.openWithin||0; o.newlyOpened+=a.newlyOpened||0; o.total+=a.total||0; o.reminderNotice+=a.reminderNotice||0; o.excessCount+=a.excessCount||0; o.minorTerminal+=a.minorTerminal||0; return o; },{closedWithin:0,closedExceed:0,openExceed:0,openWithin:0,newlyOpened:0,total:0,reminderNotice:0,excessCount:0,minorTerminal:0});
  const closedOverall=(overall.closedWithin||0)+(overall.closedExceed||0);
  overall.pctClosed=overall.total>0?+((closedOverall/overall.total)*100).toFixed(2):0;

  // Trend & TAT (all or specific company)
  const trend=[], tatBins=[];
  try{
    const list= wantAll ? [ssMap.ONWARD, ssMap.ITAM, ssMap.IREAL, ssMap.LATTE] : [(ssMap[co]||getActiveSpreadsheet_())];
    const allRows=[]; list.forEach(ss=>{ allRows.push.apply(allRows, __kpi_readTmsRowsForSS__(ss,f)); });
    if (allRows.length){
      const startMs=f&&f.startDate?new Date(f.startDate).getTime():null;
      const endMs=f&&f.endDate?(new Date(f.endDate).getTime()+24*60*60*1000-1):null;
      const nowMs=Date.now(); const SLA_MIN=60*24*2;
      const dayAgg={}; function addDay(key,fld){ if(!dayAgg[key]) dayAgg[key]={dateLabel:key,closedWithin:0,closedExceed:0,openWithin:0,openExceed:0,total:0}; dayAgg[key][fld]+=1; dayAgg[key].total+=1; }
      function ymd(d){ const dt=(d instanceof Date)?d:new Date(d); if(isNaN(dt))return null; const y=dt.getFullYear(); const m=(dt.getMonth()+1).toString().padStart(2,'0'); const da=dt.getDate().toString().padStart(2,'0'); return y+'-'+m+'-'+da; }
      const bins=[{label:'≤ 1d',min:0,max:60*24},{label:'1–2d',min:60*24,max:60*48},{label:'2–3d',min:60*48,max:60*72},{label:'3–5d',min:60*72,max:60*120},{label:'5–7d',min:60*120,max:60*168},{label:'> 7d',min:60*168,max:Infinity}], binCounts=bins.map(b=>({label:b.label,count:0}));
      allRows.forEach(r=>{
        const start=r['Start'], end=r['End']; const sMs=start?new Date(start).getTime():null; const eMs=end?new Date(end).getTime():null;
        let tat=0; const tatCell=r['Total TAT (mins)'], curTatCell=r['TAT (mins)'];
        if(end) tat=parseFloat(String(tatCell||'0').replace(/,/g,''))||0;
        else if(curTatCell!=null&&curTatCell!=='') tat=parseFloat(String(curTatCell).replace(/,/g,''))||0;
        else if(sMs) tat=Math.max(0, Math.round((nowMs-sMs)/60000));
        if(!isNaN(tat)){ for(let i=0;i<bins.length;i++){ if(tat>=bins[i].min && tat<bins[i].max){ binCounts[i].count+=1; break; } } }
        const within=tat<=SLA_MIN;
        if(eMs){ if((startMs==null||eMs>=startMs)&&(endMs==null||eMs<=endMs)){ addDay(ymd(eMs), within?'closedWithin':'closedExceed'); } }
        else if(sMs){ if((startMs==null||sMs>=startMs)&&(endMs==null||sMs<=endMs)){ addDay(ymd(sMs), within?'openWithin':'openExceed'); } }
      });
      Object.keys(dayAgg).sort().forEach(k=>{ const d=dayAgg[k]; const closed=(d.closedWithin||0)+(d.closedExceed||0); const pctClosed=d.total>0?(closed/d.total)*100:0;
        trend.push({date:d.dateLabel, dateISO:d.dateLabel, dateLabel:d.dateLabel, displayLabel:__kpi_fmtTrendLabel__(d.dateLabel), closedWithin:d.closedWithin, closedExceed:d.closedExceed, openWithin:d.openWithin, openExceed:d.openExceed, total:d.total, pctClosed:+pctClosed.toFixed(2), targetPct});
      });
      bins.forEach((b,i)=>tatBins.push({label:b.label,count:binCounts[i].count}));
    }
  }catch(err){ Logger.log('KPI trend/TAT calc error (V3): '+err); }

  return { targetPct, overall, services, rows, trend, tatBins };
}
