'use strict';
/* 2-3팀 전용 · 순수 브라우저 XLSX 파서 (외부 서버/라이브러리 없음) */
const chs=['쇼핑검색광고','파워링크','GFA'];
const EXCLUDED=new Set(['최연아','안희성']);
const TEAM_TOKEN='2-3팀(김재현)';
const DB_NAME='team23_sales_v1', RECORDS_STORE='snapshots', NOTES_STORE='daily_notes', WORK_KEY='team23_work_v2', MEMO_KEY='team23_memos_v2', LATEST_KEY='team23_latest_v1', MIGRATE_KEY='team23_daily_notes_migrated_v1';
let DATA={rows:[],meta:{}},records=[],staff=[],active=[],ownerPerf=[],low=[],team={},ownBad=0,currentPage=1,currentlyShown=[];
let currentSnapshotId=null,previousSnapshot=null,availableDates=[],dateRequest=0;const pageSize=15;
const $=id=>document.getElementById(id);
const fmt=n=>Math.round(n||0).toLocaleString('ko-KR');
const won=n=>{const a=Math.abs(n||0);if(a>=100000000){const e=Math.floor(a/100000000),m=Math.round(a%100000000/10000);return `${e}억 ${m?fmt(m)+'만':''}원`;}if(a>=10000)return fmt(Math.round(a/10000))+'만원';return fmt(Math.round(a))+'원';};
const signed=n=>(n>=0?'+':'−')+won(Math.abs(n));
const pct=n=>n===null||!isFinite(n)?'—':n.toFixed(1)+'%';
const sum=(a,k)=>a.reduce((v,r)=>v+(+r[k]||0),0);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const getCh=c=>records.filter(r=>r.channel===c);
const getStaff=o=>records.filter(r=>r.owner===o);
const groupSum=arr=>({oct:sum(arr,'oct'),target:sum(arr,'target'),actual:sum(arr,'actual'),forecast:sum(arr,'forecast'),gap:sum(arr,'gap')});
const byGap=(a,b)=>a.gap-b.gap;
const byTop=(arr,n)=>arr.filter(r=>r.target>0&&r.gap<0).sort(byGap).slice(0,n);
const stableKey=r=>`${r.channel}/${r.id}`;
function getStored(key){try{return JSON.parse(localStorage.getItem(key)||'{}')||{};}catch{return {};}}
let work={},memo={},pendingNoteSave=Promise.resolve();
function persistNotes(){
 const date=DATA.meta.asof;
 if(!date)return;
 const item={date,work:{...work},memo:{...memo}};
 pendingNoteSave=pendingNoteSave.catch(()=>{}).then(()=>putDailyNote(item));
 return pendingNoteSave.catch(e=>statusMsg(`사유·업무상태 저장 실패 (${date}): ${e.message}. 전체 백업을 내려받으세요.`, 'error'));
}
function localToday(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
function latestByDay(all){const m=new Map();for(const it of all){const date=it.data?.meta?.asof;if(date&&(!m.has(date)||it.created>m.get(date).created))m.set(date,it);}return m;}

function dateLabel(d){if(!d)return '날짜 미확인';const t=new Date(d+'T12:00:00');return isNaN(t)?d:`${t.getFullYear()}.${String(t.getMonth()+1).padStart(2,'0')}.${String(t.getDate()).padStart(2,'0')}`;}
function shortMonth(){if(!DATA.meta.asof)return '당월';return Number(DATA.meta.asof.slice(5,7))+'월';}
function prevMonth(){if(!DATA.meta.asof)return '전월';const d=new Date(DATA.meta.asof+'T12:00:00');d.setMonth(d.getMonth()-1);return (d.getMonth()+1)+'월';}
function statusMsg(s,type=''){const box=$('uploadStatus');box.textContent=s;box.className='upload-status '+type;}
function resetComputed(){records=DATA.rows.filter(r=>!EXCLUDED.has(r.owner)).map(r=>({...r,gap:r.forecast-r.target,rate:r.target>0?r.forecast/r.target*100:null}));staff=[...new Set(records.map(r=>r.owner))].filter(owner=>sum(getStaff(owner),'target')>100000).sort((a,b)=>sum(getStaff(a),'forecast')-sum(getStaff(a),'target')-(sum(getStaff(b),'forecast')-sum(getStaff(b),'target')));team=groupSum(records);active=records.filter(r=>r.target>0);low=active.filter(r=>r.gap<0);ownerPerf=staff.map(owner=>({owner,...groupSum(getStaff(owner))}));ownBad=ownerPerf.filter(p=>p.gap<0).length;}

// ZIP central directory + deflate-raw; only local browser native APIs are used.
const decoder=new TextDecoder('utf-8');
const read16=(v,o)=>v.getUint16(o,true),read32=(v,o)=>v.getUint32(o,true);
function zipIndex(buffer){
 const v=new DataView(buffer),u=new Uint8Array(buffer);let end=-1;
 for(let i=u.length-22;i>=Math.max(0,u.length-65557);i--){if(read32(v,i)===0x06054b50){end=i;break;}}
 if(end<0)throw Error('올바른 XLSX ZIP 파일이 아닙니다.');
 const count=read16(v,end+10),directory=read32(v,end+16),entries=new Map();let pos=directory;
 for(let i=0;i<count;i++){
  if(read32(v,pos)!==0x02014b50)throw Error('엑셀 압축 파일의 색인을 읽을 수 없습니다.');
  const method=read16(v,pos+10),packed=read32(v,pos+20),original=read32(v,pos+24),nl=read16(v,pos+28),ex=read16(v,pos+30),co=read16(v,pos+32),local=read32(v,pos+42);
  const name=decoder.decode(u.slice(pos+46,pos+46+nl));entries.set(name,{method,packed,original,local});pos+=46+nl+ex+co;
 }
 return {buffer,entries};
}
async function zipText(zip,name){
 const rec=zip.entries.get(name);if(!rec)throw Error(`엑셀 파일 내부에서 ${name} 시트를 찾지 못했습니다.`);
 const v=new DataView(zip.buffer),u=new Uint8Array(zip.buffer),o=rec.local;
 if(read32(v,o)!==0x04034b50)throw Error('엑셀 압축 헤더가 올바르지 않습니다.');
 const offset=o+30+read16(v,o+26)+read16(v,o+28);
 if(offset+rec.packed>u.byteLength)throw Error('엑셀 압축 데이터가 손상됐습니다.');
 const bytes=u.slice(offset,offset+rec.packed);
 if(rec.method===0)return decoder.decode(bytes);
 if(rec.method!==8)throw Error('지원되지 않는 XLSX 압축 방식입니다.');
 if(!('DecompressionStream' in window))throw Error('브라우저에 ZIP 압축 해제 기능이 없습니다. 최신 Chrome 또는 Edge를 사용하세요.');
 const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
 const array=await new Response(stream).arrayBuffer();return decoder.decode(array);
}
const xmlDoc=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('엑셀 XML을 읽을 수 없습니다.');return d;};
const byTag=(e,n)=>e.getElementsByTagNameNS('*',n);
function sheetIndex(workbookXml,relsXml){
 const wb=xmlDoc(workbookXml),rels=xmlDoc(relsXml),m={};
 for(const p of byTag(rels,'Relationship'))m[p.getAttribute('Id')]=p.getAttribute('Target');
 const out={};for(const s of byTag(wb,'sheet')){
  const rel=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
  let path=m[rel];if(!path)continue;path=path.replace(/^\//,'');if(!path.startsWith('xl/'))path='xl/'+path.replace(/^\.\//,'');out[s.getAttribute('name')]=path;
 }return out;
}
function parseStrings(xml){
 const d=xmlDoc(xml);return Array.from(byTag(d,'si'),si=>Array.from(byTag(si,'t'),t=>t.textContent).join(''));
}
function readSheet(xml,shared,predicate){
 const root=xmlDoc(xml),result=[];
 for(const row of byTag(root,'row')){
  const cells={};for(const cell of row.children){if(cell.localName!=='c')continue;
   const ref=cell.getAttribute('r')||'',col=ref.replace(/\d/g,'');if(!col)continue;
   const type=cell.getAttribute('t'),v=Array.from(cell.children).find(k=>k.localName==='v');
   let value=v?v.textContent:null;
   if(type==='inlineStr'){const is=Array.from(cell.children).find(k=>k.localName==='is');value=is?Array.from(byTag(is,'t'),t=>t.textContent).join(''):'';}
   else if(type==='s'){value=shared[Number(value)]??'';}
   else if(value!==null&&type!=='str'&&type!=='e'){const num=Number(value);if(isFinite(num)&&value!=='')value=num;}
   if(value!==null)cells[col]=value;
  }
  cells._row=Number(row.getAttribute('r'))||0;
  if(!predicate||predicate(cells))result.push(cells);
 }return result;
}
function excelDate(x){if(typeof x!=='number'||!isFinite(x)||x<40000||x>60000)return null;return new Date(Date.UTC(1899,11,30)+Math.round(x)*86400000).toISOString().slice(0,10);}
function parseNameDate(name){const m=name.match(/(?:_|-)(\d{2})(\d{2})(?:\.|_)/);if(!m)return null;return `${new Date().getFullYear()}-${m[1]}-${m[2]}`;}
const n=x=>Number(x)||0;
function buildFromSheets(sheets,fileName){
 const expected=['쇼핑검색광고','파워링크','GFA'];for(const s of expected)if(!sheets[s])throw Error(`필수 시트 '${s}'가 없습니다. 프로모션 원본 엑셀을 올려주세요.`);
 const metaRow=sheets.GFA.find(r=>r._row===2),asof=excelDate(n(metaRow?.E))||parseNameDate(fileName)||new Date().toISOString().slice(0,10),month=asof.slice(0,7);
 const channelRaw={'쇼핑검색광고':['네이버쇼핑검색'],'파워링크':['네이버웹검색'],'GFA':['네이버GFA','네이버ADVoost']};
 const raws=(sheets.raw||[]).filter(r=>String(r.C||'').includes(TEAM_TOKEN)&&!EXCLUDED.has(String(r.B||'')));
 const rawHead=(sheets.raw||[]).find(r=>r._row<=3&&excelDate(n(r.N)))||{},dates=[];
 for(const [col,value] of Object.entries(rawHead)){const day=excelDate(value);if(day&&day.startsWith(month)&&day<=asof)dates.push({col,day});}
 dates.sort((a,b)=>a.day.localeCompare(b.day));
 const dateMap=new Map();for(const q of raws){const key=`${q.B}|${String(q.F)}|${q.D}`;if(!dateMap.has(key))dateMap.set(key,[]);dateMap.get(key).push(q);}
 const rows=[];
 for(const channel of expected)for(const r of sheets[channel]){
  const owner=String(r.C??'').trim(),teamName=String(r.B??'');
  if(!teamName.includes(TEAM_TOKEN)||EXCLUDED.has(owner)||!owner||r.E===undefined||String(r.E).trim()==='')continue;
  const id=String(r.E).trim(),name=String(r.D??'').trim(),actual=n(r.H),dailySum={};let fullDays=[];
  let found=0;
  for(const kind of channelRaw[channel])for(const q of dateMap.get(`${owner}|${id}|${kind}`)||[]){found++;
   for(const {col,day} of dates)dailySum[day]=(dailySum[day]||0)+n(q[col]);
  }
  const allSum=Object.values(dailySum).reduce((a,b)=>a+b,0);
  const valid=found>0&&Math.abs(allSum-actual)<=1;
  if(valid)fullDays=dates.map(d=>({day:d.day,value:dailySum[d.day]||0}));
  const four=fullDays.filter(x=>{const wd=new Date(x.day+'T12:00:00').getDay();return wd!==0&&wd!==6;}).slice(-4),before=four.length===4?(four[0].value+four[1].value)/2:0,after=four.length===4?(four[2].value+four[3].value)/2:0;
  const trendPct=valid&&four.length===4&&before>0?(after/before-1)*100:null;
  const trend=trendPct===null?'unknown':trendPct<=-15?'down':trendPct>=15?'up':'flat';
  const advertiser=/^\d{8,}$/.test(name)?`광고주명 확인 필요 (원본값 ${name})`:name;
  rows.push({channel,owner,name:advertiser,id,oct:n(r.F),target:n(r.G),actual,forecast:n(r.I),rank:n(r.L),daily:valid?fullDays:null,trend,trendPct:trendPct===null?null:Math.round(trendPct*10)/10});
 }
 if(rows.length===0)throw Error('2-3팀(김재현) 광고계정이 발견되지 않았습니다. 원본의 소속 열을 확인하세요.');
 const unique=new Set(rows.map(r=>`${r.channel}/${r.id}`));if(unique.size!==rows.length)throw Error('매체·계정ID가 중복돼 집계가 불안정합니다. 원본을 확인해 주세요.');
 const summaryRows=[],summaryPeople=[];
 for(const r of (sheets['인원별']||[]))if(String(r.A||'').includes(TEAM_TOKEN)&&!EXCLUDED.has(String(r.B??''))){summaryPeople.push(String(r.B));summaryRows.push(r);}
 const detailPeople=[...new Set(rows.map(r=>r.owner))];
 const chart=chs.map(channel=>{const arr=rows.filter(r=>r.channel===channel);const cells=({쇼핑검색광고:['D','E'],파워링크:['I','J'],GFA:['N','O']})[channel];return {channel,detailTarget:sum(arr,'target'),detailForecast:sum(arr,'forecast'),summaryTarget:sum(summaryRows,cells[0]),summaryForecast:sum(summaryRows,cells[1])};});
 const total=sum(rows,'forecast');if(!isFinite(total))throw Error('예상매출 합계가 올바르지 않습니다.');
 return {rows,meta:{asof,month,filename:fileName,team:'서울2-3팀(김재현)',channels:chs,summaryRows:chart,amountCompared:rows.length,rawMismatch:rows.filter(r=>r.actual>0&&r.daily===null).length,memberConflict:{summaryOnly:summaryPeople.filter(x=>!detailPeople.includes(x)),detailOnly:detailPeople.filter(x=>!summaryPeople.includes(x))},excluded:[...EXCLUDED]}};
}
async function parseWorkbook(file){
 if(!/\.xls(m|x)$/i.test(file.name))throw Error('XLSX 또는 XLSM 파일만 지원합니다.');
 if(file.size>60*1024*1024)throw Error('파일이 너무 큽니다(60MB 초과).');
 const zip=zipIndex(await file.arrayBuffer());
 const workbookXml=await zipText(zip,'xl/workbook.xml'),relsXml=await zipText(zip,'xl/_rels/workbook.xml.rels'),paths=sheetIndex(workbookXml,relsXml);
 const names=['쇼핑검색광고','파워링크','GFA','인원별','raw'];
 for(const n of names.slice(0,3))if(!paths[n])throw Error(`엑셀에 '${n}' 시트가 없습니다.`);
 const shared=zip.entries.has('xl/sharedStrings.xml')?parseStrings(await zipText(zip,'xl/sharedStrings.xml')):[];
 const sheets={};for(const name of names){if(!paths[name])continue;
  const predicate=name==='raw'?(r=>r._row<=3||String(r.C||'').includes(TEAM_TOKEN)):
   name==='인원별'?(r=>r._row<3||String(r.A||'').includes(TEAM_TOKEN)):
   (r=>r._row<4||String(r.B||'').includes(TEAM_TOKEN));
  sheets[name]=readSheet(await zipText(zip,paths[name]),shared,predicate);
 }
 return buildFromSheets(sheets,file.name);
}

let storageMode='indexeddb',tempItems=[];
const FALLBACK_KEY='team23_snapshots_fallback_v1';
function fallbackList(){if(storageMode==='temporary'&&tempItems.length)return tempItems;try{const items=JSON.parse(localStorage.getItem(FALLBACK_KEY)||'null');if(Array.isArray(items))return items;}catch(e){}return tempItems;}
function fallbackWrite(items){tempItems=items;try{localStorage.setItem(FALLBACK_KEY,JSON.stringify(items));storageMode='localStorage';}catch(e){storageMode='temporary';}}
function dbOpen(){return new Promise((resolve,reject)=>{let req;try{req=indexedDB.open(DB_NAME,2);}catch(e){reject(e);return;}req.onupgradeneeded=()=>{
 if(!req.result.objectStoreNames.contains(RECORDS_STORE))req.result.createObjectStore(RECORDS_STORE,{keyPath:'id'});
 if(!req.result.objectStoreNames.contains(NOTES_STORE))req.result.createObjectStore(NOTES_STORE,{keyPath:'date'});
 };req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);});}
const NOTE_FALLBACK_KEY='team23_daily_notes_fallback_v1';
let noteTemp={};
function fallbackNotes(){if(storageMode==='temporary'&&Object.keys(noteTemp).length)return noteTemp;try{return JSON.parse(localStorage.getItem(NOTE_FALLBACK_KEY)||'{}')||{};}catch{return noteTemp;}}
function fallbackPutNote(item){const all=fallbackNotes();all[item.date]=item;noteTemp=all;try{localStorage.setItem(NOTE_FALLBACK_KEY,JSON.stringify(all));storageMode='localStorage';}catch(e){storageMode='temporary';throw Error('브라우저 저장 용량 부족 또는 저장 제한');}}
async function putDailyNote(item){try{const db=await dbOpen();await new Promise((resolve,reject)=>{const tx=db.transaction(NOTES_STORE,'readwrite');tx.objectStore(NOTES_STORE).put(item);tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>reject(tx.error);});}catch(e){fallbackPutNote(item);}}
async function allDailyNotes(){try{const db=await dbOpen();return await new Promise((resolve,reject)=>{const tx=db.transaction(NOTES_STORE,'readonly');const req=tx.objectStore(NOTES_STORE).getAll();req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);tx.oncomplete=()=>db.close();});}catch(e){return Object.values(fallbackNotes());}}
async function getDailyNote(date){try{const db=await dbOpen();return await new Promise((resolve,reject)=>{const tx=db.transaction(NOTES_STORE,'readonly');const req=tx.objectStore(NOTES_STORE).get(date);req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>reject(req.error);tx.oncomplete=()=>db.close();});}catch(e){return fallbackNotes()[date]||null;}}
async function migrateLegacyNotes(all){
 let already=false;try{already=localStorage.getItem(MIGRATE_KEY)==='1';}catch(e){}if(already||!all.length)return;
 const oldWork=getStored(WORK_KEY),oldMemo=getStored(MEMO_KEY);
 if(Object.keys(oldWork).length||Object.keys(oldMemo).length){
  let oldId='';try{oldId=localStorage.getItem(LATEST_KEY)||'';}catch(e){}
  const target=all.find(x=>x.id===oldId)||all[0];const date=target.data.meta.asof;
  const existing=await getDailyNote(date)||{date,work:{},memo:{}};
  await putDailyNote({date,work:{...oldWork,...existing.work},memo:{...oldMemo,...existing.memo}});
 }
 try{localStorage.setItem(MIGRATE_KEY,'1');}catch(e){}
}

async function storeAll(){try{const db=await dbOpen();const data=await new Promise((resolve,reject)=>{const tx=db.transaction(RECORDS_STORE,'readonly');const req=tx.objectStore(RECORDS_STORE).getAll();req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);tx.oncomplete=()=>db.close();});storageMode='indexeddb';return data.sort((a,b)=>b.created-a.created);}catch(e){return fallbackList().sort((a,b)=>b.created-a.created);}}
async function putItem(item){try{const db=await dbOpen();await new Promise((resolve,reject)=>{const tx=db.transaction(RECORDS_STORE,'readwrite');tx.objectStore(RECORDS_STORE).put(item);tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>reject(tx.error);});storageMode='indexeddb';}catch(e){const curr=fallbackList().filter(x=>x.id!==item.id);fallbackWrite([item,...curr]);}}
async function deleteItem(id){try{const db=await dbOpen();await new Promise((resolve,reject)=>{const tx=db.transaction(RECORDS_STORE,'readwrite');tx.objectStore(RECORDS_STORE).delete(id);tx.oncomplete=()=>{db.close();resolve();};tx.onerror=()=>reject(tx.error);});}catch(e){fallbackWrite(fallbackList().filter(x=>x.id!==id));}}
// 날짜별 원본을 계속 보관합니다. 기존 5개 기록 삭제 제한은 사용하지 않습니다.
async function loadItem(id,notify=false,request=null){
 await pendingNoteSave.catch(()=>{});
 const all=await storeAll(),it=all.find(x=>x.id===id);if(!it)throw Error('저장된 데이터를 찾지 못했습니다.');
 const selectedDate=it.data.meta.asof;
 const daily=await getDailyNote(selectedDate);
 if(request!==null&&request!==dateRequest)return false;
 currentSnapshotId=it.id;DATA=it.data;work={...(daily?.work||{})};memo={...(daily?.memo||{})};resetComputed();
 try{localStorage.setItem(LATEST_KEY,id);}catch(e){}
 const previous=[...latestByDay(all).values()].filter(x=>x.data.meta.asof<selectedDate).sort((a,b)=>b.data.meta.asof.localeCompare(a.data.meta.asof)||b.created-a.created);
 previousSnapshot=previous[0]||null;
 renderAll();await renderHistory();
 if(notify&&(request===null||request===dateRequest))statusMsg(`${dateLabel(selectedDate)} 기준 자료 조회 · ${DATA.meta.filename} / ${records.length}개 계정. 이 날짜에 작성한 사유·업무상태를 표시합니다.`, 'ok');
}
async function selectDate(date,notify=true){
 const request=++dateRequest;
 const all=await storeAll();if(request!==dateRequest)return false;
 const byDate=latestByDay(all),item=byDate.get(date);
 if(!item){statusMsg(`${dateLabel(date)}에 업로드된 원본이 없습니다. 해당 기준일의 엑셀을 먼저 업로드해 주세요. 다른 날짜의 실적을 대신 표시하지 않습니다.`, 'error');$('datePicker').value=DATA.meta.asof||'';return false;}
 await loadItem(item.id,notify,request);return request===dateRequest;
}
async function renderHistory(){
 let all=[];try{all=await storeAll();}catch(e){$('snapshotList').textContent='업로드 기록 저장에 접근할 수 없습니다.';return;}
 const byDate=latestByDay(all);availableDates=[...byDate.keys()].sort();
 const selected=DATA.meta.asof||'';
 $('datePicker').value=selected;$('prevDateBtn').disabled=!selected||availableDates.indexOf(selected)<=0;
 $('nextDateBtn').disabled=!selected||availableDates.indexOf(selected)>=availableDates.length-1;
 const today=localToday();
 $('todayBtn').classList.toggle('today-active',selected===today);
 $('dateMeta').textContent=selected?(selected===today?'오늘 기준 자료 조회 중':`선택: ${dateLabel(selected)}`)+` · 저장된 날짜 ${availableDates.length}일 · 선택 날짜의 사유/업무상태만 표시`: `오늘 ${dateLabel(today)} · 업로드된 날짜 ${availableDates.length}일`;
 $('snapshotList').innerHTML=availableDates.length?[...availableDates].reverse().map(date=>{
  const item=byDate.get(date);return `<button type="button" data-date="${esc(date)}" class="snapshot-btn ${date===selected?'current':''}" aria-pressed="${date===selected}" title="${esc(item.data?.meta?.filename||'')}">${dateLabel(date)}${date===today?' · 오늘':''}${date===selected?' ●':''}</button>`;
 }).join(''):'<span class="subtle">아직 업로드된 날짜가 없습니다. 엑셀을 올리면 날짜별로 누적됩니다.</span>';
 $('snapshotList').querySelectorAll('button').forEach(btn=>btn.onclick=async()=>{try{await selectDate(btn.dataset.date);}catch(e){statusMsg(e.message,'error');}});
}
async function handleUpload(file){if(!file)return;statusMsg(`${file.name} 파일을 분석하고 있습니다...`);$('xlsxInput').disabled=true;
 try{const data=await parseWorkbook(file);const now=Date.now(),id=String(now)+'_'+Math.floor(Math.random()*100000);
  const item={id,created:now,data};await putItem(item);dateRequest++;await loadItem(id);
  statusMsg(`업로드 완료 · ${dateLabel(data.meta.asof)} 기준 / ${data.rows.length}개 김재현 팀 계정 · 제외: 최연아, 안희성\n※ 과거 날짜 원본은 유지됩니다. 같은 날짜를 다시 올리면 해당 날짜에는 최신 업로드가 표시됩니다.\n※ ${storageMode==='temporary'?'이 브라우저에서는 저장이 차단되어 임시로만 표시됩니다. 반드시 전체 백업 JSON을 내려받으세요.':'이 기기의 브라우저에만 저장됩니다.'} GitHub 저장소와 다른 PC에는 자동 반영되지 않습니다.`,storageMode==='temporary'?'error':'ok');
 }catch(e){console.error(e);statusMsg('업로드 실패: '+(e?.message||String(e))+'\n기존 대시보드 데이터는 그대로 유지됩니다.','error');}finally{$('xlsxInput').disabled=false;$('xlsxInput').value='';}}
function downloadText(text,name,mime){const blob=new Blob([text],{type:mime||'application/json;charset=utf-8'}),u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(u),3000);}
async function backupAll(){try{
 await pendingNoteSave.catch(()=>{});
 const snapshot=await storeAll();const notes=await allDailyNotes();const map=new Map(notes.map(x=>[x.date,x]));
 if(DATA.meta.asof)map.set(DATA.meta.asof,{date:DATA.meta.asof,work:{...work},memo:{...memo}});
 const data={format:'team23-v2',date:new Date().toISOString(),snapshots:snapshot,dailyNotes:[...map.values()]};
 downloadText(JSON.stringify(data),'2-3팀_날짜별_매출관리_전체백업_'+localToday()+'.json');
 }catch(e){statusMsg('백업 실패: '+e.message,'error');}}
async function restoreBackup(file){if(!file)return;try{
 const payload=JSON.parse(await file.text());if(!['team23-v1','team23-v2'].includes(payload.format)||!Array.isArray(payload.snapshots)||!payload.snapshots.every(s=>s&&Array.isArray(s.data?.rows)&&s.data?.meta?.asof))throw Error('올바른 2-3팀 백업 파일이 아닙니다.');
 const existing=await storeAll();const dedup=new Map(existing.map(item=>[item.id,item]));
 for(const item of payload.snapshots)dedup.set(item.id,item);
 for(const item of payload.snapshots)await putItem(item);
 if(payload.format==='team23-v2'){
  if(!Array.isArray(payload.dailyNotes))throw Error('날짜별 사유·업무상태 백업이 손상되었습니다.');
  for(const note of payload.dailyNotes){if(note&&/^\d{4}-\d{2}-\d{2}$/.test(note.date))await putDailyNote({date:note.date,work:note.work||{},memo:note.memo||{}});}
 }else{
  const legacy=(payload.snapshots.slice().sort((a,b)=>b.created-a.created)[0]||existing[0]);
  if(legacy){const date=legacy.data.meta.asof,current=await getDailyNote(date)||{date,work:{},memo:{}};
    await putDailyNote({date,work:{...(payload.work||{}),...current.work},memo:{...(payload.memo||{}),...current.memo}});
  }
 }
 const all=await storeAll(),latest=[...latestByDay(all).values()].sort((a,b)=>b.data.meta.asof.localeCompare(a.data.meta.asof)||b.created-a.created)[0];
 if(latest)await loadItem(latest.id);statusMsg(`백업 복원 완료 · ${payload.snapshots.length}개 업로드 기록 및 날짜별 사유·업무상태 복구.`, 'ok');
 }catch(e){statusMsg('복원 실패: '+e.message,'error');}finally{$('restoreInput').value='';}}
function csvCell(v){return '"'+String(v??'').replace(/"/g,'""')+'"';}
function downloadCSV(all=false){if(!records.length)return statusMsg('먼저 엑셀을 업로드해 주세요.','error');const head=['기준일','광고주','광고계정ID','담당자','매체','전월매출','당월목표','당월누적','당월예상','예상차액','목표달성률','최근4평일추이','사유/조치메모','업무상태'];const lines=[head,...(all?records:filteredRows()).map(r=>[DATA.meta.asof,r.name,r.id,r.owner,r.channel,r.oct,r.target,r.actual,r.forecast,r.gap,pct(r.rate),pct(r.trendPct),memo[stableKey(r)]||'',work[stableKey(r)]||'미착수'])];downloadText('\uFEFF'+lines.map(x=>x.map(csvCell).join(',')).join('\r\n'),`2-3팀_${DATA.meta.asof}_사유포함_광고계정.csv`,'text/csv;charset=utf-8;');}

function card(value,label,foot,icon,cl=''){return `<div class="metric"><div class="caption"><span>${label}</span><span class="icon">${icon}</span></div><div class="number ${cl}">${value}</div><div class="foot">${foot}</div></div>`;}
function renderOverview(){
 $('asofChip').innerHTML=`<i class="dot"></i> ${dateLabel(DATA.meta.asof)} 기준`;
 $('lastLoaded').textContent=`${dateLabel(DATA.meta.asof)} · ${DATA.meta.filename}`;
 $('heroGap').textContent=signed(team.gap);$('heroPeople').textContent=`${ownBad} / ${staff.length}명`;$('heroAccounts').textContent=`${low.length} / ${active.length}개`;
 $('metricCards').innerHTML=[card(won(team.target),`${shortMonth()} 목표`,'3개 매체 · 2-3팀 소속 계정 합계','◎'),card(won(team.forecast),`${shortMonth()} 월말 예상`,dateLabel(DATA.meta.asof)+' 현재 원본 예측','◈'),card(pct(team.target>0?team.forecast/team.target*100:null),'팀 예상 달성률',signed(team.gap)+' · 목표 대비','↗',team.gap>=0?'pos':'neg'),card(won(team.actual),`${shortMonth()} 현재 누적`,`${dateLabel(DATA.meta.asof)} 기준 · ${active.length}개 목표 계정`,'◷')].join('');
 const tips={'쇼핑검색광고':'검색어·SKU별 소진, 소재와 순위 진단 후 목표 미달 계정 집중 점검','파워링크':'키워드 입찰·CPC·예산 소진을 점검해 미달액부터 회복','GFA':'GFA/ADVoost 캠페인·타겟·소재별 수수료매출과 집행 확인'};
 $('channels').innerHTML=chs.map((c,i)=>{const x=groupSum(getCh(c)),rat=x.target>0?x.forecast/x.target*100:0;return `<article class="card channel c${i}"><div class="channel-heading"><span class="name">${c}</span><span class="tag ${rat<100?'bad':''}">${rat>=100?'목표 초과':'목표 미달'}</span></div><div class="forecast-label">${shortMonth()} 예상 매출</div><div class="big">${won(x.forecast)}</div><div class="money-pair"><span class="k">목표 ${won(x.target)}</span><strong class="${x.gap<0?'neg':'pos'}">${signed(x.gap)}</strong></div><div class="progress"><span style="width:${Math.min(100,Math.max(0,rat))}%"></span></div><div class="progress-caption"><span>목표 달성률</span><b>${pct(x.target>0?rat:null)}</b></div><div class="channel-tip">${tips[c]}</div></article>`;}).join('');
}
function renderAudit(){const a=DATA.meta,compare=a.summaryRows||[],smTarget=sum(compare,'summaryTarget'),smForecast=sum(compare,'summaryForecast');
 $('auditMetrics').innerHTML=`<div class="audit-metric"><div class="cap">① 상세계정 기준 · 제외 인원 반영</div><div class="num">${won(team.forecast)}</div><div class="small">${records.length}개 계정 / 목표 ${won(team.target)}<br>서울2-3팀(김재현) + 제외 2명 적용</div></div><div class="audit-metric"><div class="cap">② '인원별' 소속 표시 기준 · 제외 인원 반영</div><div class="num">${won(smForecast)}</div><div class="small">원본 비교용 / 목표 ${won(smTarget)}<br>인원별과 상세계정의 소속 표기가 다를 수 있음</div></div>`;
 const row=compare.map(r=>`<tr><td>${esc(r.channel)}</td><td>${won(r.detailTarget)}</td><td>${won(r.detailForecast)}</td><td>${won(r.summaryTarget)}</td><td>${won(r.summaryForecast)}</td></tr>`);row.push(`<tr><td>3매체 합계</td><td>${won(team.target)}</td><td>${won(team.forecast)}</td><td>${won(smTarget)}</td><td>${won(smForecast)}</td></tr>`);$('auditRows').innerHTML=row.join('');
 const conflict=a.memberConflict||{summaryOnly:[],detailOnly:[]};$('auditNotes').innerHTML=`<b>기준 적용:</b> 매체별 계정 시트에서 2-3팀으로 기재된 계정만 집계. 최연아·안희성 항상 제외.<br><b>소속 불일치:</b> 인원별만 표시 ${esc(conflict.summaryOnly.join('·')||'없음')} / 계정 상세에만 표시 ${esc(conflict.detailOnly.join('·')||'없음')}.<br><b>추이 검증:</b> ${fmt(a.rawMismatch||0)}개 계정은 raw 일별 합계가 매체별 누적과 불일치하거나 매칭 자료가 없어 최근 평일 추이를 계산하지 않았습니다. (계정 금액은 매체별 시트 원본 사용)`;
}
function ownerRatio(owner,c){const r=groupSum(getStaff(owner).filter(x=>x.channel===c));return r.target>0?r.forecast/r.target*100:null;}
function renderStaff(){
 $('staffRows').innerHTML=ownerPerf.map(p=>{
  const problem=chs.filter(c=>ownerRatio(p.owner,c)!==null&&ownerRatio(p.owner,c)<100).length;
  const badges=chs.map(c=>{const rate=ownerRatio(p.owner,c);return `<span class="staff-pct ${rate!==null&&rate<100?'neg':'pos'}">${pct(rate)}</span>`;}).join('');
  return `<div class="staff-row"><div class="staff-who"><span class="avatar">${esc(p.owner).slice(0,1)}</span><span><span class="staff-name">${esc(p.owner)}</span><span class="staff-meta">${problem?problem+'개 매체 목표 미달':'목표 방어'}</span></span></div><span class="staff-total ${p.gap<0?'neg':'pos'}">${signed(p.gap)}</span>${badges}<button data-owner="${esc(p.owner)}" class="mini-btn">계정 보기</button></div>`;
 }).join('');
 $('staffRows').querySelectorAll('[data-owner]').forEach(b=>b.onclick=()=>{setFilters({owner:b.dataset.owner,status:'under'});});
 $('memberWarning').innerHTML=`<b>담당자 집계 정책:</b> 인원별·매체별 시트 소속 불일치와 관계없이 <b>김재현 팀 매체별 계정</b> 기준. <b>최연아·안희성은 완전 제외</b>합니다.`;
}
function renderInsights(){const maximum=Math.max(...ownerPerf.map(p=>Math.abs(p.gap)),1);
 $('ownerBars').innerHTML=ownerPerf.map(p=>`<div class="bar-line"><div class="name">${esc(p.owner)}</div><div class="area"><span style="width:${Math.max(2,Math.abs(p.gap)/maximum*100)}%;background:${p.gap<0?'#ef7a84':'#4d78f1'}"></span></div><b class="${p.gap<0?'neg':'pos'}">${signed(p.gap)}</b></div>`).join('');
 const biggestDef=active.filter(x=>x.gap<0).sort(byGap)[0],biggestPlus=active.filter(x=>x.gap>0).sort((a,b)=>b.gap-a.gap)[0];
 const chWorst=chs.map(c=>({channel:c,...groupSum(getCh(c))})).sort((a,b)=>a.gap-b.gap)[0],worst=ownerPerf[0];
 const points=[];if(worst&&worst.gap<0)points.push({title:`우선 관리 인원: ${worst.owner}`,detail:`목표 대비 예상 부족액 ${won(-worst.gap)} · 미달 매체 및 계정부터 대응`});
 if(chWorst&&chWorst.gap<0)points.push({title:`매체 집중 관리: ${chWorst.channel}`,detail:`매체 전체 목표 부족액 ${won(-chWorst.gap)} · 해당 매체 상위 미달 계정 우선`});
 if(biggestDef)points.push({title:`최대 부족 계정: ${biggestDef.name}`,detail:`${biggestDef.owner} / ${biggestDef.channel} / 예상 부족 ${won(-biggestDef.gap)}`});
 if(biggestPlus)points.push({title:`초과 실적 유지: ${biggestPlus.name}`,detail:`${biggestPlus.owner} / ${biggestPlus.channel} / 목표 초과 ${won(biggestPlus.gap)}`});
 $('insights').innerHTML=points.length?points.map((x,i)=>`<div class="insight"><span class="n">0${i+1}</span><div><b>${esc(x.title)}</b><p>${esc(x.detail)}</p></div></div>`).join(''):'<p class="muted">현재 점검 대상 없음</p>';
}
const advice={'쇼핑검색광고':'검색어·SKU 소진 / 입찰 / 상위 상품 노출 확인','파워링크':'키워드 입찰·노출·예산 소진 점검','GFA':'GFA·ADVoost 분리 후 소재·타겟 점검'};
function setFilters({owner='',channel='',status='under',query=''}){$('ownerFilter').value=owner;$('channelFilter').value=channel;$('statusFilter').value=status;$('q').value=query;currentPage=1;renderAccounts();$('accountSection').scrollIntoView({behavior:'smooth'});}
function renderFocus(){
 $('focusCards').innerHTML=chs.map(c=>`<div class="card tri-card"><div class="tri-head"><strong>${c}</strong><span class="tag ${c==='GFA'?'bad':'warn'}">부족액 TOP 5</span></div><div class="tri-sub">예상 부족액이 큰 계정 중심</div>${byTop(getCh(c),5).map((r,i)=>`<div class="focus-item"><span class="index">0${i+1}</span><div><div class="adv"><button data-key="${esc(stableKey(r))}">${esc(r.name)}</button></div><div class="meta">${esc(r.owner)} · ${esc(r.id)}</div></div><div class="gap">−${won(-r.gap)}</div></div>`).join('')}<div class="tri-foot"><b>우선 점검:</b> ${advice[c]}</div></div>`).join('');
 $('focusCards').querySelectorAll('[data-key]').forEach(b=>b.onclick=()=>{const r=records.find(x=>stableKey(x)===b.dataset.key);if(r)setFilters({owner:r.owner,channel:r.channel,query:r.id});});
}
function statusFor(r){if(r.target<=0)return '목표 없음';if(r.rate<80)return '긴급';if(r.gap<0)return '개선 필요';return '달성';}
function actionFor(r){if(r.target<=0)return '목표 및 운영 여부 확인';if(!r.actual&&r.oct>=200000)return '현재 매출 0원 · 캠페인 중단/소진 확인';return advice[r.channel];}
function filteredRows(){const q=$('q').value.trim().toLowerCase(),o=$('ownerFilter').value,c=$('channelFilter').value,s=$('statusFilter').value,sort=$('sortFilter').value;
 const arr=records.filter(r=>{if(o&&r.owner!==o||c&&r.channel!==c)return false;if(q&&!(r.name.toLowerCase().includes(q)||r.id.toLowerCase().includes(q)||(memo[stableKey(r)]||'').toLowerCase().includes(q)))return false;
  if(s==='under'&&!(r.target>0&&r.gap<0))return false;if(s==='critical'&&!(r.target>0&&r.rate<80))return false;if(s==='down'&&!(r.target>0&&r.trend==='down'))return false;if(s==='over'&&!(r.target>0&&r.gap>=0))return false;if(s==='all'&&r.target<=0)return false;if(s==='inactive'&&r.target>0)return false;return true;});
 arr.sort((a,b)=>sort==='attain'?(a.rate??Infinity)-(b.rate??Infinity):sort==='forecast'?b.forecast-a.forecast:sort==='growth'?b.gap-a.gap:sort==='trend'?(a.trendPct??Infinity)-(b.trendPct??Infinity):a.gap-b.gap);return arr;
}
function recentTrend(r){if(r.trendPct===null)return '<span class="trend-pill">자료 불충분</span>';const ar=r.trendPct>0?'↗':r.trendPct<0?'↘':'→';return `<span class="trend-pill ${r.trend}">${ar} ${r.trendPct>0?'+':''}${r.trendPct.toFixed(1)}%</span>`;}
function prevCompared(r){if(!previousSnapshot)return '';const old=previousSnapshot.data.rows.find(x=>stableKey(x)===stableKey(r));if(!old)return '';const delta=r.forecast-old.forecast;if(!delta)return '<small class="previous-data">이전 업로드 예상매출 동일</small>';return `<small class="previous-data ${delta<0?'down':'up'}">이전 업로드 대비 ${signed(delta)}</small>`;}
function workStateClass(status){return status==='완료'?'state-done':status==='진행중'?'state-in-progress':'state-not-started';}
function renderAccounts(){currentlyShown=filteredRows();const pages=Math.max(1,Math.ceil(currentlyShown.length/pageSize));currentPage=Math.max(1,Math.min(currentPage,pages));const start=(currentPage-1)*pageSize,shown=currentlyShown.slice(start,start+pageSize);
 $('accountTbody').innerHTML=shown.length?shown.map(r=>{const st=statusFor(r),key=stableKey(r),state=work[key]||'미착수',klass=st==='긴급'?'danger':st==='개선 필요'?'warning':'';
 return `<tr><td><div class="account-title">${esc(r.name)}</div><span class="account-sub">ID ${esc(r.id)}</span></td><td><span class="owner-name">${esc(r.owner)}</span></td><td><span class="ch-tag ${r.channel}">${r.channel}</span></td><td class="num">${won(r.oct)}</td><td class="num">${won(r.target)}</td><td class="num"><b>${won(r.forecast)}</b><span class="account-sub">누적 ${won(r.actual)}</span>${prevCompared(r)}</td><td class="num"><div class="gap-line ${r.gap<0?'negative':'positive'}">${signed(r.gap)}</div></td><td class="num"><span class="status ${klass}">${st} ${pct(r.rate)}</span></td><td>${recentTrend(r)}</td><td><div class="action-hint">${esc(actionFor(r))}</div></td><td><textarea class="reason-field" data-reason="${esc(key)}" aria-label="사유·조치 메모 ${esc(r.name)}" placeholder="매출 변동 사유, 광고주 피드백, 조치 계획">${esc(memo[key]||'')}</textarea></td><td><span class="work-select-wrap ${workStateClass(state)}"><select class="done-select" aria-label="${esc(r.name)} 업무상태" data-work="${esc(key)}">${['미착수','진행중','완료'].map(s=>`<option value="${s}" ${s===state?'selected':''}>${s}</option>`).join('')}</select></span></td></tr>`;
 }).join(''):'<tr><td colspan="12" class="empty">조건에 맞는 광고계정이 없습니다. 필터를 조정해 주세요.</td></tr>';
 $('accountTbody').querySelectorAll('[data-reason]').forEach(el=>{el.addEventListener('change',()=>{memo[el.dataset.reason]=el.value.slice(0,3000);persistNotes();});el.addEventListener('blur',()=>{memo[el.dataset.reason]=el.value.slice(0,3000);persistNotes();});});
 $('accountTbody').querySelectorAll('[data-work]').forEach(el=>el.onchange=()=>{work[el.dataset.work]=el.value;el.closest('.work-select-wrap').className='work-select-wrap '+workStateClass(el.value);persistNotes();});
 $('tableCount').textContent=`${fmt(currentlyShown.length)}개 계정 중 ${shown.length?fmt(start+1):'0'}–${fmt(start+shown.length)} 표시`;
 $('pageLabel').textContent=`${currentPage} / ${pages}`;$('prevBtn').disabled=currentPage<=1;$('nextBtn').disabled=currentPage>=pages;
}
function renderAll(){document.body.classList.remove('empty-mode');resetComputed();$('prevRevenueHeader').textContent=prevMonth()+' 매출';$('targetRevenueHeader').textContent=shortMonth()+' 목표';$('forecastRevenueHeader').textContent=shortMonth()+' 예상';$('ownerFilter').innerHTML='<option value="">담당자 전체</option>'+[...new Set(records.map(r=>r.owner))].sort().map(x=>`<option value="${esc(x)}">${esc(x)}</option>`).join('');
 $('q').value='';$('ownerFilter').value='';$('channelFilter').value='';$('statusFilter').value='under';$('sortFilter').value='deficit';currentPage=1;
 renderOverview();renderAudit();renderStaff();renderInsights();renderFocus();renderAccounts();
 $('methodText').innerHTML=`<b>파일:</b> ${esc(DATA.meta.filename)} / ${dateLabel(DATA.meta.asof)}<br>① 매체별 시트에서 2-3팀(김재현) 계정만 집계, <b>최연아·안희성 제외</b>. ② 원본 당월 예상−목표를 차액으로 표시. 예상값은 확정 실적이 아닙니다. ③ GFA는 GFA/ADVoost 합산. ④ 수수료 매출로, 광고주 구매 전환매출·ROAS와는 다릅니다. ⑤ raw 누적과 매체별 누적이 불일치하는 계정은 최근 4평일 추이를 표시하지 않습니다. ⑥ 날짜를 선택하면 <b>그 날짜 기준 업로드 스냅샷</b>으로 화면을 바꾸며 사유·업무상태도 날짜별로 독립 저장합니다. 아직 업로드하지 않은 날짜의 예상매출은 이전 데이터를 재활용하지 않습니다. ⑦ 업로드 이력과 메모는 <b>현재 브라우저에만 저장</b>됩니다. 다른 직원 PC와 자동 공유되지 않으니 전체 백업 JSON을 내려받으세요.`;
}
function setupEvents(){
 $('xlsxInput').addEventListener('change',e=>handleUpload(e.target.files[0]));
 const drop=$('dropZone');for(const type of ['dragenter','dragover'])drop.addEventListener(type,e=>{e.preventDefault();drop.classList.add('over');});for(const type of ['dragleave','drop'])drop.addEventListener(type,e=>{e.preventDefault();drop.classList.remove('over');});drop.addEventListener('drop',e=>handleUpload(e.dataTransfer.files[0]));
 $('datePicker').addEventListener('change',e=>selectDate(e.target.value));
 $('prevDateBtn').onclick=()=>{const i=availableDates.indexOf(DATA.meta.asof);if(i>0)selectDate(availableDates[i-1]);};
 $('nextDateBtn').onclick=()=>{const i=availableDates.indexOf(DATA.meta.asof);if(i>=0&&i<availableDates.length-1)selectDate(availableDates[i+1]);};
 $('todayBtn').onclick=()=>selectDate(localToday());
 $('backupBtn').onclick=backupAll;$('restoreBtn').onclick=()=>$('restoreInput').click();$('restoreInput').onchange=e=>restoreBackup(e.target.files[0]);$('noteCsvBtn').onclick=()=>downloadCSV(true);
 for(const id of ['q','ownerFilter','channelFilter','statusFilter','sortFilter'])$(id).addEventListener(id==='q'?'input':'change',()=>{currentPage=1;renderAccounts();});
 $('resetBtn').onclick=()=>setFilters({});$('csvBtn').onclick=downloadCSV;$('prevBtn').onclick=()=>{currentPage--;renderAccounts();};$('nextBtn').onclick=()=>{currentPage++;renderAccounts();};
}
async function boot(){setupEvents();try{
 const all=await storeAll();await migrateLegacyNotes(all);
 const latest=[...latestByDay(all).values()].sort((a,b)=>b.data.meta.asof.localeCompare(a.data.meta.asof)||b.created-a.created)[0];
 if(latest){await loadItem(latest.id);if(latest.data.meta.asof!==localToday())statusMsg(`오늘(${dateLabel(localToday())}) 자료는 아직 업로드되지 않았습니다. 마지막 저장일 ${dateLabel(latest.data.meta.asof)} 자료를 표시하고 있습니다. 날짜 선택으로 이전 기록을 열 수 있습니다.`);}else await renderHistory();
 }catch(e){console.error(e);statusMsg('브라우저에 저장된 이전 기록이 없거나 저장소 접근이 차단되었습니다. 파일을 업로드해 주세요.','error');}}
boot();
