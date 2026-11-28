(()=>{
'use strict';
const KEY='shin_rerun_reviewed_v1', VERSION='2026.11.28-r4';
const $=id=>document.getElementById(id);
const clean=v=>String(v??'').trim();
const norm=v=>clean(v).normalize('NFC').toLowerCase().replace(/\s+/g,' ');
const blank=()=>({history:[],primers:[],feedback:[],updated:'',historyName:'',primerName:''});
let db=blank(), storeError='';
try {const x=JSON.parse(localStorage.getItem(KEY)||'null');if(x&&Array.isArray(x.history)&&Array.isArray(x.primers))db={...blank(),...x,feedback:Array.isArray(x.feedback)?x.feedback:[]};}
catch(e){storeError='Không đọc được dữ liệu trước đây: '+e.message;}
const tags=(d,t)=>Array.from(d.getElementsByTagNameNS('*',t));
function persist(next){localStorage.setItem(KEY,JSON.stringify(next));db=next;}
const bad=v=>!clean(v)||/^(?:-|—|n\/a|na|none|không|chưa có)$/i.test(clean(v));
const region=p=>p>=16024&&p<=16365?'HV1':p>=73&&p<=340?'HV2':p>=438&&p<=576?'HV3':'Ngoài vùng';
const inRegion=p=>region(p)!=='Ngoài vùng';
function siteList(input){
 const out=[], re=/(^|[^A-Za-z0-9_])(\d{2,5})(?:\.(\d+))?\s*(DEL|INS|[ACGTRYSWKMBDHVN])?(?=$|[^A-Za-z0-9])/gi;
 for(const m of String(input).matchAll(re)){const p=Number(m[2]);if(!inRegion(p))continue;
  const x={pos:p,sub:m[3]||'',allele:(m[4]||'').toUpperCase()};
  if(!out.some(y=>y.pos===x.pos&&y.sub===x.sub&&y.allele===x.allele))out.push(x);
 }
 return out;
}
const STOP=new Set(['php','can','xác','xac','minh','chạy','chay','lai','lại','mồi','moi','cần','the','của','cho','trong','theo','trên','tren','vùng','vung','mẫu','mau','kiểm','tra','check','khi','tại','tai','with','from']);
function wordSet(v){return new Set(norm(v).match(/[\p{L}\p{N}]+/gu)?.filter(x=>x.length>2&&!STOP.has(x))||[]);}
function evidence(q,s){
 const query=norm(q), issue=norm(s);if(!query||!issue)return {score:0,why:''};
 if(query===issue)return {score:100,why:'Issues trùng hoàn toàn'};
 const a=siteList(q), b=siteList(s);
 if(a.length){const shared=a.map(x=>({x,y:b.find(y=>x.pos===y.pos&&x.sub===y.sub)}));
  if(shared.every(x=>x.y)){
   const alleled=shared.filter(x=>x.x.allele&&x.y.allele);
   const diff=alleled.filter(x=>x.x.allele!==x.y.allele);
   if(diff.length)return {score:55,why:'Cùng vị trí PHP, khác ký hiệu biến thể'};
   if(alleled.length===a.length)return {score:90,why:'Trùng vị trí và ký hiệu PHP'};
   return {score:75,why:'Cùng vị trí PHP'};
  }
  return {score:0,why:''};
 }
 const wa=wordSet(q),wb=wordSet(s);const same=[...wa].filter(x=>wb.has(x));
 const focus=same.filter(x=>/polyc|polyac|phase|lệch|lech|nhiễu|nhieu|signal|peak|snp|del|indel|php|hetero/i.test(x));
 if(same.length>=2||focus.length)return {score:focus.length?38:26,why:'Tương đồng mô tả Issues'};
 return {score:0,why:''};
}
async function readSheet(file,preferred){
 if(typeof JSZip==='undefined')throw Error('Không có JSZip trong Shin AI.');
 if(file.size>45*1024*1024)throw Error('Excel vượt 45 MB.');
 const zip=await JSZip.loadAsync(await file.arrayBuffer());
 const xml=async path=>{const f=zip.file(path);if(!f)throw Error('Thiếu '+path);
  const x=new DOMParser().parseFromString(await f.async('string'),'application/xml');
  if(tags(x,'parsererror').length)throw Error('Lỗi XML: '+path);return x;};
 const wb=await xml('xl/workbook.xml'),rels=await xml('xl/_rels/workbook.xml.rels'),sheets=tags(wb,'sheet');
 const sh=sheets.find(x=>norm(x.getAttribute('name'))===norm(preferred))||sheets[0];
 if(!sh)throw Error('Không thấy sheet trong Excel.');
 const rid=sh.getAttribute('r:id')||sh.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
 const rel=tags(rels,'Relationship').find(x=>x.getAttribute('Id')===rid);
 if(!rel)throw Error('Không tìm thấy relationship của sheet.');
 const target=rel.getAttribute('Target')||'';
 const path=new URL(target.startsWith('/')?target.slice(1):target.startsWith('xl/')?target:'xl/'+target,'https://shin.local/').pathname.slice(1);
 const str=n=>tags(n,'t').map(x=>x.textContent).join('');
 const ss=zip.file('xl/sharedStrings.xml')?tags(await xml('xl/sharedStrings.xml'),'si').map(str):[];
 const rows=tags(await xml(path),'row').map(row=>{
  const v={};for(const c of tags(row,'c')){const col=(c.getAttribute('r')||'').replace(/\d+/g,'');
   if(!col||col.length>2)continue;
   const kind=c.getAttribute('t'),raw=tags(c,'v')[0]?.textContent??'';
   v[col]=clean(kind==='s'?ss[Number(raw)]??'':kind==='inlineStr'?str(c):raw);
  }
  return {n:Number(row.getAttribute('r'))||0,v};
 });
 return {name:sh.getAttribute('name'),rows};
}
function parseHistory(rows){
 const head=rows.find(r=>norm(r.v.D)==='issues'&&/requirements?/.test(norm(r.v.E))&&norm(r.v.M)==='issues'&&/requirements?/.test(norm(r.v.N)));
 if(!head)throw Error('Sheet RERUN_LIMS2: không nhận ra cột Issues/Requirement (D–E và M–N).');
 const unique=new Map();let skipped=0;
 for(const row of rows){if(row.n<=head.n)continue;const v=row.v;
  const lid=[v.A,v.B].map(clean).find(x=>/^LN[_-][A-Z0-9_-]+$/i.test(x))||'';
  if(!lid)continue;
  for(const [run,ic,rc,bc] of [[2,'D','E','F'],[3,'M','N','O'],[4,'T','U','V']]){
   const issue=clean(v[ic]),requirement=clean(v[rc]);if(bad(issue)||bad(requirement)){if(!bad(issue)||!bad(requirement))skipped++;continue;}
   const r={lid,batch:clean(v[bc]),run,issue,requirement,row:row.n,source:'RERUN_LIMS2'};
   unique.set([norm(lid),run,norm(issue),norm(requirement)].join('|'),r);
  }
 }
 if(!unique.size)throw Error('Không tìm thấy quyết định rerun hợp lệ; kho cũ được giữ nguyên.');
 return {history:[...unique.values()],skipped};
}
function parsePrimers(rows){
 const h=rows.find(r=>norm(r.v.B)==='mã mồi trên hệ thống'&&norm(r.v.G)==='chiều'&&norm(r.v.I)==='trạng thái');
 if(!h)throw Error('Không nhận ra các cột B (mã mồi), G (chiều), I (trạng thái).');
 const unique=new Map();
 for(const row of rows){if(row.n<=h.n)continue;const v=row.v,name=clean(v.B);if(!name)continue;
  unique.set(norm(name),{name,position:clean(v.D),region:clean(v.E),direction:clean(v.G),note:clean(v.H),status:clean(v.I)});
 }
 if(!unique.size)throw Error('Danh sách mồi không có dữ liệu.');
 return [...unique.values()];
}
function records(data){
 const history=data.history.map(x=>({...x,source:'Lịch sử đã review'}));
 const approved=(data.feedback||[]).filter(x=>x.status==='confirmed').map(x=>({...x,run:'Reviewer',row:'Đã lưu',source:'Reviewer xác nhận'}));
 return history.concat(approved);
}
function requirementKey(value){
 const raw=norm(value);
 // Only normalize simple PCR requirements; preserve order for multi-step workflows.
 const m=raw.match(/^(pcr-bd|pcr1|pcr2)\s+([\w,-]+(?:\s*,\s*[\w-]+)*)$/i);
 if(!m)return raw.replace(/\s*,\s*/g,', ');
 const aliases={'hv2-109f':'f109','hv1-16190f':'f16190'};
 return m[1]+' '+[...new Set(m[2].split(',').map(x=>{const a=x.trim();return aliases[a]||a;}))].sort().join(', ');
}
function recommend(q,data){
 const query=clean(q),sites=siteList(query),matches=[];
 if(!query||(!sites.length&&!wordSet(query).size))return {mode:'Nhập Issues hoặc PHP cần xác minh.',matches:[],ranked:[],candidates:[],sites};
 const rr=records(data);
 for(const r of rr){const ev=evidence(query,r.issue);if(!ev.score)continue;
  const score=ev.score+(r.source==='Reviewer xác nhận'?12:0);
  matches.push({...r,score,why:ev.why});
 }
 matches.sort((a,b)=>b.score-a.score||norm(a.lid).localeCompare(norm(b.lid)));
 const top=matches[0]?.score||0,selected=matches.filter(x=>x.score>=Math.max(22,top-14)).slice(0,2000);
 const groups=new Map();
 for(const r of selected){const key=requirementKey(r.requirement);
  if(!groups.has(key))groups.set(key,{requirement:r.requirement,score:r.score,source:new Set(),lids:new Set(),cases:0});
  const a=groups.get(key);a.score=Math.max(a.score,r.score);a.source.add(r.source);a.lids.add(norm(r.lid));a.cases++;
 }
 const ranked=[...groups.values()].map(x=>({requirement:x.requirement,score:x.score,n:x.lids.size,cases:x.cases,source:[...x.source].join(' + ')}))
 .sort((a,b)=>b.score-a.score||b.n-a.n||a.requirement.localeCompare(b.requirement));
 const candidates=[],blocked=[];
 for(const t of sites)for(const p of data.primers){
  if(norm(p.status)!=='on')continue;
  if(/(?:^|[-_])16258R\b|^R16258$/i.test(p.name)){blocked.push(p.name);continue;}
  // Skip conflicting coordinates rather than silently choosing one.
  const mm=p.position.match(/^(\d{1,5})\s*([FR])?$/i),direction=norm(p.direction);
  const named=p.name.match(/(?:^|[-_])(\d{1,5})([FR])$/i);
  if(mm&&named&&Number(named[1])!==Number(mm[1]))continue;
  if(!mm)continue;const anchor=Number(mm[1]);let dir='';
  if(['forward','f','xuôi','xuoi'].includes(direction))dir='F';
  if(['reverse','r','ngược','nguoc'].includes(direction))dir='R';
  if(!dir)dir=(mm[2]||'').toUpperCase();
  if(!dir||(mm[2]&&mm[2].toUpperCase()!==dir))continue;
  const distance=((dir==='F'?t.pos-anchor:anchor-t.pos)+16569)%16569;
  if(distance>0&&distance<=700)candidates.push({pos:t.pos,name:p.name,direction:dir,anchor,distance,note:p.note});
 }
 candidates.sort((a,b)=>a.pos-b.pos||a.distance-b.distance);
 return {mode:!rr.length?'Chưa có quyết định được nạp.':!selected.length?'Chưa thấy ca tương tự; cần reviewer quyết định.':selected[0].why,
 matches:selected,ranked,candidates:candidates.slice(0,150),sites,blocked:[...new Set(blocked)]};
}
function renderTable(host,heads,rows,choose){
 host.replaceChildren();if(!rows.length){host.textContent='Chưa có dữ liệu phù hợp.';return;}
 const t=document.createElement('table'),tr=t.createTHead().insertRow(),body=t.createTBody();
 for(const h of heads){const th=document.createElement('th');th.textContent=h;tr.append(th);}
 rows.forEach((row,i)=>{const tr=body.insertRow();row.forEach(v=>{const c=tr.insertCell();c.textContent=String(v??'');});
  if(choose){const cell=tr.insertCell(),btn=document.createElement('button');btn.type='button';btn.textContent='Chọn';btn.onclick=()=>choose(i);cell.append(btn);}
 });if(choose){const c=document.createElement('th');c.textContent='';tr.append(c);}host.append(t);
}
function saveText(name,content,type){const url=URL.createObjectURL(new Blob([content],{type}));const a=document.createElement('a');a.href=url;a.download=name;document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function exportFeedback(){
 const head=['Updated','LID','Batch','Issues','Requirement','Status','Note'];
 const esc=v=>{let s=String(v??'');if(/^[=+\-@]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"';};
 const rows=(db.feedback||[]).map(x=>[x.updated,x.lid,x.batch,x.issue,x.requirement,x.status,x.note]);
 saveText('Shin_Rerun_Reviewer_Feedback.csv','\uFEFF'+[head,...rows].map(row=>row.map(esc).join(',')).join('\r\n'),'text/csv;charset=utf-8');
}
function mount(){
 const nav=$('agentFeatureNavInner'),learn=$('workspaceLearn'),main=learn?.parentElement;if(!nav||!main||$('shinRerunAI'))return;
 const host=document.createElement('section');host.id='workspacePrimerRerun';host.className='workspace-pane';host.style.setProperty('--sl-line','#d9e5e3');host.style.setProperty('--sl-blue','#276d93');main.append(host);
 const tab=document.createElement('button');tab.type='button';tab.className='agent-feature-tab nav-feature-card';tab.dataset.target='primer-rerun';tab.setAttribute('aria-label','Primer Rerun');tab.innerHTML='<span class="nav-card-heading"><svg class="nav-card-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round"><path d="M4 3c0 8 16 10 16 18M20 3C20 11 4 13 4 21M7 5h10M8 19h8M8 9h8M8 15h8"/></svg><span>Primer Rerun</span></span>';
 const learnTab=nav.querySelector('[data-target="learn"]');if(learnTab)learnTab.before(tab);else nav.append(tab);
 const navStyle=document.createElement('style');navStyle.textContent='body:not([data-view="primer-rerun"]) #workspacePrimerRerun{display:none!important}body[data-view="primer-rerun"] #workspacePrimerRerun{display:block!important}body[data-view="primer-rerun"] .workspace-pane:not(#workspacePrimerRerun){display:none!important}';document.head.append(navStyle);
 const sync=()=>{const active=document.body.dataset.view==='primer-rerun';tab.classList.toggle('is-active',active);tab.setAttribute('aria-selected',String(active));if(active){tab.setAttribute('aria-current','page');const title=$('shin-sidebar-title')?.querySelector('h2');if(title)title.textContent='Primer Rerun';}else tab.removeAttribute('aria-current');};
 tab.addEventListener('click',()=>{document.body.dataset.view='primer-rerun';nav.querySelectorAll('[data-target]').forEach(t=>{if(t!==tab){t.classList.remove('is-active');t.setAttribute('aria-selected','false');t.removeAttribute('aria-current');}});try{history.replaceState(null,'','#primer-rerun');}catch(_){}sync();window.scrollTo({top:0,behavior:'auto'});});
 nav.addEventListener('click',e=>{const t=e.target.closest('[data-target]');if(t&&t!==tab&&location.hash==='#primer-rerun'){try{history.replaceState(null,'','#'+t.dataset.target);}catch(_){}}});
 new MutationObserver(sync).observe(document.body,{attributes:true,attributeFilter:['data-view']});
 if(location.hash==='#primer-rerun')tab.click();else sync();
 const style=document.createElement('style');style.id='srStyle';style.textContent='#shinRerunAI .sr-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}#shinRerunAI .sr-grid label{display:flex;flex-direction:column;gap:4px}#shinRerunAI input[type=text],#shinRerunAI textarea,#shinRerunAI select{width:100%;max-width:100%}#shinRerunAI .sr-actions{display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin:10px 0}#shinRerunAI .sr-muted{font-size:12px;opacity:.82}#shinRerunAI .sr-table{overflow-x:auto;max-height:330px;overflow-y:auto}#shinRerunAI .sr-table table{min-width:560px}#shinRerunAI hr{margin:20px 0;border:0;border-top:1px solid #c5d5d5}';
 document.head.append(style);
 const card=document.createElement('section');card.id='shinRerunAI';card.className='sl-card';
 card.innerHTML='<h2>Primer Rerun · Chọn mồi chạy lại</h2>'+
 '<p>Tra cứu từ các quyết định trong sheet <b>RERUN_LIMS2</b>, danh mục mồi và quyết định reviewer đã xác nhận. Đây là hỗ trợ lựa chọn, không tự thay thế review chromatogram.</p>'+
 '<div class="sr-grid"><label>Lịch sử rerun đã review (.xlsx)<input id="srHistory" type="file" accept=".xlsx,.xlsm"></label><label>Danh sách mồi LIMS2 (.xlsx)<input id="srPrimers" type="file" accept=".xlsx,.xlsm"></label></div>'+
 '<div class="sr-actions"><button id="srImport" type="button">Cập nhật dữ liệu học</button><button id="srBackup" type="button">Sao lưu JSON</button><label>Khôi phục JSON <input id="srRestore" type="file" accept=".json"></label></div>'+
 '<p id="srStatus" role="status" class="sr-muted"></p><p class="sr-muted">Dữ liệu Excel, LID và phản hồi chỉ lưu trong trình duyệt hiện tại, không tự tải lên GitHub. Nhập lại Excel khi sheet nguồn thay đổi; các phản hồi reviewer đã lưu vẫn được giữ.</p><hr>'+
 '<label>Issues / PHP cần xác minh<textarea id="srQuery" rows="2" placeholder="Ví dụ: 152Y?; 309C; 573.1C; HV2 polyC; hoặc mô tả Issues"></textarea></label>'+
 '<div class="sr-actions"><button id="srPredict" type="button">Tra cứu, đề xuất</button><button id="srCopy" type="button">Sao chép phương án đầu</button></div>'+
 '<p id="srResult" role="status"></p><h3>Phương án trong dữ liệu đã review</h3><div id="srRank" class="sr-table"></div>'+
 '<h3>Mồi ON có hướng đọc về PHP</h3><p class="sr-muted">Chỉ sàng lọc theo tọa độ và chiều mồi (tối đa 700 nt); bỏ qua mồi có tọa độ mâu thuẫn hoặc không xác định, KHÔNG chứng minh mồi khuếch đại phù hợp, độ phủ thực tế hoặc khả năng xác minh PHP. Mồi HV1-16258R không được tự đề xuất.</p><div id="srCandidates" class="sr-table"></div>'+
 '<h3>Ca tương tự (tối đa 100)</h3><div id="srCases" class="sr-table"></div><hr>'+
 '<h3>Kết luận reviewer và học lại</h3><p class="sr-muted">Chỉ bản ghi đánh dấu “Đã xác nhận” mới được dùng cho lần gợi ý tiếp theo. Bản ghi chờ review không tác động gợi ý.</p>'+
 '<div class="sr-grid"><label>LID<input id="srLid" type="text" placeholder="LN_26_XXXX"></label><label>Batch (không bắt buộc)<input id="srBatch" type="text" placeholder="MS_..."></label><label>Phương án / mồi được chọn<input id="srChosen" type="text" placeholder="Tên mồi hoặc Requirement reviewer chọn"></label><label>Trạng thái<select id="srDecision"><option value="pending">Chờ review</option><option value="confirmed">Đã xác nhận bởi reviewer</option></select></label></div>'+
 '<label>Ghi chú reviewer<textarea id="srNote" rows="2" placeholder="Lý do chọn mồi / kết quả sau rerun"></textarea></label>'+
 '<div class="sr-actions"><button id="srSave" type="button">Lưu kết luận</button><button id="srExport" type="button">Xuất phản hồi CSV</button></div>'+
 '<p id="srFeedbackStatus" role="status"></p><div id="srFeedback" class="sr-table"></div>';
 host.prepend(card);
 const navSections=document.createElement('div');navSections.className='sr-actions';navSections.setAttribute('aria-label','Các phần Primer Rerun');
 const sections=[['Chọn mồi','srQuery'],['Lịch sử rerun','srCases'],['Danh mục mồi','srCatalog'],['Dữ liệu học','srHistory'],['Reviewer','srLid']];
 for(const [label,id] of sections){const b=document.createElement('button');b.type='button';b.textContent=label;b.onclick=()=>$(id)?.scrollIntoView({behavior:'smooth',block:'center'});navSections.append(b);}
 card.querySelector('h2').after(navSections);
 const catalog=document.createElement('details');catalog.id='srCatalog';catalog.innerHTML='<summary>Danh mục mồi LIMS2 (tất cả trạng thái)</summary><div id="srCatalogTable" class="sr-table"></div>';card.append(catalog);
 navSections.querySelectorAll('button')[2].addEventListener('click',()=>{catalog.open=true;});
 const historyAll=document.createElement('details');historyAll.innerHTML='<summary>Tra cứu toàn bộ lịch sử RERUN_LIMS2</summary><input id="srAllSearch" type="search" placeholder="LID, Issues, Requirement, Batch"><p id="srAllCount"></p><div id="srAllTable" class="sr-table"></div>';card.append(historyAll);
 navSections.querySelectorAll('button')[1].onclick=()=>{historyAll.open=true;historyAll.scrollIntoView({behavior:'smooth',block:'start'});};
 function catalogWarning(x){
 const a=clean(x.position).match(/^(\d{1,5})\s*[FR]?$/i),b=clean(x.name).match(/(?:^|[-_])(\d{1,5})[FR]$/i);
 return a&&b&&Number(a[1])!==Number(b[1])?' · Kiểm tra tọa độ':!a?' · Tọa độ cần xác minh':!clean(x.direction)?' · Thiếu chiều':'';
 }
 function refreshCatalog(){renderTable($('srCatalogTable'),['Mồi','Tọa độ','Vùng','Chiều','Trạng thái'],db.primers.map(x=>[x.name,x.position,x.region,x.direction,(x.status||'Chưa xác định')+catalogWarning(x)]));}
 function refreshAll(){const q=norm($('srAllSearch').value);const a=db.history.filter(x=>!q||norm([x.lid,x.batch,x.issue,x.requirement].join(' ')).includes(q));$('srAllCount').textContent=a.length+' ca phù hợp · hiển thị tối đa 200';renderTable($('srAllTable'),['LID','Batch','Issues','Requirement'],a.slice(0,200).map(x=>[x.lid,x.batch,x.issue,x.requirement]));}
 $('srAllSearch').addEventListener('input',refreshAll);
 function status(message=''){refreshCatalog();refreshAll();
  const approved=(db.feedback||[]).filter(x=>x.status==='confirmed').length;
  const pending=(db.feedback||[]).length-approved;
  $('srStatus').textContent=message+' · '+db.history.length+' ca lịch sử · '+db.primers.length+' mồi · '+approved+' reviewer xác nhận · '+pending+' chờ review · Cập nhật: '+(db.updated||'chưa có')+' · '+[db.historyName,db.primerName].filter(Boolean).join(' / ');
 }
 let top='';
 function feedbackTable(){
  const rows=(db.feedback||[]).slice().reverse().slice(0,50);
  renderTable($('srFeedback'),['LID','Issues','Quyết định','Trạng thái','Ngày cập nhật'],rows.map(x=>[x.lid,x.issue,x.requirement,x.status==='confirmed'?'Đã xác nhận':'Chờ review',x.updated]),i=>{
   const x=rows[i];$('srLid').value=x.lid||'';$('srBatch').value=x.batch||'';$('srQuery').value=x.issue||'';$('srChosen').value=x.requirement||'';$('srDecision').value=x.status||'pending';$('srNote').value=x.note||'';run(false);$('srFeedbackStatus').textContent='Đã nạp kết luận để chỉnh sửa. Bấm Lưu kết luận để ghi đè.';
  });
 }
 function run(fill=true){
  const r=recommend($('srQuery').value,db);top=r.ranked[0]?.requirement||'';
  $('srResult').textContent=r.mode+' · '+r.matches.length+' ca tương ứng · '+(r.sites.map(x=>x.pos+(x.sub?'.'+x.sub:'')+(x.allele||'')+' ('+region(x.pos)+')').join(', ')||'không xác định được vị trí')+'. Tần suất trong lịch sử không phải xác suất đúng; reviewer cần đối chiếu trace, phạm vi và SOP.';
  renderTable($('srRank'),['Requirement / lựa chọn','LID độc lập','Mức khớp','Nguồn'],r.ranked.map(x=>[x.requirement,x.n,x.score>=100?'Rất gần':x.score>=75?'Cùng PHP':'Tham khảo',x.source]),i=>{$('srChosen').value=r.ranked[i].requirement;});
  renderTable($('srCandidates'),['PHP','Tên mồi','Chiều','Tọa độ','Cách nt','Ghi chú'],r.candidates.map(x=>[x.pos,x.name,x.direction,x.anchor,x.distance,x.note]),i=>{$('srChosen').value=r.candidates[i].name;});
  renderTable($('srCases'),['LID','Batch','Run','Issues','Requirement','Loại khớp','Nguồn'],r.matches.slice(0,100).map(x=>[x.lid,x.batch,x.run,x.issue,x.requirement,x.why,x.source]));
  if(fill&&!$('srChosen').value&&top)$('srChosen').value=top;
 }
 $('srImport').onclick=async()=>{const b=$('srImport');b.disabled=true;
  try {const h=$('srHistory').files[0],p=$('srPrimers').files[0];if(!h&&!p)throw Error('Chọn ít nhất một Excel.');
   const next={...db};let skipped=0;
   if(h){const data=await readSheet(h,'RERUN_LIMS2');const res=parseHistory(data.rows);next.history=res.history;next.historyName=h.name+' / '+data.name;skipped=res.skipped;}
   if(p){const data=await readSheet(p,'Sheet1');next.primers=parsePrimers(data.rows);next.primerName=p.name+' / '+data.name;}
   next.updated=new Date().toLocaleString('vi-VN');persist(next);
   status('Đã nhập dữ liệu; bỏ '+skipped+' cặp Issues/Requirement thiếu thông tin.');run();
  }catch(e){status('Lỗi: '+e.message+' (kho cũ chưa bị thay đổi)');}finally{b.disabled=false;}
 };
 $('srPredict').onclick=()=>run();
 $('srQuery').addEventListener('input',()=>{$('srChosen').value='';});
 $('srCopy').onclick=async()=>{try{if(!top)throw Error('Chưa có phương án từ lịch sử.');
  await navigator.clipboard.writeText(top);status('Đã sao chép: '+top);
 }catch(e){status(e.message);}};
 $('srSave').onclick=()=>{try{
   const lid=clean($('srLid').value),issue=clean($('srQuery').value),requirement=clean($('srChosen').value);
   if(!lid||!issue||!requirement)throw Error('Cần LID, Issues/PHP và phương án chọn.');
   if(!/^LN[_-][A-Z0-9_-]+$/i.test(lid))throw Error('LID không đúng dạng LN_...');
   const row={lid,issue,requirement,batch:clean($('srBatch').value),note:clean($('srNote').value),status:$('srDecision').value,updated:new Date().toLocaleString('vi-VN')};
   const feedback=(db.feedback||[]).filter(x=>!(norm(x.lid)===norm(lid)&&norm(x.issue)===norm(issue)));
   feedback.push(row);persist({...db,feedback});
   $('srFeedbackStatus').textContent=row.status==='confirmed'?'Đã lưu quyết định reviewer; được dùng cho gợi ý tiếp theo.':'Đã lưu ở trạng thái chờ review; chưa dùng để gợi ý.';
   status();feedbackTable();run(false);
  }catch(e){$('srFeedbackStatus').textContent='Không lưu được: '+e.message;}
 };
 $('srExport').onclick=()=>exportFeedback();
 $('srBackup').onclick=()=>saveText('Shin_Rerun_AI_Backup.json',JSON.stringify({format:KEY,version:VERSION,data:db},null,2),'application/json;charset=utf-8');
 $('srRestore').onchange=async()=>{try{
   const f=$('srRestore').files[0];if(!f)return;const x=JSON.parse(await f.text());
   if(x.format!==KEY||!Array.isArray(x.data?.history)||!Array.isArray(x.data?.primers)||!x.data.history.every(r=>typeof r.issue==='string'&&typeof r.requirement==='string')||!x.data.primers.every(r=>typeof r.name==='string')||(x.data.feedback&&!Array.isArray(x.data.feedback)))throw Error('File sao lưu không hợp lệ.');
   persist({...blank(),...x.data,feedback:Array.isArray(x.data.feedback)?x.data.feedback:[]});
   status('Đã khôi phục kho dữ liệu.');feedbackTable();run();
  }catch(e){status('Lỗi khôi phục: '+e.message);}};
 status(storeError);feedbackTable();run();
}
mount();if(!$('shinRerunAI'))document.addEventListener('DOMContentLoaded',mount,{once:true});
})();