const ATTENDANT_FIELDS = ['Atendente','Atendente responsável','Atendente Responsável','Responsável'];
const KEEP = ['Protocolo','Recebido em','Status','Secretaria','Assunto','Manifestação','Origem de cadastro','Resposta'];
const FINAL_FIELDS = ['Finalizado em','Finalizados em'];
const ACTIVE_STATUSES = ['Aberto','Agendado','Em andamento','Pendente'];
const CURRENT_MODE = 'single-file-v4';
const REMOVE = ['Encaminhado','Solicitante','CPFCNPJ','Telefone 01','Telefone 02','Telefone 03','E-mail','Lido','CEP','Endereço','Número','Bairro','Referência','Região','Unidade','Grupo','Latitude','Longitude','Visível','Categoria','Funcionário público','Liberado pelo ouvidor','Protocolo anterior','Arquivos','Urgente','Dados sigilosos','Opinião','Justificativa','Prorrogado','Retificado','Cartão SUS','Finalizado em','Identificador','Materiais','Status da Associação','Reincidente'];
const REQUIRED = KEEP;
const DB_NAME = 'ouvidoria-base-local';
const DB_VERSION = 1;
const DB_STORE = 'base';
const DB_KEY = 'atual';
let state = { rows: [], filtered: [], periodRows: [], fileName: '', lastUpdated: '', duplicates: new Set() };
let metricScope = 'all';
const $ = id => document.getElementById(id);
const dropZone = $('dropZone'), input = $('fileInput');

$('chooseFile').addEventListener('click', e => { e.stopPropagation(); input.click(); });
dropZone.addEventListener('click', e => { if(e.target !== $('chooseFile')) input.click(); });
dropZone.addEventListener('keydown', e => { if(e.key === 'Enter' || e.key === ' ') input.click(); });
['dragenter','dragover'].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.add('dragover'); }));
['dragleave','drop'].forEach(type => dropZone.addEventListener(type, e => { e.preventDefault(); dropZone.classList.remove('dragover'); }));
dropZone.addEventListener('drop', e => { const file = e.dataTransfer.files[0]; if(file) processFile(file); });
input.addEventListener('change', e => { if(e.target.files[0]) processFile(e.target.files[0]); e.target.value = ''; });

function notify(message, kind='success') { $('alertArea').innerHTML = `<div class="alert ${kind}">${escapeHtml(message)}</div>`; }
function escapeHtml(value='') { return String(value).replace(/[&<>'"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#039;','"':'&quot;'}[c])); }
function normalize(value) { return value === null || value === undefined ? '' : String(value).trim(); }
function parseDate(value) {
  const text = normalize(value); if(!text) return null;
  const match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if(match) return new Date(Number(match[3]), Number(match[2])-1, Number(match[1]), Number(match[4]||0), Number(match[5]||0), Number(match[6]||0));
  const d = new Date(value); return Number.isNaN(d.getTime()) ? null : d;
}
function dateKey(value) { const d = parseDate(value); return d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` : ''; }
function displayDate(value) { const d = parseDate(value); return d ? d.toLocaleDateString('pt-BR') + (normalize(value).includes(':') ? ' ' + d.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'}) : '') : normalize(value) || '—'; }
function safeClass(value) { return normalize(value).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,'-').replace(/[^a-z0-9-]/g,''); }

async function processFile(file) {
  $('uploadStatus').textContent = 'Lendo arquivo…'; $('alertArea').innerHTML = '';
  try {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type:'array', cellDates:true });
    if(!workbook.SheetNames.includes('Dados')) throw new Error('A planilha precisa conter uma aba chamada “Dados”.');
    const sheet = workbook.Sheets['Dados'];
    const matrix = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'' });
    if(!matrix.length) throw new Error('A aba “Dados” está vazia.');
    const headers = matrix[0].map(normalize);
    const missing = REQUIRED.filter(c => !headers.includes(c));
    if(missing.length) throw new Error(`Colunas essenciais ausentes: ${missing.join(', ')}.`);
    const indexes = Object.fromEntries(KEEP.map(c => [c, headers.indexOf(c)]));
    const finalIndex = headers.findIndex(c => FINAL_FIELDS.includes(c));
    const attendantIndex = headers.findIndex(c => ATTENDANT_FIELDS.includes(c));
    const bairroIndex = headers.indexOf('Bairro');
    const importedRows = matrix.slice(1).filter(row => row.some(v => normalize(v))).map((row, idx) => {
      const item = {}; KEEP.forEach(c => item[c] = normalize(row[indexes[c]])); item._atendente = attendantIndex >= 0 ? normalize(row[attendantIndex]) : ''; item._bairro = bairroIndex >= 0 ? normalize(row[bairroIndex]) : ''; item._finalizadoEm = finalIndex >= 0 ? normalize(row[finalIndex]) : ''; item._index = idx; item._dateKey = dateKey(item['Recebido em']); return item;
    }).sort((a,b) => (b._dateKey || '').localeCompare(a._dateKey || '') || a._index-b._index);
    const operationalRows = limitToOperationalWindow(importedRows);
    state.rows = operationalRows.map((row, idx) => ({...row, _index: idx}));
    state.lastUpdated = new Date().toISOString(); state.fileName = file.name; await saveBase(); recomputeDuplicates();
    populateFilters(); setPeriod('month'); applyFilters();
    $('workspace').classList.remove('hidden'); activatePage('overviewPage'); $('uploadStatus').textContent = `${state.rows.length} registros operacionais na base local`;
    $('settingsFileName').textContent = file.name; $('settingsFileMeta').textContent = `${state.rows.length} registros · última atualização ${new Date(state.lastUpdated).toLocaleString('pt-BR')}`;
    notify(`Importação concluída. A base anterior foi substituída. Foram mantidos ${state.rows.length} registros, incluindo 5 protocolos anteriores ao aberto mais antigo.`, 'success');
  } catch(err) { $('uploadStatus').textContent = 'Falha na leitura'; notify(err.message || 'Não foi possível ler o arquivo.', 'error'); }
}
function isCompleted(r) { return normalize(r.Status) === 'Concluído' && !!parseDate(r._finalizadoEm); }
function isOpen(r) { return ACTIVE_STATUSES.includes(normalize(r.Status)); }
function limitToOperationalWindow(rows) {
  const oldestOpenIndex = rows.reduce((last, row, index) => isOpen(row) ? index : last, -1);
  if (oldestOpenIndex < 0) return rows;
  return rows.slice(0, oldestOpenIndex + 6);
}
function ageDays(r) { const d=parseDate(r['Recebido em']); if(!d) return null; return Math.floor((new Date(new Date().getFullYear(),new Date().getMonth(),new Date().getDate()) - new Date(d.getFullYear(),d.getMonth(),d.getDate())) / 86400000); }
function deadlineState(r) { if(!isOpen(r)) return ''; const age=ageDays(r); if(age === null) return ''; if(age > 30) return 'overdue'; if(age >= 25) return 'due-soon'; return ''; }
function deadlineDate(r) { const d=parseDate(r['Recebido em']); if(!d) return null; const due=new Date(d.getFullYear(),d.getMonth(),d.getDate()); due.setDate(due.getDate()+30); return due; }
function deadlineDateKey(r) { const d=deadlineDate(r); return d ? `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}` : ''; }
function deadlineDateLabel(r) { const d=deadlineDate(r); return d ? d.toLocaleDateString('pt-BR') : 'sem data'; }
function deadlineLabel(r) { const age=ageDays(r), state=deadlineState(r); if(state==='overdue') return `Atrasado há ${age} dias`; if(state==='due-soon') return `Vence em ${Math.max(0,30-age)} dias`; return 'Dentro do prazo'; }
function recomputeDuplicates() { const counts = {}; state.rows.forEach(r => { const p = r.Protocolo || '(sem protocolo)'; counts[p] = (counts[p]||0)+1; }); state.duplicates = new Set(Object.keys(counts).filter(p => p !== '(sem protocolo)' && counts[p] > 1)); }
function openBaseDB() { return new Promise((resolve,reject) => { const request = indexedDB.open(DB_NAME, DB_VERSION); request.onupgradeneeded = () => { if(!request.result.objectStoreNames.contains(DB_STORE)) request.result.createObjectStore(DB_STORE); }; request.onsuccess = () => resolve(request.result); request.onerror = () => reject(request.error); }); }
function saveBase() { return openBaseDB().then(db => new Promise((resolve,reject) => { const tx = db.transaction(DB_STORE,'readwrite'); tx.objectStore(DB_STORE).put({ rows: state.rows, lastUpdated: state.lastUpdated, fileName: state.fileName, mode: CURRENT_MODE }, DB_KEY); tx.oncomplete = () => { db.close(); resolve(); }; tx.onerror = () => { db.close(); reject(tx.error); }; })); }
function clearStoredBase() { return openBaseDB().then(db => new Promise((resolve,reject) => { const tx=db.transaction(DB_STORE,'readwrite'); tx.objectStore(DB_STORE).delete(DB_KEY); tx.oncomplete=()=>{db.close();resolve();}; tx.onerror=()=>{db.close();reject(tx.error);}; })); }
function loadBase() { return openBaseDB().then(db => new Promise((resolve,reject) => { const tx = db.transaction(DB_STORE,'readonly'); const request = tx.objectStore(DB_STORE).get(DB_KEY); request.onsuccess = () => { db.close(); resolve(request.result || null); }; request.onerror = () => { db.close(); reject(request.error); }; })); }
async function restoreBase() { try { const saved = await loadBase(); if(!saved?.rows?.length) return; const operationalRows = limitToOperationalWindow(saved.rows); state.rows = operationalRows.map((row, idx) => ({...row, _index: idx})); state.lastUpdated = saved.lastUpdated || ''; state.fileName = saved.fileName || ''; if(saved.mode !== CURRENT_MODE) { await clearStoredBase(); notify('A base anterior foi descartada porque acumulava versões históricas. Carregue a planilha atual do dia.', 'warning'); return; } if(operationalRows.length !== saved.rows.length) await saveBase(); recomputeDuplicates(); populateFilters(); setPeriod('month'); applyFilters(); $('workspace').classList.remove('hidden'); activatePage('overviewPage'); $('uploadStatus').textContent = `${state.rows.length} registros operacionais na base local`; $('settingsFileName').textContent = state.fileName || 'Base local da Ouvidoria'; $('settingsFileMeta').textContent = `${state.rows.length} registros · janela operacional com 5 protocolos anteriores ao aberto mais antigo · última atualização ${state.lastUpdated ? new Date(state.lastUpdated).toLocaleString('pt-BR') : 'não informada'}`;  } catch(e) { notify('Não foi possível restaurar a base local neste navegador.', 'warning'); } }
function unique(field) { return [...new Set(state.rows.map(r => r[field]).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'pt-BR')); }
function populateFilters() {
  [['statusFilter','Status'],['secretaryFilter','Secretaria'],['originFilter','Origem de cadastro']].forEach(([id,field]) => { const el=$(id), first=el.options[0].text; el.innerHTML=`<option value="">${first}</option>`; unique(field).forEach(v => el.insertAdjacentHTML('beforeend',`<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`)); });
  const dates = state.rows.map(r=>r._dateKey).filter(Boolean).sort(); if(dates.length){ $('dateFrom').min=dates[0]; $('dateFrom').max=dates[dates.length-1]; $('dateTo').min=dates[0]; $('dateTo').max=dates[dates.length-1]; }
}
['searchInput','dateFrom','dateTo','statusFilter','secretaryFilter','originFilter','deadlineFilter'].forEach(id => $(id).addEventListener('input', applyFilters));
$('periodFilter').addEventListener('change',()=>{ setPeriod($('periodFilter').value); applyFilters(); });
function activatePage(pageId) { document.querySelectorAll('.nav-item').forEach(x=>x.classList.toggle('active',x.dataset.page===pageId)); document.querySelectorAll('.page-section').forEach(page=>page.classList.toggle('hidden',page.id!==pageId)); const titles={overviewPage:'Painel diário',recordsPage:'Manifestações do arquivo',attentionPage:'Protocolos que exigem atenção',settingsPage:'Configurações'}; $('pageTitle').textContent=titles[pageId]||'Painel diário'; $('periodToolbar').classList.toggle('hidden',pageId==='settingsPage'||pageId==='attentionPage'); }
document.querySelectorAll('.nav-item').forEach(btn=>btn.addEventListener('click',()=>activatePage(btn.dataset.page)));
document.addEventListener('click',e=>{const btn=e.target.closest('.metric-filter');if(btn){metricScope=btn.dataset.metric;renderAll();}});
activatePage('settingsPage');
['dateFrom','dateTo'].forEach(id => $(id).addEventListener('change',()=>{ $('periodFilter').value='custom'; applyFilters(); }));
$('clearFilters').addEventListener('click',()=>{ ['searchInput'].forEach(id=>$(id).value=''); $('periodFilter').value='month'; setPeriod('month'); ['statusFilter','secretaryFilter','originFilter','deadlineFilter'].forEach(id=>$(id).value=''); applyFilters(); });
$('clearBase').addEventListener('click', async ()=>{ await clearStoredBase(); state.rows=[]; state.filtered=[]; state.duplicates=new Set(); $('workspace').classList.remove('hidden'); $('uploadStatus').textContent='Aguardando arquivo'; $('settingsFileName').textContent='Nenhuma planilha carregada'; $('settingsFileMeta').textContent='Carregue a planilha diária para ver a quantidade de registros e a última atualização.'; activatePage('settingsPage'); notify('Base local limpa. Carregue a planilha atual do dia.', 'success'); });
function inputDate(d) { return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function setPeriod(period) { const today=new Date(); today.setHours(0,0,0,0); if(period==='all'||period==='custom'){ if(period==='all'){ $('dateFrom').value=''; $('dateTo').value=''; } return; } let from=new Date(today), to=new Date(today); if(period==='month'){ from=new Date(today.getFullYear(),today.getMonth(),1); } else if(period==='3months'){ from=new Date(today); from.setMonth(from.getMonth()-3); } else if(period==='6months'){ from=new Date(today); from.setMonth(from.getMonth()-6); } else if(period==='year'){ from=new Date(today.getFullYear(),0,1); } $('dateFrom').value=inputDate(from); $('dateTo').value=inputDate(to); }
setPeriod('month');
function applyFilters() {
  const q = $('searchInput').value.toLowerCase().trim(), from=$('dateFrom').value, to=$('dateTo').value, status=$('statusFilter').value, secretary=$('secretaryFilter').value, origin=$('originFilter').value, deadline=$('deadlineFilter').value;
  state.periodRows = state.rows.filter(r => (!from||r._dateKey>=from) && (!to||r._dateKey<=to));
  state.filtered = state.periodRows.filter(r => { const hay=KEEP.map(k=>r[k]).join(' ').toLowerCase(); return (!q||hay.includes(q)) && (!status||r.Status===status) && (!secretary||r.Secretaria===secretary) && (!origin||r['Origem de cadastro']===origin) && (!deadline||deadlineState(r)===deadline); });
  renderAll();
}
function countBy(field, rows=state.rows) { return rows.reduce((acc,r)=>{const k=r[field]||'Não informado';acc[k]=(acc[k]||0)+1;return acc;},{}); }
function renderAll() {
  const rows=state.rows, periodRows=state.periodRows, filtered=state.filtered, open=periodRows.filter(isOpen), completed=periodRows.filter(isCompleted), dueSoon=rows.filter(r=>deadlineState(r)==='due-soon'), overdue=rows.filter(r=>deadlineState(r)==='overdue');
  const viewRows=metricScope==='open'?filtered.filter(isOpen):(metricScope==='completed'?filtered.filter(isCompleted):filtered);
  $('metrics').innerHTML = [['Protocolados',periodRows.length,'blue','all'],['Em aberto',open.length,'orange','open'],['Concluído',completed.length,'green','completed']].map(x=>`<article class="metric ${x[2]} ${metricScope===x[3]?'selected':''}"><div class="metric-label">${x[0]}</div><div class="metric-value">${x[1]}</div><button class="metric-filter ${metricScope===x[3]?'active':''}" data-metric="${x[3]}" type="button">Filtrar</button></article>`).join('');
  $('dueSoonCount').textContent=dueSoon.length; $('overdueCount').textContent=overdue.length; renderDeadlineLists(dueSoon, overdue);
  renderEvolution(periodRows, metricScope); renderRanking('secretaryRanking','Secretaria',viewRows); renderRanking('neighborhoodRanking','_bairro',viewRows); renderRanking('subjectRanking','Assunto',viewRows); renderTable(viewRows); $('resultCount').textContent=`${viewRows.length} de ${periodRows.length} registros`;
}
function renderDeadlineLists(dueSoon, overdue) {
  const dueGroups = {}; dueSoon.sort((a,b)=>deadlineDateKey(a).localeCompare(deadlineDateKey(b))).forEach(r => { const key=deadlineDateKey(r)||'sem-data'; (dueGroups[key] ||= []).push(r); });
  $('dueSoonList').innerHTML = Object.keys(dueGroups).length ? Object.entries(dueGroups).map(([key,items]) => `<div class="deadline-group"><div class="deadline-group-title"><span>${key==='sem-data'?'Sem data':escapeHtml(deadlineDateLabel(items[0]))}</span><span>${items.length}</span></div>${items.map(r=>`<div class="deadline-item"><strong>${escapeHtml(r.Protocolo||'Sem protocolo')}</strong><small>${escapeHtml(r.Secretaria||'Sem secretaria')}</small></div>`).join('')}</div>`).join('') : '<div class="deadline-empty">Nenhum protocolo vence nos próximos 5 dias.</div>';
  const secGroups = {}; overdue.forEach(r => { const key=r.Secretaria||'Sem secretaria'; (secGroups[key] ||= []).push(r); });
  $('overdueList').innerHTML = Object.keys(secGroups).length ? Object.entries(secGroups).sort((a,b)=>b[1].length-a[1].length).map(([sec,items]) => `<div class="deadline-group"><div class="deadline-group-title"><span>${escapeHtml(sec)}</span><span>${items.length}</span></div>${items.sort((a,b)=>(ageDays(b)||0)-(ageDays(a)||0)).map(r=>`<div class="deadline-item"><strong>${escapeHtml(r.Protocolo||'Sem protocolo')}</strong><small>${escapeHtml(deadlineLabel(r))}</small></div>`).join('')}</div>`).join('') : '<div class="deadline-empty">Nenhum protocolo está atrasado há mais de 30 dias.</div>';
}
function renderEvolution(rows, scope='all') {
  const dates=rows.map(r=>r._dateKey).filter(Boolean).sort(); if(!dates.length){$('evolutionChart').innerHTML='<div class="ranking-empty">Nenhum protocolo no período selecionado.</div>';return;}
  const start=dates[0], end=dates[dates.length-1], span=Math.round((new Date(end)-new Date(start))/86400000), monthly=span>45, buckets={};
  const field=scope==='open'?'open':scope==='completed'?'completed':'all'; const labelsByScope={all:['Protocolados','#2d72c7','legend-blue'],open:['Em aberto','#df8b27','legend-orange'],completed:['Concluído','#2b9b67','legend-green']}; const [seriesLabel,color,legendClass]=labelsByScope[scope];
  rows.forEach(r=>{let key=r._dateKey;if(monthly)key=key.slice(0,7);const bucket=(buckets[key] ||= {all:0,open:0,completed:0});bucket.all++;if(isOpen(r))bucket.open++;if(isCompleted(r))bucket.completed++;});
  const keys=Object.keys(buckets).sort(), W=760,H=230,L=42,R=16,T=18,B=38,max=Math.max(1,...keys.map(k=>buckets[k][field])), x=i=>L+(i/(Math.max(1,keys.length-1)))*(W-L-R), y=v=>(H-B)-(v/max)*(H-B-T);
  const line=()=>{const pts=keys.map((k,i)=>`${x(i).toFixed(1)},${y(buckets[k][field]).toFixed(1)}`).join(' ');return `<polyline points="${pts}" fill="none" stroke="${color}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>`+keys.map((k,i)=>`<circle cx="${x(i)}" cy="${y(buckets[k][field])}" r="3.5" fill="${color}"/>`).join('');};
  const labels=keys.map((k,i)=>i%Math.max(1,Math.ceil(keys.length/8))===0?`<text x="${x(i)}" y="${H-12}" text-anchor="middle">${k.slice(monthly?0:5).replace('-', '/')}</text>`:'').join('');
  $('evolutionChart').innerHTML=`<div class="chart-legend"><span><i class="${legendClass}"></i>${seriesLabel}</span></div><svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Evolução de ${seriesLabel}"><line x1="${L}" y1="${H-B}" x2="${W-R}" y2="${H-B}" class="chart-axis"/><line x1="${L}" y1="${T}" x2="${L}" y2="${H-B}" class="chart-axis"/><text x="${L-8}" y="${T+4}" text-anchor="end">${max}</text>${line()}${labels}</svg>`;
}
function renderRanking(id, field, rows) { const counts={}; rows.forEach(r=>{ const value=normalize(r[field])||'Não informado'; counts[value]=(counts[value]||0)+1; }); const entries=Object.entries(counts).sort((a,b)=>b[1]-a[1]||a[0].localeCompare(b[0],'pt-BR')); $(id).innerHTML=entries.length?entries.map(([label,n],i)=>`<div class="ranking-row"><div><strong title="${escapeHtml(label)}">${i+1}. ${escapeHtml(label)}</strong><small>${((n/rows.length)*100).toFixed(1).replace('.',',')}% do período</small></div><span class="ranking-count">${n}</span></div>`).join(''):'<div class="ranking-empty">Nenhum registro no período.</div>'; }
function renderBars(id, counts, total) { const top=Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,5); $(id).innerHTML=top.length?top.map(([label,n])=>`<div class="bar-row"><span class="bar-label" title="${escapeHtml(label)}">${escapeHtml(label)}</span><span class="bar-track"><span class="bar-fill" style="width:${Math.max(3,n/total*100)}%"></span></span><span class="bar-value">${n}</span></div>`).join(''):'<span class="muted">Sem dados suficientes.</span>'; }
function renderTable(rows) { const body=$('recordsBody'); $('emptyTable').classList.toggle('hidden',rows.length>0); body.innerHTML=rows.slice(0,300).map(r=>`<tr><td>${escapeHtml(r.Protocolo||'—')}</td><td>${displayDate(r['Recebido em'])}</td><td><span class="status-chip ${safeClass(r.Status)}">${escapeHtml(r.Status||'Sem status')}</span></td><td><span class="cell-truncate" title="${escapeHtml(r.Secretaria)}">${escapeHtml(r.Secretaria||'—')}</span></td><td><span class="cell-truncate" title="${escapeHtml(r._atendente)}">${escapeHtml(r._atendente||'—')}</span></td><td><span class="cell-truncate" title="${escapeHtml(r.Assunto)}">${escapeHtml(r.Assunto||'—')}</span></td><td>${escapeHtml(r['Origem de cadastro']||'—')}</td><td><span class="deadline-chip ${deadlineState(r)}">${escapeHtml(deadlineLabel(r))}</span></td><td class="${r.Resposta?'response-yes':'response-no'}">${r.Resposta?'Sim':'Não'}</td><td><button class="view-button" data-index="${r._index}">Ver</button></td></tr>`).join(''); body.querySelectorAll('.view-button').forEach(btn=>btn.addEventListener('click',()=>openDetail(state.rows.find(r=>r._index===Number(btn.dataset.index))))); }
function openDetail(r) { if(!r)return; $('detailContent').innerHTML=`<div class="detail"><h2>${escapeHtml(r.Protocolo||'Sem protocolo')}</h2><p class="detail-meta">Recebido em ${displayDate(r['Recebido em'])} · <span class="status-chip ${safeClass(r.Status)}">${escapeHtml(r.Status||'Sem status')}</span></p><div class="detail-grid"><div class="detail-field"><label>Secretaria</label><strong>${escapeHtml(r.Secretaria||'Não informado')}</strong></div><div class="detail-field"><label>Atendente</label><strong>${escapeHtml(r._atendente||'Não informado')}</strong></div><div class="detail-field"><label>Assunto</label><strong>${escapeHtml(r.Assunto||'Não informado')}</strong></div><div class="detail-field"><label>Origem de cadastro</label><strong>${escapeHtml(r['Origem de cadastro']||'Não informado')}</strong></div></div><div class="detail-field"><label>Manifestação</label><div class="detail-block ${r.Manifestação?'':'empty'}">${escapeHtml(r.Manifestação||'Sem manifestação informada.')}</div></div><div class="detail-field" style="margin-top:18px"><label>Resposta</label><div class="detail-block ${r.Resposta?'':'empty'}">${escapeHtml(r.Resposta||'Ainda não há resposta registrada.')}</div></div></div>`; $('detailDialog').showModal(); }
$('closeDialog').addEventListener('click',()=> $('detailDialog').close()); $('detailDialog').addEventListener('click',e=>{if(e.target.id==='detailDialog')$('detailDialog').close()});
function exportRows(rows) { return rows.map(r=>({...Object.fromEntries(KEEP.map(k=>[k,r[k]])), Atendente:r._atendente||''})); }
function download(blob,name) { const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),800); }
$('exportCsv').addEventListener('click',()=>{const csv=XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(exportRows(state.filtered)),{FS:';'});download(new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}),`ouvidoria-tratada-${new Date().toISOString().slice(0,10)}.csv`);});
$('exportXlsx').addEventListener('click',()=>{const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(exportRows(state.filtered)),'Dados tratados');XLSX.writeFile(wb,`ouvidoria-tratada-${new Date().toISOString().slice(0,10)}.xlsx`);});
$('printReport').addEventListener('click',()=>window.print());
window.addEventListener('DOMContentLoaded', restoreBase);
