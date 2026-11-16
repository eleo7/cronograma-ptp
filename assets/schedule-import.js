/* ExcelJS 4.4.0 is distributed locally under MIT; see vendor/exceljs-LICENSE. */
const importAssetBase = new URL('.', document.currentScript.src);
let importSheets = [], importRows = [], importBusy = false;
const importNormalize = value => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase().replace(/[^A-Z0-9]/g, '');
const importText = value => {
  if (value == null) return '';
  if (typeof value === 'object' && !(value instanceof Date)) {
    if ('formula' in value || 'sharedFormula' in value) return importText(value.result);
    if (value.richText) return value.richText.map(t => t.text).join('');
    return value.text || '';
  }
  return value;
};
function openScheduleImport() {
  if (importBusy) return;
  importSheets = []; importRows = [];
  document.getElementById('scheduleImportFile').value = '';
  document.getElementById('scheduleImportNewType').value = '';
  document.getElementById('scheduleImportPreview').innerHTML = '';
  document.getElementById('scheduleImportMessage').textContent = 'Selecione um arquivo para analisar. Nenhum dado será alterado antes da confirmação.';
  document.getElementById('scheduleImportConfirm').disabled = true;
  document.getElementById('scheduleImportOverlay').classList.add('active');
}
function closeScheduleImport() {
  if (!importBusy) document.getElementById('scheduleImportOverlay').classList.remove('active');
}
function parseScheduleCSV(text) {
  const first = text.replace(/^\uFEFF/, '').split(/\r?\n/)[0];
  const delimiter = [';', ',', '\t'].sort((a,b)=>first.split(b).length-first.split(a).length)[0];
  const rows = []; let row=[], field='', quoted=false;
  text=text.replace(/^\uFEFF/,'');
  for (let i=0;i<text.length;i++) {
    const c=text[i];
    if (c==='"') { if (quoted && text[i+1]==='"') { field+='"';i++; } else quoted=!quoted; }
    else if (c===delimiter && !quoted) { row.push(field);field=''; }
    else if ((c==='\n'||c==='\r') && !quoted) { if(c==='\r'&&text[i+1]==='\n')i++;row.push(field);rows.push(row);row=[];field=''; }
    else field+=c;
  }
  if(quoted)throw Error('CSV com aspas não fechadas. Exporte novamente a planilha.');
  if(field||row.length){row.push(field);rows.push(row);}
  return rows;
}
async function loadScheduleExcel() {
  if(window.ExcelJS)return;
  await new Promise((resolve,reject)=>{const script=document.createElement('script');script.src=new URL('vendor/exceljs-4.4.0.min.js',importAssetBase).href;script.onload=resolve;script.onerror=()=>{script.remove();reject(Error('Não foi possível carregar o leitor Excel. Tente novamente.'));};document.head.append(script);});
}
async function readScheduleFile(input) {
  if(importBusy)return;
  const file=input.files[0];if(!file)return;
  importBusy=true;importSheets=[];importRows=[];
  const message=document.getElementById('scheduleImportMessage');
  document.getElementById('scheduleImportConfirm').disabled=true;
  document.getElementById('scheduleImportPreview').innerHTML='';
  message.textContent='Lendo planilha…';
  try {
    if(file.size>5*1024*1024)throw Error('O limite é 5 MB. Divida a planilha em arquivos menores.');
    if(/\.csv$/i.test(file.name)) {
      const bytes=await file.arrayBuffer();let text=new TextDecoder('utf-8').decode(bytes);
      if(text.includes('\uFFFD'))text=new TextDecoder('windows-1252').decode(bytes);
      importSheets=[{name:file.name,rows:parseScheduleCSV(text),date1904:false}];
    } else if(/\.xlsx$/i.test(file.name)) {
      await loadScheduleExcel();const workbook=new ExcelJS.Workbook();await workbook.xlsx.load(await file.arrayBuffer());
      importSheets=workbook.worksheets.slice(0,1).map(sheet=>{
        const rows=[];
        sheet.eachRow(row=>{
          const values=[];
          row.eachCell((cell,column)=>{
            const value=importText(cell.value);
            if(value==null||String(value).trim()==='')return;
            if(column>200)throw Error('Há dados além da coluna 200. Use uma planilha com até 200 colunas de dados.');
            values[column-1]=value;
          });
          if(!values.length)return;
          values.sourceLine=row.number;
          rows.push(values);
          if(rows.length>20000)throw Error('Use no máximo 20.000 linhas com dados. Células vazias ou apenas formatadas não são contadas.');
        });
        return {name:sheet.name,rows,date1904:!!workbook.properties.date1904};
      });
    } else throw Error('Use Excel .xlsx ou CSV. Para arquivos .xls, salve como .xlsx no Excel.');
    if(!importSheets.length)throw Error('A planilha não contém abas.');
    if(importSheets.some(s=>s.rows.length>20000))throw Error('Use no máximo 20.000 linhas por aba.');
  } catch(error) { importSheets=[];message.textContent=error.message || 'Não foi possível ler o arquivo.'; }
  finally { importBusy=false; }
  if(importSheets.length)previewScheduleSheet();
}
function scheduleDate(value,date1904=false) {
  value=importText(value);
  let y,m,d;
  if(value instanceof Date) { if(isNaN(value))return null;y=value.getUTCFullYear();m=value.getUTCMonth()+1;d=value.getUTCDate(); }
  else if(typeof value==='number' && Number.isFinite(value)) {
    if(value<1||(!date1904&&Math.floor(value)===60))return null;
    const date=new Date(Date.UTC(date1904?1904:1899,date1904?0:11,date1904?1:30)+Math.floor(value)*86400000);
    y=date.getUTCFullYear();m=date.getUTCMonth()+1;d=date.getUTCDate();
  } else {
    const text=String(value).trim();let match=text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?$/);
    if(match){d=+match[1];m=+match[2];y=+match[3];}
    else {match=text.match(/^(\d{4})-(\d{2})-(\d{2})(?:T\d{2}:\d{2}.*)?$/);if(!match)return null;y=+match[1];m=+match[2];d=+match[3];}
  }
  const checked=new Date(Date.UTC(y,m-1,d));
  if(y<1900||y>9999||checked.getUTCFullYear()!==y||checked.getUTCMonth()!==m-1||checked.getUTCDate()!==d)return null;
  return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}
function scheduleRegion(value) {
  const norm=importNormalize(value);if(!norm)return null;
  const exact=REGIONS.filter(r=>[r.id,r.name].some(v=>importNormalize(v)===norm));
  if(exact.length===1)return exact[0];
  const alias=v=>{let key=importNormalize(v).replace(/^\d+/,'').replace(/^SE/,'');return key==='GOI'?'GOYA':key;};
  const matches=REGIONS.filter(r=>[r.id,r.name].some(v=>alias(v)===alias(value)));
  return matches.length===1?matches[0]:null;
}
function buildSchedulePlan(rows,date1904=false) {
  const aliases={team:['EQUIPE','EQUIPES'],region:['SE','REGIAO'],date:['DATADAPROGRAMACAO','DATAPROGRAMACAO','DATADOPROGRAMACAO'],type:['TIPODADEMANDA','TIPODEEQUIPE']};
  let header=-1,columns={};
  for(let i=0;i<Math.min(30,rows.length);i++) {
    const names=rows[i].map(importNormalize);const found={};
    for(const [key,options] of Object.entries(aliases))found[key]=names.findIndex(n=>options.includes(n));
    if(found.team>=0&&found.region>=0&&found.date>=0){header=i;columns=found;break;}
  }
  if(header<0)throw Error('Cabeçalhos não encontrados nas primeiras 30 linhas. Use EQUIPE, SE (ou REGIÃO) e DATA DA PROGRAMAÇÃO.');
  const existing=new Set(state.events.map(e=>JSON.stringify([importNormalize(state.teams.find(t=>t.id===e.teamId)?.name||e.teamId),e.region,e.date])));
  const report=[], additions=[],newTeams=new Map();
  const selectedType=document.getElementById('scheduleImportNewType').value;
  for(let i=header+1;i<rows.length;i++) {
    const row=rows[i];if(!row.some(v=>String(v??'').trim()))continue;
    const name=String(importText(row[columns.team])??'').trim();const rawRegion=String(importText(row[columns.region])??'').trim();const rawDate=importText(row[columns.date]);
    const entry={line:row.sourceLine||i+1,name,region:rawRegion,date:'',reason:''};report.push(entry);
    if(!name||!rawRegion||rawDate==null||rawDate===''){entry.reason='Ignorada: campo obrigatório vazio';continue;}
    const date=scheduleDate(rawDate,date1904);if(!date){entry.reason='Ignorada: data inválida';continue;}entry.date=date;
    const region=scheduleRegion(rawRegion);if(!region){entry.reason='Ignorada: região não reconhecida';continue;}entry.region=region.name;
    const key=JSON.stringify([importNormalize(name),region.id,date]);
    if(existing.has(key)){entry.reason='Ignorada: programação já existente ou repetida';continue;}
    const matches=state.teams.filter(t=>importNormalize(t.name)===importNormalize(name));
    if(matches.length>1){entry.reason='Ignorada: nome de equipe ambíguo';continue;}
    const team=matches[0];
    let type=team?.type||newTeams.get(importNormalize(name))?.type;
    if(!team&&!type) {
      const demand=columns.type>=0?importNormalize(importText(row[columns.type])):'';
      type=demand.startsWith('PODA')?'PODA':demand.startsWith('DEGRADACAO')?'DEGRADAÇÃO':selectedType;
      if(!type){entry.reason='Ignorada: equipe nova sem tipo; selecione o tipo abaixo';continue;}
      newTeams.set(importNormalize(name),{name,type});
    }
    existing.add(key);additions.push({name,teamId:team?.id,region:region.id,date});entry.reason='Adicionar';
  }
  return {report,additions,newTeams};
}
function previewScheduleSheet() {
  if(importBusy||!importSheets.length)return;
  document.getElementById('scheduleImportConfirm').disabled=true;
  document.getElementById('scheduleImportPreview').innerHTML='';
  const sheet=importSheets[0];
  try {
    const plan=buildSchedulePlan(sheet.rows,sheet.date1904);importRows=sheet.rows;
    document.getElementById('scheduleImportMessage').textContent=`${plan.additions.length} programação(ões) para adicionar · ${plan.newTeams.size} equipe(s) nova(s) · ${plan.report.length-plan.additions.length} linha(s) ignorada(s). Nenhum registro existente será alterado.`;
    document.getElementById('scheduleImportPreview').innerHTML=`<div class="week-table-wrap"><table class="week-table"><thead><tr><th>Linha</th><th>Equipe</th><th>Região</th><th>Data</th><th>Resultado</th></tr></thead><tbody>${plan.report.slice(0,200).map(r=>`<tr><td>${r.line}</td><td>${escapeHtml(r.name)}</td><td>${escapeHtml(r.region)}</td><td>${r.date?formatDateBR(r.date):'—'}</td><td>${escapeHtml(r.reason)}</td></tr>`).join('')}</tbody></table></div>${plan.report.length>200?'<p>Prévia limitada a 200 linhas; todas as linhas foram analisadas.</p>':''}`;
    document.getElementById('scheduleImportConfirm').disabled=!plan.additions.length;
  }catch(error){importRows=[];document.getElementById('scheduleImportMessage').textContent=error.message;}
}
function confirmScheduleImport() {
  if(importBusy||!importRows.length)return;
  const sheet=importSheets[0];
  // Recheck current data at confirmation, including records added after preview.
  const plan=buildSchedulePlan(importRows,sheet.date1904);
  if(!plan.additions.length){previewScheduleSheet();return;}
  const previousTeams=state.teams,previousEvents=state.events;
  const previousStoredTeams=localStorage.getItem('ptp_teams'),previousStoredEvents=localStorage.getItem('ptp_events');
  const teams=[...state.teams];
  for(const item of plan.newTeams.values())teams.push({id:genId(),name:item.name,type:item.type,color:TEAM_COLORS[teams.length%TEAM_COLORS.length]});
  const events=plan.additions.map(item=>({id:genId(),teamId:item.teamId||teams.find(t=>importNormalize(t.name)===importNormalize(item.name)).id,region:item.region,date:item.date,category:'normal',description:'',source:'spreadsheet'}));
  try {
    state.teams=teams;state.events=[...state.events,...events];saveState();
  }catch(error){
    state.teams=previousTeams;state.events=previousEvents;
    try {for(const [key,value] of [['ptp_teams',previousStoredTeams],['ptp_events',previousStoredEvents]]){if(value===null)localStorage.removeItem(key);else localStorage.setItem(key,value);}}catch(_){}
    document.getElementById('scheduleImportMessage').textContent='Não foi possível salvar. Verifique o espaço disponível no navegador e tente novamente.';return;
  }
  state.currentDate=new Date(events[0].date+'T12:00:00');
  importRows=[];closeScheduleImport();render();switchTab('calendario');
  if(document.getElementById('dayExpandOverlay').classList.contains('active'))renderDayExpand();
  showToast(`${events.length} programação(ões) adicionada(s). Registros anteriores preservados.`);
}
