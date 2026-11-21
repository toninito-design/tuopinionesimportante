const $ = id => document.getElementById(id);
const trim = value => String(value ?? '').replace(/\s+/g,' ').trim();
const norm = value => trim(value).normalize('NFKD').replace(/\p{M}/gu,'').toLowerCase();
const uniq = values => [...new Set(values)];
const levels = ['5º Primaria','6º Primaria','1º ESO','2º ESO','3º ESO','4º ESO','1º BAC','2º BAC'];
const sorted = values => uniq(values).filter(Boolean).sort((a,b) => {
  const levelOrder = levels.indexOf(a) - levels.indexOf(b);
  return levelOrder || a.localeCompare(b,'es',{numeric:true});
});
const keyFor = field => JSON.stringify([norm(field.q),subjectKey(field.m)]);
const sameQuestion = (a,b) => norm(a) === norm(b);
const ALL = '';
const fmt = new Intl.NumberFormat('es-ES',{maximumFractionDigits:1});
const excluded = new Set(['no curso la asignatura','no la curso','no lo he utilizado','no aplica']);
const ordinal = ['Nada','Nunca','Poco','Bastante','Mucho'];
const colors = ['#4867bc','#a64d9d','#21826b','#ad5d1b','#5b6990','#a64152'];
const responseColors = {Nada:'#c93238',Nunca:'#c93238',Poco:'#e7b800',Bastante:'#2870ce',Mucho:'#248347'};
const answerColor = (answer,index) => responseColors[answer] || colors[index%colors.length];
const subjectKey = value => {
  const subject=norm(value).replace(/[.]/g,'');
  if (['ingles','english'].includes(subject)) return 'english';
  if (['spoken','sopken','spoken english','ingles oral','cambridge','cambridge english','ingles cambridge'].includes(subject)) return 'spoken english';
  return subject;
};
const levelKey = value => norm(value).replace(/(\d)\s*[.ºª°o]+/g,'$1').replace(/\bbach(?:illerato)?\b/g,'bac').replace(/\bprim\b/g,'primaria');
const groupKey = value => norm(value).replace(/^grupo\s+/,'');
const rosterKey = (matter,level,group) => JSON.stringify([subjectKey(matter),levelKey(level),groupKey(group)]);
const questionFields = data => data.fields.filter(field=>!field.identity);
let catalog = [], activeTab = 'import', draft, renderCount = 0, report;
const cache = new Map();

function node(tag,value,cls) {
  const element = document.createElement(tag);
  if (value !== undefined) element.textContent = value;
  if (cls) element.className = cls;
  return element;
}
function status(message = '', error = false) {
  $('status').textContent = message;
  $('status').classList.toggle('error',error);
}
function options(id,items,placeholder) {
  const select = $(id);
  const previous = select.value;
  select.replaceChildren(...(placeholder ? [new Option(placeholder,ALL)] : []), ...items.map(item =>
    Array.isArray(item) ? new Option(item[1],item[0]) : new Option(item)));
  if ([...select.options].some(item=>item.value===previous)) select.value = previous;
  select.disabled = select.options.length < 2;
}
async function api(id,method='GET',body) {
  const endpoint = id && method === 'GET' ? '/api/editions?id='+encodeURIComponent(id) : '/api/editions';
  const response = await fetch(endpoint,{method,headers:body?{'Content-Type':'application/json'}:undefined,body:body?JSON.stringify(body):undefined});
  const data = await response.json();
  if (!response.ok) throw Error(data.error);
  return data;
}
async function edition(id) {
  if (!cache.has(id)) {
    const data=await api(id);
    for (const field of data.fields) {
      if (!field.m && /participacion oral/.test(norm(field.q)) && /spoken english/.test(norm(field.q))) {
        field.m='Spoken English';
        field.teacher='spoken';
      }
    }
    cache.set(id,data);
  }
  return cache.get(id);
}
function editionLabel(item) {return item.year+(item.label?' · '+item.label:'');}
async function refresh() {
  catalog = await api();
  options('edition',catalog.map(item=>[item.id,editionLabel(item)]));
  $('history').replaceChildren(...catalog.map(item=>node('li',editionLabel(item)+' · '+item.count+' respuestas')));
  if (!catalog.length) $('history').append(node('li','Todavía no hay ediciones guardadas.'));
}
function tab(name) {
  activeTab = name;
  renderCount++;
  for (const button of document.querySelectorAll('[data-tab]')) {
    button.classList.toggle('active',button.dataset.tab===name);
    button.setAttribute('aria-pressed',String(button.dataset.tab===name));
  }
  $('import').hidden = name !== 'import';
  $('analysis').hidden = name === 'import';
  $('teacher-info').hidden = name !== 'teachers';
  $('view-label').hidden = name !== 'results';
  if (name !== 'import') update('edition');
}
function questionLabel(header) {
  return trim(header.replace(/\s+\/\s+.*$/,'')).slice(0,500);
}
function teacherType(matter) {
  const m = subjectKey(matter);
  return m==='english'?'english':m==='spoken english'?'spoken':m==='cambridge'?'cambridge':'';
}
function readBook(bytes) {
  const book = XLSX.read(bytes,{type:'array'});
  const match = book.SheetNames.map(name => XLSX.utils.sheet_to_json(book.Sheets[name],{header:1,defval:'',raw:false}))
    .find(rows => rows[0]?.some(h=>/^selecciona el nivel/i.test(trim(h))) && rows[0]?.some(h=>/^selecciona el grupo/i.test(trim(h))));
  if (!match) throw Error('No encuentro una hoja de Forms con columnas de nivel y grupo.');
  const [header,...body] = match;
  const headers = header.map(trim);
  const find = pattern => headers.findIndex(h=>pattern.test(norm(h)));
  const level = find(/^selecciona el nivel/), group = find(/^selecciona el grupo/);
  const english = headers.findIndex(h=>/(profesor|docente).*imparte.*\b(english|ingles)\b/.test(norm(h)) && !/spoken|cambridge/.test(norm(h)));
  const spoken = find(/(profesor|docente).*imparte.*(spoken english|spoken|ingles oral)/);
  const cambridge = find(/^(?:[¿ ]*)(?:que |cual es |selecciona )?(?:profesor|docente).*imparte.*cambridge/);
  const subjects = new Map(Object.entries(schema.mapping).map(([name,value])=>[trim(name),value]));
  const exact = schema.headers.length === headers.length && schema.headers.every((h,i)=>trim(h)===headers[i]);
  const metadata = /^(id|nombre|name|email|correo electronico|hora de inicio|hora de finalizacion|hora de la ultima modificacion|start time|completion time|last modified time)$/;
  const fields = headers.flatMap((h,i)=>{
    if (!h || [level,group,english,spoken].includes(i) || metadata.test(norm(h))) return [];
    if (i===cambridge) return [{i,h,q:'Docente indicado en Cambridge',m:'',identity:true,ambiguous:false}];
    const mapped = (exact ? subjects.get(h) : undefined) ||
      (/participacion oral/.test(norm(h)) && /spoken english/.test(norm(h)) ? [questionLabel(h),'Spoken English'] : undefined);
    const ambiguous = !mapped && (subjects.has(h) || (!/[¿?]|\s+\/\s+/.test(h) && h.length<32));
    const [q,m] = mapped || [ambiguous?'':questionLabel(h),''];
    return [{i,h,q,m,ambiguous}];
  });
  const rows = body.filter(row=>trim(row[level]) && trim(row[group]));
  if (!rows.length) throw Error('No hay filas con nivel y grupo.');
  return {rows,fields,level,group,english,spoken,exact};
}
function review() {
  const totals = new Map();
  for (const row of draft.rows) {
    const label = trim(row[draft.level])+' '+trim(row[draft.group]);
    totals.set(label,(totals.get(label)||0)+1);
  }
  $('classification').textContent = draft.rows.length+' respuestas clasificadas. '+sorted([...totals.keys()]).map(label=>label+' ('+totals.get(label)+')').join(' · ');
  $('mapping').replaceChildren(...draft.fields.map(field=>{
    const outer = node('div');
    const title = node('strong',field.h);
    const row = node('div',undefined,'row');
    for (const [prop,label] of [['q','Pregunta'],['m','Materia']]) {
      const wrapper = node('label',label,'wide');
      const input = node('input');
      input.value = field[prop];
      input.maxLength = 500;
      input.placeholder = prop==='q'?'Escribe el enunciado o deja vacío para excluir':'Opcional';
      input.oninput = ()=>{field[prop]=trim(input.value);field.reviewed=true;};
      wrapper.append(input);
      row.append(wrapper);
    }
    const ignore = node('label','Omitir columna');
    const check = node('input');
    check.type = 'checkbox';
    check.onchange = ()=>{field.ignored=check.checked};
    ignore.append(check);
    row.append(ignore);
    outer.append(title,row);
    return outer;
  }));
  $('mapping-details').open = draft.fields.some(field=>field.ambiguous);
  $('review').hidden = false;
}
async function readFile() {
  const file = $('file').files[0];
  if (!file) throw Error('Selecciona el Excel exportado desde Forms.');
  if (file.size > 15_000_000) throw Error('El Excel supera los 15 MB.');
  const bytes = await file.arrayBuffer();
  const hash = await crypto.subtle.digest('SHA-256',bytes);
  draft = {...readBook(bytes),id:Array.from(new Uint8Array(hash),n=>n.toString(16).padStart(2,'0')).join('')};
  review();
}
async function saveFile() {
  if (!draft) throw Error('Revisa el Excel primero.');
  const year = trim($('year').value).replace(/-/g,'–');
  if (!/^\d{4}–\d{2}$/.test(year) || (Number(year.slice(0,4))+1)%100 !== Number(year.slice(-2))) {
    throw Error('Introduce un año académico válido, por ejemplo 2025–26.');
  }
  if (draft.fields.some(field=>field.ambiguous && !field.ignored && !field.q)) {
    $('mapping-details').open = true;
    throw Error('Identifica las columnas pendientes o marca «Omitir columna».');
  }
  const sourced = draft.fields.filter(field=>field.q && !field.ignored);
  const fields = sourced.map(field=>({q:field.q,m:field.m,teacher:teacherType(field.m),...(field.identity?{identity:true}:{})}));
  if (!fields.length || uniq(fields.map(keyFor)).length !== fields.length) {
    throw Error('Cada columna debe tener una combinación única de pregunta y materia.');
  }
  const rows = draft.rows.map(row=>[draft.level,draft.group,draft.english,draft.spoken,...sourced.map(field=>field.i)]
    .map((index,i)=>trim(row[index]).slice(0,i<4?160:500)));
  await api(null,'POST',{id:draft.id,data:{year,label:trim($('label').value),fields,rows,roster:[]}});
  const id = draft.id;
  draft = undefined;
  $('file').value = '';
  $('review').hidden = true;
  await refresh();
  $('edition').value = id;
  status('Edición guardada en el histórico.');
  tab('results');
}
function relevantFields(data) {
  return questionFields(data).filter(field=>sameQuestion(field.q,$('question').value));
}
function matches(row,level,group) {return (!level||row[0]===level)&&(!group||row[1]===group);}
function category(raw) {
  if (!raw) return '';
  const first = raw.split(/\s+\/\s+/)[0];
  return ordinal.find(item=>norm(item)===norm(first)) || first;
}
function distribution(data,field,filter) {
  const counts = new Map();
  let skip = 0;
  if (!field) return {counts,skip,missing:true};
  const i = data.fields.indexOf(field)+4;
  for (const row of data.rows) {
    if (!filter(row)) continue;
    const answer = category(row[i]);
    if (!answer || excluded.has(norm(answer))) {skip++;continue;}
    counts.set(answer,(counts.get(answer)||0)+1);
  }
  return {counts,skip};
}
function teacherFor(data,field,row) {
  const type = teacherType(field.m);
  const cambridge = data.fields.findIndex(item=>item.identity);
  const spoken = row[3];
  const alternative = cambridge>=0 ? row[cambridge+4] : '';
  const validName = value => !!value && !/^(no |ningun|sin docente|no se|no aplica)/.test(norm(value));
  const candidates = [spoken,alternative].filter(validName);
  const direct = type==='english' ? row[2] : type==='spoken' && new Set(candidates.map(norm)).size===1 ? candidates[0] : '';
  if (type==='spoken' && new Set(candidates.map(norm)).size>1) return '';
  const key=rosterKey(field.m,row[0],row[1]);
  const names=uniq(data.roster.filter(item=>rosterKey(...item.slice(0,3))===key).map(item=>item[3]));
  if (validName(direct)) {
    return names.find(name=>norm(name)===norm(direct)) || direct;
  }
  return names.length===1 ? names[0] : '';
}
function getField(data,q,m) {
  return questionFields(data).find(field=>sameQuestion(field.q,q)&&subjectKey(field.m)===subjectKey(m));
}
async function chartData(data,field) {
  const level = $('level').value, group = $('group').value, view = $('view').value;
  const filter = row=>matches(row,level,group);
  if (activeTab==='teachers') {
    const names = sorted(data.rows.filter(filter).map(row=>teacherFor(data,field,row)));
    const series=names.map(name=>[name,distribution(data,field,row=>filter(row)&&teacherFor(data,field,row)===name)]);
    const unknown=distribution(data,field,row=>filter(row)&&!teacherFor(data,field,row));
    if (unknown.counts.size) series.push(['Sin docente identificado',unknown]);
    return series;
  }
  if (view==='year') {
    return Promise.all(catalog.map(async item=>{
      const other = await edition(item.id);
      return [editionLabel(item),distribution(other,getField(other,field.q,field.m),filter)];
    }));
  }
  if (view==='subject') return relevantFields(data).map(item=>[item.m||'General',distribution(data,item,filter)]);
  if (view==='group') return sorted(data.rows.filter(row=>matches(row,level,ALL)).map(row=>row[1]))
    .map(item=>['Grupo '+item,distribution(data,field,row=>matches(row,level,item))]);
  return sorted(data.rows.filter(row=>matches(row,ALL,group)).map(row=>row[0]))
    .map(item=>[item,distribution(data,field,row=>matches(row,item,group))]);
}
function display(series,field,teacherMode) {
  const charts = $('charts');
  charts.replaceChildren(node('h2',field.q));
  if (!series.length) {
    charts.append(node('p',teacherMode?'Asigna profesorado a la materia, nivel y grupo para ver sus respuestas.':'No hay respuestas para la selección.'));
    return;
  }
  const categories = uniq(series.flatMap(([,data])=>[...data.counts.keys()])).sort((a,b)=>
    (ordinal.indexOf(a)+1||99)-(ordinal.indexOf(b)+1||99)||a.localeCompare(b,'es'));
  const context = [
    $('edition').selectedOptions[0]?.textContent,
    $('subject').parentElement.hidden ? '' : field.m,
    $('level').parentElement.hidden ? '' : ($('level').value || 'Todos los niveles'),
    $('group').parentElement.hidden ? '' : ($('group').value ? 'Grupo '+$('group').value : 'Todos los grupos')
  ].filter(Boolean);
  report = {series,field,categories,context,teacherMode};
  $('pdf').disabled = false;
  charts.append(node('p',context.join(' · '),'muted'));
  const legend = node('div',undefined,'legend');
  categories.forEach((label,i)=>{
    const item = node('span');
    const swatch = node('i',undefined,'swatch');
    swatch.style.background = answerColor(label,i);
    item.append(swatch,label);
    legend.append(item);
  });
  charts.append(node('p','Porcentaje de respuestas válidas (0–100 %).','muted'),legend);
  const combined=!teacherMode && $('view').value==='level' && categories.length>0;
  if (combined) {
    const container=node('div',undefined,'level-scroll');
    const graph=node('div',undefined,'level-chart');
    graph.style.setProperty('--level-count',series.length);
    graph.setAttribute('role','img');
    graph.setAttribute('aria-label','Columnas agrupadas por nivel. Consulta los recuentos debajo.');
    for (const [label,data] of series) {
      const cohort=node('div');
      const bars=node('div',undefined,'bars');
      const n=validCount(data);
      categories.forEach((answer,i)=>{
        const count=data.counts.get(answer)||0;
        const bar=node('div',undefined,'bar');
        bar.style.height=(n?count/n*100:0)+'%';
        bar.style.background=answerColor(answer,i);
        bar.title=answer+' · '+(n?fmt.format(count/n*100)+' %':'Sin datos');
        bars.append(bar);
      });
      cohort.append(bars,node('small',label));
      graph.append(cohort);
    }
    container.append(graph);
    charts.append(container);
  }
  const plots = node('div',undefined,'plots');
  const comparable = teacherMode && categories.length && categories.every(label=>ordinal.includes(label));
  if (comparable) charts.append(node('p','Bastante + Mucho resume las dos categorías más altas. Interpreta ese porcentaje según la pregunta y el número de respuestas.','muted'));
  if (teacherMode) charts.append(node('p','El bloque «Sin docente identificado» muestra respuestas del grupo sin atribuirlas a una persona.','muted'));
  for (const [name,{counts,skip,missing}] of series) {
    const plot = node('article',undefined,'plot');
    plot.append(node('strong',name));
    const n = [...counts.values()].reduce((sum,value)=>sum+value,0);
    if (missing || !n) {
      plot.append(node('p',missing?'Pregunta no disponible en esta edición.':'Sin respuestas válidas'+(skip?' · '+skip+' fuera del cálculo':'')+'.'));
      plots.append(plot);
      continue;
    }
    const bars = node('div',undefined,'bars');
    bars.setAttribute('aria-hidden','true');
    const details = node('details');
    details.append(node('summary',n+' válidas'+(skip?' · '+skip+' fuera del cálculo':'')));
    categories.forEach((label,i)=>{
      const count = counts.get(label)||0;
      const bar = node('div',undefined,'bar');
      bar.style.height = count/n*100+'%';
      bar.style.background = answerColor(label,i);
      bar.title = label+' · '+fmt.format(count/n*100)+' % ('+count+')';
      bars.append(bar);
      details.append(node('p',bar.title));
    });
    if (!combined) plot.append(bars);
    plot.append(details);
    if (comparable && name!=='Sin docente identificado') plot.append(node('p','Bastante + Mucho · '+fmt.format(((counts.get('Bastante')||0)+(counts.get('Mucho')||0))/n*100)+' %'));
    plots.append(plot);
  }
  charts.append(plots);
}
async function update(changed) {
  const version = ++renderCount;
  report = undefined;
  $('pdf').disabled = true;
  const id = $('edition').value;
  if (!id) {
    $('charts').replaceChildren(node('p','Importa una edición para consultar los gráficos.'));
    $('analysis').querySelector('.row').hidden = true;
    $('teacher-info').hidden = true;
    return;
  }
  $('analysis').querySelector('.row').hidden = false;
  $('teacher-info').hidden = activeTab !== 'teachers';
  try {
    const data = await edition(id);
    if (version!==renderCount) return;
    if (changed==='edition') {
      const eligible = questionFields(data).filter(field=>activeTab!=='teachers'||field.m);
      options('question',uniq(eligible.map(field=>field.q)),eligible.length?undefined:'Sin preguntas por materia');
      options('level',sorted(data.rows.map(row=>row[0])),'Todos');
      updateRosterPanel(data,id);
    }
    if (changed==='edition'||changed==='level') {
      options('group',sorted(data.rows.filter(row=>matches(row,$('level').value,ALL)).map(row=>row[1])),'Todos');
    }
    if (changed==='edition'||changed==='question') {
      const fields = relevantFields(data);
      options('subject',fields.map(field=>[field.m,field.m||'General']));
    }
    $('subject').parentElement.hidden = activeTab==='results' && $('view').value==='subject';
    const progress=activeTab==='results' && $('view').value==='progress';
    $('metric-label').hidden=!progress;
    $('level').parentElement.hidden = activeTab==='results' && (progress || $('view').value==='level');
    $('group').parentElement.hidden = activeTab==='results' && $('view').value==='group';

    const field = getField(data,$('question').value,$('subject').value);
    if (!field) {
      $('charts').replaceChildren(node('p','Esta edición no tiene preguntas vinculadas a una materia.'));
      return;
    }
    if (progress) {await showProgress(field,version);return;}
    const series = await chartData(data,field);
    if(version===renderCount) display(series,field,activeTab==='teachers');
  } catch(error) {
    if (version===renderCount) status(error.message,true);
  }
}
function updateRosterPanel(data,id) {
  const current=catalog.find(item=>item.id===id);
  const earlier=catalog.filter(item=>item.year<current.year);
  options('roster-source',earlier.map(item=>[item.id,editionLabel(item)]));
  $('roster-reuse').hidden=!earlier.length || !!data.roster.length;
  $('roster-copy').hidden=true;
  $('roster-info').textContent=data.roster.length ? data.roster.length+' asignaciones guardadas en esta edición.' : 'Esta edición todavía no tiene un Excel de profesorado.';
}
function parseRoster(bytes) {
  const book=XLSX.read(bytes,{type:'array'});
  const rows=XLSX.utils.sheet_to_json(book.Sheets[book.SheetNames[0]],{header:1,raw:false,defval:''})
    .map(row=>Array.from({length:4},(_,i)=>trim(row[i]))).filter(row=>row.some(Boolean));
  if (rows.length && /^(asignatura|materia|area)/.test(norm(rows[0][0])) && /^(curso|nivel)/.test(norm(rows[0][1]))) rows.shift();
  const bad=rows.findIndex(row=>row.some(value=>!value||value.length>160));
  if (bad>=0) throw Error('Completa las cuatro columnas en la fila de datos '+(bad+1)+'.');
  if (!rows.length) throw Error('El archivo no contiene asignaciones.');
  const distinct=new Map();
  for (const [matter,level,group,name] of rows) {
    distinct.set(rosterKey(matter,level,group)+'|'+norm(name),[matter,level,group,name]);
  }
  return [...distinct.values()];
}
async function storeRoster(id,roster,data) {
  const available=new Set(data.rows.map(row=>JSON.stringify([levelKey(row[0]),groupKey(row[1])])));
  const known=new Set(questionFields(data).map(field=>subjectKey(field.m)));
  const matched=roster.filter(([m,l,g])=>known.has(subjectKey(m))&&available.has(JSON.stringify([levelKey(l),groupKey(g)]))).length;
  if (data.roster.length && !window.confirm('¿Sustituir las asignaciones de esta edición por las del nuevo Excel?')) return;
  await api(null,'PUT',{id,data:roster});
  cache.delete(id);
  updateRosterPanel(await edition(id),id);
  status(roster.length+' asignaciones guardadas. '+matched+' coinciden con materias y grupos de la encuesta.');
  if ($('edition').value===id) await update('roster');
}
async function importRoster() {
  const id=$('edition').value;
  const file=$('roster-file').files[0];
  if (!file) throw Error('Selecciona el Excel de profesorado.');
  if (file.size>10_000_000) throw Error('El archivo supera 10 MB.');
  await storeRoster(id,parseRoster(await file.arrayBuffer()),await edition(id));
}
async function copyRoster() {
  const id=$('edition').value;
  const previous=await edition($('roster-source').value);
  if (!previous.roster.length) throw Error('Esa edición no tiene asignaciones guardadas.');
  await storeRoster(id,previous.roster,await edition(id));
}
function onClick(id,handler) {
  $(id).onclick=async()=>{
    $(id).disabled=true;
    try {status();await handler();}
    catch(error){status(error.message,true);}
    finally{$(id).disabled=false;}
  };
}
$('pdf').onclick=()=>{try{downloadReport();}catch(error){status(error.message,true);}};
onClick('read',readFile);
onClick('save',saveFile);
onClick('roster-upload',importRoster);
onClick('roster-confirm',copyRoster);
$('roster-reuse').onclick=()=>{$('roster-copy').hidden=!$('roster-copy').hidden;};
$('file').onchange=()=>{draft=undefined;$('review').hidden=true;};
for(const id of ['edition','question','subject','level','group','view','metric']) $(id).onchange=()=>update(id);

for(const button of document.querySelectorAll('[data-tab]')) button.onclick=()=>tab(button.dataset.tab);
tab('import');
refresh().catch(error=>status(error.message,true));
function downloadReport() {
  if (!report) return;
  if(report.trend) return downloadTrend(report.trend);
  const {series,field,categories,context,teacherMode} = report;
  const doc = new window.jspdf.jsPDF();
  const margin = 16;
  let y = 18;
  function room(height) {
    if (y+height > 279) {doc.addPage();y=18;}
  }
  function write(value,size=10,bold=false) {
    doc.setFont('helvetica',bold?'bold':'normal');
    doc.setFontSize(size);
    for (const line of doc.splitTextToSize(value,178)) {
      room(6);
      doc.text(line,margin,y);
      y+=size*.42+1;
    }
  }
  write(teacherMode?'Encuesta del alumnado · Profesorado':'Encuesta del alumnado · Resultados',15,true);
  write(field.q,13,true);
  write(context.join(' · '));
  write('Porcentaje de respuestas válidas (0–100 %).');
  for (const [name,{counts,skip,missing}] of series) {
    const n = [...counts.values()].reduce((sum,value)=>sum+value,0);
    room(64);
    y+=4;
    write(name,12,true);
    if (!n) {
      write(missing?'Pregunta no disponible en esta edición.':'Sin respuestas válidas · '+skip+' fuera del cálculo.');
      continue;
    }
    write(n+' válidas'+(skip?' · '+skip+' fuera del cálculo':''));
    const width = Math.min(16,145/Math.max(categories.length,1));
    const baseline = y+30;
    doc.setDrawColor(180,192,204);
    doc.line(margin,baseline,194,baseline);
    categories.forEach((label,i)=>{
      const height = 30*(counts.get(label)||0)/n;
      if (!height) return;
      doc.setFillColor(answerColor(label,i));
      doc.rect(margin+2+i*(width+2),baseline-height,width,height,'F');
    });
    y=baseline+5;
    categories.forEach((label,i)=>{
      const count=counts.get(label)||0;
      room(7);
      doc.setFillColor(answerColor(label,i));
      doc.rect(margin,y-3,3,3,'F');
      doc.setFontSize(9);
      doc.text(doc.splitTextToSize(label+' · '+fmt.format(count/n*100)+' % ('+count+')',172),margin+5,y);
      y+=doc.splitTextToSize(label,142).length*4+1;
    });
    if (teacherMode && name!=='Sin docente identificado' && categories.every(label=>ordinal.includes(label))) {
      write('Bastante + Mucho · '+fmt.format(((counts.get('Bastante')||0)+(counts.get('Mucho')||0))/n*100)+' %');
    }
  }
  doc.save('encuesta-'+(teacherMode?'profesorado':'resultados')+'.pdf');
}
const levelColors=['#0065af','#a13763','#4e7e30','#8553a2','#a34d19','#007e79','#594e9a','#a14137'];
const validCount = distribution => [...distribution.counts.values()].reduce((a,b)=>a+b,0);
function trendPoint(distribution,metric) {
  const n=validCount(distribution);
  const values=metric==='top'?['Bastante','Mucho']:[metric];
  return n ? {n,value:values.reduce((sum,key)=>sum+(distribution.counts.get(key)||0),0)*100/n} : null;
}
async function showProgress(field,version) {
  const group=$('group').value;
  const rows=await Promise.all([...catalog].reverse().map(async entry=>[entry,await edition(entry.id)]));
  if(version!==renderCount) return;
  const entries=rows.filter(([,data])=>getField(data,field.q,field.m));
  const byLevel=new Map(entries.flatMap(([,data])=>data.rows.filter(row=>!group||groupKey(row[1])===groupKey(group)).map(row=>[levelKey(row[0]),row[0]])));
  const available=sorted([...byLevel.values()]);
  const distributions=available.map(level=>entries.map(([,data])=>
    distribution(data,getField(data,field.q,field.m),row=>levelKey(row[0])===levelKey(level)&&(!group||groupKey(row[1])===groupKey(group)))));
  const labels=sorted(distributions.flatMap(values=>values.flatMap(value=>[...value.counts.keys()])));
  const hasTop=labels.includes('Bastante')||labels.includes('Mucho');
  options('metric',[...(hasTop?[['top','Bastante + Mucho']]:[]),...labels]);
  const metric=$('metric').value;
  const metricLabel=$('metric').selectedOptions[0]?.textContent;
  const names=entries.map(([entry])=>editionLabel(entry));
  const series=available.map((level,i)=>({level,color:levelColors[i%levelColors.length],
    points:distributions[i].map(value=>trendPoint(value,metric))})).filter(line=>line.points.some(Boolean));
  const title=field.q+(field.m?' · '+field.m:'');
  const context=(metricLabel||'')+' · '+(group?'Grupo '+group:'Todos los grupos');
  const charts=$('charts');
  charts.replaceChildren(node('h2',title),node('p',context,'muted'));
  if(!series.length) {charts.append(node('p','Sin respuestas válidas para estos filtros.'));return;}
  if(names.length===1) charts.append(node('p','Hay una edición comparable. La línea de evolución aparecerá cuando haya dos o más ediciones con datos.'));
  const trend={title,context,names,series};
  report={trend};
  $('pdf').disabled=false;
  charts.append(trendSvg(trend));
  const legend=node('div',undefined,'legend');
  for(const {level,color} of series) {
    const label=node('span'), swatch=node('i',undefined,'swatch');
    swatch.style.background=color;label.append(swatch,level);legend.append(label);
  }
  charts.append(legend);
  const details=node('details');
  details.append(node('summary','Ver porcentajes y respuestas válidas'));
  names.forEach((name,i)=>{
    details.append(node('strong',name));
    for(const {level,points} of series) {
      const point=points[i];
      details.append(node('p',level+' · '+(point?fmt.format(point.value)+' % ('+point.n+' válidas)':'Sin datos')));
    }
  });
  charts.append(details);
}
function xFor(i,n,left,width) {return n===1?left+width/2:left+i*width/(n-1);}
function trendSvg({names,series}) {
  const ns='http://www.w3.org/2000/svg';
  const svg=document.createElementNS(ns,'svg');
  svg.setAttribute('class','trend');
  svg.setAttribute('viewBox','0 0 740 290');
  svg.setAttribute('role','img');
  svg.setAttribute('aria-label','Evolución por nivel, con porcentajes de 0 a 100');
  function add(tag,attributes,text) {
    const element=document.createElementNS(ns,tag);
    Object.entries(attributes).forEach(([key,value])=>element.setAttribute(key,value));
    if(text)element.textContent=text;
    svg.append(element);
    return element;
  }
  const y=value=>238-value*2.15;
  const x=i=>xFor(i,names.length,65,610);
  for(const value of [0,25,50,75,100]){
    add('line',{x1:53,x2:705,y1:y(value),y2:y(value),class:'grid'});
    add('text',{x:45,y:y(value)+4,'text-anchor':'end'},value+' %');
  }
  const labelEvery=Math.ceil(names.length/7);
  names.forEach((name,i)=>{
    if(i%labelEvery===0) add('text',{x:x(i),y:260,'text-anchor':'middle'},name.split(' · ')[0]);
  });
  for(const {level,color,points} of series) points.forEach((point,i)=>{
    if(!point)return;
    if(points[i-1]) add('line',{x1:x(i-1),y1:y(points[i-1].value),x2:x(i),y2:y(point.value),stroke:color,'stroke-width':2.5});
    const circle=add('circle',{cx:x(i),cy:y(point.value),r:4,fill:color});
    const title=document.createElementNS(ns,'title');
    title.textContent=level+' · '+names[i]+' · '+fmt.format(point.value)+' % ('+point.n+' válidas)';
    circle.append(title);
  });
  return svg;
}
function downloadTrend({title,context,names,series}) {
  const doc=new window.jspdf.jsPDF();
  const header=doc.splitTextToSize(title,175);
  doc.setFontSize(13);
  doc.text(header,16,18);
  const top=26+header.length*6;
  doc.setFontSize(10);
  doc.text(context,16,top);
  const x=i=>xFor(i,names.length,35,150), y=value=>top+67-value*.55;
  for(const value of [0,25,50,75,100]){
    doc.setDrawColor('#cbd8e3');
    doc.line(25,y(value),192,y(value));
    doc.text(value+' %',12,y(value)+1);
  }
  for(const {color,points} of series) points.forEach((point,i)=>{
    if(!point)return;
    doc.setDrawColor(color);
    doc.setFillColor(color);
    if(points[i-1]) doc.line(x(i-1),y(points[i-1].value),x(i),y(point.value));
    doc.circle(x(i),y(point.value),1.2,'F');
  });
  const labelEvery=Math.ceil(names.length/6);
  names.forEach((name,i)=>{
    if(i%labelEvery===0)doc.text(name.split(' · ')[0],x(i),y(0)+6,{align:'center'});
  });
  let position=y(0)+14;
  for(const {level,color,points} of series){
    if(position>270){doc.addPage();position=18;}
    doc.setFillColor(color);
    doc.rect(16,position-3,3,3,'F');
    doc.text(level,22,position);
    position+=6;
    points.forEach((point,i)=>{
      if(position>270){doc.addPage();position=18;}
      doc.text(level+' · '+names[i]+' · '+(point?fmt.format(point.value)+' % ('+point.n+' válidas)':'Sin datos'),22,position);
      position+=5;
    });
    position+=3;
  }
  doc.save('encuesta-evolucion.pdf');
}
