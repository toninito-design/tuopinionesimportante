const storageName = 'encuesta-alumnado:'+location.pathname.replace(/(?:index\.html)?\/?$/,'/');
let browserDatabase;
function openHistory() {
  if (!browserDatabase) {
    browserDatabase = new Promise((resolve,reject) => {
      const request = indexedDB.open(storageName,1);
      request.onupgradeneeded=()=>request.result.createObjectStore('editions',{keyPath:'id'});
      request.onsuccess=()=>resolve(request.result);
      request.onerror=()=>reject(Error('No se puede abrir el histórico de este navegador.'));
    });
  }
  return browserDatabase;
}
const recordId = id => typeof id==='string' && /^[a-f0-9]{64}$/.test(id);

async function transaction(mode,action) {
  const database=await openHistory();
  return new Promise((resolve,reject)=>{
    const tx=database.transaction('editions',mode);
    let value, issue;
    const fail=message=>{issue=Error(message);tx.abort();};
    tx.oncomplete=()=>resolve(value);
    tx.onabort=()=>reject(issue || Error('El navegador no pudo guardar o leer el histórico.'));
    action(tx.objectStore('editions'),result=>{value=result;},fail);
  });
}

async function api(id,method='GET',body) {
  if(method==='GET' && !id) {
    return transaction('readonly',(store,done)=>{
      const result=[];
      const cursor=store.openCursor();
      cursor.onsuccess=()=>{
        const item=cursor.result;
        if(!item){
          done(result.sort((a,b)=>b.year.localeCompare(a.year)||a.label.localeCompare(b.label)));
          return;
        }
        const {id,data}=item.value;
        result.push({id,year:data.year,label:data.label,count:data.rows.length});
        item.continue();
      };
    });
  }
  const key=id||body?.id;
  if(!recordId(key))throw Error('Edición no válida.');
  return transaction(method==='GET'?'readonly':'readwrite',(store,done,fail)=>{
    const request=store.get(key);
    request.onsuccess=()=>{
      const existing=request.result;
      if(method==='POST'){
        if(existing){fail('Este Excel ya está guardado en el histórico.');return;}
        store.add({id:key,data:body.data});
        done({id:key});
        return;
      }
      if(!existing){fail('Edición no encontrada.');return;}
      if(method==='PUT' || method==='PATCH'){
        const property=method==='PUT'?'roster':'curriculum';
        store.put({id:key,data:{...existing.data,[property]:body.data}});
        done({id:key});
        return;
      }
      done(existing.data);
    };
  });
}

function validBackup(record) {
  const strings = list => Array.isArray(list) && list.every(value=>typeof value==='string');
  const {id,data}=record||{};
  return recordId(id) && data && typeof data.year==='string' && /^\d{4}–\d{2}$/.test(data.year) &&
    typeof data.label==='string' && Array.isArray(data.fields) && data.fields.length>0 &&
    data.fields.every(field=>field&&typeof field.q==='string'&&typeof field.m==='string') &&
    Array.isArray(data.rows) && data.rows.every(row=>strings(row)&&row.length===data.fields.length+4) &&
    (data.curriculum===undefined || (Array.isArray(data.curriculum) && data.curriculum.every(row=>Array.isArray(row)&&row.length===3&&typeof row[0]==='string'&&typeof row[1]==='string'&&typeof row[2]==='boolean'))) &&
    Array.isArray(data.roster) && data.roster.every(row=>strings(row)&&row.length===4);
}

async function exportHistory() {
  const editions = await transaction('readonly',(store,done)=>{
    const request=store.getAll();
    request.onsuccess=()=>done(request.result);
  });
  if(!editions.length)throw Error('Todavía no hay ediciones para guardar.');
  const url=URL.createObjectURL(new Blob([JSON.stringify({version:1,editions})],{type:'application/json'}));
  const link=document.createElement('a');
  link.href=url;
  link.download='encuesta-historico.json';
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(()=>URL.revokeObjectURL(url),1000);
}
async function restoreHistory() {
  const file=$('backup-file').files[0];
  if(!file)throw Error('Selecciona una copia del histórico.');
  const contents=JSON.parse(await file.text());
  if(contents?.version!==1 || !Array.isArray(contents.editions) || !contents.editions.every(validBackup)) {
    throw Error('La copia del histórico no tiene el formato esperado.');
  }
  const count=await transaction('readwrite',(store,done)=>{
    let added=0;
    for(const record of contents.editions){
      const request=store.add(record);
      request.onsuccess=()=>{added++;done(added);};
      request.onerror=event=>{
        if(request.error.name==='ConstraintError'){event.preventDefault();event.stopPropagation();}
      };
    }
  });
  await refresh();
  $('backup-file').value='';
  status((count||0)+' ediciones añadidas al histórico.');
}
