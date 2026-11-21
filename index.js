
const LIMIT = 7_000_000;
const text = (value, size = 500) => typeof value === 'string' && value.length <= size;
const output = (data, status = 200) => Response.json(data, {status, headers: {'Cache-Control':'no-store'}});
function validEdition(data) {
  if (!data || !/^\d{4}–\d{2}$/.test(data.year) || !text(data.label,80) ||
      !Array.isArray(data.fields) || !data.fields.length || data.fields.length > 500 ||
      !Array.isArray(data.rows) || !data.rows.length || data.rows.length > 15000) return false;
  return data.fields.every(field => field && text(field.q) && field.q.trim() && text(field.m) &&
    ['', 'english','spoken','cambridge'].includes(field.teacher)) &&
    data.rows.every(row => Array.isArray(row) && row.length === 4 + data.fields.length &&
      row.every((value, index) => text(value, index < 4 ? 160 : 500))) &&
    Array.isArray(data.roster) && data.roster.length <= 3000 &&
    data.roster.every(row => Array.isArray(row) && row.length === 4 && row.every(value => text(value,160)));
}
export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (url.pathname === '/' && request.method === 'GET') return new Response(PAGE,{
      headers:{'Content-Type':'text/html; charset=utf-8','X-Content-Type-Options':'nosniff',
        'Content-Security-Policy':"default-src 'none'; script-src 'self' 'unsafe-inline'; style-src 'unsafe-inline'; connect-src 'self'; base-uri 'none'; form-action 'none'"}
    });
    if (['/xlsx.js','/pdf.js'].includes(url.pathname) && request.method === 'GET') return new Response(url.pathname === '/xlsx.js' ? XLSX_SOURCE : PDF_SOURCE,{
      headers:{'Content-Type':'text/javascript; charset=utf-8','Cache-Control':'public, max-age=86400','X-Content-Type-Options':'nosniff'}
    });
    if (url.pathname === '/descargar/encuesta-alumnado.zip' && request.method === 'GET') {
      if (!request.headers.get('oai-authenticated-user-id')) return output({error:'Inicia sesión para descargar el código.'},401);
      return new Response(Uint8Array.from(atob(SOURCE_ZIP), character => character.charCodeAt(0)), {
        headers: {
          'Content-Type':'application/zip',
          'Content-Disposition':'attachment; filename="encuesta-alumnado.zip"',
          'Cache-Control':'private, no-store',
          'X-Content-Type-Options':'nosniff'
        }
      });
    }
    if (url.pathname !== '/api/editions') return output({error:'Página no encontrada.'},404);
    if (!env.BUCKET) return output({error:'El almacenamiento del histórico no está disponible.'},503);
    const identity = request.headers.get('oai-authenticated-user-id');
    if (!identity) return output({error:'Inicia sesión para acceder al histórico.'},401);
    const key = await crypto.subtle.digest('SHA-256',new TextEncoder().encode(identity));
    const prefix = Array.from(new Uint8Array(key), byte=>byte.toString(16).padStart(2,'0')).join('')+'/editions/';
    try {
      if (request.method === 'GET') {
        const id = url.searchParams.get('id');
        if (id) {
          if (!/^[a-f0-9]{64}$/.test(id)) return output({error:'Edición no válida.'},400);
          const edition = await env.BUCKET.get(prefix+id);
          return edition ? output(await edition.json()) : output({error:'Edición no encontrada.'},404);
        }
        const result = [];
        let cursor;
        do {
          const page = await env.BUCKET.list({prefix,include:['customMetadata'],cursor});
          result.push(...page.objects.map(object=>({id:object.key.slice(prefix.length),...object.customMetadata})));
          cursor = page.truncated ? page.cursor : undefined;
        } while (cursor);
        return output(result.sort((a,b)=>b.year.localeCompare(a.year)||a.label.localeCompare(b.label)));
      }
      if (!['POST','PUT'].includes(request.method) || request.headers.get('Origin') !== url.origin) return output({error:'Operación no permitida.'},403);
      if (Number(request.headers.get('Content-Length')) > LIMIT) return output({error:'El archivo supera el tamaño admitido.'},413);
      const body = await request.text();
      if (body.length > LIMIT) return output({error:'El archivo supera el tamaño admitido.'},413);
      const {id,data} = JSON.parse(body);
      if (request.method === 'PUT') {
        if (!/^[a-f0-9]{64}$/.test(id)) return output({error:'Edición no válida.'},400);
        const existing = await env.BUCKET.get(prefix+id);
        if (!existing) return output({error:'Edición no encontrada.'},404);
        const updated = {...await existing.json(),roster:data};
        if (!validEdition(updated)) return output({error:'Comprueba las asignaciones de profesorado.'},400);
        const saved = await env.BUCKET.put(prefix+id,JSON.stringify(updated),{
          onlyIf:{etagMatches:existing.etag},customMetadata:existing.customMetadata
        });
        return saved ? output({id}) : output({error:'La edición ha cambiado. Actualiza y vuelve a intentarlo.'},409);
      }
      if (!/^[a-f0-9]{64}$/.test(id) || !validEdition(data)) return output({error:'Comprueba el año, las columnas y los datos.'},400);
      const saved = await env.BUCKET.put(prefix+id, JSON.stringify(data), {
        onlyIf:{etagDoesNotMatch:'*'},customMetadata:{year:data.year,label:data.label,count:String(data.rows.length)}
      });
      return saved ? output({id},201) : output({error:'Este Excel ya está guardado en el histórico.'},409);
    } catch (error) {
      return output({error:error instanceof SyntaxError ? 'El archivo enviado no es válido.' : 'No se pudo acceder al histórico. Vuelve a intentarlo.'}, error instanceof SyntaxError ? 400 : 500);
    }
  }
};
