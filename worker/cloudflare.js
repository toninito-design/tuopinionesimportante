import site from '../dist/server/index.js';

let cachedKeys;

function decodePart(value) {
  return Uint8Array.from(atob(value.replace(/-/g, '+').replace(/_/g, '/')), c => c.charCodeAt(0));
}

function decodeJson(value) {
  return JSON.parse(new TextDecoder().decode(decodePart(value)));
}

async function publicKey(issuer, kid) {
  if (cachedKeys?.issuer !== issuer || cachedKeys.expires < Date.now() || !cachedKeys.keys.some(key => key.kid === kid)) {
    const response = await fetch(issuer + '/cdn-cgi/access/certs');
    if (!response.ok) throw new Error('Access keys unavailable');
    const {keys} = await response.json();
    if (!Array.isArray(keys)) throw new Error('Invalid key set');
    cachedKeys = {issuer,keys,expires:Date.now()+300_000};
  }
  const key = cachedKeys.keys.find(key => key.kid === kid && key.kty === 'RSA');
  if (!key) throw new Error('Unknown signing key');
  return crypto.subtle.importKey('jwk',key,{name:'RSASSA-PKCS1-v1_5',hash:'SHA-256'},false,['verify']);
}

async function identity(request, env) {
  const issuer = env.ACCESS_ISSUER;
  const audience = env.ACCESS_AUD;
  if (!/^https:\/\/[a-z0-9-]+\.cloudflareaccess\.com$/.test(issuer) || !audience) return null;
  const token = request.headers.get('Cf-Access-Jwt-Assertion');
  if (!token || token.length > 16000) return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;
  const [head,payload,signature] = parts;
  const header = decodeJson(head);
  const claims = decodeJson(payload);
  const now = Date.now()/1000;
  if (header.alg !== 'RS256' || !header.kid || claims.iss !== issuer ||
      !Array.isArray(claims.aud) || !claims.aud.includes(audience) ||
      !Number.isFinite(claims.exp) || claims.exp <= now || (claims.nbf && claims.nbf > now) ||
      typeof claims.sub !== 'string' || !claims.sub || typeof claims.email !== 'string' || !claims.email) return null;
  const key = await publicKey(issuer,header.kid);
  const valid = await crypto.subtle.verify('RSASSA-PKCS1-v1_5',key,decodePart(signature),new TextEncoder().encode(head+'.'+payload));
  return valid ? claims.sub : null;
}

export default {
  async fetch(request,env,ctx) {
    let user;
    try { user = await identity(request,env); }
    catch { user = null; }
    if (!user) return new Response('Acceso privado. Inicia sesión con Cloudflare Access.',{
      status:401,headers:{'Cache-Control':'no-store','Content-Type':'text/plain; charset=utf-8'}
    });
    const headers = new Headers(request.headers);
    headers.set('oai-authenticated-user-id',user);
    return site.fetch(new Request(request,{headers}),env,ctx);
  }
};
