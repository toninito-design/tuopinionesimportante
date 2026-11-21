from pathlib import Path
import re

root = Path(__file__).resolve().parent.parent
worker = root / 'worker'
page = (worker / 'page.html').read_text()
page = page.replace('<meta charset="utf-8">', '''<meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; script-src 'self' 'unsafe-inline'; style-src 'unsafe-inline'; img-src data: blob:; connect-src 'none'; base-uri 'none'; form-action 'none'">''')
page = re.sub(r'<a href="/descargar/encuesta-alumnado\.zip".*?</a>', '', page)
page = page.replace('src="/xlsx.js"', 'src="./worker/xlsx.js"').replace('src="/pdf.js"', 'src="./worker/pdf.js"')
page = page.replace('<h2 style="margin-top:22px">Ediciones guardadas</h2><ul id="history"></ul>', '''<h2 style="margin-top:22px">Ediciones guardadas</h2>
<p class="muted">Los archivos se procesan en este navegador. Guarda una copia para recuperar tus ediciones si cambias de dispositivo o borras los datos del navegador.</p>
<ul id="history"></ul>
<details><summary>Copia del histórico</summary>
<div class="row"><button id="backup-save">Descargar copia</button><label>Archivo de copia<input type="file" id="backup-file" accept=".json"></label><button id="backup-restore">Restaurar copia</button></div>
</details>''')
client = (worker / 'client.js').read_text()
start = client.index('async function api(')
end = client.index('async function edition(', start)
client = client[:start] + client[end:]
storage = (root / 'pages' / 'storage.js').read_text()
page = page.replace('SCHEMA', (worker / 'schema.json').read_text().replace('<', '\\u003c'),1)
page = page.replace('CLIENT', storage + '\n' + client + '''
onClick('backup-save',exportHistory);
onClick('backup-restore',restoreHistory);
''',1)
(root / 'index.html').write_text(page)
