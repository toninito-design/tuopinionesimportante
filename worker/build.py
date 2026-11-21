import base64
import json
import subprocess
from pathlib import Path

here = Path(__file__).resolve().parent
page = (here / 'page.html').read_text()
page = page.replace('SCHEMA', (here / 'schema.json').read_text().replace('<', '\\u003c'))
page = page.replace('CLIENT', (here / 'client.js').read_text())
dist = here.parent / 'dist'
(dist / 'server').mkdir(parents=True, exist_ok=True)
manifest = here.parent / '.openai' / 'hosting.json'
if manifest.is_file():
    (dist / '.openai').mkdir(exist_ok=True)
    (dist / '.openai' / 'hosting.json').write_bytes(manifest.read_bytes())
source_zip = subprocess.check_output(
    ['git', 'archive', '--format=zip', '--prefix=encuesta-alumnado/', 'HEAD'],
    cwd=here.parent,
)
(dist / 'server' / 'index.js').write_text(
    (here / 'index.js').read_text()
    + '\nconst PAGE = ' + json.dumps(page, ensure_ascii=False)
    + ';\nconst XLSX_SOURCE = ' + json.dumps((here / 'xlsx.js').read_text(), ensure_ascii=False)
    + ';\nconst PDF_SOURCE = ' + json.dumps((here / 'pdf.js').read_text(), ensure_ascii=False)
    + ';\nconst SOURCE_ZIP = ' + json.dumps(base64.b64encode(source_zip).decode('ascii')) + ';\n'
)
