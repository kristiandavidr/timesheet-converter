/* Runs the Python converter (py/*.py) in Pyodide, off the main thread.
   Messages in:  {id, cmd, args}   Messages out: {id, ok, result | error} and {type:'status', text} */
const PYODIDE_VERSION = '0.27.7';
const APP_VERSION = self.location.search.replace(/^\?v=/, '') || 'dev';
const PY_FILES = ['config.py', 'parse.py', 'roster.py', 'build.py', 'api.py'];
const WHEELS = ['et_xmlfile-2.0.0-py3-none-any.whl', 'openpyxl-3.1.5-py2.py3-none-any.whl'];

let ready = null;
const status = (text) => self.postMessage({ type: 'status', text });

async function boot() {
  status('Downloading Python (first time only, about 15 MB)…');
  importScripts(`https://cdn.jsdelivr.net/pyodide/v${PYODIDE_VERSION}/full/pyodide.js`);
  const py = await loadPyodide();
  status('Loading openpyxl…');
  // openpyxl is not part of the Pyodide distribution: install the pinned wheels shipped in vendor/
  await py.loadPackage(['micropip']);
  const wheels = WHEELS.map((w) => new URL(`vendor/${w}`, self.location.href).href);
  py.globals.set('_wheels', py.toPy(wheels));
  await py.runPythonAsync('import micropip\nawait micropip.install(_wheels, deps=False)');
  status('Loading converter…');
  py.FS.mkdirTree('/app/py');
  py.FS.mkdirTree('/work');
  for (const f of PY_FILES) {
    const r = await fetch(`py/${f}?v=${APP_VERSION}`);
    if (!r.ok) throw new Error(`Could not load py/${f} (${r.status})`);
    py.FS.writeFile(`/app/py/${f}`, await r.text());
  }
  py.runPython(`
import sys, warnings
warnings.filterwarnings('ignore')
sys.path.insert(0, '/app/py')
import api
`);
  return py;
}

function writeFile(py, path, buf) {
  if (buf) py.FS.writeFile(path, new Uint8Array(buf));
  else { try { py.FS.unlink(path); } catch (e) { /* not there */ } }
}

// Python exceptions come back as long tracebacks; keep the last line (the message)
function cleanError(e) {
  const msg = String(e && e.message || e);
  const lines = msg.trim().split('\n').filter(Boolean);
  const last = lines[lines.length - 1] || msg;
  return last.replace(/^\w*(Error|Exception):\s*/, '');
}

self.onmessage = async (ev) => {
  const { id, cmd, args } = ev.data;
  try {
    if (!ready) ready = boot();
    const py = await ready;
    if (cmd === 'init') { self.postMessage({ id, ok: true, result: 'ready' }); return; }
    const api = py.globals.get('api');
    let result;
    if (cmd === 'load') {
      writeFile(py, '/work/in.csv', args.csv);
      result = api.load('/work/in.csv');
    } else if (cmd === 'convert') {
      writeFile(py, '/work/template.xlsx', args.template);
      const summary = api.convert('/work/template.xlsx', '/work/out.xlsx', JSON.stringify(args.edits));
      const bytes = py.FS.readFile('/work/out.xlsx');
      self.postMessage({ id, ok: true, result: { summary, bytes } }, [bytes.buffer]);
      api.destroy();
      return;
    } else {
      throw new Error('unknown command ' + cmd);
    }
    api.destroy();
    self.postMessage({ id, ok: true, result });
  } catch (e) {
    if (ready && !(await ready.then(() => true, () => false))) ready = null; // failed boot: allow a retry
    self.postMessage({ id, ok: false, error: cleanError(e), detail: String(e && e.message || e) });
  }
};
