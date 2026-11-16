/* Timesheet Converter: drop the Jibble CSV, click Export. The conversion is Python (py/*.py) in worker.js. */
'use strict';

const APP_VERSION = '1.2.0';
const TEMPLATE_URL = 'template/format-timesheet-clean.xlsx';
const MON = ['', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
const MON_SHORT = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'];

const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* ---- worker ---- */
const worker = new Worker(`worker.js?v=${APP_VERSION}`);
let rpcId = 0;
const pending = new Map();
worker.onmessage = (ev) => {
  const m = ev.data;
  if (m.type === 'status') return;
  const p = pending.get(m.id);
  if (!p) return;
  pending.delete(m.id);
  if (m.ok) p.resolve(m.result); else { console.error(m.detail); p.reject(new Error(m.error)); }
};
function rpc(cmd, args = {}, transfer = []) {
  return new Promise((resolve, reject) => {
    const id = ++rpcId;
    pending.set(id, { resolve, reject });
    worker.postMessage({ id, cmd, args }, transfer);
  });
}
const engine = rpc('init');           // start loading Python right away
engine.catch(() => {});

/* ---- state ---- */
let A = null;                        // what Python read from the CSV
let seq = 0;

function fmt(iso, year = true) {
  const [y, m, d] = iso.split('-').map(Number);
  return `${d} ${MON[m]}${year ? ' ' + y : ''}`;
}
function fileName() {
  const [y1, m1, d1] = A.period.start.split('-').map(Number);
  const [y2, m2, d2] = A.period.end.split('-').map(Number);
  return `Timesheet ${d1} ${MON_SHORT[m1]}${y1 !== y2 ? ' ' + y1 : ''} - ${d2} ${MON_SHORT[m2]} ${y2}.xlsx`;
}
function setMsg(html, kind = '') { $('msg').className = 'msg ' + kind; $('msg').innerHTML = html; }
const spin = (t) => `<span class="spinner"></span>${esc(t)}`;

/* ---- 1. import ---- */
async function useFile(f) {
  const my = ++seq;
  A = null;
  $('export').disabled = true;
  $('summary').innerHTML = '';
  setMsg('');
  $('drop').classList.add('has');
  $('name').textContent = f.name;
  $('info').className = 'hint';
  $('info').innerHTML = spin('Reading…');
  try {
    await engine.catch(() => { throw new Error('The converter could not load. Check the internet connection and reload the page.'); });
    const res = await rpc('load', { csv: await f.arrayBuffer() });
    if (my !== seq) return;
    A = JSON.parse(res);
    if (A.period.ndays > 31) throw new Error('This file covers more than 31 days. Export one month from Jibble.');
    $('info').innerHTML = `${A.people.length} people &middot; ${esc(fmt(A.period.start))} &ndash; ${esc(fmt(A.period.end))}`;
    $('export').disabled = false;
  } catch (e) {
    if (my !== seq) return;
    A = null;
    $('info').className = 'hint err';
    $('info').textContent = e.message;
  }
}

const drop = $('drop'), input = $('file');
input.addEventListener('change', () => { if (input.files[0]) useFile(input.files[0]); input.value = ''; });
drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); if (e.dataTransfer.files[0]) useFile(e.dataTransfer.files[0]); });
window.addEventListener('dragover', (e) => e.preventDefault());   // a missed drop must not open the file in the tab
window.addEventListener('drop', (e) => e.preventDefault());

/* ---- 2. export ---- */
let url = null;
$('export').addEventListener('click', async () => {
  if (!A) return;
  $('export').disabled = true;
  $('summary').innerHTML = '';
  setMsg(spin('Building the timesheet…'));
  try {
    const r = await fetch(`${TEMPLATE_URL}?v=${APP_VERSION}`);
    if (!r.ok) throw new Error('Could not load the template.');
    const template = await r.arrayBuffer();
    const res = await rpc('convert', { template, edits: {} }, [template]);
    const S = JSON.parse(res.summary);
    if (url) URL.revokeObjectURL(url);
    url = URL.createObjectURL(new Blob([res.bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url; a.download = fileName(); a.click();
    setMsg(`Done: <b>${esc(fileName())}</b> is in your Downloads. <a href="${url}" download="${esc(fileName())}">Download again</a>`, 'ok');
    showSummary(S);
  } catch (e) {
    setMsg(esc(e.message), 'err');
  } finally {
    $('export').disabled = false;
  }
});

/* ---- what to check ---- */
function byPerson(list, f) {
  const g = new Map();
  for (const x of list) { if (!g.has(x.name)) g.set(x.name, []); g.get(x.name).push(x); }
  return `<ul>${[...g].map(([n, xs]) => `<li><b>${esc(n.toUpperCase())}</b>: ${xs.map(f).join(', ')}</li>`).join('')}</ul>`;
}
function showSummary(S) {
  const T = S.totals, I = A.issues;
  const tiles = [['People', S.n], ['Hadir', T.hadir], ['WFH', T.wfh], ['Lembur', T.lembur]];
  let html = `<div class="tiles">${tiles.map(([k, v]) => `<div class="tile"><div class="k">${k}</div><div class="v">${v}</div></div>`).join('')}</div>`;
  if (I.missing_out.length) {
    html += `<details><summary>Left blank: no clock-out<span>${I.missing_out.length}</span></summary>${byPerson(I.missing_out, (x) => esc(fmt(x.date, false)))}</details>`;
  }
  if (I.long_shifts.length) {
    html += `<details><summary>Shifts over 14 hours<span>${I.long_shifts.length}</span></summary>${byPerson(I.long_shifts, (x) => `${esc(fmt(x.date, false))} (${x.hours} h)`)}</details>`;
  }
  $('summary').innerHTML = html;
}
