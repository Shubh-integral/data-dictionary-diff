(function(){
  'use strict';
  const $ = id => document.getElementById(id);
  const MARKUP = ['standard_field','risk_type','risk_status','Integral transformation','protege transform code','parent_column','qid','qid_sort_order','Notes'];
  const KEY = ['data_source_id','source_table','column_name','example_values'];

  let files = [];      // {name, size, sheets, rowCount}
  let tabs = [];       // {id, file, sheet, rows(aoa), nonEmpty, parsed}
  let result = null;
  let view = 'new';
  let downloads = null;
  const MAX_FILE_BYTES = 20 * 1024 * 1024;
  const MAX_TOTAL_BYTES = 50 * 1024 * 1024;
  const MAX_TOTAL_ROWS = 500002;
  const PARSE_TIMEOUT_MS = 30000;
  let loadGeneration = 0;
  let intakeQueue = Promise.resolve();
  const activeParsers = new Set();

  if (window.claude && typeof window.claude.use === 'function') {
    window.claude.use('downloads').then(d => { downloads = d; updateDlHint(); }).catch(()=>{});
  }
  const inClaude = !!(window.claude && typeof window.claude.use === 'function');
  function updateDlHint(){
    $('dlHint').textContent = (inClaude && !downloads) ? 'Download isn’t available in this view. Use “Copy worklist as TSV” and paste into a sheet.' : '';
  }

  // ---------- file intake ----------
  const drop = $('drop'), input = $('fileInput');
  drop.addEventListener('click', () => input.click());
  drop.addEventListener('keydown', e => { if (e.key==='Enter'||e.key===' ') { e.preventDefault(); input.click(); } });
  ['dragenter','dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => queueFiles(e.dataTransfer.files));
  input.addEventListener('change', () => { queueFiles(input.files); input.value=''; });

  function queueFiles(list){
    const incoming = Array.from(list), generation = loadGeneration;
    intakeQueue = intakeQueue.then(() => generation === loadGeneration ? addFiles(incoming, generation) : undefined)
      .catch(e => { if (generation === loadGeneration) setMsg('loadMsg', {kind:'err', text:e.message || 'Could not load files.'}); });
  }

  function parseFile(file){
    return new Promise((resolve, reject) => {
      const worker = createDictionaryParserWorker();
      let settled = false;
      const cancel = () => finish(new Error('File loading cancelled.'));
      const finish = (error, value) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer); worker.terminate(); activeParsers.delete(cancel);
        error ? reject(error) : resolve(value);
      };
      const timer = setTimeout(() => finish(new Error('Parsing took longer than 30 seconds. Try a smaller file.')), PARSE_TIMEOUT_MS);
      activeParsers.add(cancel);
      worker.onmessage = ({data}) => data.error ? finish(new Error(data.error)) : finish(null, data);
      worker.onerror = e => { e.preventDefault(); finish(new Error('The file parser could not finish. Try a smaller or valid spreadsheet.')); };
      worker.onmessageerror = () => finish(new Error('The file parser returned unreadable data.'));
      file.arrayBuffer().then(buffer => { if (!settled) worker.postMessage(buffer, [buffer]); }).catch(e => finish(e));
    });
  }

  async function addFiles(list, generation){
    const errs = [];
    $('runBtn').disabled = true;
    show('cancelLoadBtn');
    setMsg('loadMsg', {kind:'info', text:'Reading files locally…'});
    for (const f of list) {
      if (generation !== loadGeneration) return;
      if (files.length >= 5) { errs.push(`Skipped ${f.name}: up to five files at a time.`); continue; }
      if (!/\.(xlsx|xls|xlsm|csv|tsv|txt)$/i.test(f.name)) { errs.push(`Skipped ${f.name}: unsupported file type.`); continue; }
      if (f.size > MAX_FILE_BYTES) { errs.push(`Skipped ${f.name}: the maximum file size is 20 MB.`); continue; }
      if (files.reduce((sum, file) => sum + file.size, 0) + f.size > MAX_TOTAL_BYTES) { errs.push(`Skipped ${f.name}: selected files exceed the 50 MB total limit.`); continue; }
      let name = f.name;
      let suffix = 2;
      while (files.some(x => x.name === name)) name = f.name.replace(/(\.[^.]+)?$/, ` (${suffix++})$1`);
      try {
        const parsed = await parseFile(f);
        if (generation !== loadGeneration) return;
        if (files.reduce((sum, file) => sum + file.rowCount, 0) + parsed.rowCount > MAX_TOTAL_ROWS) throw new Error('Selected files exceed the 500,000-row total safety limit.');
        files.push({name, size: f.size, sheets: parsed.sheets, rowCount: parsed.rowCount});
      } catch(e) {
        if (generation !== loadGeneration) return;
        errs.push(`Could not read ${f.name}: ${e.message || e}. Is it a real spreadsheet or CSV?`);
      }
    }
    if (generation !== loadGeneration) return;
    setMsg('loadMsg', errs.length ? {kind:'err', text:'Some files were not added', items:errs} : null);
    try { rebuild(); }
    finally { $('runBtn').disabled = false; hide('cancelLoadBtn'); }
  }

  function removeFile(name){
    files = files.filter(f => f.name !== name);
    rebuild();
  }

  function rebuild(){
    $('fileList').innerHTML = files.map(f => `<li>${esc(f.name)} (${f.sheets.length} tab${f.sheets.length===1?'':'s'}) <button title="Remove" data-rm="${esc(f.name)}">×</button></li>`).join('');
    $('fileList').querySelectorAll('button').forEach(b => b.addEventListener('click', () => removeFile(b.dataset.rm)));
    tabs = [];
    for (const f of files) for (const s of f.sheets) {
      const rows = s.rows;
      const nonEmpty = rows.filter(r => r.some(c => String(c).trim() !== '')).length;
      tabs.push({id: tabs.length, file:f.name, sheet:s.name, rows, nonEmpty});
    }
    hide('results'); result = null;
    if (!files.length) {
      hide('pick');
      for (const id of ['baseSel','newSel','baseHint','newHint','readoutTable','tableWrap','stats','resMsg','pickMsg','dlHint']) $(id).replaceChildren();
      return;
    }
    show('pick');
    populateSelects();
  }

  // ---------- auto-detect ----------
  function scoreBaseline(t){
    const n = (t.file+' '+t.sheet).toLowerCase();
    let s = 0;
    if (/final/.test(n)) s += 3;
    if (/master/.test(n)) s += 3;
    if (/baseline/.test(n)) s += 4;
    if (/transform/.test(n)) s += 2;
    if (/\bnew\b|incoming|updated|v2|draft/.test(n)) s -= 4;
    if (t.parsed && t.parsed.hasMarkup) s += 3;      // has transformation columns
    if (t.nonEmpty < 2) s -= 10;
    return s;
  }
  function scoreNew(t){
    const n = (t.file+' '+t.sheet).toLowerCase();
    let s = 0;
    if (/\bnew\b|incoming|updated|v2|draft|latest/.test(n)) s += 4;
    if (/final|master|baseline/.test(n)) s -= 3;
    if (t.parsed && !t.parsed.hasMarkup) s += 2;    // un-marked-up
    if (t.nonEmpty < 2) s -= 10;
    return s;
  }

  function populateSelects(){
    for (const t of tabs) t.parsed = parseTab(t);
    const usable = tabs.filter(t => t.parsed.rows.length > 0);
    const opt = t => `<option value="${t.id}">${esc(t.file)} › ${esc(t.sheet)} (${t.parsed.rows.length} rows)</option>`;
    const html = '<option value="">— choose —</option>' + tabs.map(opt).join('');
    $('baseSel').innerHTML = html; $('newSel').innerHTML = html;

    let base = null, nw = null, autoB = false, autoN = false;
    if (usable.length) {
      const bs = [...usable].sort((x,y)=>scoreBaseline(y)-scoreBaseline(x));
      base = bs[0];
      // with two files, the new tab should come from the other file
      let pool = usable.filter(t => t !== base);
      if (files.length === 2 && pool.some(t => t.file !== base.file)) pool = pool.filter(t => t.file !== base.file);
      const ns = pool.sort((x,y)=>scoreNew(y)-scoreNew(x));
      nw = ns[0] || null;
      // weak evidence for the new tab (only 'has no markup') isn't enough unless it's the only other choice
      if (nw && scoreNew(nw) < 2 && usable.length !== 2) nw = null;
      // if both scores are flat (e.g. two single-tab files with neutral names), prefer the one with markup as baseline
      if (base && nw && scoreBaseline(base) === scoreBaseline(nw) && !base.parsed.hasMarkup && nw.parsed.hasMarkup) { const t = base; base = nw; nw = t; }
      autoB = base && scoreBaseline(base) > 0; autoN = nw && scoreNew(nw) > 0;
    }
    $('baseSel').value = base ? base.id : ''; $('newSel').value = nw ? nw.id : '';
    toggle('baseAuto', autoB); toggle('newAuto', autoN);
    describe();
    renderReadout();
    const msgs = [];
    if (usable.length === 1) msgs.push('Only one tab with usable rows so far (' + usable[0].sheet + '). Add the other file, or check the readout below if a tab should have been readable.');
    if (tabs.some(t => t.parsed.problem && t.parsed.rows.length === 0 && t.nonEmpty > 2)) msgs.push('Some tabs could not be read; see “What was found in each tab”.');
    if (files.length === 2) msgs.push(`Two files loaded: ${files[0].name} and ${files[1].name}. Check the two dropdowns point at the right tabs, then Compare.`);
    if (!usable.length) msgs.push('No tab contains any rows.');
    setMsg('pickMsg', msgs.length ? {kind:'warn', text:'', items:msgs} : null);
  }
  function renderReadout(){
    const h = ['file','tab','rows','layout','tables','source','markup','skipped','status'];
    let html = '<table><thead><tr>' + h.map(c=>`<th>${c}</th>`).join('') + '</tr></thead><tbody>';
    for (const t of tabs) {
      const p = t.parsed;
      html += `<tr><td>${esc(t.file)}</td><td>${esc(t.sheet)}</td><td>${p.rows.length}</td><td>${esc(p.layout||'—')}</td><td>${p.tables.size}</td><td>${esc([...p.sources].join(', ')||'—')}</td><td>${p.hasMarkup?'yes':'no'}</td><td>${p.bad + p.dupes.length}</td><td class="${p.problem?'bad':''}">${esc(p.problem || 'ok')}</td></tr>`;
    }
    $('readoutTable').innerHTML = html + '</tbody></table>';
  }
  $('baseSel').addEventListener('change', () => { hide('baseAuto'); describe(); });
  $('newSel').addEventListener('change', () => { hide('newAuto'); describe(); });

  function describe(){
    const b = tabs[$('baseSel').value], n = tabs[$('newSel').value];
    $('baseHint').textContent = b ? hintFor(b) : '';
    $('newHint').textContent = n ? hintFor(n) : '';
  }
  function hintFor(t){
    const p = t.parsed;
    const bits = [];
    if (p.problem) return 'Problem: ' + p.problem;
    bits.push(p.layout);
    bits.push(`${p.tables.size} tables`);
    bits.push(p.hasMarkup ? 'has markup' : 'no markup columns');
    if (p.sources.size) bits.push('source: ' + [...p.sources].join(', '));
    return bits.join(' · ');
  }

  // ---------- parsing ----------
  function norm(s){ return String(s==null?'':s).trim(); }
  function key(s){ return norm(s).toLowerCase(); }

  function parseTab(t){
    const rows = t.rows.filter(r => r.some(c => norm(c) !== ''));
    let header = null, headerRow = -1, start = 0;
    // look for a header row anywhere in the first 20 non-empty rows (title rows, notes, blank rows above are fine)
    for (let i = 0; i < Math.min(rows.length, 20); i++) {
      const h = rows[i].map(c => key(c));
      const joined = h.join(' ');
      if (joined.includes('table') && joined.includes('column')) { header = h; headerRow = i; start = i + 1; break; }
    }
    const idx = {};
    let layout = '';
    if (header) {
      const find = (...pats) => header.findIndex(c => c && pats.some(p => p.test(c)));
      idx.src = find(/^data_source_id$/, /source_id/, /^source$/, /data ?source/);
      idx.tbl = find(/^source_table$/, /table_name/, /^table$/, /table/);
      idx.col = find(/^column_name$/, /^column$/, /field_name/, /column/);
      idx.ex  = find(/^example_values$/, /example/, /sample/);
      MARKUP.forEach((m,i) => { idx['m'+i] = header.findIndex(c => c === m.toLowerCase()); });
      if (idx.m7 < 0) idx.m7 = header.findIndex(c => c.includes('transform') && c.includes('code'));
      if (idx.m3 < 0) idx.m3 = header.findIndex(c => c.includes('transform') && !c.includes('code'));
      layout = 'header in row ' + (headerRow + 1);
    } else {
      // no header: decide whether column A is a source id or already the table name
      const body = rows.slice(start);
      const widths = body.map(r => r.filter(c => norm(c) !== '').length);
      const maxW = widths.reduce((max, width) => Math.max(max, width), 0);
      const colA = new Set(body.map(r => key(r[0])).filter(Boolean));
      const aLooksLikeSource = colA.size <= 3 && body.length > 3;
      if (maxW <= 3 && !aLooksLikeSource) { idx.src = -1; idx.tbl = 0; idx.col = 1; idx.ex = 2; layout = 'no header, columns A–C = table, column, samples'; }
      else { idx.src = 0; idx.tbl = 1; idx.col = 2; idx.ex = 3; layout = 'no header, columns A–D = source, table, column, samples'; }
      MARKUP.forEach((m,i) => { idx['m'+i] = idx.src === 0 ? 4+i : -1; });
    }
    const hasMarkup = MARKUP.some((m,i) => idx['m'+i] >= 0 && rows.slice(start).some(r => norm(r[idx['m'+i]]) !== ''));
    const out = [], tables = new Set(), sources = new Set(), dupes = []; let bad = 0;
    const seen = new Set();
    for (const r of rows.slice(start)) {
      const rec = {
        data_source_id: idx.src >= 0 ? key(r[idx.src]) : '', source_table: key(r[idx.tbl]), column_name: key(r[idx.col]),
        example_values: idx.ex >= 0 ? norm(r[idx.ex]) : '',
      };
      MARKUP.forEach((m,i) => { rec[m] = idx['m'+i] >= 0 ? norm(r[idx['m'+i]]) : ''; });
      if (!rec.source_table || !rec.column_name) { bad++; continue; }
      const k = rec.source_table + '|' + rec.column_name;
      if (seen.has(k)) { dupes.push(k); continue; }
      seen.add(k);
      tables.add(rec.source_table); if (rec.data_source_id) sources.add(rec.data_source_id);
      out.push(rec);
    }
    const problem = !out.length ? (rows.length ? 'no usable rows: could not find table and column names' : 'empty')
                  : (idx.tbl < 0 || idx.col < 0) ? 'header found but no table/column columns' : '';
    return {rows: out, header: !!header, headerRow, layout, hasMarkup, tables, sources, dupes, bad, problem};
  }

  // ---------- diff ----------
  $('runBtn').addEventListener('click', run);
  function resetInputs(){
    loadGeneration++;
    for (const cancel of [...activeParsers]) cancel();
    files = []; rebuild(); setMsg('loadMsg', null); $('runBtn').disabled = false; hide('cancelLoadBtn');
  }
  $('resetBtn').addEventListener('click', resetInputs);
  $('cancelLoadBtn').addEventListener('click', resetInputs);

  function run(){
    const b = tabs[$('baseSel').value], n = tabs[$('newSel').value];
    const errs = [];
    if (!b) errs.push('Pick a baseline tab.');
    if (!n) errs.push('Pick a new-dictionary tab.');
    if (b && n && b === n) errs.push('Baseline and new are the same tab. Pick two different ones.');
    if (b && !b.parsed.rows.length) errs.push('The baseline tab has no usable rows (needs table and column in every row).');
    if (n && !n.parsed.rows.length) errs.push('The new tab has no usable rows (needs table and column in every row).');
    if (errs.length) { setMsg('pickMsg', {kind:'err', text:'Can’t compare yet', items:errs}); return; }
    setMsg('pickMsg', null);

    let base = b.parsed.rows, nw = n.parsed.rows;
    const notes = [];
    // source scoping: if the new DD names a source, only compare against that source in baseline
    const newSrc = n.parsed.sources, baseSrc = b.parsed.sources;
    if (newSrc.size && baseSrc.size > 1) {
      base = base.filter(r => newSrc.has(r.data_source_id));
      notes.push(`Baseline narrowed to source${newSrc.size>1?'s':''} ${[...newSrc].join(', ')} (${base.length} rows) because the new file only has that source.`);
      if (!base.length) notes.push('No baseline rows match that source, so everything will show as new. Check the data_source_id column.');
    } else if (!newSrc.size && baseSrc.size > 1) {
      notes.push(`The new file has no data_source_id, and the baseline has ${baseSrc.size} sources (${[...baseSrc].join(', ')}). Comparing against all of them.`);
    }
    if (n.parsed.bad) notes.push(`${n.parsed.bad} row${n.parsed.bad>1?'s':''} in the new file skipped (missing table or column name).`);
    if (b.parsed.bad) notes.push(`${b.parsed.bad} row${b.parsed.bad>1?'s':''} in the baseline skipped (missing table or column name).`);
    if (n.parsed.dupes.length) notes.push(`${n.parsed.dupes.length} duplicate table/column pair${n.parsed.dupes.length>1?'s':''} in the new file ignored (e.g. ${n.parsed.dupes[0]}).`);
    if (b.parsed.dupes.length) notes.push(`${b.parsed.dupes.length} duplicate pair${b.parsed.dupes.length>1?'s':''} in the baseline ignored.`);
    if (!n.parsed.header) notes.push('New file: ' + n.parsed.layout + '.');
    if (!b.parsed.header) notes.push('Baseline: ' + b.parsed.layout + '.');

    const bTables = new Set(base.map(r => r.source_table));
    const bKeys = new Map(base.map(r => [r.source_table+'|'+r.column_name, r]));
    const nTables = new Set(nw.map(r => r.source_table));
    const nKeys = new Set(nw.map(r => r.source_table+'|'+r.column_name));

    const all = nw.map(r => {
      const k = r.source_table+'|'+r.column_name;
      let st = 'EXISTS';
      if (!bTables.has(r.source_table)) st = 'NEW TABLE';
      else if (!bKeys.has(k)) st = 'NEW COLUMN';
      const o = {...r, Status: st};
      if (st === 'EXISTS') { const br = bKeys.get(k); MARKUP.forEach(m => { o[m] = br[m]; }); if (!o.data_source_id) o.data_source_id = br.data_source_id; }
      return o;
    });
    const onlyNew = all.filter(r => r.Status !== 'EXISTS');
    const removed = base.filter(r => !nKeys.has(r.source_table+'|'+r.column_name))
      .map(r => ({...r, Status: nTables.has(r.source_table) ? 'REMOVED COLUMN' : 'REMOVED TABLE'}));

    const srcLabel = ([...newSrc][0] || [...baseSrc][0] || 'dd');
    result = {all, onlyNew, removed, notes, src: srcLabel, baseName: b.file+' › '+b.sheet, newName: n.file+' › '+n.sheet,
              baseCount: base.length, newCount: nw.length, hasMarkup: b.parsed.hasMarkup};
    if (!n.parsed.rows.length) return;
    if (removed.length === base.length && base.length) notes.push('Every baseline row shows as removed. The two tabs may be from different sources, or table names are spelled differently.');
    render();
  }

  // ---------- render ----------
  function render(){
    const r = result;
    show('results');
    const c = s => r.all.filter(x => x.Status === s).length;
    $('stats').innerHTML = `
      <div class="stat nt"><b>${c('NEW TABLE')}</b><span>new table rows</span></div>
      <div class="stat nc"><b>${c('NEW COLUMN')}</b><span>new columns</span></div>
      <div class="stat rm"><b>${r.removed.length}</b><span>removed</span></div>
      <div class="stat ex"><b>${c('EXISTS')}</b><span>already in baseline</span></div>
      <div class="stat plain"><b>${r.baseCount}</b><span>baseline rows</span></div>
      <div class="stat plain"><b>${r.newCount}</b><span>new-file rows</span></div>`;
    const newTables = [...new Set(r.onlyNew.filter(x=>x.Status==='NEW TABLE').map(x=>x.source_table))];
    const newColTables = [...new Set(r.onlyNew.filter(x=>x.Status==='NEW COLUMN').map(x=>x.source_table))];
    const remTables = [...new Set(r.removed.filter(x=>x.Status==='REMOVED TABLE').map(x=>x.source_table))];
    const items = [...r.notes];
    if (newTables.length) items.unshift('New tables: ' + newTables.join(', '));
    if (newColTables.length) items.unshift('Tables with new columns: ' + newColTables.join(', '));
    if (remTables.length) items.unshift('Removed tables: ' + remTables.join(', '));
    setMsg('resMsg', items.length ? {kind: r.notes.length ? 'warn' : 'ok', text:'', items} : null);
    renderTable();
    updateDlHint();
    $('results').scrollIntoView({behavior:'smooth', block:'start'});
  }

  $('viewTabs').querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
    view = b.dataset.v; $('viewTabs').querySelectorAll('button').forEach(x => x.classList.toggle('on', x===b)); renderTable();
  }));

  function currentRows(){
    if (view === 'new') return result.onlyNew;
    if (view === 'removed') return result.removed;
    return result.all;
  }
  function columnsFor(){
    const cols = [...KEY, 'Status'];
    if (result.hasMarkup || view !== 'new') cols.push(...MARKUP);
    return cols;
  }
  function renderTable(){
    const rows = currentRows();
    if (!rows.length) { $('tableWrap').innerHTML = `<div class="empty">${view==='new' ? 'Nothing new. Every table and column in the new file already exists in the baseline.' : view==='removed' ? 'Nothing removed.' : 'No rows.'}</div>`; return; }
    const cols = view === 'new' ? [...KEY, 'Status'] : [...KEY, 'Status', ...MARKUP];
    const LIMIT = 500;
    let h = '<table><thead><tr>' + cols.map(c => `<th>${esc(c)}</th>`).join('') + '</tr></thead><tbody>';
    for (const r of rows.slice(0, LIMIT)) {
      h += `<tr class="${r.Status.replace(' ','_')}">` + cols.map(c => c==='Status' ? `<td class="st"><span>${esc(r[c])}</span></td>` : `<td title="${esc(r[c])}">${esc(r[c])}</td>`).join('') + '</tr>';
    }
    h += '</tbody></table>';
    if (rows.length > LIMIT) h += `<div class="empty">Showing first ${LIMIT} of ${rows.length}. The download has everything.</div>`;
    $('tableWrap').innerHTML = h;
  }

  // ---------- export ----------
  function buildWorkbook(){
    const r = result;
    const wb = XLSX.utils.book_new();
    const stamp = new Date();
    const pad = n => String(n).padStart(2,'0');
    const date = `${stamp.getFullYear()}-${pad(stamp.getMonth()+1)}-${pad(stamp.getDate())}`;
    const time = `${pad(stamp.getHours())}${pad(stamp.getMinutes())}`;
    const c = s => r.all.filter(x => x.Status === s).length;
    const summary = [
      ['Baseline', r.baseName], ['New file', r.newName], ['Run at', `${date} ${time.slice(0,2)}:${time.slice(2)}`], [],
      ['Baseline rows', r.baseCount], ['New-file rows', r.newCount], [],
      ['NEW TABLE rows', c('NEW TABLE')], ['NEW COLUMN rows', c('NEW COLUMN')], ['EXISTS rows', c('EXISTS')], ['REMOVED rows', r.removed.length], [],
      ['New tables', [...new Set(r.onlyNew.filter(x=>x.Status==='NEW TABLE').map(x=>x.source_table))].join(', ') || '-'],
      ['Tables with new columns', [...new Set(r.onlyNew.filter(x=>x.Status==='NEW COLUMN').map(x=>x.source_table))].join(', ') || '-'],
      ['Removed tables', [...new Set(r.removed.filter(x=>x.Status==='REMOVED TABLE').map(x=>x.source_table))].join(', ') || '-'],
      [], ['Notes'], ...r.notes.map(n => [n]),
    ];
    const ws0 = XLSX.utils.aoa_to_sheet(summary); ws0['!cols'] = [{wch:26},{wch:90}];
    XLSX.utils.book_append_sheet(wb, ws0, 'Summary');
    const cols = [...KEY, 'Status', ...MARKUP];
    const toSheet = rows => {
      const ws = XLSX.utils.json_to_sheet(rows.map(x => Object.fromEntries(cols.map(k => [k, x[k] ?? '']))), {header: cols});
      ws['!cols'] = cols.map(k => ({wch: Math.min(60, Math.max(12, ...[k, ...rows.slice(0,200).map(x => String(x[k]??''))].map(s => s.length+2)))}));
      ws['!autofilter'] = {ref: XLSX.utils.encode_range({s:{c:0,r:0}, e:{c:cols.length-1, r:Math.max(rows.length,1)}})};
      return ws;
    };
    XLSX.utils.book_append_sheet(wb, toSheet(r.onlyNew), 'New');
    XLSX.utils.book_append_sheet(wb, toSheet(r.removed), 'Removed');
    XLSX.utils.book_append_sheet(wb, toSheet(r.all), 'AllNew');
    const name = `${cap(r.src)}_DD_Diff_${date}_${time}_${r.onlyNew.length}new_${r.removed.length}removed.xlsx`;
    return {wb, name};
  }

  $('dlBtn').addEventListener('click', async () => {
    if (!result) return;
    let wb, name, data;
    try {
      ({wb, name} = buildWorkbook());
      data = XLSX.write(wb, {type:'array', bookType:'xlsx'});
    } catch(e) {
      $('dlHint').textContent = `Could not create Excel download: ${e.message || e}. Your comparison is still available; try copying the worklist.`;
      return;
    }
    if (!downloads) {
      // outside the Claude viewer a plain download works
      try {
        const a = document.createElement('a');
        a.href = URL.createObjectURL(new Blob([data], {type:'application/octet-stream'}));
        a.download = name; document.body.appendChild(a); a.click(); a.remove();
        setTimeout(() => URL.revokeObjectURL(a.href), 2000);
        $('dlHint').textContent = `Saved as ${name} (if nothing happened, use Copy worklist as TSV).`;
      } catch(e) { $('dlHint').textContent = 'Download isn’t available here. Use Copy worklist as TSV.'; }
      return;
    }
    try {
      await downloads.save({filename: name, data: new Blob([data])});
      $('dlHint').textContent = `Saved ${name}`;
    } catch(e) {
      const code = e && e.code;
      $('dlHint').textContent = code === 'declined' ? 'Download cancelled.' : code === 'rate_limited' ? 'A download prompt is already open. Try again in a moment.' : 'Download isn’t available here. Use Copy worklist as TSV.';
    }
  });

  $('copyBtn').addEventListener('click', async () => {
    if (!result) return;
    const cols = [...KEY, 'Status'];
    const tsv = [cols.map(clipboardCell).join('\t'), ...result.onlyNew.map(r => cols.map(c => clipboardCell(r[c])).join('\t'))].join('\n');
    try { await navigator.clipboard.writeText(tsv); $('dlHint').textContent = `Copied ${result.onlyNew.length} rows. Formula-like values are escaped; line breaks are replaced with spaces. Use Excel download to preserve original values.`; }
    catch(e) { $('dlHint').textContent = 'Clipboard blocked by the browser. Select the table and copy it instead.'; }
  });

  // ---------- nav ----------
  $('nav').querySelectorAll('a').forEach(a => a.addEventListener('click', e => {
    e.preventDefault();
    $('nav').querySelectorAll('a').forEach(x => x.classList.toggle('on', x === a));
    $('view-tool').classList.toggle('hidden', a.dataset.sec !== 'tool');
    $('view-help').classList.toggle('hidden', a.dataset.sec !== 'help');
    window.scrollTo({top:0});
  }));

  // ---------- helpers ----------
  function clipboardCell(value){
    let text = String(value ?? '').replace(/[\u0000-\u001F\u007F-\u009F\u2028\u2029]/g, ' ')
      .replace(/[\u200B-\u200F\u202A-\u202E\u2060-\u206F]/g, '');
    if (/^[\s\uFEFF]*[=+\-@＝＋－＠]/u.test(text)) text = "'" + text;
    return '"' + text.replace(/"/g, '""') + '"';
  }
  function esc(s){ return String(s==null?'':s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
  function cap(s){ return s ? s.charAt(0).toUpperCase()+s.slice(1) : s; }
  function show(id){ $(id).classList.remove('hidden'); }
  function hide(id){ $(id).classList.add('hidden'); }
  function toggle(id, on){ $(id).classList.toggle('hidden', !on); }
  function setMsg(id, m){
    const el = $(id);
    if (!m) { el.innerHTML = ''; return; }
    const cls = m.kind === 'err' ? 'err' : m.kind === 'warn' ? 'warn' : 'info';
    el.innerHTML = `<div class="msg ${cls}">${m.text ? esc(m.text) : ''}${m.items && m.items.length ? '<ul>' + m.items.map(i => `<li>${esc(i)}</li>`).join('') + '</ul>' : ''}</div>`;
  }
})();
