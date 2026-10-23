const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const crypto = require('node:crypto');
const assert = require('node:assert/strict');
const { spawn } = require('node:child_process');
const http = require('node:http');
const repoRoot = path.resolve(__dirname, '..');
const XLSX = require('../vendor/xlsx-0.20.3.full.min.js');

const root = fs.mkdtempSync(path.join(os.tmpdir(), 'dd-security-review-'));
const downloads = path.join(root, 'downloads');
fs.mkdirSync(downloads);
const output = { evidenceDirectory: root, tests: [], requests: [], errors: [] };
const serverRequests = [];
const server = http.createServer((request, response) => {
  serverRequests.push({method:request.method,url:request.url});
  let pathname;
  try { pathname = decodeURIComponent(new URL(request.url,'http://localhost').pathname); }
  catch { response.writeHead(400); response.end(); return; }
  if (!pathname.startsWith('/data-dictionary-diff/')) { response.writeHead(404); response.end(); return; }
  const relative = pathname.slice('/data-dictionary-diff/'.length) || 'index.html';
  const file = path.resolve(repoRoot, relative);
  if (!file.startsWith(repoRoot + path.sep) || relative.split('/').some(part=>part.startsWith('.'))) { response.writeHead(403); response.end(); return; }
  const types={'.html':'text/html','.js':'text/javascript','.css':'text/css','.woff2':'font/woff2'};
  try { const data=fs.readFileSync(file); response.writeHead(200,{'Content-Type':types[path.extname(file)]||'application/octet-stream'}); response.end(data); }
  catch { response.writeHead(404); response.end(); }
});
const pause = ms => new Promise(resolve => setTimeout(resolve, ms));
async function until(check, timeout = 20000) {
  const end = Date.now() + timeout;
  while (Date.now() < end) {
    const v = await check();
    if (v) return v;
    await pause(100);
  }
  throw new Error('Timed out');
}
const chrome = spawn('/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', [
  '--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check',
  '--disable-extensions', '--disable-background-networking',
  `--user-data-dir=${path.join(root, 'chrome-profile')}`, '--remote-debugging-port=0', 'about:blank',
], { stdio: ['ignore', 'ignore', 'pipe'] });
let chromeStderr = '';
chrome.stderr.on('data', b => { chromeStderr += b; });
let socket;
let sequence = 0;
const pending = new Map();
let phase = 'startup';
async function main() {
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  const origin=`http://127.0.0.1:${server.address().port}`;
  output.url=origin+'/data-dictionary-diff/';
  const activePort = path.join(root, 'chrome-profile', 'DevToolsActivePort');
  await until(() => fs.existsSync(activePort));
  const port = fs.readFileSync(activePort, 'utf8').split('\n')[0];
  const targets = await fetch(`http://127.0.0.1:${port}/json/list`).then(r => r.json());
  socket = new WebSocket(targets.find(t => t.type === 'page').webSocketDebuggerUrl);
  await new Promise((resolve, reject) => { socket.addEventListener('open', resolve, { once: true }); socket.addEventListener('error', reject, { once: true }); });
  socket.addEventListener('message', e => {
    const m = JSON.parse(e.data);
    if (m.id) {
      const p = pending.get(m.id);
      if (!p) return;
      pending.delete(m.id);
      m.error ? p.reject(new Error(JSON.stringify(m.error))) : p.resolve(m.result);
    } else if (m.method === 'Network.requestWillBeSent') {
      output.requests.push({phase, url: m.params.request.url, method: m.params.request.method, type: m.params.type});
    } else if (m.method === 'Runtime.exceptionThrown') {
      output.errors.push({phase, message: m.params.exceptionDetails.exception?.description || m.params.exceptionDetails.text});
    }
  });
  const cdp = (method, params = {}) => new Promise((resolve, reject) => {
    const id = ++sequence;
    pending.set(id, {resolve, reject});
    socket.send(JSON.stringify({id, method, params}));
  });
  const evaluate = async expression => {
    const result = await cdp('Runtime.evaluate', {expression, awaitPromise: true, returnByValue: true});
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.exception?.description || result.exceptionDetails.text);
    return result.result?.value;
  };
  await cdp('Runtime.enable');
  await cdp('Network.enable');
  await cdp('Page.enable');
  await cdp('Browser.setDownloadBehavior', {behavior: 'allow', downloadPath: downloads});
  await cdp('Page.navigate', {url: output.url});
  await until(() => evaluate('document.readyState === "complete" && typeof XLSX === "object" && !!document.getElementById("fileInput")'));
  output.libraryVersion = await evaluate('XLSX.version');
  assert.equal(output.libraryVersion, '0.20.3');
  output.tests.push({name: 'Local app loaded with patched parser, scripts and CSS integrity checks', result: 'pass'});
  await pause(1500);
  const marker = '<img src=x onerror=window.__ddSecurityXss=1>';
  const filename = 'baseline-"><img src=x onerror=window.__ddSecurityXss=1>.csv';
  const baseline = 'data_source_id,source_table,column_name,example_values,Notes\ndemo,users,id,1,baseline-note\ndemo,users,email,example@example.invalid,baseline-email-note\n';
  const incoming = 'data_source_id,source_table,column_name,example_values\ndemo,users,id,1\ndemo,users,email,' + marker + '\ndemo,users,new_field,=1+1\ndemo,users,multiline,"safe\n=1+1"\ndemo,users,html,' + marker + '\n';
  phase = 'csv-import-compare';
  await evaluate(`(() => { const dt = new DataTransfer(); dt.items.add(new File([${JSON.stringify(baseline)}], ${JSON.stringify(filename)}, {type:'text/csv'})); dt.items.add(new File([${JSON.stringify(incoming)}], 'new.csv', {type:'text/csv'})); const input=document.getElementById('fileInput'); input.files=dt.files; input.dispatchEvent(new Event('change', {bubbles:true})); })()`);
  await until(() => evaluate('document.getElementById("baseSel").options.length === 3 && document.getElementById("newSel").value !== "" && !document.getElementById("runBtn").disabled'));
  await evaluate('document.getElementById("runBtn").click()');
  const ui = await evaluate(`({stats:document.getElementById('stats').innerText, xss:!!window.__ddSecurityXss, insertedImages:document.querySelectorAll('#tableWrap img, #fileList img, #readoutTable img, #resMsg img').length, text:document.getElementById('tableWrap').innerText})`);
  assert.equal(ui.xss, false);
  assert.equal(ui.insertedImages, 0);
  assert.ok(ui.text.includes(marker));
  assert.ok(ui.stats.includes('3\nnew columns'));
  output.tests.push({name:'CSV import and comparison; HTML payload in filename/cells', result:'pass', evidence:ui});
  phase = 'clipboard';
  await evaluate(`Object.defineProperty(navigator.clipboard, 'writeText', {configurable:true,value:async text => {window.__ddCopiedText=text;}}); document.getElementById('copyBtn').click()`);
  const copied = await until(() => evaluate('window.__ddCopiedText'));
  assert.ok(copied.includes('"\'=1+1"'));
  assert.ok(copied.includes('"safe =1+1"'));
  assert.equal(copied.split('\n').length,4);
  output.tests.push({name:'Clipboard escapes formula-like values and keeps each record on one line', result:'pass', evidence:copied});
  phase = 'xlsx-download';
  await evaluate('document.getElementById("dlBtn").click()');
  const download = await until(() => fs.readdirSync(downloads).find(f=>f.endsWith('.xlsx')));
  const workbook = XLSX.read(fs.readFileSync(path.join(downloads, download)), {type:'buffer',cellFormula:true});
  const sheet = workbook.Sheets.New;
  const cells = Object.entries(sheet).filter(([key])=>!key.startsWith('!')).map(([address,cell])=>({address,type:cell.t,value:cell.v,formula:cell.f}));
  assert.ok(cells.some(c=>c.value==='=1+1' && c.type==='s' && !c.formula));
  assert.equal(cells.filter(c=>c.formula).length,0);
  const all = XLSX.utils.sheet_to_json(workbook.Sheets.AllNew, {defval:''});
  assert.equal(all.find(r=>r.column_name==='id').Notes,'baseline-note');
  output.tests.push({name:'Real Excel download keeps formula-looking CSV values as text and baseline markup',result:'pass',evidence:cells});
  phase = 'xlsx-import';
  const fixture=XLSX.utils.book_new();
  const base=XLSX.utils.aoa_to_sheet([['source_table','column_name','example_values'],['users','id','1']]);
  const next=XLSX.utils.aoa_to_sheet([['source_table','column_name','example_values'],['users','id','1'],['users','calculated',2],['users','html',marker]]);
  next.C3={t:'n',f:'1+1',v:2};
  next.C4.l={Target:'https://example.invalid/should-not-be-requested'};
  XLSX.utils.book_append_sheet(fixture,base,'Baseline');
  XLSX.utils.book_append_sheet(fixture,next,'New');
  const bytes=Array.from(XLSX.write(fixture,{type:'buffer',bookType:'xlsx'}));
  await evaluate(`(() => {document.getElementById('resetBtn').click(); const dt=new DataTransfer(); dt.items.add(new File([new Uint8Array(${JSON.stringify(bytes)})], 'review.xlsx', {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})); const input=document.getElementById('fileInput');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await until(() => evaluate('document.getElementById("fileList").textContent.includes("review.xlsx") && document.getElementById("baseSel").options.length===3 && document.getElementById("newSel").value!=="" && !document.getElementById("runBtn").disabled'));
  await evaluate('document.getElementById("runBtn").click()');
  const xlsxResult=await evaluate(`({text:document.getElementById('tableWrap').innerText,xss:!!window.__ddSecurityXss,links:document.querySelectorAll('#tableWrap a').length})`);
  assert.ok(xlsxResult.text.includes('calculated'));
  assert.equal(xlsxResult.xss,false);
  assert.equal(xlsxResult.links,0);
  output.tests.push({name:'XLSX cached formula and hyperlink import',result:'pass',evidence:xlsxResult});
  await pause(1000);
  const importRequests=output.requests.filter(r=>r.phase!=='startup' && /^https?:/.test(r.url));
  const externalRequests=output.requests.filter(r=>/^https?:/.test(r.url) && new URL(r.url).origin!==origin);
  assert.equal(externalRequests.length,0);
  assert.ok(serverRequests.every(r=>r.method==='GET'));
  output.tests.push({name:'No external requests or uploads during startup/import/compare/copy/download',result:'pass',evidence:importRequests});

  phase='csp-containment';
  const containment=await evaluate(`(async()=>{
    window.__ddViolations=[]; document.addEventListener('securitypolicyviolation',e=>window.__ddViolations.push({directive:e.effectiveDirective,blocked:e.blockedURI}));
    let pageFetchBlocked=false;try{await fetch(${JSON.stringify(origin+'/should-block-page')});}catch{pageFetchBlocked=true;}
    const inline=document.createElement('script');inline.textContent='window.__ddInlineRan=true';document.body.appendChild(inline);
    const workerResult=await new Promise((resolve,reject)=>{const url=URL.createObjectURL(new Blob(['self.onmessage=async()=>{try{await fetch('+JSON.stringify(${JSON.stringify(origin+'/should-block-worker')})+');postMessage(false);}catch{postMessage(true);}}'],{type:'text/javascript'}));const w=new Worker(url);URL.revokeObjectURL(url);w.onmessage=e=>{w.terminate();resolve(e.data);};w.onerror=reject;w.postMessage(null);});
    return {pageFetchBlocked,workerFetchBlocked:workerResult,inlineScriptBlocked:!window.__ddInlineRan};
  })()`);
  assert.equal(containment.pageFetchBlocked,true); assert.equal(containment.workerFetchBlocked,true); assert.equal(containment.inlineScriptBlocked,true);
  assert.equal(serverRequests.some(r=>r.url.includes('should-block')),false);
  output.tests.push({name:'CSP blocks page and blob-worker fetches and injected inline scripts',result:'pass',evidence:containment});

  phase='headerless-size-test';
  await evaluate(`(() => {document.getElementById('resetBtn').click();const csv=Array.from({length:150000},(_,i)=>'users,field_'+i+',sample').join('\\n');window.__ddLargeBytes=new TextEncoder().encode(csv).length;window.__ddHeartbeats=0;window.__ddHeartbeatTimer=setInterval(()=>window.__ddHeartbeats++,20);const dt=new DataTransfer();dt.items.add(new File([csv],'headerless.csv',{type:'text/csv'}));const input=document.getElementById('fileInput');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await until(()=>evaluate('document.getElementById("fileList").textContent.includes("headerless.csv") && !document.getElementById("runBtn").disabled'),30000);
  const large=await evaluate(`({bytes:window.__ddLargeBytes,heartbeats:window.__ddHeartbeats,rows:document.getElementById('baseSel').textContent,errors:document.getElementById('loadMsg').textContent})`);
  await evaluate('clearInterval(window.__ddHeartbeatTimer)');
  assert.ok(large.rows.includes('150000 rows')); assert.equal(large.errors,''); assert.ok(large.heartbeats>0);
  output.tests.push({name:'Previously failing 3.8 MB/150,000-row headerless CSV succeeds with responsive worker parsing',result:'pass',evidence:large});

  phase='oversize-file';
  await evaluate(`(()=>{document.getElementById('resetBtn').click();const dt=new DataTransfer();dt.items.add(new File([new Uint8Array(20*1024*1024+1)],'oversize.csv'));const input=document.getElementById('fileInput');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await until(()=>evaluate('document.getElementById("loadMsg").textContent.includes("20 MB")'));
  assert.equal(await evaluate('document.getElementById("fileList").textContent'),'');
  output.tests.push({name:'Oversized file rejected visibly without loading it',result:'pass'});

  phase='timeout';
  await evaluate(`(()=>{window.__ddOriginalFactory=createDictionaryParserWorker;window.__ddOriginalTimeout=window.setTimeout;window.__ddTerminated=0;window.createDictionaryParserWorker=()=>{const url=URL.createObjectURL(new Blob(['self.onmessage=()=>{}'],{type:'text/javascript'}));const w=new Worker(url);URL.revokeObjectURL(url);const stop=w.terminate.bind(w);w.terminate=()=>{window.__ddTerminated++;stop();};return w;};window.setTimeout=(fn,ms,...args)=>window.__ddOriginalTimeout(fn,ms===30000?50:ms,...args);const dt=new DataTransfer();dt.items.add(new File(['table,column\\nusers,id'],'timeout.csv'));const input=document.getElementById('fileInput');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await until(()=>evaluate('document.getElementById("loadMsg").textContent.includes("30 seconds")'));
  assert.equal(await evaluate('window.__ddTerminated'),1);
  output.tests.push({name:'Worker timeout terminates stalled parser and reports error (test clock accelerated)',result:'pass'});
  await evaluate('window.setTimeout=window.__ddOriginalTimeout');

  phase='cancel';
  await evaluate(`(()=>{const dt=new DataTransfer();dt.items.add(new File(['table,column\\nusers,id'],'cancel.csv'));const input=document.getElementById('fileInput');input.files=dt.files;input.dispatchEvent(new Event('change',{bubbles:true}));})()`);
  await until(()=>evaluate('document.getElementById("runBtn").disabled'));
  assert.equal(await evaluate('document.getElementById("cancelLoadBtn").classList.contains("hidden")'),false);
  await evaluate('document.getElementById("cancelLoadBtn").click()');
  await pause(150);
  assert.equal(await evaluate('window.__ddTerminated'),2);
  assert.equal(await evaluate('document.getElementById("fileList").textContent'),'');
  assert.equal(await evaluate('document.getElementById("tableWrap").textContent'),'');
  assert.equal(await evaluate('document.getElementById("loadMsg").textContent'),'');
  await evaluate('window.createDictionaryParserWorker=window.__ddOriginalFactory');
  output.tests.push({name:'Start over cancels an active parser and clears previous rendered data',result:'pass'});
  assert.equal(output.errors.length,0);
  output.reviewedHtmlSha256=crypto.createHash('sha256').update(fs.readFileSync(path.join(repoRoot,'index.html'))).digest('hex');
  output.serverRequests=serverRequests;
  output.finished=true;
}
main().catch(error=>{output.failure=String(error.stack||error);output.chromeStderr=chromeStderr.slice(-3000);process.exitCode=1;}).finally(async()=>{
  fs.writeFileSync(path.join(root,'results.json'),JSON.stringify(output,null,2));
  console.log(JSON.stringify(output,null,2));
  if(socket)socket.close();
  chrome.kill('SIGTERM');
  server.close();
  setTimeout(()=>process.exit(process.exitCode||0),500);
});
