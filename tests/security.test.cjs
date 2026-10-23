const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {createHash} = require('node:crypto');
const root = path.resolve(__dirname, '..');
const XLSX = require('../vendor/xlsx-0.20.3.full.min.js');
const appSource = fs.readFileSync(path.join(root, 'app.js'), 'utf8');

function app(overrides = {}) {
  const elements = new Map();
  const get = id => {
    if (!elements.has(id)) elements.set(id, {
      innerHTML: '', textContent: '', value: '', disabled: false, handlers: {},
      classList: {add(){},remove(){},toggle(){}},
      addEventListener(name, fn){ this.handlers[name] = fn; },
      querySelectorAll(){ return []; }, scrollIntoView(){},
      replaceChildren(){ this.innerHTML = ''; this.textContent = ''; this.value = ''; },
    });
    return elements.get(id);
  };
  const context = vm.createContext({window: {}, document: {getElementById:get}, XLSX, setTimeout, clearTimeout, ...overrides});
  const expose = `globalThis.review = {parseTab, clipboardCell, buildWorkbook, parseFile, addFiles,
    setTabs(value){tabs=value;}, setResult(value){result=value;}, getResult(){return result;}, getFiles(){return files;}, run};`;
  vm.runInContext(appSource.replace(/\}\)\(\);\s*$/, `${expose}\n})();`), context);
  return {api:context.review, get, context};
}

function worker(buffer) {
  let result;
  const self = {postMessage(value){ result = value; }};
  const context = vm.createContext({self, XLSX, importScripts(){}, DataView, Number, Error, JSON});
  vm.runInContext(fs.readFileSync(path.join(root, 'parser-worker.js'), 'utf8'), context);
  context.dictionaryParserWorker('unused-in-unit-tests');
  self.onmessage({data:buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)});
  return JSON.parse(JSON.stringify(result));
}

function compare(baseRows, newRows) {
  const harness = app();
  const tabs = [{rows:baseRows,file:'baseline.csv',sheet:'Baseline'}, {rows:newRows,file:'new.csv',sheet:'New'}];
  for (const tab of tabs) tab.parsed = harness.api.parseTab(tab);
  harness.api.setTabs(tabs);
  harness.get('baseSel').value = '0'; harness.get('newSel').value = '1';
  harness.api.run();
  return {result:JSON.parse(JSON.stringify(harness.api.getResult())), ...harness};
}

test('patched parser, no remote assets or inline executable code, and matching integrity hashes', () => {
  assert.equal(XLSX.version, '0.20.3');
  const html = fs.readFileSync(path.join(root, 'index.html'), 'utf8');
  const css = fs.readFileSync(path.join(root, 'styles.css'), 'utf8');
  assert.match(html, /connect-src 'none'/);
  assert.match(html, /worker-src blob:/);
  assert.match(html, /form-action 'none'/);
  assert.doesNotMatch(html, /unsafe-inline|unsafe-eval|\bon\w+\s*=|<script>\s*\S|(?:src|href)="https?:/i);
  assert.doesNotMatch(css, /url\(['"]?https?:|@import/i);
  const tags = html.match(/<(?:script|link)\b[^>]*>/g);
  let checked = 0;
  for (const tag of tags) {
    const resource = tag.match(/(?:src|href)="([^"]+\.(?:js|css))"/);
    if (!resource) continue;
    const expected = createHash('sha384').update(fs.readFileSync(path.join(root, resource[1]))).digest('base64');
    assert.ok(tag.includes(`integrity="sha384-${expected}"`), resource[1]); checked++;
  }
  assert.equal(checked, 4);
});

test('comparison preserves normalization, new/removed statuses, and baseline markup', () => {
  const header = ['data_source_id','source_table','column_name','example_values','risk_type','Notes'];
  const {result,api} = compare(
    [header,[' Demo ',' Users ',' ID ','baseline','QI','baseline note'],['demo','users','old','x','','removed'],['demo','gone','id','x','','']],
    [header,['DEMO','USERS','ID','new sample','incoming','incoming note'],['demo','users','new','x','',''],['demo','added','id','x','','']]
  );
  assert.deepEqual(result.all.map(r=>[r.source_table,r.column_name,r.Status]), [['users','id','EXISTS'],['users','new','NEW COLUMN'],['added','id','NEW TABLE']]);
  assert.deepEqual(result.removed.map(r=>[r.source_table,r.column_name,r.Status]), [['users','old','REMOVED COLUMN'],['gone','id','REMOVED TABLE']]);
  assert.equal(result.all[0].risk_type,'QI'); assert.equal(result.all[0].Notes,'baseline note');
  assert.equal(result.all[0].example_values,'new sample');
  const wb = api.buildWorkbook().wb;
  assert.deepEqual(Array.from(wb.SheetNames),['Summary','New','Removed','AllNew']);
  assert.equal(XLSX.utils.sheet_to_json(wb.Sheets.AllNew)[0].Notes,'baseline note');
});

test('comparison preserves source scoping, title-row detection, duplicates and incomplete-row handling', () => {
  const header = ['source','table','column','samples','Notes'];
  const {result} = compare(
    [['Title'],[],header,['alpha','users','id','1','alpha markup'],['beta','orders','id','2','beta markup']],
    [header,['beta','orders','id','3',''],['beta','orders','id','duplicate',''],['beta','','missing','x','']]
  );
  assert.equal(result.baseCount,1); assert.equal(result.newCount,1);
  assert.equal(result.all[0].Status,'EXISTS'); assert.equal(result.all[0].Notes,'beta markup');
  assert.equal(result.removed.length,0);
  assert.ok(result.notes.some(note=>note.includes('duplicate')));
  assert.ok(result.notes.some(note=>note.includes('skipped')));
});

test('headerless layouts retain positional matching and handle 150,000 rows without argument overflow', () => {
  const {api} = app();
  const three = api.parseTab({rows:[['users','id','1'],['orders','id','2']]});
  assert.equal(three.rows[0].source_table,'users'); assert.equal(three.rows[0].column_name,'id');
  const four = api.parseTab({rows:[['demo','users','id','1'],['demo','users','name','x'],['demo','orders','id','2'],['demo','orders','total','3']]});
  assert.equal(four.rows[0].data_source_id,'demo'); assert.equal(four.rows[0].source_table,'users');
  const large = api.parseTab({rows:Array.from({length:150000},(_,i)=>['users',`field_${i}`,'sample'])});
  assert.equal(large.rows.length,150000);
});

test('clipboard encodes formulas, Unicode variants, quotes and row/cell separators; workbook strings stay unchanged', () => {
  const {api} = app();
  const dangerous = ['=1+1','+1+1','-1+1','@SUM(1)','  =1+1','\uFEFF=1+1','＝1+1','＋1','－1','＠SUM(1)','\n=1+1','\u0000=1+1','\u200B=1+1','\u202E=1+1'];
  for (const value of dangerous) {
    const field = api.clipboardCell(value);
    const parsed = XLSX.utils.sheet_to_json(XLSX.read(`v\n${field}`, {type:'string',raw:true}).Sheets.Sheet1,{header:1});
    assert.ok(parsed[1][0].startsWith("'"), value);
    assert.doesNotMatch(field, /[\r\n\t]/);
  }
  assert.equal(api.clipboardCell('a"b\tc\nd'),'"a""b c d"');
  const {api:exportApi} = compare([['table','column','sample'],['users','id','1']], [['table','column','sample'],['users','new','=1+1'],['users','other','safe\n=1+1']]);
  const saved = XLSX.read(XLSX.write(exportApi.buildWorkbook().wb,{type:'buffer',bookType:'xlsx'}),{type:'buffer',cellFormula:true});
  assert.equal(saved.Sheets.New.D2.t,'s'); assert.equal(saved.Sheets.New.D2.v,'=1+1'); assert.equal(saved.Sheets.New.D2.f,undefined);
  assert.equal(saved.Sheets.New.D3.v,'safe\n=1+1');
});

test('worker reads CSV and XLSX cached values without carrying formula or hyperlink metadata', () => {
  const csv = worker(Buffer.from('source_table,column_name,example_values\nusers,new,=1+1\n'));
  assert.equal(csv.sheets[0].rows[1][2],'=1+1');
  const wb=XLSX.utils.book_new(); const sheet=XLSX.utils.aoa_to_sheet([['table','column','sample'],['users','calc',2],['users','link','text']]);
  sheet.C2.f='1+1'; sheet.C3.l={Target:'https://example.invalid/'}; XLSX.utils.book_append_sheet(wb,sheet,'New');
  const parsed=worker(XLSX.write(wb,{type:'buffer',bookType:'xlsx'}));
  assert.equal(parsed.sheets[0].rows[1][2],'2'); assert.equal(parsed.sheets[0].rows[2][2],'text');
  assert.doesNotMatch(JSON.stringify(parsed),/example\.invalid|"f":|"l":/);
});

test('worker rejects oversized dimensions and expanded ZIPs before returning partial data', () => {
  const wb=XLSX.utils.book_new(); const sheet=XLSX.utils.aoa_to_sheet([['table','column'],['users','id']]);
  sheet['!ref']='A1:B250002'; XLSX.utils.book_append_sheet(wb,sheet,'Large');
  const large=worker(XLSX.write(wb,{type:'buffer',bookType:'xlsx'}));
  assert.match(large.error,/250,000/); assert.equal(large.sheets,undefined);
  const small=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(small,XLSX.utils.aoa_to_sheet([['table','column'],['users','id']]),'Data');
  const zipped=Buffer.from(XLSX.write(small,{type:'buffer',bookType:'xlsx',compression:true}));
  const central=zipped.indexOf(Buffer.from([0x50,0x4b,0x01,0x02])); assert.ok(central>0);
  zipped.writeUInt32LE(101*1024*1024,central+24);
  assert.match(worker(zipped).error,/100 MB/);
});

test('Excel format limits produce a visible error instead of an uncaught rejection', async () => {
  const harness = compare([['table','column','sample'],['users','id','1']], [['table','column','sample'],['users','new','x'.repeat(40000)]]);
  await harness.get('dlBtn').handlers.click();
  assert.match(harness.get('dlHint').textContent,/Could not create Excel download/);
  assert.equal(harness.api.getResult().onlyNew.length,1);
});

test('oversized files are rejected before reading; parser timeout terminates worker', async () => {
  const harness=app();
  await harness.api.addFiles([{name:'oversize.csv',size:21*1024*1024,arrayBuffer(){throw new Error('must not read');}}],0);
  assert.match(harness.get('loadMsg').innerHTML,/20 MB/); assert.equal(harness.api.getFiles().length,0);
  let timeout,terminated=false;
  const slow=app({createDictionaryParserWorker:()=>({terminate(){terminated=true;},postMessage(){}}),setTimeout:fn=>{timeout=fn;return 1;},clearTimeout(){}});
  const pending=slow.api.parseFile({arrayBuffer:async()=>new ArrayBuffer(1)});
  timeout();
  await assert.rejects(pending,/30 seconds/); assert.equal(terminated,true);
});
