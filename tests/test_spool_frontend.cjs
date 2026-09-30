const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
function setup() {
  class Element {
    constructor() { this.listeners = {}; this.children = []; this.style = {}; this.value = ''; this.textContent = ''; this.classList = {add(){}, remove(){}}; }
    set innerHTML(v) { this.html = v; this.children = []; }
    get innerHTML() { return this.html || ''; }
    setAttribute(k, v) { this[k]=v; }
    set textContent(v) { this.text=v; this.children=[]; }
    get textContent() { return this.text || ""; }
    addEventListener(e, f) { this.listeners[e] = f; }
    appendChild(e) { this.children.push(e); }
    querySelector() { return this.remove ||= new Element(); }
    querySelectorAll() { return []; }
  }
  const elements = new Map();
  const el = id => { if (!elements.has(id)) elements.set(id, new Element()); return elements.get(id); };
  let complete, reject;
  const requests = [];
  const context = vm.createContext({document: {getElementById: el, querySelector: el, querySelectorAll: () => [], createElement: () => new Element()},
    FormData, runBOM(){}, syncRunProjectUI(){}, escHtml: String, showToast(){},
    fetchJson(url, options) { requests.push(options.body); return new Promise((yes,no) => {complete=yes; reject=no;}); }});
  const source = fs.readFileSync(path.join(__dirname,'../FabBOMTool/static/app.js'),'utf8');
  vm.runInContext(source.slice(0,source.indexOf('async function runBOM()')),context);
  return {el, requests, select(files) {el('spool-file-input').files=files; el('spool-file-input').listeners.change();},
    run: () => el('spool-run-btn').listeners.click(),
    finish: name => complete({rows:[{'Fab Number':name}],row_count:1,diagnostics:[],output_filename:name+'.xlsx'}),
    respond: data => complete(data),
    fail: () => reject(new Error('failed')), context};
}
const pdf = (name, text) => new File([text],name,{type:'application/pdf'});
test('sequential runs upload only the current PDF and clear old output before sending',async()=>{
  const h=setup(); h.select([pdf('first.pdf','FIRST')]); let p=h.run();
  assert.equal(await h.requests[0].get('pdfs').text(),'FIRST'); h.finish('FIRST'); await p;
  assert.equal(h.el('spool-summary-card').style.display,'block');
  h.select([pdf('second.pdf','SECOND')]); p=h.run();
  assert.equal(h.el('spool-summary-card').style.display,'none');
  assert.equal(h.el('spool-download').innerHTML,'');
  assert.equal(h.requests[1].getAll('pdfs').length,1);
  assert.equal(h.requests[1].get('pdfs').name,'second.pdf');
  assert.equal(await h.requests[1].get('pdfs').text(),'SECOND');
  h.finish('SECOND'); await p;
  assert.equal(h.el('spool-summary').textContent,'1 PDFs processed · 1 rows extracted · Excel: SECOND.xlsx');
  assert.match(h.el('spool-download').innerHTML,/Download Excel/);
  assert.match(h.el('spool-download').innerHTML,/SECOND.xlsx/);
});
test('same name and size replaces the original; removal and drop replace selection',async()=>{
  const h=setup(); h.select([pdf('same.pdf','OLD')]); h.select([pdf('same.pdf','NEW')]);
  let p=h.run(); assert.equal(await h.requests[0].get('pdfs').text(),'NEW'); h.finish('NEW'); await p;
  h.el('spool-file-list').children[0].remove.listeners.click();
  assert.equal(h.el('spool-run-btn').disabled,true);
  h.el('spool-dropzone').listeners.drop({preventDefault(){},dataTransfer:{files:[pdf('drop.pdf','DROP')]}});
  p=h.run(); assert.equal(h.requests[1].getAll('pdfs').length,1); assert.equal(h.requests[1].get('pdfs').name,'drop.pdf'); h.finish('DROP'); await p;
});
test('selection changes suppress old responses and prevent overlapping requests',async()=>{
  const h=setup(); h.select([pdf('a.pdf','A')]); const p=h.run(); h.select([pdf('b.pdf','B')]);
  await h.run(); assert.equal(h.requests.length,1); h.finish('A'); await p;
  assert.equal(h.el('spool-summary-card').style.display,'none');
  const next=h.run(); h.fail(); await next;
  assert.equal(h.el('spool-run-err').textContent,'failed');
  assert.equal(h.el('spool-summary-card').style.display,'none');
  assert.equal(h.el('spool-run-btn').disabled,false);
});
test('BOM selection remains independent and additive',()=>{
  const h=setup(); vm.runInContext('addFiles([{name:"bom1.pdf",size:1}]); addFiles([{name:"bom2.pdf",size:2}]);',h.context);
  h.select([pdf('spool.pdf','S')]);
  assert.equal(vm.runInContext('selectedFiles.length',h.context),2);
});
test('combined history renders run types, legacy fallback, filenames and downloads',async()=>{
  const h=setup();
  h.context.fetchJson=async()=>[
    {id:2,run_type:'Spool Drawing Reader',export_file:'spool.xlsx',pdf_count:1,pdf_filenames:['drawing.pdf'],row_count:7},
    {id:1,export_file:'bom.xlsx',pdf_count:1,total_inches:42,row_count:3}
  ];
  const source=fs.readFileSync(path.join(__dirname,'../FabBOMTool/static/app.js'),'utf8');
  vm.runInContext(source.slice(source.indexOf('async function loadHistory()'),source.indexOf('async function deleteRun(')),h.context);
  await vm.runInContext('loadHistory()',h.context);
  const table=h.el('history-body').children[0];
  assert.match(table.innerHTML,/<th>Run Type<\/th>/);
  const rows=table.children[0].children;
  assert.match(rows[0].innerHTML,/Spool Reader/);
  assert.match(rows[0].innerHTML,/drawing.pdf/);
  assert.match(rows[0].innerHTML,/\/api\/download\/spool.xlsx/);
  assert.match(rows[1].innerHTML,/BOM Analyzer/);
  assert.match(rows[1].innerHTML,/42.00/);
  assert.match(rows[1].innerHTML,/\/api\/download\/bom.xlsx/);
});

test('compact summary shows counts and diagnostics without needing extracted rows',()=>{
  const h=setup();
  vm.runInContext('renderSpoolSummary({row_count:12, output_filename:"batch.xlsx", diagnostics:[{file:"a.pdf", warnings:["Missing title"]},{file:"b.pdf",error:"Unreadable"}]},2)',h.context);
  assert.equal(h.el('spool-summary').textContent,'2 PDFs processed · 12 rows extracted · Excel: batch.xlsx');
  assert.match(h.el('spool-diagnostics').textContent,/a.pdf: Missing title/);
  assert.match(h.el('spool-diagnostics').textContent,/b.pdf: Unreadable/);
  vm.runInContext('renderSpoolSummary({row_count:0, diagnostics:[]},1)',h.context);
  assert.equal(h.el('spool-diagnostics').textContent,'No warnings or errors.');
  assert.equal(h.el('spool-download').innerHTML,'');
});

function setupOverlay() {
  const h = setup();
  const source = fs.readFileSync(path.join(__dirname,'../FabBOMTool/static/app.js'),'utf8');
  vm.runInContext(source.slice(source.indexOf('// Overlay owns separate selections')), h.context);
  h.choose = (side, files) => {h.el(`overlay-${side}-input`).files=files; h.el(`overlay-${side}-input`).listeners.change();};
  h.compare = () => h.el('overlay-run-btn').listeners.click();
  return h;
}
test('Overlay keeps OLD/NEW and Standard selections independent and replaces each set',async()=>{
  const h=setupOverlay(); h.select([pdf('standard.pdf','STANDARD')]);
  h.choose('old',[pdf('old1.pdf','OLD1')]);
  assert.equal(h.el('overlay-run-btn').disabled,true);
  h.choose('new',[pdf('new1.pdf','NEW1'),pdf('new2.pdf','NEW2')]);
  h.choose('old',[pdf('old2.pdf','OLD2')]);
  const p=h.compare();
  assert.deepEqual(h.requests[0].getAll('old_pdfs').map(f=>f.name),['old2.pdf']);
  assert.deepEqual(h.requests[0].getAll('new_pdfs').map(f=>f.name),['new1.pdf','new2.pdf']);
  assert.equal(await h.requests[0].get('old_pdfs').text(),'OLD2');
  assert.equal(vm.runInContext('spoolFiles[0].name',h.context),'standard.pdf');
  h.respond({row_count:4,counts:{Added:1,Removed:1,Modified:1,Unchanged:1},output_filename:'delta.xlsx',diagnostics:[]}); await p;
  assert.match(h.el('overlay-status').textContent,/1 OLD PDFs · 2 NEW PDFs · 1 Added · 1 Removed · 1 Modified · 1 Unchanged/);
  assert.equal(h.el('overlay-download').children[0].href,'/api/download/delta.xlsx');
  h.choose('new',[pdf('next.pdf','NEXT')]);
  assert.equal(h.el('overlay-download').children.length,0);
});
test('Overlay suppresses stale requests, supports removal, and switches modes',async()=>{
  const h=setupOverlay(); h.choose('old',[pdf('old.pdf','A')]); h.choose('new',[pdf('new.pdf','B')]);
  const p=h.compare(); h.choose('new',[pdf('next.pdf','C')]); await h.compare();
  assert.equal(h.requests.length,1); h.respond({row_count:1,output_filename:'stale.xlsx'}); await p;
  assert.equal(h.el('overlay-download').children.length,0);
  h.el('overlay-new-files').children[0].children[1].listeners.click();
  assert.equal(h.el('overlay-run-btn').disabled,true);
  h.el('spool-mode').listeners.change({target:{value:'overlay'}});
  assert.equal(h.el('spool-standard-panel').hidden,true); assert.equal(h.el('spool-overlay-panel').hidden,false);
  h.el('spool-mode').listeners.change({target:{value:'standard'}});
  assert.equal(h.el('spool-standard-panel').hidden,false); assert.equal(h.el('spool-overlay-panel').hidden,true);
});
test('Overlay request errors clear processing state and allow retry',async()=>{
  const h=setupOverlay();h.choose('old',[pdf('a.pdf','A')]);h.choose('new',[pdf('b.pdf','B')]);
  const p=h.compare();h.fail();await p;
  assert.equal(h.el('overlay-error').textContent,'failed');assert.equal(h.el('overlay-status').textContent,'');
  assert.equal(h.el('overlay-run-btn').disabled,false);
});
test('Overlay History badge displays both labeled filename sets',async()=>{
  const h=setup();h.context.fetchJson=async()=>[{id:1,run_type:'Spool Overlay',pdf_count:2,run_metadata:{old_filenames:['old.pdf'],new_filenames:['new.pdf']}}];
  const source=fs.readFileSync(path.join(__dirname,'../FabBOMTool/static/app.js'),'utf8');
  vm.runInContext(source.slice(source.indexOf('async function loadHistory()'),source.indexOf('async function deleteRun(')),h.context);
  await vm.runInContext('loadHistory()',h.context);
  const html=h.el('history-body').children[0].children[0].children[0].innerHTML;
  assert.match(html,/Spool Overlay/);assert.match(html,/OLD: old.pdf/);assert.match(html,/NEW: new.pdf/);assert.match(html,/—/);
});
