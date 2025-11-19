const $ = (sel)=>document.querySelector(sel);
const API_BASE_INPUT = $('#apiBase');
const API_BASE = () => {
  const v = API_BASE_INPUT ? API_BASE_INPUT.value.trim() : '';
  return v || window.location.origin;
};

let lastReport = null;

/* Removed click-to-toggle JS for help bubbles.
   They now open via CSS on :hover / :focus / :focus-within. */

// Drag & drop (template file)
const drop = $('#drop'), fileInput = $('#file'), fileInfo = $('#fileInfo');
if (drop){
  ['dragenter','dragover'].forEach(ev => drop.addEventListener(ev, (e)=>{e.preventDefault();e.stopPropagation();drop.classList.add('drag');}));
  ['dragleave','drop'].forEach(ev => drop.addEventListener(ev, (e)=>{e.preventDefault();e.stopPropagation();drop.classList.remove('drag');}));
  drop.addEventListener('drop', (e)=>{
    const f = e.dataTransfer.files?.[0];
    if(!f) return;
    fileInput.files = e.dataTransfer.files;
    fileInfo.textContent = `${f.name} (${fmtSize(f.size)})`;
  });
}
if (fileInput){
  fileInput.addEventListener('change', ()=>{
    const f = fileInput.files?.[0];
    fileInfo.textContent = f ? `${f.name} (${fmtSize(f.size)})` : 'No file selected.';
  });
}

// Clear
$('#clear')?.addEventListener('click', ()=>{
  if (fileInput){ fileInput.value = ''; fileInfo.textContent='No file selected.'; }
  $('#text').value=''; $('#data').value='';
  $('#varFile').value=''; $('#varList').value=''; $('#varSheet').value=''; $('#varColumn').value='';
  $('#results').style.display='none'; lastReport=null; $('#download').classList.add('hidden');
});

// Run
$('#run')?.addEventListener('click', async ()=>{
  try {
    const req = await buildRequest();
    const res = await fetch(`${API_BASE()}/validate`, req);

    let body;
    const ct = res.headers.get('content-type') || '';
    if (ct.includes('application/json')) {
      body = await res.json();
    } else {
      const txt = await res.text();
      body = { ok:false, error: `Server returned ${res.status}.`, raw: txt };
    }

    lastReport = body;
    renderResults(body, res.status);
    $('#download').classList.toggle('hidden', false);
  } catch (err) {
    console.error(err);
    showToast('Request failed. See console.');
  }
});

// Download JSON
$('#download')?.addEventListener('click', ()=>{
  if(!lastReport) return;
  const blob = new Blob([JSON.stringify(lastReport, null, 2)], {type:'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = `validation-${new Date().toISOString().slice(0,19).replace(/[:T]/g,'-')}.json`;
  a.click();
  URL.revokeObjectURL(a.href);
});

function fmtSize(n){
  if(n<1024) return `${n} B`;
  if(n<1024*1024) return `${(n/1024).toFixed(1)} KB`;
  return `${(n/1024/1024).toFixed(1)} MB`;
}

// Parse "Catalog list" textarea as JSON array or newline/comma-separated list
function parseCatalogList(text){
  const t = (text||'').trim();
  if(!t) return null;
  if(t.startsWith('[')){
    try { 
      const arr = JSON.parse(t);
      if(Array.isArray(arr)) return arr;
    } catch(e){ throw new Error('Catalog list: invalid JSON array'); }
    return null;
  }
  const items = t.split(/[\n,]+/).map(s=>s.trim()).filter(Boolean);
  return items.length ? items : null;
}

async function buildRequest(){
  const tFile = fileInput?.files?.[0];
  const vFile = $('#varFile')?.files?.[0];
  const scan = $('#scan').value;
  const preview = $('#preview').checked;
  const permissive = $('#permissive').checked;
  const encoding = $('#encoding').value;
  const varMode = $('#varMode')?.value || 'head';
  const varSheet = $('#varSheet')?.value.trim();
  const varColumn = $('#varColumn')?.value.trim();
  const varHasHeader = $('#varHasHeader')?.checked;
  const varCsvEnc = $('#varCsvEnc')?.value || 'utf-8';

  const dataTxt = $('#data').value.trim();
  let dataObj = null;
  if (dataTxt){
    try { dataObj = JSON.parse(dataTxt); }
    catch(e){ throw new Error('Sample data must be valid JSON'); }
  }

  const cd = {
    var_start: $('#varStart').value,
    var_end: $('#varEnd').value,
    block_start: $('#blkStart').value,
    block_end: $('#blkEnd').value
  };
  const hasCD = Object.values(cd).every(v => v && v.length>0);

  const varListParsed = parseCatalogList($('#varList')?.value);

  const useForm = Boolean(tFile || vFile);

  if (useForm){
    const form = new FormData();

    if (tFile) {
      form.append('file', tFile);
    } else {
      const text = $('#text').value;
      if(!text.trim()) throw new Error('Provide a file or paste template text.');
      form.append('text', text);
    }

    if (dataObj) form.append('data_json', JSON.stringify(dataObj));
    if (hasCD) form.append('custom_delims', JSON.stringify(cd));
    form.append('scan', scan);
    form.append('preview', String(preview));
    form.append('permissive', String(permissive));
    form.append('encoding', encoding);

    if (vFile) {
      if (!varColumn) throw new Error('Please fill "Column" for the catalog file (.xlsx/.csv).');
      const name = vFile.name.toLowerCase();
      form.append('varlist_file', vFile);
      form.append('varlist_mode', varMode);
      form.append('varlist_column', varColumn);
      if (name.endsWith('.xlsx') && varSheet) form.append('varlist_sheet', varSheet);
      if (name.endsWith('.csv')) form.append('varlist_encoding', varCsvEnc);
      form.append('varlist_has_header', String(varHasHeader));
    } else if (varListParsed) {
      form.append('varlist_json', JSON.stringify(varListParsed));
      form.append('varlist_mode', varMode);
    }
    return { method:'POST', body: form };
  } else {
    const text = $('#text').value;
    if(!text.trim()) throw new Error('Provide a file or paste template text.');
    const payload = {
      template: text, scan, preview, permissive,
      ...(dataObj ? {data: dataObj} : {}),
      ...(hasCD ? {custom_delims: cd} : {}),
      ...(varListParsed ? {varlist: varListParsed} : {}),
      varlist_mode: varMode
    };
    return {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify(payload)
    };
  }
}

function renderResults(r, status){
  $('#results').style.display = 'block';
  const summary = $('#summary');
  const issues = $('#issues');
  const extras = $('#extras');

  if (r && r.error) {
    summary.innerHTML = `<span class="chip err"><span class="dot err"></span>Error (${status})</span>`;
    issues.innerHTML = issueCard({ title:'Request error', message:r.error, excerpt:r.raw ? String(r.raw).slice(0,1000) : undefined });
    extras.innerHTML = '';
    return;
  }

  const errors = Array.isArray(r.errors) ? r.errors : [];
  const theMissing = Array.isArray(r.missing_from_data) ? r.missing_from_data : [];
  const unknowns = Array.isArray(r.unknown_variables) ? r.unknown_variables : [];
  const disallowed = errors.filter(e => /^disallowed_/.test(e.type));
  const syntaxErrs = errors.filter(e => e.type === 'syntax' || e.type === 'undefined' || e.type === 'render' || e.type === 'scan_unclosed_block');

  summary.innerHTML = `
    <span class="chip ${r.ok ? 'ok':'err'}"><span class="dot ${r.ok ? 'ok' : 'err'}"></span>${r.ok ? 'OK' : 'Issues found'} (${status})</span>
    <span class="chip"><span class="dot err"></span>${syntaxErrs.length} syntax/render</span>
    <span class="chip"><span class="dot warn"></span>${theMissing.length} missing vars</span>
    <span class="chip"><span class="dot warn"></span>${unknowns.length} unknown vs catalog</span>
    <span class="chip"><span class="dot warn"></span>${disallowed.length} policy</span>
    <span class="chip"><span class="dot ok"></span>${(r.used_tags||[]).length} tags</span>
    <span class="chip"><span class="dot ok"></span>${(r.used_filters||[]).length} filters</span>
  `;

  const cards = [];

  syntaxErrs.forEach((e,i)=>{
    cards.push(issueCard({
      title:`Syntax/Render #${i+1}`,
      message:e.message||'',
      location:e.docx_location,
      anchor:e.anchor,
      excerpt:e.excerpt,
      hint:e.hint
    }));
  });

  disallowed.forEach((e,i)=>{
    cards.push(issueCard({
      title:`Policy #${i+1} (${e.type.replace('disallowed_','')})`,
      message:e.message||'',
      items:e.items
    }));
  });

  theMissing.forEach((m,i)=>{
    const occ = Array.isArray(m.occurs_at) ? m.occurs_at : [];
    cards.push(issueCard({
      title:`Missing variable #${i+1}`,
      message:`${m.path}` + (m.suggestion ? ` — did you mean <code class="inline">${esc(m.suggestion)}</code>?` : ''),
      occurrences: occ
    }));
  });

  unknowns.forEach((u,i)=>{
    const occ = Array.isArray(u.occurs_at) ? u.occurs_at : [];
    const extra = u.suggestion ? ` — did you mean <code class="inline">${esc(u.suggestion)}</code>?` : '';
    cards.push(issueCard({
      title:`Unknown variable #${i+1}`,
      message:`${u.name} (checked key: <code class="inline">${esc(u.catalog_key)}</code>)${extra}`,
      occurrences: occ
    }));
  });

  issues.innerHTML = cards.join('') || `<div class="muted">No issues 🎉</div>`;

  // Extras
  extras.innerHTML = `
    <details open>
      <summary><strong>Expected variables</strong> (${(r.expected_variables||[]).length})</summary>
      <div style="margin-top:8px">${(r.expected_variables||[]).map(v=>`<span class="tag">${esc(v)}</span>`).join(' ') || '<span class="muted">none</span>'}</div>
    </details>
    <details style="margin-top:8px">
      <summary><strong>Unused data paths</strong> (${(r.unused_data_paths||[]).length})</summary>
      <div style="margin-top:8px; max-height:180px; overflow:auto">${(r.unused_data_paths||[]).map(v=>`<span class="tag">${esc(v)}</span>`).join(' ') || '<span class="muted">none</span>'}</div>
    </details>
    <details style="margin-top:8px">
      <summary><strong>Usage</strong></summary>
      <div class="kv" style="margin-top:8px">
        <div class="muted">Tags</div><div>${(r.used_tags||[]).map(v=>`<span class="tag">${esc(v)}</span>`).join(' ') || '<span class="muted">none</span>'}</div>
        <div class="muted">Filters</div><div>${(r.used_filters||[]).map(v=>`<span class="tag">${esc(v)}</span>`).join(' ') || '<span class="muted">none</span>'}</div>
        <div class="muted">Tests</div><div>${(r.used_tests||[]).map(v=>`<span class="tag">${esc(v)}</span>`).join(' ') || '<span class="muted">none</span>'}</div>
        <div class="muted">Calls</div><div>${(r.used_calls||[]).map(v=>`<span class="tag">${esc(v)}</span>`).join(' ') || '<span class="muted">none</span>'}</div>
      </div>
    </details>
  `;

  document.querySelectorAll('[data-copy]').forEach(btn=>{
    btn.addEventListener('click', async ()=>{
      const txt = btn.getAttribute('data-copy') || '';
      if(!txt) return;
      try { await navigator.clipboard.writeText(txt); showToast('Anchor copied'); }
      catch(e){ showToast('Copy failed'); }
    });
  });
}

function issueCard({title, message, location, anchor, excerpt, hint, items, occurrences}){
  const loc = location ? `<span class="tag"><strong>Location</strong> ${esc(location)}</span>` : '';
  const anc = anchor ? `<button class="btn ghost" data-copy="${escAttr(anchor)}" title="Copy anchor">Copy anchor</button>` : '';
  const itemsHtml = items ? `<div class="muted" style="margin-top:6px">Items: ${items.map(x=>`<code class="inline">${esc(x)}</code>`).join(' ')}</div>` : '';
  const hintHtml = hint ? `<div class="hint-box"><span class="hint-ico">💡</span><div>${esc(hint)}</div></div>` : '';
  const excerptHtml = excerpt ? `<pre>${esc(excerpt)}</pre>` : '';
  let occHtml = '';
  if (occurrences && occurrences.length){
    occHtml = `<div class="muted" style="margin-top:6px">Occurs at:</div><div class="list" style="margin-top:6px">` +
      occurrences.map(o=>{
        const loc = o.docx_location ? `<span class="tag">${esc(o.docx_location)}</span>` : '';
        const anc = o.anchor ? `<button class="btn ghost" data-copy="${escAttr(o.anchor)}" title="Copy anchor">Copy</button>` : '';
        const line = o.line ? `<span class="tag">line ${o.line}</span>` : '';
        return `<div class="rowline">${loc}${line}${anc}</div>`;
      }).join('') + `</div>`;
  }
  return `
    <div class="issue">
      <div class="rowline">
        <h4>${esc(title)}</h4>
        <span class="muted">${msgToHtml(message)}</span>
        <span class="right"></span>
        ${loc}${anc}
      </div>
      ${itemsHtml}
      ${hintHtml}
      ${excerptHtml}
      ${occHtml}
    </div>
  `;
}

function esc(s){ return (s??'').toString().replace(/[&<>"']/g, m=>({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[m])); }
function escAttr(s){ return esc(s).replace(/\n/g,' '); }
function msgToHtml(m){
  return esc(m).replace(/'([^']+)'/g, "<code class='inline'>$1</code>");
}

let toastTimer=null;
function showToast(msg){
  const t = $('#toast'); t.textContent=msg; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer=setTimeout(()=>t.classList.remove('show'), 1600);
}

// ===== THEME TOGGLER ======================================================
(function initTheme(){
  const KEY = 'ui-theme'; // 'light' | 'dark'
  const btn = document.getElementById('themeBtn');
  const root = document.documentElement;

  const saved = localStorage.getItem(KEY);
  let theme = saved || (window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark');
  apply(theme);

  if (btn) {
    btn.addEventListener('click', () => {
      theme = (root.getAttribute('data-theme') === 'light') ? 'dark' : 'light';
      apply(theme);
      localStorage.setItem(KEY, theme);
    });
  }

  function apply(t){
    root.setAttribute('data-theme', t);
    if (btn){
      const isLight = t === 'light';
      btn.setAttribute('aria-pressed', String(isLight));
      btn.textContent = isLight ? '🌞' : '🌙';
      btn.title = isLight ? 'Switch to dark' : 'Switch to light';
    }
  }
})();
