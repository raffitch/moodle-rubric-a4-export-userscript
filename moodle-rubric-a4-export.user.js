// ==UserScript==
// @name         Moodle Rubric - A4 Export + Quick Grade
// @namespace    https://github.com/raffitch/moodle-rubric-a4-export-userscript
// @version      4.3.4
// @description  A4 export fits width via grid and can auto-scale to ONE page height before print; shows points, highlights selected, per-criterion remarks, Overall Feedback (HTML stripped), reads Current grade from gradebook link. Removes "Due date ..." and any time stamps near the student name. Includes quota shield.
// @author       raffitch
// @license      MIT
// @homepageURL  https://github.com/raffitch/moodle-rubric-a4-export-userscript
// @supportURL   https://github.com/raffitch/moodle-rubric-a4-export-userscript/issues
// @match        https://moodle.didi.ac.ae/mod/assign/*
// @run-at       document-end
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/raffitch/moodle-rubric-a4-export-userscript/main/moodle-rubric-a4-export.user.js
// @updateURL    https://raw.githubusercontent.com/raffitch/moodle-rubric-a4-export-userscript/main/moodle-rubric-a4-export.meta.js
// ==/UserScript==

(function () {
  'use strict';
  if (window.__rtRubricInit) return;
  window.__rtRubricInit = true;

  const PAGE_OK = () => document.body && document.body.id === 'page-mod-assign-grader';
  const LS_KEY = 'rtGradeTokens';

  /* ---------- STORAGE SHIELD (quota + noisy key) ---------- */
  (function storageShield(){
    const BAD_KEY = 'RESUME_TOKEN_STORE_KEY';
    const _setItem = Storage.prototype.setItem;
    try {
      Storage.prototype.setItem = function(k, v) {
        try {
          if (k === BAD_KEY) return; // swallow huge writes
          return _setItem.call(this, k, v);
        } catch (e) {
          if (k === LS_KEY) {
            try { sessionStorage.setItem(k, v); return; } catch(_){}
            try { window.__rtMemStore = window.__rtMemStore || {}; window.__rtMemStore[k] = v; return; } catch(__){}
          }
          if (k !== BAD_KEY) throw e;
        }
      };
    } catch { /* ignore */ }
  })();

  const safeGet = (k) => { try{const v=localStorage.getItem(k); if(v!=null) return v;}catch(_){} try{const v=sessionStorage.getItem(k); if(v!=null) return v;}catch(_){} try{return (window.__rtMemStore||{})[k]||null;}catch(_){return null;} };
  const safeSet = (k,val) => { try{localStorage.setItem(k,val);return;}catch(_){} try{sessionStorage.setItem(k,val);return;}catch(_){} try{window.__rtMemStore=window.__rtMemStore||{}; window.__rtMemStore[k]=val;return;}catch(_){} };

  /* ------------------- Token helpers ------------------- */
  const normalizeSep = (s)=>String(s).replace(/[،؛;\n\r\t]+/g,',').replace(/\s*,\s*/g,',').replace(/,+/g,',').replace(/^,|,$/g,'');
  const parseTokens  = (raw)=> !raw?[] : (normalizeSep(raw).includes(',')? normalizeSep(raw).split(','): normalizeSep(raw).trim().split(/\s+/)).map(t=>t.trim()).filter(Boolean);
  const validToken   = (tok)=>/^[A-F](?:[+-])?$/.test(String(tok).trim().toUpperCase()) || /^(NS|NA|N\/A|ABS|AB)$/i.test(String(tok));
  const canonicalToken=(tok)=>{const t=String(tok).trim().toUpperCase(); if(/^(NS|NA|N\/A|ABS|AB)$/.test(t)) return 'NS'; const m=t.match(/^([A-F])([+-])?$/); return m?(m[1]+(m[2]||'')):t;};
  function extractToken(s){
    if(!s) return ''; const text=String(s).trim();
    let m=text.match(/^\(\s*([A-F][+-]?|NS)\s*:/i); if(m&&m[1]) return canonicalToken(m[1]);
    m=text.match(/^\s*([A-F][+-]?|NS)\s*[:\-–—]/i); if(m&&m[1]) return canonicalToken(m[1]);
    if(/^\s*(NS|NA|N\/A|ABS|AB)\s*$/i.test(text)) return 'NS';
    if(/^\s*NOT\s*SUBMITTED\s*$/i.test(text)) return 'NS';
    m=text.match(/([A-F])\s*([+-])?/i); return m?(m[1].toUpperCase()+((m[2]||'').toUpperCase())):'';
  }

  /* ------------------- Wait for rubric ------------------- */
  function waitForRubricStable(maxMs=10000){
    return new Promise((resolve,reject)=>{
      const t0=performance.now();
      (function poll(){
        if(!PAGE_OK()) return reject(new Error('Not on grader page'));
        const rubric=document.querySelector('.gradingform_rubric');
        const crits=rubric? rubric.querySelectorAll('.criteria .criterion'):null;
        const hasInputs=rubric? rubric.querySelector('input[type="radio"],input[type="checkbox"]'):null;
        if(rubric && crits && crits.length>0 && hasInputs){ setTimeout(()=>resolve(rubric),150); return; }
        if(performance.now()-t0>maxMs) return reject(new Error('Rubric did not appear'));
        requestAnimationFrame(poll);
      })();
    });
  }

  /* ------------------- Compact CSS (grader page) ------------------- */
  const CSS = `
    .gradingform_rubric .levels .level .score { display: none !important; }
    .gradingform_rubric .levels .level { padding: .25rem .4rem !important; line-height: 1.2; }
    .gradingform_rubric .criteria .criterion .description,
    .gradingform_rubric .criteria .criterion .criteriondescription {
      white-space: nowrap !important; overflow: hidden !important; text-overflow: ellipsis !important;
      max-width: 60ch; display: block; margin: 0 !important;
    }
    .gradingform_rubric .criteria .criterion { padding: .2rem 0 !important; }
    .gradingform_rubric .criteria .criterion .remark,
    .gradingform_rubric .criteria .criterion .criterionremark,
    .gradingform_rubric .criteria .criterion textarea { display: none !important; }
  `;
  function injectCSS(){ if(document.getElementById('rtRubricCompactCSS_SAFE')) return; const st=document.createElement('style'); st.id='rtRubricCompactCSS_SAFE'; st.textContent=CSS; document.head.appendChild(st); }

  /* ------------------- Tokenize visible labels ------------------- */
  function transformTokensOnce(root){
    root.querySelectorAll('.gradingform_rubric .levels .level .definition').forEach(d=>{
      const full=d.getAttribute('data-rt-full')||d.textContent||'';
      if(!d.hasAttribute('data-rt-full')){ d.setAttribute('data-rt-full', full); d.setAttribute('title', full.trim()); }
      d.textContent = extractToken(full) || '·';
    });
  }

  /* ------------------- Rubric mapping ------------------- */
  const getCriteria=()=>Array.from(document.querySelectorAll('.gradingform_rubric .criteria .criterion')).filter(cr=>cr.querySelector('.levels .level input[type="radio"], .levels .level input[type="checkbox"]'));
  const countCriteria=()=>getCriteria().length;

  function extractPointsFromLevel(level){
    const sEl=level.querySelector('.score'); let raw=(sEl?sEl.textContent:'')||'';
    raw=raw.trim();
    if(!raw){
      const def=level.querySelector('.definition'); const t=(def? def.getAttribute('data-rt-full')||def.textContent:'')||'';
      const m=String(t).match(/(\d+(?:\.\d+)?)\s*(?:pt|pts|point|points)?/i); if(m) raw=m[1];
    }
    const m2=raw.match(/(\d+(?:\.\d+)?)/); return m2?m2[1]:'';
  }

  function mapLevelsForCriterion(criterionEl){
    const levels=Array.from(criterionEl.querySelectorAll('.levels .level'));
    return levels.map(level=>{
      const def=level.querySelector('.definition');
      const full=(def ? (def.getAttribute('data-rt-full')||def.textContent) : '')||'';
      const tok=extractToken(full);
      const pts=extractPointsFromLevel(level);
      const input=level.querySelector('input[type="radio"], input[type="checkbox"]');
      let label=null;
      if(input){
        if(input.id) label=level.querySelector(`label[for="${input.id}"]`) || document.querySelector(`label[for="${input.id}"]`);
        if(!label) label=input.closest('label');
        if(!label) label=level.querySelector('label');
      }
      return { level, token: tok, input, label, full, points: pts };
    });
  }

  /* ------------------- Reliable selection ------------------- */
  function triggerMouseSequence(el){ ['mousedown','mouseup','click'].forEach(type=>{ el.dispatchEvent(new MouseEvent(type,{bubbles:true,cancelable:true,view:window})); }); }
  function ensureSelected(entry){
    if(!entry || !entry.input) return false;
    try{ entry.input.scrollIntoView({block:'center',inline:'center'});}catch(_){}
    if(entry.label){ try{ triggerMouseSequence(entry.label); if(entry.input.checked) return true; }catch(_){} }
    try{ triggerMouseSequence(entry.input); if(entry.input.checked) return true; }catch(_){}
    if(entry.level){ try{ triggerMouseSequence(entry.level); if(entry.input.checked) return true; }catch(_){} }
    try{ entry.input.checked=true; entry.input.dispatchEvent(new Event('change',{bubbles:true})); entry.input.dispatchEvent(new Event('input',{bubbles:true})); return entry.input.checked; }catch(_){}
    return entry.input.checked===true;
  }

  /* ------------------- Metadata & scraping ------------------- */
  const stripTimeAnywhere = (s) => String(s||'').replace(/\b\d{1,2}:\d{2}\s*(?:AM|PM)?\b/ig,'').replace(/\s{2,}/g,' ').trim();
  const stripDueDateAnywhere = (s) => String(s||'')
    .replace(/\bDue\s*date\s*:\s*[^|,\n]+/ig,'')
    .replace(/\bDue\s*:\s*[^|,\n]+/ig,'')
    .replace(/\s{2,}/g,' ')
    .trim();

  function sanitizeName(s){
    return stripTimeAnywhere(
      stripDueDateAnywhere(
        String(s||'')
          .replace(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/ig,'')
          .replace(/^\s*[:\-–—]\s*/,'')
          .replace(/\s+/g,' ')
          .trim()
      )
    );
  }

  function getCourseAndAssignment(){
    const info=document.querySelector('[data-region="grading-navigation"] [data-region="assignment-info"]') || document;
    const links=info.querySelectorAll('a[href*="/course/view.php"], a[href*="/mod/assign/view.php"]');
    let course='', assignment='';
    links.forEach(a=>{
      const href=a.getAttribute('href')||''; const t=(a.textContent||'').trim();
      if(/\/course\/view\.php/i.test(href)) course=t||course;
      if(/\/mod\/assign\/view\.php/i.test(href)) assignment=t||assignment;
    });
    return { course, assignment };
  }

  function getStudentName(){
    // 1) Change-user select
    const sel=document.getElementById('change-user-select');
    if(sel && sel.selectedIndex>=0){
      const opt=sel.options[sel.selectedIndex];
      if(opt){ const val=sanitizeName(opt.textContent||''); if(val) return val; }
    }
    // 2) User info panel
    const ui=document.querySelector('[data-region="user-info"]');
    if(ui){
      const a=ui.querySelector('a[href*="/user/view.php"]'); if(a && a.textContent.trim()) return sanitizeName(a.textContent);
      const txt=(ui.textContent||'').replace(/\s+/g,' ').trim();
      const m=txt.match(/name:\s*(.+?)\s*(?:email:|id:|$)/i); if(m && m[1]) return sanitizeName(m[1]);
      const parts=txt.split(/\s{2,}|\|/).map(s=>s.trim()).filter(Boolean); if(parts.length) return sanitizeName(parts[0]);
    }
    return '';
  }

  function getDateStr(){ try{ return new Date().toLocaleDateString(undefined,{year:'numeric',month:'short',day:'2-digit'});}catch{ return String(new Date()).replace(/\s+\d{2}:\d{2}.*$/,''); } }

  function htmlToPlainText(html){
    const div=document.createElement('div'); div.innerHTML=html;
    div.querySelectorAll('p, div, li, br').forEach(el=>{
      if(el.tagName.toLowerCase()==='br'){ el.replaceWith(document.createTextNode('\n')); }
      else { el.appendChild(document.createTextNode('\n')); }
    });
    return (div.textContent||'').replace(/\n{3,}/g,'\n\n').trim();
  }

  function getOverallFeedback(){
    const panel=document.querySelector('[data-region="grade"]') || document;
    const ta=panel.querySelector('textarea[id*="feedback"], textarea[name*="feedback"], textarea[name*="[text]"]');
    if(ta && ta.value && ta.value.trim() && !/^add a comment/i.test(ta.value)) return htmlToPlainText(ta.value.trim());

    const rich=panel.querySelector('.editor_atto_content, .editor-content, .feedbacktext, [id^="id_assignfeedbackcomments_editoreditable"]');
    if(rich){ const raw=rich.innerHTML||rich.textContent||''; const txt=htmlToPlainText(raw); if(txt && !/^add a comment/i.test(txt)) return txt; }

    const hidden=panel.querySelector('input[name*="feedback"][name$="[text]"], textarea[name*="feedback"][name$="[text]"]');
    if(hidden && hidden.value && hidden.value.trim()){ const txt=htmlToPlainText(hidden.value); if(txt) return txt; }

    const possibles=panel.querySelectorAll('.comment, .feedback, .feedbacktext, .editor_atto, .commentarea, .mod_assign_feedback');
    for(const el of possibles){ const raw=el.innerHTML||el.textContent||''; const txt=htmlToPlainText(raw); if(txt && !/^add a comment/i.test(txt)) return txt; }
    return '';
  }

  function getCurrentGradeInGradebook(){
    const a = document.querySelector('a[href*="/grade/report/grader/index.php"]');
    if (a && a.textContent.trim()) return a.textContent.trim();
    const panel=document.querySelector('[data-region="grade"]') || document;
    const gSpan=panel.querySelector('.grade, .currentgrade, [data-grade]');
    if (gSpan) return (gSpan.textContent || gSpan.getAttribute('data-grade') || '').trim();
    return '';
  }

  /* ------------------- Collect rubric + compute overall if needed ------------------- */
  function collectRubricData(){
    const crits=getCriteria(); const items=[]; let totalPoints=0; let maxPoints=0;

    crits.forEach((cr, idx)=>{
      const titleNode=cr.querySelector('.description, .criteriondescription') || cr.querySelector('.descriptiontext') || cr;
      const title=(titleNode? titleNode.textContent : `Criterion ${idx+1}`).trim();
      const levels=mapLevelsForCriterion(cr);
      const checkedIndex=levels.findIndex(l=>l.input && l.input.checked);

      const levelsData=levels.map((l,i)=>({
        token: canonicalToken(l.token || ''),
        points: l.points || '',
        description: l.full || '',
        selected: i===checkedIndex
      }));

      // points math
      const numeric=(s)=> (s && !isNaN(parseFloat(s))) ? parseFloat(s) : 0;
      const levelPoints=levels.map(x=>numeric(x.points));
      const maxForCrit = levelPoints.length ? Math.max.apply(null, levelPoints) : 0;
      maxPoints  += maxForCrit;
      if (checkedIndex >= 0) totalPoints += numeric(levels[checkedIndex].points);

      // remark (if any)
      let remark=''; const remarkNode=cr.querySelector('textarea, .criterionremark, .remark');
      if(remarkNode) remark=(remarkNode.value || remarkNode.textContent || '').trim();

      items.push({ index: idx+1, title, remark, levels: levelsData });
    });

    // Overall grade (this grading form) – scrape visible field; fallback to computed points
    let overall=''; const gp=document.querySelector('[data-region="grade"]') || document;
    const overallEl = gp.querySelector('input[type="text"][name*="grade"], input[name$="[grade]"], input[type="number"][name*="grade"], .gradevalue, .grade');
    if (overallEl) overall = (overallEl.value || overallEl.textContent || '').trim();
    if (!overall && maxPoints>0) overall = `${totalPoints.toFixed(2)} / ${maxPoints.toFixed(2)} pts`;

    const overallFeedback=getOverallFeedback();
    const currentGradebook=getCurrentGradeInGradebook();

    return { items, overall, overallFeedback, currentGradebook };
  }

  /* ------------------- Build A4 (Landscape) HTML using GRID and One-Page Scaling ------------------- */
  const escapeHTML=(s)=>String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

  function buildReportHTML(){
    const { course, assignment }=getCourseAndAssignment();
    let student = getStudentName();
    student = sanitizeName(student); // removes due date + time + emails

    const when=getDateStr();
    const { items, overall, overallFeedback, currentGradebook } = collectRubricData();

    const critBlocks = items.map(it=>{
      const levels = it.levels.map(l=>{
        const pts = l.points ? ` <span class="pts">(${l.points} pts)</span>` : '';
        return `
          <div class="level ${l.selected?'sel':''}">
            <div class="tok">${l.token || '—'}${pts}</div>
            <div class="ldesc">${escapeHTML(l.description || '') || '—'}</div>
          </div>`;
      }).join('');
      return `
        <div class="criterion">
          <div class="cmeta">
            <div class="cidx">${it.index}</div>
            <div class="ctitle">
              <div class="ct">${escapeHTML(it.title)}</div>
              ${it.remark ? `<div class="cr">Remark: ${escapeHTML(it.remark)}</div>` : ''}
            </div>
          </div>
          <div class="levels">${levels}</div>
        </div>`;
    }).join('');

    return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>Rubric – ${student ? student + ' – ' : ''}${assignment || 'Assignment'}</title>
<style>
  :root{ --page-width: 277mm; --page-height: 190mm; }
  @page { size: A4 landscape; margin: 10mm; }
  html, body { background:#fff; }
  body { font-family: system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; color:#111; font-size: 9px; line-height: 1.33; margin: 0; }
  h1 { font-size: 14px; margin: 0 0 6px 0; }
  h2 { font-size: 12px; margin: 0 0 6px 0; }

  .page { width: var(--page-width); min-height: var(--page-height); margin: 0 auto; overflow: hidden; }
  .page + .page { break-before: page; page-break-before: always; }

  .rubric-shell { position: relative; width: var(--page-width); }
  .rubric-content { position: absolute; top: 0; left: 0; width: var(--page-width); transform-origin: top left; }

  .meta { margin: 0 0 6px 0; font-size: 9px; color:#333;
          display:grid; grid-template-columns: repeat(2, minmax(0,1fr)); gap: 2px 16px; }
  .meta div { margin:1px 0; } .meta strong { font-weight:700; }

  /* Criterion block layout (no wide table) */
  .criterion { border-top: 1px solid #e6e6e6; padding: 5px 0 4px 0; }
  .criterion:first-of-type { border-top: 0; }
  .cmeta { display:grid; grid-template-columns: 28px 1fr; gap: 8px; margin-bottom: 4px; }
  .cidx { font-weight: 700; }
  .ct { font-weight: 700; margin-bottom: 2px; }
  .cr { color:#444; font-style: italic; }

  /* Levels as grid that wraps to fit width */
  .levels { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 6px; }
  .level { border:1px solid #eee; border-radius: 4px; padding: 4px; break-inside: avoid; }
  .tok { font-weight: 800; margin-bottom: 2px; text-align: center; }
  .tok .pts { font-weight: 600; opacity: .85; }
  .ldesc { white-space: normal; word-break: break-word; overflow-wrap: anywhere; }
  .sel { background:#eaf5ea; outline:1.4px solid #22a322; outline-offset:-1.4px; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  .rubric-page { -webkit-print-color-adjust: exact; print-color-adjust: exact; }

  .blocks { margin-top: 6px; display:grid; grid-template-columns:1fr; gap:6px 16px; }
  .block  { border:1px solid #ddd; padding:6px; border-radius:6px; }
  .block h3 { margin:0 0 4px 0; font-size:10px; border-bottom:1px solid #eee; padding-bottom:3px; }
  .block .content { min-height:48px; white-space: pre-wrap; font-size: 9px; }

  @media print { .noprint{ display:none !important; } }
  .toolbar.noprint { position: sticky; top: 0; background:#fff; padding:4px 0; margin-bottom:4px; border-bottom:1px solid #ddd; display:flex; gap:8px; align-items:center; }
  .toolbar button { padding:4px 7px; border:1px solid #ccc; background:#f9f9f9; cursor:pointer; border-radius:6px; font-size:9px; }
  .toolbar label { font-size: 9px; display:inline-flex; align-items:center; gap:4px; }
</style>
</head>
<body>
  <div class="page rubric-page">
    <div class="rubric-shell" id="rubricShell">
      <div class="rubric-content" id="rubricContent">
        <h1>Rubric Report (Full Breakdown)</h1>
        <div class="meta">
          <div><strong>Course:</strong> ${escapeHTML(course || '—')}</div>
          <div><strong>Assignment:</strong> ${escapeHTML(assignment || '—')}</div>
          <div><strong>Student:</strong> ${escapeHTML(student || '—')}</div>
          <div><strong>Generated:</strong> ${escapeHTML(when)}</div>
          <div><strong>Current grade in gradebook:</strong> ${escapeHTML(currentGradebook || '—')}</div>
          <div><strong>Overall grade (this grading form):</strong> ${escapeHTML(overall || '—')}</div>
        </div>

        ${critBlocks}
      </div>
    </div>
  </div>

  <div class="page feedback-page">
    <h1>Overall Feedback</h1>
    <div class="blocks">
      <div class="block">
        <div class="content">${escapeHTML(overallFeedback || '') || ' '}</div>
      </div>
    </div>
  </div>

  <script>
    (function(){
      function getRubricNodes(){
        return {
          shell: document.getElementById('rubricShell'),
          content: document.getElementById('rubricContent')
        };
      }
      function getPagePx(){
        // Measure the printable box (A4 landscape minus 10mm margins) in real px
        var probe = document.createElement('div');
        probe.style.position = 'absolute';
        probe.style.visibility = 'hidden';
        probe.style.width = '277mm';  // 297mm - 20mm margins
        probe.style.height = '190mm'; // 210mm - 20mm margins
        document.body.appendChild(probe);
        var rect = probe.getBoundingClientRect();
        probe.remove();
        return { width: rect.width || 0, height: rect.height || 0 };
      }
      function syncShell(scale){
        var nodes = getRubricNodes();
        if (!nodes.shell || !nodes.content) return;
        var height = nodes.content.scrollHeight || 0;
        if (scale && scale !== 1) height = Math.ceil(height * scale);
        nodes.shell.style.height = height + 'px';
      }
      function autoFitToOnePage(){
        try{
          var nodes = getRubricNodes();
          if (!nodes.shell || !nodes.content) return;
          document.documentElement.classList.add('fit');
          nodes.content.style.transform = 'scale(1)';
          syncShell(1);
          var page = getPagePx();
          var contentHeight = nodes.content.scrollHeight || 0;
          var contentWidth = nodes.content.scrollWidth || 0;
          if (contentHeight <= 0 || contentWidth <= 0) return;
          var scaleH = page.height / contentHeight;
          var scaleW = page.width / contentWidth;
          var s = Math.min(scaleH, scaleW);
          if (s < 1) s = Math.max(0.5, s * 0.98); else s = 1;
          nodes.content.style.transform = 'scale(' + s + ')';
          syncShell(s);
        }catch(e){ /* ignore */ }
      }
      function resetFit(){
        try{
          var nodes = getRubricNodes();
          if (!nodes.shell || !nodes.content) return;
          document.documentElement.classList.remove('fit');
          nodes.content.style.transform = 'scale(1)';
          syncShell(1);
        }catch(e){ /* ignore */ }
      }
      window.__rtAutoFit = autoFitToOnePage;
      window.__rtResetFit = resetFit;

      var fitToggle = document.getElementById('fitToggle');
      if (fitToggle) fitToggle.checked = true;

      setTimeout(function(){ try{ resetFit(); if (fitToggle && fitToggle.checked) autoFitToOnePage(); }catch(e){} }, 0);

      if (fitToggle) {
        fitToggle.addEventListener('change', function(){
          if (fitToggle.checked) autoFitToOnePage(); else resetFit();
        });
      }

            setTimeout(function(){ try{window.focus();}catch(e){} }, 50);
    })();
  </script>
</body>
</html>`;
  }

  /* ------------------- Export helpers (popup + inline + iframe) ------------------- */
  function openReportWindowSyncOrFallback() {
    const html = buildReportHTML();
    openReportInline(html);
    return true;
  }
  function openReportInline(html) {
    try { if (window.__rtHidePanel) window.__rtHidePanel(); } catch(_){ }
    let root = document.getElementById('rtInlineExportRoot');
    if (!root) {
      root = document.createElement('div');
      root.id = 'rtInlineExportRoot';
      root.style.position = 'fixed';
      root.style.inset = '0';
      root.style.zIndex = '2147483646';
      root.style.background = 'rgba(0,0,0,.35)';
      root.style.backdropFilter = 'blur(1px)';
      document.body.appendChild(root);
    } else {
      root.innerHTML = '';
      root.style.display = 'block';
    }

    const shadow = root.shadowRoot || root.attachShadow({ mode: 'open' });
    const wrap = document.createElement('div');
    const style = document.createElement('style');
    style.textContent = `
      :host { all: initial; }
      .sheet {
        box-sizing: border-box;
        background: #fff; color: #111;
        width: min(1400px, 95vw);
        height: min(95vh, 95vh);
        margin: 2.5vh auto;
        border-radius: 10px;
        box-shadow: 0 16px 60px rgba(0,0,0,.35);
        overflow: hidden; display: flex; flex-direction: column;
        border: 1px solid rgba(0,0,0,.15);
      }
      .bar {
        display: flex; gap: 8px; align-items: center;
        padding: 8px; border-bottom: 1px solid #e0e0e0;
        background: #fafafa;
      }
      .bar button { padding: 6px 10px; border: 1px solid #ccc; border-radius: 8px; background: #fff; cursor: pointer; font: 12px system-ui; }
      .bar label { font: 12px system-ui; display:inline-flex; align-items:center; gap:6px; }
      .frame { flex: 1 1 auto; border: 0; width: 100%; }
      @media print { #rtInlineExportRoot { display: none !important; } }
    `;
    wrap.innerHTML = `
      <div class="sheet">
        <div class="bar">
          <label style="display:inline-flex;align-items:center;gap:6px;"><input id="overlayFit" type="checkbox" checked/> Fit to 1 page</label>
          <button id="overlayPrint" class="primary" type="button">Print / Save as PDF</button>
          <div style="margin-left:auto;font:12px system-ui;color:#555">A4 landscape — rubric auto-fit</div>
          <button id="close" type="button">Close</button>
        </div>
        <iframe class="frame" id="reportFrame" referrerpolicy="no-referrer"></iframe>
      </div>
    `;
    shadow.innerHTML = '';
    shadow.append(style, wrap);



    const $ = (sel) => shadow.querySelector(sel);
    const frame = $('#reportFrame');
    const withFrame = (cb) => { try{ cb(frame.contentWindow, frame.contentDocument || frame.contentWindow.document); }catch(_){ } };
    const closeBtn = $('#close');
    const fitToggle = $('#overlayFit');
    const printBtn = $('#overlayPrint');
    const closeInline = () => {
      try { root.style.display = 'none'; } catch(_){ }
      try { if (window.__rtShowPanel) window.__rtShowPanel(); } catch(_){ }
    };
    window.__rtCloseInline = closeInline;
    const runFit = () => {
      withFrame((win) => {
        try{
          if (fitToggle && !fitToggle.checked && win.__rtResetFit) { win.__rtResetFit(); return; }
          if (win.__rtAutoFit) win.__rtAutoFit();
        }catch(_){ }
      });
    };
    if (closeBtn) closeBtn.addEventListener('click', closeInline);
    if (fitToggle) fitToggle.addEventListener('change', () => { runFit(); });
    if (printBtn) printBtn.addEventListener('click', () => {
      runFit();
      withFrame((win) => { try{ win.focus(); win.print(); }catch(_){ } });
    });

    function writeToIframe() {
      const doc = frame.contentDocument || frame.contentWindow?.document;
      if (!doc) return false;
      doc.open(); doc.write(html); doc.close(); runFit(); return true;
    }
    if (!writeToIframe()) frame.addEventListener('load', () => { if (writeToIframe()) runFit(); }, { once: true });
  }

  function printViaIframe(html) {
    const iframe = document.createElement('iframe');
    iframe.style.position = 'fixed'; iframe.style.right = '0'; iframe.style.bottom = '0';
    iframe.style.width = '0'; iframe.style.height = '0'; iframe.style.border = '0';
    document.body.appendChild(iframe);
    const doc = iframe.contentDocument || iframe.contentWindow.document;
    doc.open(); doc.write(html); doc.close();
    setTimeout(() => { try { iframe.contentWindow.focus(); iframe.contentWindow.print(); } catch(_){} setTimeout(() => { iframe.remove(); }, 500); }, 100);
  }

  /* ------------------- Shadow Panel ------------------- */
  function ensurePanel(){
    if(document.getElementById('rtPanelHost')) return;
    const host=document.createElement('div'); host.id='rtPanelHost';
    host.style.position='fixed'; host.style.left='16px'; host.style.bottom='16px'; host.style.zIndex='2147483647';
    document.body.appendChild(host);
    const shadow=host.attachShadow({mode:'open'});

    const wrap=document.createElement('div'); const style=document.createElement('style');
    style.textContent=`
      :host { all: initial; }
      .panel { box-sizing:border-box; width:min(46vw,620px); max-width:95vw; background:#fff; border:1px solid rgba(0,0,0,.15);
               border-radius:12px; box-shadow:0 8px 30px rgba(0,0,0,.15); padding:10px; font:13px/1.4 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; color:#111; }
      h4 { margin:0 0 6px 0; font-size:13px; font-weight:600; }
      textarea { width:100%; min-height:40px; max-height:120px; resize:vertical; padding:8px; border:1px solid rgba(0,0,0,.2); border-radius:8px; font:inherit; }
      .row { margin-top:8px; display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
      .btn { padding:6px 10px; border-radius:8px; border:1px solid rgba(0,0,0,.2); background:#f8f8f8; cursor:pointer; font-weight:600; }
      .primary { background:#0d6efd; color:#fff; border-color:#0d6efd; }
      .accent  { background:#28a745; color:#fff; border-color:#28a745; }
      .status { margin-left:auto; font-weight:600; }
      .error { margin-top:6px; color:#b00020; white-space:pre-wrap; }
      .hint { margin-top:6px; color:#555; font-size:12px; }
    `;
    wrap.innerHTML = `
      <div class="panel">
        <h4>Quick grade (comma/space separated: A, A-, B+, NS …)</h4>
        <textarea id="tokens" placeholder="A, A-, A, A-, NS"></textarea>
        <div class="row">
          <button id="apply" class="btn primary" type="button">Apply</button>
          <button id="clear" class="btn"  type="button">Clear</button>
          <button id="rescan" class="btn" type="button" title="Re-scan rubric">Rescan</button>
          <button id="export" class="btn accent" type="button" title="Export A4 (grid-fit + one-page)">Export A4</button>
          <span class="status">Criteria: <span id="count">…</span></span>
        </div>
        <div id="error" class="error"></div>
        <div class="hint">Export wraps to fit width; Fit to 1 page scales rubric to one page; feedback prints on page 2.</div>
      </div>
    `;
    shadow.append(style, wrap);

    window.__rtHidePanel = () => { try{ host.style.display = "none"; }catch(_){ } };
    window.__rtShowPanel = () => { try{ host.style.display = "block"; }catch(_){ } };

    const $=(sel)=>shadow.querySelector(sel);
    const tokensEl=$('#tokens'); const errorEl=$('#error'); const countEl=$('#count');
    const setCount=()=>{ countEl.textContent=String(countCriteria()); };
    const saved=safeGet(LS_KEY); if(saved) tokensEl.value=saved; setCount();

    $('#clear').addEventListener('click',()=>{
      tokensEl.value=''; try{localStorage.removeItem(LS_KEY);}catch(_){} try{sessionStorage.removeItem(LS_KEY);}catch(_){} if(window.__rtMemStore) delete window.__rtMemStore[LS_KEY]; errorEl.textContent='';
    });
    $('#rescan').addEventListener('click',()=>{ try{ transformTokensOnce(document); setCount(); errorEl.textContent=''; }catch(_){} });
    $('#apply').addEventListener('click',()=>{
      errorEl.textContent=''; const raw=tokensEl.value; safeSet(LS_KEY, raw);
      const tokens=parseTokens(raw).map(canonicalToken); const nCrit=countCriteria();
      if(tokens.length!==nCrit){ errorEl.textContent=`Token count (${tokens.length}) does not match criteria (${nCrit}). Parsed: [${tokens.join(', ')}]`; return; }
      for(let i=0;i<tokens.length;i++){ if(!validToken(tokens[i])){ errorEl.textContent=`Invalid token at position ${i+1}: "${tokens[i]}".`; return; } }
      const crits=getCriteria(); const missing=[];
      for(let i=0;i<crits.length;i++){
        const want=canonicalToken(tokens[i]); const levels=mapLevelsForCriterion(crits[i]);
        let entry=levels.find(l=>canonicalToken(l.token)===want);
        if(!entry){
          const defs=Array.from(crits[i].querySelectorAll('.levels .level .definition'));
          entry=defs.map((d,idx)=>({ token:canonicalToken((d.textContent||'').trim()), input:levels[idx]?.input, label:levels[idx]?.label, level:levels[idx]?.level })).find(x=>x.token===want);
        }
        if(!entry || !entry.input){ missing.push(i+1); continue; }
        const ok=ensureSelected(entry); if(!ok) missing.push(i+1);
      }
      errorEl.textContent = missing.length ? `Could not select for criteria: ${missing.join(', ')}. (Try clicking once in the rubric then Apply again.)` : 'Grades applied ✔';
    });
    $('#export').addEventListener('click',()=>{ try{ transformTokensOnce(document); }catch(_){} try{ if (window.__rtHidePanel) window.__rtHidePanel(); }catch(_){ } openReportWindowSyncOrFallback(); });

    const changeSel=document.querySelector('#change-user-select');
    if(changeSel){ changeSel.addEventListener('change',()=>{ setTimeout(()=>{ try{ transformTokensOnce(document); setCount(); }catch(_){} },600); }); }
  }

  /* ------------------- Boot ------------------- */
  async function start(){
    if(!PAGE_OK()) return;
    try { await waitForRubricStable(); injectCSS(); transformTokensOnce(document); ensurePanel(); setTimeout(()=>{ try{ transformTokensOnce(document); }catch(_){} },500); }
    catch(_){}
  }
  if(document.readyState==='loading'){ document.addEventListener('DOMContentLoaded', start); } else { start(); }
})();
