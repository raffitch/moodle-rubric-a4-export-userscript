// ==UserScript==
// @name         Moodle Rubric - A4 Export + Quick Grade
// @namespace    https://github.com/raffitch/moodle-rubric-a4-export-userscript
// @version      4.4.7
// @description  A4 rubric export preview with fit/orientation/font-size controls; highlights selected levels; quick grade tokens; shows gradebook grade and feedback; strips due dates/timestamps; includes quota shield.
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
  function applyCompactCSS(enabled){
    const id='rtRubricCompactCSS_SAFE';
    const existing=document.getElementById(id);
    if(enabled){
      if(existing) return existing;
      const st=document.createElement('style'); st.id=id; st.textContent=CSS; document.head.appendChild(st); return st;
    }
    if(existing){ existing.remove(); }
    return null;
  }
  function injectCSS(){ applyCompactCSS(true); }

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
    .replace(/\bDue\s*date\s*:\s*[^|\n\r]+/ig,'')
    .replace(/\bDue\s*:\s*[^|\n\r]+/ig,'')
    .replace(/\s{2,}/g,' ')
    .trim();

  function sanitizeName(s){
    const base=String(s||'')
      .replace(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/ig,'')
      .replace(/^\s*[:\-–—]\s*/,'')
      .replace(/\s+/g,' ')
      .trim();
    return stripTimeAnywhere(stripDueDateAnywhere(base));
  }

  function getCourseAndAssignment(){
    const nav=document.querySelector('[data-region="grading-navigation"]');
    const info=nav? nav.querySelector('[data-region="assignment-info"]') : null;
    const stripLabel=(txt,label)=>String(txt||'').replace(new RegExp('^'+label+'\s*:\s*','i'),'').trim();
    let course='', assignment='';

    if(info){
      const courseLink=info.querySelector('a[href*="/course/view.php"]');
      if(courseLink){ course = stripLabel(courseLink.textContent || courseLink.title || '', 'Course'); }

      const assignLink=info.querySelector('a[href*="/mod/assign/view.php"]:not([href*="action=grading"])') || info.querySelector('a[href*="/mod/assign/view.php"]');
      if(assignLink){ assignment = stripLabel(assignLink.textContent || assignLink.title || '', 'Assignment'); }
    }

    if(!course || !assignment){
      const links=(info || document).querySelectorAll('a[href*="/course/view.php"], a[href*="/mod/assign/view.php"]');
      links.forEach(a=>{
        const href=a.getAttribute('href')||''; const t=(a.textContent||a.getAttribute('title')||'').trim();
        if(!course && /\/course\/view\.php/i.test(href)) course=stripLabel(t,'Course');
        if(/\/mod\/assign\/view\.php/i.test(href)){
          const clean=stripLabel(t,'Assignment');
          if(!assignment || (!/action=grading/i.test(href) && clean)) assignment = clean || assignment;
        }
      });
    }
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

    const selectedSum = maxPoints>0 ? `${totalPoints.toFixed(2)} / ${maxPoints.toFixed(2)} pts` : '';
    if (!overall && selectedSum) overall = selectedSum;

    const overallFeedback=getOverallFeedback();
    const currentGradebook=getCurrentGradeInGradebook();

    return { items, overall, selectedSum, totalPoints, maxPoints, overallFeedback, currentGradebook };
  }

  /* ------------------- Build A4 (Landscape) HTML using GRID and One-Page Scaling ------------------- */
  const escapeHTML=(s)=>String(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

    function buildReportHTML(){
    const { course, assignment }=getCourseAndAssignment();
    let student = getStudentName();
    student = sanitizeName(student);

    const { items, selectedSum, overallFeedback, currentGradebook } = collectRubricData();

    const critBlocks = items.map(it=>{
      const levels = it.levels.map(l=>{
        const ptsText = l.points !== '' ? String(l.points) : '';
        const pts = ptsText ? ` <span class="pts">(${escapeHTML(ptsText)} pts)</span>` : '';
        const tok = escapeHTML(l.token || '—');
        const desc = escapeHTML(l.description || '') || '—';
        return `
              <div class="level ${l.selected?'sel':''}">
                <div class="tok">${tok}${pts}</div>
                <div class="ldesc">${desc}</div>
              </div>`;
      }).join('');
      return `
          <div class="criterion">
            <div class="cmeta">
              <div class="cidx">${it.index}</div>
              <div class="ctitle">
                <div class="ct">${escapeHTML(it.title)}</div>
              </div>
            </div>
            <div class="levels">
${levels}
            </div>
          </div>`;
    }).join('\n');

    const feedbackBlock = overallFeedback ? escapeHTML(overallFeedback) : '';
    const gradebookText = escapeHTML(currentGradebook || '—');
    const selectedText = escapeHTML(selectedSum || '—');
    const courseText = escapeHTML(course || '—');
    const assignmentText = escapeHTML(assignment || '—');
    const studentText = escapeHTML(student || '—');

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Rubric Export Preview${studentText ? ' — ' + studentText : ''}</title>
  <style>
    :root {
      --page-width: 210mm;
      --page-height: 297mm;
      --fit-scale: 1;
      --level-min: 110px;
      --level-gap: 5px;
      --desc-fs: 6px;
    }

    @page {
      size: auto;
      margin: 0mm;
    }

    html,
    body {
      background: #f3f4f6;
    }

    body {
      margin: 0;
      font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
      color: #111;
      font-size: 12px;
      line-height: 1.35;
    }

    .appbar {
      position: sticky;
      top: 0;
      z-index: 5;
      background: rgba(255, 255, 255, .9);
      backdrop-filter: blur(6px);
      border-bottom: 1px solid #e5e7eb;
      padding: 10px 14px;
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
    }

    .appbar .title {
      font-weight: 800;
    }

    .appbar label {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 12px;
      color: #111;
    }

    .appbar button {
      padding: 7px 10px;
      border: 1px solid #d1d5db;
      background: #fff;
      border-radius: 10px;
      cursor: pointer;
      font-weight: 650;
    }

    .appbar .pill {
      margin-left: auto;
      font-size: 12px;
      color: #374151;
      background: #fff;
      border: 1px solid #e5e7eb;
      border-radius: 999px;
      padding: 6px 10px;
    }

    .appbar input[type="range"] {
      width: 160px;
    }

    .stage {
      padding: 18px;
      display: flex;
      justify-content: center;
    }

    .page {
      width: var(--page-width);
      min-height: var(--page-height);
      height: var(--page-height);
      background: #fff;
      border-radius: 14px;
      box-shadow: 0 18px 50px rgba(0, 0, 0, .12);
      border: 1px solid rgba(0, 0, 0, .07);
      overflow: hidden;
      display: flex;
      align-items: flex-start;
      justify-content: center;
    }

    .rubric-shell {
      width: var(--page-width);
      transform: scale(var(--fit-scale));
      transform-origin: top left;
    }

    .rubric-content {
      padding: 12px;
    }

    h1 {
      font-size: 16px;
      margin: 0 0 8px 0;
    }

    h2 {
      font-size: 13px;
      margin: 14px 0 8px 0;
    }

    .meta {
      margin: 0 0 10px 0;
      font-size: 11px;
      color: #374151;
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 2px 16px;
    }

    .meta strong {
      font-weight: 750;
      color: #111;
    }

    .criterion {
      border-top: 1px solid #e5e7eb;
      padding: 8px 0 7px 0;
    }

    .criterion:first-of-type {
      border-top: 0;
    }

    .cmeta {
      display: grid;
      grid-template-columns: 28px 1fr;
      gap: 10px;
      margin-bottom: 6px;
    }

    .cidx {
      font-weight: 900;
      color: #111;
    }

    .ct {
      font-weight: 900;
    }

    .levels {
      display: grid;
      width: 100%;
      grid-template-columns: repeat(auto-fit, minmax(var(--level-min), 1fr));
      grid-auto-flow: row dense;
      gap: var(--level-gap);
      align-items: stretch;
    }

    .level {
      border: 1px solid #e5e7eb;
      border-radius: 10px;
      padding: 7px;
      break-inside: avoid;
      background: #fff;
    }

    .tok {
      font-weight: 950;
      margin-bottom: 4px;
      text-align: center;
      font-size: 12px;
    }

    .pts {
      font-weight: 700;
      opacity: .9;
      font-size: 11px;
    }

    .ldesc {
      white-space: normal;
      word-break: break-word;
      overflow-wrap: anywhere;
      color: #111;
      font-size: var(--desc-fs);
      line-height: 1.25;
    }

    .sel {
      background: #d8f3d8;
      outline: 1.6px solid #16a34a;
      outline-offset: -1.6px;
      box-shadow: inset 0 0 0 1px rgba(22, 163, 74, .35);
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }

    body.mode-selected .level:not(.sel) {
      display: none !important;
    }

    body.mode-selected .levels {
      grid-template-columns: 1fr !important;
    }

    body.mode-selected .tok {
      text-align: left;
    }

    body.mode-full .level:not(.sel) {
      padding: 6px;
      font-size: 10.5px;
    }

    body.mode-full .level:not(.sel) .tok {
      font-size: 11px;
    }

    body.mode-full .level:not(.sel) .ldesc {
      line-height: 1.25;
    }

    .feedback {
      margin-top: 20px;
      page-break-before: always;
      break-before: page;
    }

    .fbox {
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      padding: 10px;
      white-space: pre-wrap;
      color: #111;
      background: #fff;
    }

    body.fit-active .levels {
      display: grid !important;
      grid-template-columns: repeat(6, 1fr) !important;
      gap: 4px !important;
    }

    body.fit-active .level {
      min-width: 0;
      padding: 5px;
      border-radius: 6px;
    }

    body.fit-active .level .tok {
      font-size: 12px;
      margin-bottom: 2px;
    }

    body.fit-active .level .ldesc {
      font-size: var(--desc-fs);
      line-height: 1.25;
    }

    body.fit-active .level .pts {
      font-size: 11px;
    }

    @media print {

      html,
      body {
        background: #fff;
        margin: 0 !important;
        padding: 0 !important;
        width: 100% !important;
        height: 100% !important;
        overflow: visible !important;
      }

      .appbar,
      .stage {
        display: none !important;
      }

      .printwrap {
        display: block !important;
        position: static !important;
        width: 100% !important;
        height: auto !important;
        margin: 0 !important;
        padding: 0 !important;
        transform: none !important;
      }

      .printwrap .page {
        box-shadow: none !important;
        border: 0 !important;
        border-radius: 0 !important;
        margin: 0 !important;
        padding: 0 !important;
        background: none !important;
        width: 100% !important;
        height: auto !important;
        min-height: auto !important;
        transform: none !important;
      }

      .rubric-shell {
        transform: none !important;
        width: 100% !important;
        margin: 0 !important;
      }

      .levels {
        display: grid !important;
        grid-template-columns: repeat(auto-fit, minmax(90px, 1fr)) !important;
        gap: 5px !important;
      }
    }

    .printwrap {
      display: none;
    }
  </style>
</head>

<body class="mode-full">
  <div class="appbar">
    <div class="title">Rubric Export Preview</div>
    <label><input id="toggleMode" type="checkbox" checked /> Show all levels</label>
    <label><input id="toggleFit" type="checkbox" checked /> Fit to A4</label>
    <label>
      <select id="selOrientation">
        <option value="landscape">A4 Landscape</option>
        <option value="portrait" selected>A4 Portrait</option>
      </select>
    </label>
    <label>Font Size
      <input id="fontSizeRange" type="range" min="4" max="14" value="6" step="0.5" style="width:100px" />
      <span id="fontSizeVal">6px</span>
    </label>
    <button id="btnPrint" type="button">Print / Save PDF</button>
  </div>

  <div class="stage">
    <div class="page">
      <div class="rubric-shell" id="rubricShell">
        <div class="rubric-content" id="rubricContent">
          <div class="meta">
            <div><strong>Course:</strong> ${courseText}</div>
            <div><strong>Assignment:</strong> ${assignmentText}</div>
            <div><strong>Student:</strong> ${studentText}</div>
            <div><strong>Current grade in gradebook:</strong> ${gradebookText}</div>
            <div><strong>Assignment grade (selected sum):</strong> ${selectedText}</div>
          </div>

          ${critBlocks}

          <div class="feedback">
            <h2>Overall Feedback</h2>
            <div class="fbox">${feedbackBlock || ' '}</div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <div class="printwrap">
    <div class="page">
      <div class="rubric-shell" id="printRubricShell">
        <div class="rubric-content" id="printRubricContent"></div>
      </div>
    </div>
  </div>

  <script>
    (function () {
      const $ = (id) => document.getElementById(id);
      const toggleFit = $('toggleFit');
      const toggleMode = $('toggleMode');
      const selOrientation = $('selOrientation');
      const btnPrint = $('btnPrint');

      const fontSizeRange = $('fontSizeRange');
      const fontSizeVal = $('fontSizeVal');

      const A4_W = 297;
      const A4_H = 210;

      function syncPrintDOM() {
        const src = $('rubricContent');
        const dst = $('printRubricContent');
        if (!src || !dst) return;
        dst.innerHTML = src.innerHTML;
      }

      function clamp(v, min, max) { return Math.min(max, Math.max(min, v)); }

      function setScale(scale) {
        const s = Math.max(scale || 1, 0.5);
        document.documentElement.style.setProperty('--fit-scale', s);
      }

      const nextFrame = () => new Promise(r => requestAnimationFrame(() => r()));

      async function fitStable() {
        if (toggleFit) {
          document.body.classList.toggle('fit-active', toggleFit.checked);
        }

        if (!toggleFit || !toggleFit.checked) {
          setScale(1);
          return;
        }

        try {
          const content = $('rubricContent');
          if (!content) return;

          setScale(1);

          await nextFrame();
          await nextFrame();

          const contentW = content.scrollWidth;
          if (!contentW) { setScale(1); return; }

          const root = getComputedStyle(document.documentElement);
          const pageW_mm = parseFloat(root.getPropertyValue('--page-width')) || A4_W;

          const PPCM = 3.7795;
          const availW = pageW_mm * PPCM;

          const safeW = availW - 40;

          let s = safeW / contentW;

          if (s > 1) s = 1;

          setScale(s);
        } catch (e) { console.error(e); }
      }

      function updateOrientation() {
        const val = selOrientation ? selOrientation.value : 'portrait';
        const isLand = val === 'landscape';

        const w = isLand ? A4_W : A4_H;
        const h = isLand ? A4_H : A4_W;

        document.documentElement.style.setProperty('--page-width', w + 'mm');
        document.documentElement.style.setProperty('--page-height', h + 'mm');

        const styleId = 'page-orientation-style';
        let styleEl = document.getElementById(styleId);
        if (!styleEl) {
          styleEl = document.createElement('style');
          styleEl.id = styleId;
          document.head.appendChild(styleEl);
        }
        styleEl.innerHTML = '@page { size: A4 ' + val + '; margin: 0; }';

        fitStable();
      }

      function setLayout(opts) {
        if (opts && typeof opts.minWidth === 'number') {
          const mw = clamp(opts.minWidth, 90, 320);
          document.documentElement.style.setProperty('--level-min', mw + 'px');
        }
      }

      function setMode(mode) {
        document.body.classList.toggle('mode-selected', mode === 'selected');
        document.body.classList.toggle('mode-full', mode === 'full');
      }

      if (toggleMode) {
        toggleMode.addEventListener('change', async () => {
          setMode(toggleMode.checked ? 'full' : 'selected');
          await fitStable();
        });
      }

      if (toggleFit) {
        toggleFit.addEventListener('change', async () => {
          await fitStable();
        });
      }

      if (selOrientation) {
        selOrientation.addEventListener('change', () => {
          updateOrientation();
        });
      }

      if (fontSizeRange) {
        const updateFS = () => {
          const val = fontSizeRange.value;
          document.documentElement.style.setProperty('--desc-fs', val + 'px');
          if (fontSizeVal) fontSizeVal.textContent = val + 'px';
          fitStable();
        };
        fontSizeRange.addEventListener('input', updateFS);
        updateFS();
      }

      if (btnPrint) {
        btnPrint.addEventListener('click', async () => {
          syncPrintDOM();
          await fitStable();
          window.print();
        });
      }

      window.addEventListener('beforeprint', async () => {
        syncPrintDOM();
        await fitStable();
      });

      (document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve())
        .then(() => {
          updateOrientation();
        });

      setMode('full');
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
    }
    root.style.display = 'block';

    const shadow = root.shadowRoot || root.attachShadow({ mode: 'open' });
    const style = document.createElement('style');
    style.textContent = `
      :host { all: initial; }
      .backdrop { position: fixed; inset: 0; padding: 18px; box-sizing: border-box; display: flex; align-items: center; justify-content: center; }
      .framewrap { position: relative; background: #fff; color: #111; width: min(1400px, 96vw); height: min(95vh, 95vh); border-radius: 14px; box-shadow: 0 18px 60px rgba(0,0,0,.35); overflow: hidden; border: 1px solid rgba(0,0,0,.12); display: flex; }
      iframe { border: 0; width: 100%; height: 100%; }
      .close { position: absolute; top: 10px; right: 10px; z-index: 2; background: rgba(0,0,0,.65); color: #fff; border: 0; border-radius: 999px; width: 34px; height: 34px; cursor: pointer; font-size: 18px; font-weight: 700; line-height: 1; display: grid; place-items: center; }
      @media print { :host { display: none !important; } }
    `;
    const wrap = document.createElement('div');
    wrap.className = 'backdrop';
    wrap.innerHTML = `
      <div class="framewrap">
        <button class="close" type="button" aria-label="Close preview">×</button>
        <iframe class="frame" id="reportFrame" referrerpolicy="no-referrer"></iframe>
      </div>
    `;
    shadow.innerHTML = '';
    shadow.append(style, wrap);

    const frame = shadow.querySelector('#reportFrame');
    const closeBtn = shadow.querySelector('.close');
    const closeInline = () => {
      try { root.style.display = 'none'; } catch(_){ }
      try { if (window.__rtShowPanel) window.__rtShowPanel(); } catch(_){ }
    };
    window.__rtCloseInline = closeInline;

    if (closeBtn) closeBtn.addEventListener('click', closeInline);
    shadow.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeInline(); });
    wrap.addEventListener('click', (e) => { if (e.target === wrap) closeInline(); });

    function writeToIframe() {
      const doc = frame?.contentDocument || frame?.contentWindow?.document;
      if (!doc) return false;
      doc.open(); doc.write(html); doc.close();
      return true;
    }
    if (!writeToIframe()) frame?.addEventListener('load', () => { writeToIframe(); }, { once: true });
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
      .shell { position: relative; display: inline-block; padding-left: 70px; padding-top: 12px; }
      .blob {
        position: absolute; left: -10px; top: -14px;
        width: 60px; height: 60px; border-radius: 999px;
        background: linear-gradient(135deg, #6b8bff, #9b7bff);
        box-shadow: 0 10px 30px rgba(0,0,0,.25);
        color: #fff; font: 20px/1 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;
        display: grid; place-items: center; cursor: pointer; border: 0; z-index: 2;
        transition: transform 0.1s ease, box-shadow 0.1s ease;
      }
      .blob:hover { box-shadow: 0 12px 36px rgba(0,0,0,.32); transform: translateY(-1px); }
      .panel {
        box-sizing:border-box; width:auto; max-width: 360px; min-width: 240px;
        background:#fff; border:1px solid rgba(0,0,0,.15);
        border-radius:14px; box-shadow:0 12px 36px rgba(0,0,0,.2);
        padding:10px; font:13px/1.4 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif; color:#111;
      }
      .panel.hidden { display: none; }
      .header { display:flex; align-items:center; gap:8px; }
      .title { font-size:13px; font-weight:700; }
      .spacer { flex:1; }
      textarea { width:100%; max-width: 320px; min-height:40px; max-height:120px; resize:vertical; padding:8px; border:1px solid rgba(0,0,0,.2); border-radius:8px; font:inherit; }
      .row { margin-top:8px; display:flex; gap:8px; align-items:center; flex-wrap:wrap; }
      .btn { padding:6px 10px; border-radius:8px; border:1px solid rgba(0,0,0,.2); background:#f8f8f8; cursor:pointer; font-weight:600; }
      .primary { background:#0d6efd; color:#fff; border-color:#0d6efd; }
      .accent  { background:#28a745; color:#fff; border-color:#28a745; }
      .status { margin-left:auto; font-weight:600; }
      .error { margin-top:6px; color:#b00020; white-space:pre-wrap; }
      label.chk { display:inline-flex; align-items:center; gap:6px; font-size:12px; user-select:none; }
      .icon-btn { border:1px solid rgba(0,0,0,.2); background:#f8f8f8; border-radius:999px; width:28px; height:28px; display:grid; place-items:center; cursor:pointer; }
    `;
    wrap.innerHTML = `
      <div class="shell">
        <button class="blob" id="blobToggle" type="button" title="Open tools">✦</button>
        <div class="panel" id="panel">
          <div class="header">
            <div class="title">Insert Comma Seperated Grades</div>
            <div class="spacer"></div>
            <label class="chk"><input type="checkbox" id="compactToggle" checked> Compact rubric</label>
            <button class="icon-btn" id="minimize" type="button" title="Minimize">—</button>
          </div>
          <textarea id="tokens" placeholder="A, A-, A, A-, NS"></textarea>
          <div class="row">
            <button id="apply" class="btn primary" type="button">Apply</button>
            <button id="rescan" class="btn" type="button" title="Re-scan rubric">Rescan</button>
            <button id="export" class="btn accent" type="button" title="Open A4 preview with fit/orientation controls">Export A4</button>
            <span class="status">Criteria: <span id="count">…</span></span>
          </div>
          <div id="error" class="error"></div>
        </div>
      </div>
    `;
    shadow.append(style, wrap);

    window.__rtHidePanel = () => { try{ host.style.display = "none"; }catch(_){ } };
    window.__rtShowPanel = () => { try{ host.style.display = "block"; }catch(_){ } };

    const $=(sel)=>shadow.querySelector(sel);
    const tokensEl=$('#tokens'); const errorEl=$('#error'); const countEl=$('#count'); const compactToggle=$('#compactToggle');
    const blobBtn=$('#blobToggle'); const panelEl=$('#panel'); const minimizeBtn=$('#minimize');
    const setCount=()=>{ countEl.textContent=String(countCriteria()); };
    const saved=safeGet(LS_KEY); if(saved) tokensEl.value=saved; setCount();
    if(compactToggle){ compactToggle.checked=true; compactToggle.addEventListener('change',()=>{ applyCompactCSS(!!compactToggle.checked); }); applyCompactCSS(true); }

    const setExpanded=(flag)=>{
      if(!panelEl || !blobBtn) return;
      panelEl.classList.toggle('hidden', !flag);
      blobBtn.style.display = flag ? 'none' : 'grid';
    };
    if(blobBtn) blobBtn.addEventListener('click', ()=>setExpanded(true));
    if(minimizeBtn) minimizeBtn.addEventListener('click', ()=>setExpanded(false));
    setExpanded(true);

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
