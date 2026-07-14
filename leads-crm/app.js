(() => {
  'use strict';

  const STATUSES = ['New', 'Contacted', 'Qualified', 'Proposal Sent', 'Won', 'Lost'];
  const TAG_CLASS = {
    'New': 'tag-accent',
    'Contacted': 'tag-accent-2',
    'Qualified': 'tag-outline',
    'Proposal Sent': 'tag-outline',
    'Won': 'tag-neutral',
    'Lost': 'tag-neutral'
  };

  const SAMPLE_LEADS = [
    { id: 2, timestamp: '2026-07-10T14:22:00Z', name: 'Maya Whitfield', email: 'maya@brightfield.co', phone: '(415) 555-0148', company: 'Brightfield Co', message: 'Interested in a quote for the downtown remodel.', status: 'New', notes: '' },
    { id: 3, timestamp: '2026-07-09T09:05:00Z', name: 'Dean Ostrowski', email: 'dean.o@ostro.com', phone: '(312) 555-0199', company: 'Ostro Manufacturing', message: 'Can you call me this week?', status: 'Contacted', notes: 'Left voicemail Tues.' },
    { id: 4, timestamp: '2026-07-08T18:40:00Z', name: 'Priya Nandan', email: 'priya@nandanlaw.com', phone: '(646) 555-0122', company: 'Nandan Law', message: 'Looking for ongoing service, not one-off.', status: 'Qualified', notes: '' },
    { id: 5, timestamp: '2026-07-06T11:15:00Z', name: 'Colin Reyes', email: 'colin@reyesbuild.net', phone: '(720) 555-0177', company: 'Reyes Build', message: '', status: 'Proposal Sent', notes: 'Sent proposal 7/6, following up Friday.' },
    { id: 6, timestamp: '2026-07-01T08:00:00Z', name: 'Ingrid Solberg', email: 'ingrid@solberg.io', phone: '(206) 555-0133', company: 'Solberg Studio', message: 'Not a fit right now, revisit in Q4.', status: 'Lost', notes: '' }
  ];

  const SCRIPT_CODE = `/**
 * Leads CRM — Google Apps Script backend
 * Paste into Extensions → Apps Script on the Sheet StaticForms writes to,
 * then Deploy → New deployment → Web app (Execute as: Me, Access: Anyone).
 */
const SHEET_NAME = 'Sheet1';
const COLS = { timestamp:1, name:2, email:3, phone:4, company:5, message:6, status:7, notes:8 };

function _sheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.getSheets()[0];
  const header = sh.getRange(1,1,1,8).getValues()[0];
  if (!header[6]) sh.getRange(1,7).setValue('Status');
  if (!header[7]) sh.getRange(1,8).setValue('Notes');
  return sh;
}
function _json(obj){return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}

function doGet(e) {
  const sh = _sheet();
  const last = sh.getLastRow();
  if (last < 2) return _json({leads:[]});
  const rows = sh.getRange(2,1,last-1,8).getValues();
  const leads = rows.map((r,i)=>({
    id:i+2, timestamp: r[0]? new Date(r[0]).toISOString():'',
    name:r[1]||'', email:r[2]||'', phone:r[3]||'', company:r[4]||'', message:r[5]||'',
    status:r[6]||'New', notes:r[7]||''
  })).filter(l=>l.name||l.email);
  return _json({leads});
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || '{}');
  const sh = _sheet();
  if (body.action === 'update' && body.id) {
    const row = Number(body.id);
    if (body.status !== undefined) sh.getRange(row, COLS.status).setValue(body.status);
    if (body.notes !== undefined) sh.getRange(row, COLS.notes).setValue(body.notes);
    return _json({ok:true});
  }
  return _json({ok:false, error:'Unknown action'});
}`;

  function fmtDate(iso) {
    if (!iso) return '—';
    const d = new Date(iso);
    if (isNaN(d)) return '—';
    return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function escapeHtml(v) {
    return String(v == null ? '' : v)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  const root = document.getElementById('app');

  const state = {
    webAppUrl: '',
    urlDraft: '',
    showSetup: false,
    leads: [],
    usingSample: true,
    loading: false,
    error: '',
    search: '',
    statusFilter: 'All',
    sortKey: 'timestamp',
    sortDir: 'desc',
    selectedId: null,
    statusDraft: 'New',
    notesDraft: '',
    saving: false
  };

  function setState(patch) {
    Object.assign(state, typeof patch === 'function' ? patch(state) : patch);
    render();
  }

  function fetchLeads(url) {
    setState({ loading: true, error: '' });
    fetch(url)
      .then(r => r.json())
      .then(data => {
        setState({ leads: data.leads || [], usingSample: false, loading: false });
      })
      .catch(err => {
        setState({ error: 'Could not reach the Web App (' + err.message + '). Showing sample data instead.', leads: SAMPLE_LEADS, usingSample: true, loading: false });
      });
  }

  function saveUrl() {
    const url = state.urlDraft.trim();
    localStorage.setItem('crm_webapp_url', url);
    setState({ webAppUrl: url, showSetup: false });
    if (url) fetchLeads(url);
    else setState({ leads: SAMPLE_LEADS, usingSample: true });
  }
  function clearUrl() {
    localStorage.removeItem('crm_webapp_url');
    setState({ webAppUrl: '', urlDraft: '', leads: SAMPLE_LEADS, usingSample: true, showSetup: false });
  }
  function toggleSetup() { setState(s => ({ showSetup: !s.showSetup })); }
  function refresh() { if (state.webAppUrl) fetchLeads(state.webAppUrl); }

  function sortBy(key) {
    setState(s => ({
      sortKey: key,
      sortDir: s.sortKey === key && s.sortDir === 'asc' ? 'desc' : 'asc'
    }));
  }

  function openLead(id) {
    const lead = state.leads.find(l => l.id === id);
    if (!lead) return;
    setState({ selectedId: id, statusDraft: lead.status || 'New', notesDraft: lead.notes || '' });
  }
  function closeDialog() { setState({ selectedId: null }); }

  function saveLead() {
    const id = state.selectedId;
    const { statusDraft, notesDraft, webAppUrl, usingSample } = state;
    setState(s => ({
      leads: s.leads.map(l => l.id === id ? { ...l, status: statusDraft, notes: notesDraft } : l)
    }));
    if (usingSample || !webAppUrl) { setState({ selectedId: null }); return; }
    setState({ saving: true });
    fetch(webAppUrl, {
      method: 'POST',
      body: JSON.stringify({ action: 'update', id, status: statusDraft, notes: notesDraft })
    }).then(() => setState({ saving: false, selectedId: null }))
      .catch(err => setState({ saving: false, error: 'Save failed: ' + err.message }));
  }

  function computeVals() {
    const s = state;
    const configured = !!s.webAppUrl;

    let leads = s.leads.filter(l => {
      const q = s.search.trim().toLowerCase();
      const matchesQ = !q || [l.name, l.email, l.company].some(v => (v || '').toLowerCase().includes(q));
      const matchesStatus = s.statusFilter === 'All' || l.status === s.statusFilter;
      return matchesQ && matchesStatus;
    });

    leads = leads.slice().sort((a, b) => {
      const av = a[s.sortKey] || '', bv = b[s.sortKey] || '';
      const cmp = av < bv ? -1 : av > bv ? 1 : 0;
      return s.sortDir === 'asc' ? cmp : -cmp;
    });

    const visibleLeads = leads.map(l => ({
      ...l,
      tagClass: TAG_CLASS[l.status] || 'tag-neutral',
      dateLabel: fmtDate(l.timestamp)
    }));

    const counts = { total: s.leads.length, New: 0, Contacted: 0, Won: 0 };
    s.leads.forEach(l => { if (counts[l.status] !== undefined) counts[l.status]++; });

    const selectedRaw = s.leads.find(l => l.id === s.selectedId) || null;
    const selected = selectedRaw ? { ...selectedRaw, dateLabel: fmtDate(selectedRaw.timestamp) } : null;

    const sortMarks = {};
    ['name', 'company', 'status', 'timestamp'].forEach(k => {
      sortMarks[k] = s.sortKey === k ? (s.sortDir === 'asc' ? ' ↑' : ' ↓') : '';
    });

    return {
      configured,
      showSetup: s.showSetup,
      setupLabel: s.showSetup ? 'Close setup' : (configured ? 'Settings' : 'Connect your sheet'),
      statusLine: s.loading ? 'Loading…' : (s.usingSample ? 'Sample data' : ('Synced · ' + s.leads.length + ' leads')),
      statCards: [
        { label: 'Total leads', value: counts.total },
        { label: 'New', value: counts.New },
        { label: 'Contacted', value: counts.Contacted },
        { label: 'Won', value: counts.Won }
      ],
      hasError: !!s.error,
      errorText: s.error,
      sortMarks,
      visibleLeads,
      noResults: visibleLeads.length === 0,
      selected,
      saving: s.saving,
      saveLabel: s.saving ? 'Saving…' : 'Save'
    };
  }

  function renderNav(v) {
    return `
      <div class="nav" style="border-bottom:1px solid var(--color-divider); padding:var(--space-4) var(--space-6);">
        <div class="nav-brand">Leads CRM</div>
        <div style="margin-left:auto; display:flex; align-items:center; gap:var(--space-3);">
          ${v.configured ? `
            <span class="text-muted" style="font-size:12px;">${escapeHtml(v.statusLine)}</span>
            <button class="btn btn-secondary" data-action="refresh">Refresh</button>
          ` : ''}
          <button class="btn btn-ghost" data-action="toggle-setup">${escapeHtml(v.setupLabel)}</button>
        </div>
      </div>
    `;
  }

  function renderSetup(v) {
    if (!v.showSetup) return '';
    return `
      <div style="max-width:920px; margin:var(--space-6) auto; padding:0 var(--space-6);">
        <div class="card blueprint elev-sm" style="padding:var(--space-6); gap:var(--space-4);">
          <i class="corner tl"></i><i class="corner tr"></i><i class="corner bl"></i><i class="corner br"></i>
          <h3 style="margin:0;">Connect your Google Sheet</h3>
          <p class="card-body" style="opacity:1; margin:0;">
            Deploy the included Apps Script as a Web App (Execute as: Me, Access: Anyone),
            then paste its URL below. Nothing is stored anywhere but this browser.
          </p>
          <div class="field">
            <label>Apps Script Web App URL</label>
            <input class="input" type="text" placeholder="https://script.google.com/macros/s/…/exec" data-bind="urlDraft" value="${escapeHtml(state.urlDraft)}">
          </div>
          <div style="display:flex; gap:var(--space-2);">
            <button class="btn btn-primary" data-action="save-url">Save &amp; connect</button>
            ${v.configured ? `<button class="btn btn-secondary" data-action="clear-url">Disconnect</button>` : ''}
          </div>
          <div class="hr"></div>
          <h5 style="margin:0;">Apps Script code (paste into the Sheet's Extensions → Apps Script)</h5>
          <pre style="background:var(--color-surface); border:1px solid var(--color-divider); padding:var(--space-3); font-size:12px; overflow:auto; max-height:280px; white-space:pre-wrap;">${escapeHtml(SCRIPT_CODE)}</pre>
        </div>
      </div>
    `;
  }

  function renderMain(v) {
    return `
      <div style="max-width:1200px; margin:0 auto; padding:var(--space-6);">
        <div style="display:grid; grid-template-columns:repeat(4,1fr); gap:var(--space-4); margin-bottom:var(--space-6);">
          ${v.statCards.map(stat => `
            <div class="card blueprint" style="padding:var(--space-4);">
              <i class="corner tl"></i><i class="corner tr"></i><i class="corner bl"></i><i class="corner br"></i>
              <div class="card-kicker">${escapeHtml(stat.label)}</div>
              <div style="font-family:var(--font-heading); font-weight:600; font-size:32px;">${escapeHtml(stat.value)}</div>
            </div>
          `).join('')}
        </div>

        <div style="display:flex; gap:var(--space-3); align-items:center; margin-bottom:var(--space-4); flex-wrap:wrap;">
          <div class="field" style="flex:1; min-width:220px; margin:0;">
            <input class="input" type="text" placeholder="Search name, email, company…" data-bind="search" value="${escapeHtml(state.search)}">
          </div>
          <select class="input" style="width:auto; min-width:160px;" data-bind="statusFilter">
            <option value="All" ${state.statusFilter === 'All' ? 'selected' : ''}>All statuses</option>
            ${STATUSES.map(s => `<option value="${escapeHtml(s)}" ${state.statusFilter === s ? 'selected' : ''}>${escapeHtml(s)}</option>`).join('')}
          </select>
        </div>

        ${v.hasError ? `
          <div class="card" style="border-color:var(--color-accent); margin-bottom:var(--space-4); padding:var(--space-3);">
            <div class="card-body" style="opacity:1;">${escapeHtml(v.errorText)}</div>
          </div>
        ` : ''}

        <div style="border:1px solid var(--color-divider); overflow:auto;">
          <table class="table">
            <thead>
              <tr>
                <th data-sort-key="name" style="cursor:pointer;">Name${v.sortMarks.name}</th>
                <th data-sort-key="company" style="cursor:pointer;">Company${v.sortMarks.company}</th>
                <th>Email</th>
                <th>Phone</th>
                <th data-sort-key="status" style="cursor:pointer;">Status${v.sortMarks.status}</th>
                <th data-sort-key="timestamp" style="cursor:pointer;">Received${v.sortMarks.timestamp}</th>
              </tr>
            </thead>
            <tbody>
              ${v.visibleLeads.map(lead => `
                <tr data-lead-id="${lead.id}" style="cursor:pointer;">
                  <td style="font-weight:500;">${escapeHtml(lead.name)}</td>
                  <td>${escapeHtml(lead.company)}</td>
                  <td class="text-muted">${escapeHtml(lead.email)}</td>
                  <td class="text-muted">${escapeHtml(lead.phone)}</td>
                  <td><span class="tag ${lead.tagClass}">${escapeHtml(lead.status)}</span></td>
                  <td class="text-muted">${escapeHtml(lead.dateLabel)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
          ${v.noResults ? `<div class="text-muted" style="padding:var(--space-6); text-align:center;">No leads match.</div>` : ''}
        </div>
      </div>
    `;
  }

  function renderDialog(v) {
    if (!v.selected) return '';
    const sel = v.selected;
    return `
      <div class="dialog-backdrop" data-action="close-backdrop">
        <div class="dialog blueprint elev-lg">
          <i class="corner tl"></i><i class="corner tr"></i><i class="corner bl"></i><i class="corner br"></i>
          <div class="dialog-title">${escapeHtml(sel.name)}</div>
          <div class="dialog-body" style="display:flex; flex-direction:column; gap:var(--space-3);">
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:var(--space-3);">
              <div><div class="text-muted" style="font-size:11px;">EMAIL</div><div>${escapeHtml(sel.email)}</div></div>
              <div><div class="text-muted" style="font-size:11px;">PHONE</div><div>${escapeHtml(sel.phone)}</div></div>
              <div><div class="text-muted" style="font-size:11px;">COMPANY</div><div>${escapeHtml(sel.company)}</div></div>
              <div><div class="text-muted" style="font-size:11px;">RECEIVED</div><div>${escapeHtml(sel.dateLabel)}</div></div>
            </div>
            ${sel.message ? `
              <div>
                <div class="text-muted" style="font-size:11px; margin-bottom:4px;">SUBMITTED MESSAGE</div>
                <div style="background:var(--color-surface); border:1px solid var(--color-divider); padding:var(--space-3); font-size:13px;">${escapeHtml(sel.message)}</div>
              </div>
            ` : ''}
            <div class="field" style="margin:0;">
              <label>Status</label>
              <select class="input" data-bind="statusDraft">
                ${STATUSES.map(s => `<option value="${escapeHtml(s)}" ${state.statusDraft === s ? 'selected' : ''}>${escapeHtml(s)}</option>`).join('')}
              </select>
            </div>
            <div class="field" style="margin:0;">
              <label>Notes</label>
              <textarea class="input" rows="4" data-bind="notesDraft">${escapeHtml(state.notesDraft)}</textarea>
            </div>
          </div>
          <div class="dialog-actions">
            <button class="btn btn-secondary" data-action="close-dialog">Cancel</button>
            <button class="btn btn-primary" data-action="save-lead" ${v.saving ? 'disabled' : ''}>${escapeHtml(v.saveLabel)}</button>
          </div>
        </div>
      </div>
    `;
  }

  function render() {
    const v = computeVals();
    root.innerHTML = `
      <div style="min-height:100vh; background:var(--color-bg); color:var(--color-text); font-family:var(--font-body);">
        ${renderNav(v)}
        ${renderSetup(v)}
        ${renderMain(v)}
        ${renderDialog(v)}
      </div>
    `;
    bindEvents();
  }

  function bindEvents() {
    const byAction = (name) => root.querySelector(`[data-action="${name}"]`);

    byAction('refresh')?.addEventListener('click', refresh);
    byAction('toggle-setup')?.addEventListener('click', toggleSetup);
    byAction('save-url')?.addEventListener('click', saveUrl);
    byAction('clear-url')?.addEventListener('click', clearUrl);
    byAction('close-dialog')?.addEventListener('click', closeDialog);
    byAction('save-lead')?.addEventListener('click', saveLead);

    root.querySelector('[data-bind="urlDraft"]')?.addEventListener('input', e => setState({ urlDraft: e.target.value }));
    root.querySelector('[data-bind="search"]')?.addEventListener('input', e => setState({ search: e.target.value }));
    root.querySelector('[data-bind="statusFilter"]')?.addEventListener('change', e => setState({ statusFilter: e.target.value }));
    root.querySelector('[data-bind="statusDraft"]')?.addEventListener('change', e => setState({ statusDraft: e.target.value }));
    root.querySelector('[data-bind="notesDraft"]')?.addEventListener('input', e => setState({ notesDraft: e.target.value }));

    root.querySelectorAll('[data-sort-key]').forEach(el => {
      el.addEventListener('click', () => sortBy(el.dataset.sortKey));
    });
    root.querySelectorAll('tr[data-lead-id]').forEach(tr => {
      tr.addEventListener('click', () => openLead(Number(tr.dataset.leadId)));
    });

    const backdrop = root.querySelector('.dialog-backdrop');
    if (backdrop) {
      backdrop.addEventListener('click', e => { if (e.target === e.currentTarget) closeDialog(); });
      root.querySelector('.dialog')?.addEventListener('click', e => e.stopPropagation());
    }
  }

  function init() {
    const saved = localStorage.getItem('crm_webapp_url') || '';
    state.webAppUrl = saved;
    state.urlDraft = saved;
    if (saved) {
      fetchLeads(saved);
    } else {
      state.leads = SAMPLE_LEADS;
      state.usingSample = true;
      render();
    }
  }

  init();
})();
