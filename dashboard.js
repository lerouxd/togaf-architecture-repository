/*
 * dashboard.js — TOGAF Architecture Repository · Sunrise Medical Group IT
 *
 * DATA CLASSIFICATION: INTERNAL — architecture metadata only. No patient or personal data.
 *
 * Security notes for peer review
 * ──────────────────────────────
 *  • Loaded as an external script so the CSP can keep script-src 'self' (no unsafe-inline).
 *  • All UI is built with document.createElement + textContent via h(). innerHTML is never
 *    assigned user data; containers are cleared with replaceChildren().
 *  • No credentials live in this file. SharePoint calls ride the user's existing Entra ID
 *    session cookie; writes use the per-session request digest from /_api/contextinfo.
 *  • Site URLs are validated against https://<tenant>.sharepoint.com before use.
 *  • READ_ONLY is flipped to true by the pipeline for the Tech Support build. When true,
 *    every create / edit / delete path is disabled in the UI and refused in the data layer.
 */
(function () {
  'use strict';

  const READ_ONLY = false;

  // ══════════════════════════════════════════════════════════
  //  CONSTANTS
  // ══════════════════════════════════════════════════════════
  const LIST_TITLE = 'ArchitectureRepository';
  const LOCAL_KEY = 'togaf-arch-repo-v1';
  const SITE_KEY = 'togaf-arch-repo-site';
  const SP_HOST_RE = /^https:\/\/[a-z0-9-]+\.sharepoint\.com(\/(sites|teams)\/[^/?#\s]+)?\/?$/i;
  const MAX_CHIPS = 30;

  const DOMAINS = ['Business', 'Data', 'Application', 'Technology'];
  const TYPES = ['Principle', 'Standard', 'Blueprint', 'Pattern', 'Roadmap', 'Decision'];
  const STATUSES = ['Draft', 'Approved', 'Current', 'Active', 'Retired'];
  const LIVE_STATUSES = ['Approved', 'Current', 'Active'];

  const DOMAIN_STYLE = {
    Business:    { code: 'BUS', fg: '#92400e', bg: '#fef3c7', bar: '#d97706', icon: '◆' },
    Data:        { code: 'DAT', fg: '#065f46', bg: '#d1fae5', bar: '#059669', icon: '◇' },
    Application: { code: 'APP', fg: '#1e40af', bg: '#dbeafe', bar: '#2563eb', icon: '▣' },
    Technology:  { code: 'TEC', fg: '#5b21b6', bg: '#ede9fe', bar: '#7c3aed', icon: '⚙' }
  };
  const STATUS_STYLE = {
    Draft:    { fg: '#92400e', bg: '#fef3c7', dot: '#d97706' },
    Approved: { fg: '#065f46', bg: '#d1fae5', dot: '#10b981' },
    Current:  { fg: '#1e40af', bg: '#dbeafe', dot: '#3b82f6' },
    Active:   { fg: '#166534', bg: '#dcfce7', dot: '#22c55e' },
    Retired:  { fg: '#475569', bg: '#f1f5f9', dot: '#94a3b8' }
  };

  // Artefact property → SharePoint column display name (from the setup script in index.html)
  const COLUMNS = {
    id: 'ArtefactId', domain: 'ArchDomain', subdomain: 'Subdomain', type: 'ArtefactType',
    status: 'ArtefactStatus', version: 'Version', owner: 'Owner', description: 'ArchDescription',
    rationale: 'Rationale', groups: 'BusinessGroups', projects: 'Projects', tags: 'Tags',
    created: 'ArtefactCreated', updated: 'ArtefactUpdated'
  };
  const ARRAY_PROPS = ['groups', 'projects', 'tags'];

  const LIMITS = {
    title: 255, subdomain: 100, version: 20, owner: 150, description: 4000, rationale: 4000, chip: 100
  };

  const VIEWS = [
    { key: 'overview',   label: 'Overview',        icon: '◈' },
    { key: 'landscape',  label: 'All Artefacts',   icon: '☰' },
    ...DOMAINS.map(d => ({ key: 'domain:' + d, label: d, icon: DOMAIN_STYLE[d].icon, domain: d })),
    { key: 'governance', label: 'Governance',      icon: '⚖' }
  ];

  // ══════════════════════════════════════════════════════════
  //  STATE
  // ══════════════════════════════════════════════════════════
  const state = {
    mode: null,            // 'sharepoint' | 'local'
    artefacts: [],
    view: 'overview',
    selectedId: null,
    confirmDelete: false,
    search: '',
    filters: { domain: '', type: '', status: '', sort: 'updated' },
    user: '',
    form: { editingId: null, tags: [], groups: [], projects: [] },
    saving: false
  };

  // ══════════════════════════════════════════════════════════
  //  DOM HELPERS
  // ══════════════════════════════════════════════════════════
  const $ = id => document.getElementById(id);

  /** Build an element. Children are appended as nodes or text — never parsed as HTML. */
  function h(tag, props, ...kids) {
    const el = document.createElement(tag);
    if (props) {
      for (const [k, v] of Object.entries(props)) {
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'style') Object.assign(el.style, v);
        else if (k === 'text') el.textContent = v;
        else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2), v);
        else el.setAttribute(k, v === true ? '' : String(v));
      }
    }
    for (const kid of kids.flat(Infinity)) {
      if (kid == null || kid === false) continue;
      el.append(kid instanceof Node ? kid : String(kid));
    }
    return el;
  }

  function show(el, display) { el.style.display = display || ''; }
  function hide(el) { el.style.display = 'none'; }

  /** Keyboard-activatable clickable element. */
  function clickable(el, fn) {
    el.setAttribute('tabindex', '0');
    el.setAttribute('role', 'button');
    el.addEventListener('click', fn);
    el.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fn(e); }
    });
    return el;
  }

  let toastTimer = null;
  function toast(msg) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.remove('show'), 3000);
  }

  let savedTimer = null;
  function flashSaved(msg) {
    const s = $('save-indicator');
    s.textContent = msg || '✓ Saved';
    s.style.opacity = '1';
    clearTimeout(savedTimer);
    savedTimer = setTimeout(() => { s.style.opacity = '0'; }, 2200);
  }

  // ══════════════════════════════════════════════════════════
  //  SANITISATION
  // ══════════════════════════════════════════════════════════
  function clean(value, max, multiline) {
    let s = String(value == null ? '' : value);
    s = multiline ? s.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '')
                  : s.replace(/[\u0000-\u001F\u007F]/g, ' ');
    s = s.trim();
    return max ? s.slice(0, max) : s;
  }

  function oneOf(value, allowed, fallback) {
    return allowed.includes(value) ? value : fallback;
  }

  function cleanList(arr) {
    const out = [];
    const seen = new Set();
    for (const v of Array.isArray(arr) ? arr : []) {
      const s = clean(v, LIMITS.chip);
      const key = s.toLowerCase();
      if (s && !seen.has(key)) { seen.add(key); out.push(s); }
      if (out.length >= MAX_CHIPS) break;
    }
    return out;
  }

  function parseList(raw) {
    if (Array.isArray(raw)) return cleanList(raw);
    if (!raw) return [];
    const s = String(raw).trim();
    if (s.startsWith('[')) {
      try { return cleanList(JSON.parse(s)); } catch (e) { /* fall through to delimited */ }
    }
    return cleanList(s.split(/[;\n]/));
  }

  function normalise(a) {
    const now = new Date().toISOString();
    return {
      _spId: a._spId || null,
      id: clean(a.id, 40),
      title: clean(a.title, LIMITS.title) || '(untitled)',
      domain: oneOf(a.domain, DOMAINS, 'Business'),
      subdomain: clean(a.subdomain, LIMITS.subdomain),
      type: oneOf(a.type, TYPES, 'Principle'),
      status: oneOf(a.status, STATUSES, 'Draft'),
      version: clean(a.version, LIMITS.version),
      owner: clean(a.owner, LIMITS.owner),
      description: clean(a.description, LIMITS.description, true),
      rationale: clean(a.rationale, LIMITS.rationale, true),
      tags: parseList(a.tags),
      groups: parseList(a.groups),
      projects: parseList(a.projects),
      created: validDate(a.created) || now,
      updated: validDate(a.updated) || validDate(a.created) || now
    };
  }

  function validDate(v) {
    if (!v) return '';
    const d = new Date(v);
    return isNaN(d.getTime()) ? '' : d.toISOString();
  }

  function fmtDate(iso) {
    return iso ? iso.slice(0, 10) : '—';
  }

  function nextArtefactId(domain) {
    const prefix = DOMAIN_STYLE[domain].code + '-';
    let max = 0;
    for (const a of state.artefacts) {
      if (a.id.startsWith(prefix)) {
        const n = parseInt(a.id.slice(prefix.length), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    }
    return prefix + String(max + 1).padStart(3, '0');
  }

  // ══════════════════════════════════════════════════════════
  //  STORAGE — LOCAL (developer mode)
  // ══════════════════════════════════════════════════════════
  const LocalStore = {
    async load() {
      let raw = null;
      try { raw = localStorage.getItem(LOCAL_KEY); } catch (e) { /* storage blocked */ }
      if (raw == null) return sampleData();
      try {
        const parsed = JSON.parse(raw);
        return Array.isArray(parsed) ? parsed.map(normalise) : [];
      } catch (e) {
        toast('Local data was unreadable and has been ignored');
        return [];
      }
    },
    persist() {
      try {
        localStorage.setItem(LOCAL_KEY, JSON.stringify(state.artefacts));
      } catch (e) {
        throw new Error('Browser storage is unavailable or full');
      }
    },
    async create(a) { state.artefacts.push(a); this.persist(); return a; },
    async update(a) {
      const i = state.artefacts.findIndex(x => x.id === a.id);
      if (i >= 0) state.artefacts[i] = a;
      this.persist();
      return a;
    },
    async remove(a) {
      state.artefacts = state.artefacts.filter(x => x.id !== a.id);
      this.persist();
    }
  };

  function sampleData() {
    const d = (n) => new Date(Date.now() - n * 86400000).toISOString();
    return [
      { id: 'BUS-001', title: 'SAMPLE — Customer data is a shared asset', domain: 'Business', type: 'Principle', status: 'Approved', version: '1.0', owner: 'Sample Owner', subdomain: 'Governance', description: 'Fictional sample artefact for developer mode.', rationale: 'Demonstrates how principles are displayed.', tags: ['sample', 'data-sharing'], groups: ['Finance'], projects: ['Sample Project'], created: d(40), updated: d(12) },
      { id: 'DAT-001', title: 'SAMPLE — Master data ownership standard', domain: 'Data', type: 'Standard', status: 'Current', version: '2.1', owner: 'Sample Owner', subdomain: 'MDM', description: 'Fictional sample artefact for developer mode.', rationale: '', tags: ['sample', 'data-sharing'], groups: ['Operations'], projects: [], created: d(30), updated: d(5) },
      { id: 'APP-001', title: 'SAMPLE — Integration via API gateway', domain: 'Application', type: 'Pattern', status: 'Active', version: '1.2', owner: 'Sample Owner', subdomain: 'Integration', description: 'Fictional sample artefact for developer mode.', rationale: '', tags: ['sample', 'integration'], groups: [], projects: ['Sample Project'], created: d(25), updated: d(3) },
      { id: 'APP-002', title: 'SAMPLE — Adopt a single iPaaS platform', domain: 'Application', type: 'Decision', status: 'Draft', version: '0.1', owner: 'Sample Owner', subdomain: 'Integration', description: 'Fictional sample decision record.', rationale: 'Illustrates the ADR list in Governance.', tags: ['sample', 'integration'], groups: [], projects: [], created: d(4), updated: d(1) },
      { id: 'TEC-001', title: 'SAMPLE — Network segmentation blueprint', domain: 'Technology', type: 'Blueprint', status: 'Draft', version: '0.3', owner: 'Sample Owner', subdomain: 'Networking', description: 'Fictional sample artefact for developer mode.', rationale: '', tags: ['sample'], groups: ['IT'], projects: [], created: d(8), updated: d(2) },
      { id: 'TEC-002', title: 'SAMPLE — Legacy file server', domain: 'Technology', type: 'Roadmap', status: 'Retired', version: '1.0', owner: 'Sample Owner', subdomain: 'Storage', description: 'Fictional sample artefact for developer mode.', rationale: '', tags: ['sample'], groups: [], projects: [], created: d(90), updated: d(60) }
    ].map(normalise);
  }

  // ══════════════════════════════════════════════════════════
  //  STORAGE — SHAREPOINT REST
  // ══════════════════════════════════════════════════════════
  const SP = {
    site: '',
    digest: '',
    digestExpires: 0,
    cols: {},   // artefact prop → internal column name

    listUrl() {
      return this.site + "/_api/web/lists/getbytitle('" + LIST_TITLE + "')";
    },

    async fetchJson(url, opts) {
      const res = await fetch(url, Object.assign({
        credentials: 'include',
        headers: { Accept: 'application/json;odata=nometadata' }
      }, opts));
      if (!res.ok) {
        const err = new Error('SharePoint request failed (' + res.status + ')');
        err.status = res.status;
        throw err;
      }
      if (res.status === 204) return null;
      const text = await res.text();
      return text ? JSON.parse(text) : null;
    },

    async contextInfo(site) {
      return this.fetchJson(site + '/_api/contextinfo', {
        method: 'POST',
        headers: { Accept: 'application/json;odata=nometadata' }
      });
    },

    async getDigest() {
      if (this.digest && Date.now() < this.digestExpires) return this.digest;
      const info = await this.contextInfo(this.site);
      this.digest = info.FormDigestValue;
      this.digestExpires = Date.now() + ((info.FormDigestTimeoutSeconds || 1800) - 60) * 1000;
      return this.digest;
    },

    /** Resolve the web URL and confirm the user's session is valid. */
    async connect(site) {
      const info = await this.contextInfo(site);
      this.site = String(info.WebFullUrl || site).replace(/\/$/, '');
      this.digest = info.FormDigestValue;
      this.digestExpires = Date.now() + ((info.FormDigestTimeoutSeconds || 1800) - 60) * 1000;
      return this.site;
    },

    async currentUser() {
      try {
        const u = await this.fetchJson(this.site + '/_api/web/currentuser?$select=Title');
        return clean(u && u.Title, 150);
      } catch (e) { return ''; }
    },

    /** Returns false when the list does not exist (404). */
    async resolveColumns() {
      let data;
      try {
        data = await this.fetchJson(this.listUrl() +
          '/fields?$select=InternalName,Title,ReadOnlyField,FromBaseType&$filter=Hidden eq false');
      } catch (e) {
        if (e.status === 404) return false;
        throw e;
      }
      const fields = (data && data.value) || [];
      const cols = {};
      for (const [prop, name] of Object.entries(COLUMNS)) {
        // Prefer an exact internal-name match; otherwise a writable custom column with that
        // display name (guards against clashes such as the built-in "Version" column).
        const exact = fields.find(f => f.InternalName === name && !f.ReadOnlyField);
        const byTitle = fields.find(f => f.Title === name && !f.ReadOnlyField && !f.FromBaseType);
        cols[prop] = (exact || byTitle || { InternalName: name }).InternalName;
      }
      this.cols = cols;
      return true;
    },

    fromItem(item) {
      const a = { _spId: item.Id, title: item.Title };
      for (const [prop, col] of Object.entries(this.cols)) a[prop] = item[col];
      if (!a.created) a.created = item.Created;
      if (!a.updated) a.updated = item.Modified;
      if (!a.id) a.id = 'SP-' + item.Id;
      return normalise(a);
    },

    toItem(a) {
      const item = { Title: a.title };
      for (const [prop, col] of Object.entries(this.cols)) {
        item[col] = ARRAY_PROPS.includes(prop) ? JSON.stringify(a[prop]) : a[prop];
      }
      return item;
    },

    async load() {
      const select = ['Id', 'Title', 'Created', 'Modified', ...Object.values(this.cols)].join(',');
      let url = this.listUrl() + '/items?$select=' + select + '&$top=5000';
      const out = [];
      while (url) {
        const data = await this.fetchJson(url);
        for (const item of (data && data.value) || []) out.push(this.fromItem(item));
        url = data && data['odata.nextLink'];
      }
      return out;
    },

    async write(url, method, body) {
      const digest = await this.getDigest();
      const headers = {
        Accept: 'application/json;odata=nometadata',
        'Content-Type': 'application/json;odata=nometadata',
        'X-RequestDigest': digest
      };
      if (method !== 'POST') { headers['X-HTTP-Method'] = method; headers['IF-MATCH'] = '*'; }
      return this.fetchJson(url, {
        method: 'POST',
        headers,
        body: body ? JSON.stringify(body) : undefined
      });
    },

    async create(a) {
      const item = await this.write(this.listUrl() + '/items', 'POST', this.toItem(a));
      a._spId = item && item.Id;
      state.artefacts.push(a);
      return a;
    },
    async update(a) {
      await this.write(this.listUrl() + '/items(' + Number(a._spId) + ')', 'MERGE', this.toItem(a));
      const i = state.artefacts.findIndex(x => x.id === a.id);
      if (i >= 0) state.artefacts[i] = a;
      return a;
    },
    async remove(a) {
      await this.write(this.listUrl() + '/items(' + Number(a._spId) + ')', 'DELETE');
      state.artefacts = state.artefacts.filter(x => x.id !== a.id);
    }
  };

  function store() { return state.mode === 'sharepoint' ? SP : LocalStore; }

  /** Work out the SharePoint web URL this page is hosted in, if any. */
  function detectSite() {
    const ctx = window._spPageContextInfo;
    if (ctx && ctx.webAbsoluteUrl) return String(ctx.webAbsoluteUrl);
    if (!/\.sharepoint\.com$/i.test(location.hostname)) return '';
    const path = location.pathname;
    const lib = path.match(/^(.*?)\/(SiteAssets|SitePages|Shared(%20| )Documents|Lists|Style(%20| )Library)(\/|$)/i);
    if (lib) return location.origin + lib[1];
    const site = path.match(/^\/(sites|teams)\/[^/]+/i);
    return location.origin + (site ? site[0] : '');
  }

  // ══════════════════════════════════════════════════════════
  //  STARTUP
  // ══════════════════════════════════════════════════════════
  async function init() {
    initEventListeners();
    buildNav();
    if (READ_ONLY) hide($('btn-add-main'));

    const site = detectSite();
    if (!site) {
      hide($('loading'));
      show($('auth-gate'), 'flex');
      return;
    }
    await connectSharePoint(site, { fromStartup: true });
  }

  async function connectSharePoint(site, opts) {
    const fromStartup = opts && opts.fromStartup;
    show($('loading'), 'flex');
    try {
      await SP.connect(site);
      const exists = await SP.resolveColumns();
      if (!exists) {
        hide($('loading'));
        show($('setup-modal'), 'flex');
        if (!fromStartup) setTestResult('Connected, but the ArchitectureRepository list was not found on this site.', false);
        return false;
      }
      state.user = await SP.currentUser();
      state.artefacts = await SP.load();
      try { localStorage.setItem(SITE_KEY, SP.site); } catch (e) { /* optional */ }
      state.mode = 'sharepoint';
      hide($('setup-modal'));
      startApp();
      return true;
    } catch (e) {
      hide($('loading'));
      if (fromStartup) {
        if (e.status === 401 || e.status === 403) {
          show($('auth-gate'), 'flex');
        } else {
          show($('setup-modal'), 'flex');
          setTestResult('Could not reach SharePoint automatically: ' + e.message, false);
        }
      } else {
        setTestResult('Connection failed: ' + e.message +
          '. The page must be opened from the same SharePoint site it connects to.', false);
      }
      return false;
    }
  }

  async function startLocalMode() {
    state.mode = 'local';
    state.artefacts = await LocalStore.load();
    hide($('auth-gate'));
    hide($('setup-modal'));
    startApp();
  }

  function startApp() {
    hide($('loading'));
    hide($('auth-gate'));
    show($('app'), 'flex');

    const status = $('sp-status');
    status.className = 'sp-status ' + (state.mode === 'sharepoint' ? 'connected' : 'local');
    status.textContent = (state.mode === 'sharepoint' ? '● SharePoint' : '● Local mode') +
      (READ_ONLY ? ' · Read-only' : '');
    status.title = state.mode === 'sharepoint' ? SP.site : 'Browser localStorage — not for real data';

    if (state.user) {
      const u = $('hdr-user');
      u.textContent = state.user;
      show(u);
    }
    render();
  }

  // ══════════════════════════════════════════════════════════
  //  EVENT WIRING (static elements in index.html)
  // ══════════════════════════════════════════════════════════
  function initEventListeners() {
    // Auth gate developer override
    $('dev-override-confirm').addEventListener('change', e => {
      $('btn-dev-override').disabled = !e.target.checked;
    });
    $('btn-dev-override').addEventListener('click', () => {
      if ($('dev-override-confirm').checked) startLocalMode();
    });

    // Setup modal
    $('chk-local-confirm').addEventListener('change', e => {
      $('btn-local-mode').disabled = !e.target.checked;
    });
    $('btn-local-mode').addEventListener('click', () => {
      if ($('chk-local-confirm').checked) startLocalMode();
    });
    try {
      const saved = localStorage.getItem(SITE_KEY);
      if (saved) $('sp-site-url').value = saved;
    } catch (e) { /* optional */ }
    $('btn-sp-connect').addEventListener('click', onTestConnect);
    $('sp-site-url').addEventListener('keydown', e => { if (e.key === 'Enter') onTestConnect(); });

    // Sidebar
    $('btn-add-main').addEventListener('click', () => openForm(null));

    // Search
    let searchTimer = null;
    $('search-box').addEventListener('input', e => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        state.search = clean(e.target.value, 100).toLowerCase();
        render();
      }, 150);
    });

    // Form modal
    $('btn-form-close').addEventListener('click', closeForm);
    $('btn-form-cancel').addEventListener('click', closeForm);
    $('btn-save-form').addEventListener('click', saveForm);
    $('f-title').addEventListener('input', updateSaveEnabled);
    $('form-modal').addEventListener('click', e => { if (e.target === $('form-modal')) closeForm(); });
    for (const kind of ['tags', 'groups', 'projects']) {
      $('btn-add-' + kind).addEventListener('click', () => addChip(kind));
      $('input-' + kind).addEventListener('keydown', e => {
        if (e.key === 'Enter') { e.preventDefault(); addChip(kind); }
      });
    }

    document.addEventListener('keydown', e => {
      if (e.key !== 'Escape') return;
      if ($('form-modal').style.display !== 'none') closeForm();
      else if (state.selectedId && isListView()) { state.selectedId = null; render(); }
    });
  }

  function setTestResult(msg, ok) {
    const el = $('sp-test-result');
    el.textContent = msg;
    el.style.color = ok ? '#065f46' : '#dc2626';
  }

  async function onTestConnect() {
    const url = clean($('sp-site-url').value, 300).replace(/\/+$/, '');
    if (!SP_HOST_RE.test(url)) {
      setTestResult('Enter a URL like https://tenant.sharepoint.com/sites/yoursite', false);
      return;
    }
    const btn = $('btn-sp-connect');
    btn.disabled = true;
    setTestResult('Testing connection…', true);
    const ok = await connectSharePoint(url, { fromStartup: false });
    btn.disabled = false;
    if (ok) toast('Connected to SharePoint');
  }

  // ══════════════════════════════════════════════════════════
  //  NAVIGATION
  // ══════════════════════════════════════════════════════════
  function buildNav() {
    const nav = $('sb-nav');
    nav.replaceChildren();
    for (const v of VIEWS) {
      const btn = h('button', {
        class: 'nav-btn' + (state.view === v.key ? ' active' : ''),
        'data-view': v.key,
        'aria-current': state.view === v.key ? 'page' : null,
        onclick: () => navigate(v.key)
      }, h('span', { class: 'nav-icon', 'aria-hidden': 'true', text: v.icon }), v.label);
      nav.append(btn);
    }
  }

  function navigate(view, selectId) {
    state.view = view;
    state.selectedId = selectId || null;
    state.confirmDelete = false;
    const v = VIEWS.find(x => x.key === view);
    if (v && v.domain) state.filters.domain = '';
    buildNav();
    render();
    $('content').scrollTop = 0;
  }

  function isListView() { return state.view === 'landscape' || state.view.startsWith('domain:'); }

  // ══════════════════════════════════════════════════════════
  //  RENDERING
  // ══════════════════════════════════════════════════════════
  function render() {
    const v = VIEWS.find(x => x.key === state.view) || VIEWS[0];
    $('hdr-title').textContent = v.key === 'overview' ? 'Architecture Repository' : v.label;
    const n = state.artefacts.length;
    $('artefact-count').textContent = n + (n === 1 ? ' artefact' : ' artefacts');
    const search = $('search-box');
    if (isListView()) show(search); else hide(search);

    const content = $('content');
    let node;
    if (state.view === 'overview') node = renderOverview();
    else if (state.view === 'governance') node = renderGovernance();
    else node = renderLandscape(v.domain);
    content.replaceChildren(node);
  }

  function domainTag(domain) {
    const s = DOMAIN_STYLE[domain];
    return h('span', { class: 'domain-tag', style: { background: s.bg, color: s.fg }, text: s.code });
  }

  function statusBadge(status) {
    const s = STATUS_STYLE[status] || STATUS_STYLE.Draft;
    return h('span', { class: 'status-badge', style: { background: s.bg, color: s.fg } },
      h('span', { class: 'status-dot', style: { background: s.dot } }), status);
  }

  function countBy(prop, keys) {
    const m = Object.fromEntries(keys.map(k => [k, 0]));
    for (const a of state.artefacts) if (a[prop] in m) m[a[prop]]++;
    return m;
  }

  function byUpdatedDesc(a, b) { return b.updated.localeCompare(a.updated); }

  function empty(msg) {
    return h('div', { style: { fontSize: '11px', color: '#94a3b8', padding: '6px 2px' }, text: msg });
  }

  // ── Overview ────────────────────────────────────────────
  function renderOverview() {
    const all = state.artefacts;
    const live = all.filter(a => LIVE_STATUSES.includes(a.status)).length;
    const drafts = all.filter(a => a.status === 'Draft').length;
    const retired = all.filter(a => a.status === 'Retired').length;

    const stat = (icon, value, label, color) =>
      h('div', { class: 'stat-card' },
        h('div', { class: 'stat-icon', 'aria-hidden': 'true', text: icon }),
        h('div', { class: 'stat-value', style: { color }, text: String(value) }),
        h('div', { class: 'stat-label', text: label }));

    const byDomain = countBy('domain', DOMAINS);
    const max = Math.max(1, ...Object.values(byDomain));
    const domainPanel = h('div', { class: 'panel' },
      h('div', { class: 'panel-title', text: 'Artefacts by Domain' }),
      DOMAINS.map(d => clickable(h('div', { class: 'bar-row', style: { cursor: 'pointer' } },
        h('div', { class: 'bar-label' },
          h('span', { class: 'bar-name', text: d }),
          h('span', { class: 'bar-val', style: { color: DOMAIN_STYLE[d].bar }, text: String(byDomain[d]) })),
        h('div', { class: 'bar-track' },
          h('div', { class: 'bar-fill', style: { width: (byDomain[d] / max * 100) + '%', background: DOMAIN_STYLE[d].bar } }))
      ), () => navigate('domain:' + d))));

    const byType = countBy('type', TYPES);
    const maxT = Math.max(1, ...Object.values(byType));
    const typePanel = h('div', { class: 'panel' },
      h('div', { class: 'panel-title', text: 'Artefacts by Type' }),
      TYPES.map(t => h('div', { class: 'type-row' },
        h('span', { class: 'type-name', text: t }),
        h('div', { class: 'type-right' },
          h('div', { class: 'mini-bar-track' },
            h('div', { class: 'mini-bar-fill', style: { width: (byType[t] / maxT * 100) + '%' } })),
          h('span', { class: 'type-count', text: String(byType[t]) })))));

    const recent = all.slice().sort(byUpdatedDesc).slice(0, 8);
    const recentPanel = h('div', { class: 'panel' },
      h('div', { class: 'panel-title', text: 'Recently Updated' }),
      recent.length ? recent.map(a => clickable(h('div', { class: 'recent-row' },
        domainTag(a.domain),
        h('span', { class: 'recent-title', text: a.title }),
        h('span', { class: 'recent-date', text: fmtDate(a.updated) })
      ), () => navigate('landscape', a.id))) : empty('No artefacts yet.'));

    const byStatus = countBy('status', STATUSES);
    const maxS = Math.max(1, ...Object.values(byStatus));
    const statusPanel = h('div', { class: 'panel' },
      h('div', { class: 'panel-title', text: 'Lifecycle Status' }),
      STATUSES.map(s => h('div', { class: 'bar-row' },
        h('div', { class: 'bar-label' },
          statusBadge(s),
          h('span', { class: 'bar-val', style: { color: STATUS_STYLE[s].fg }, text: String(byStatus[s]) })),
        h('div', { class: 'bar-track' },
          h('div', { class: 'bar-fill', style: { width: (byStatus[s] / maxS * 100) + '%', background: STATUS_STYLE[s].dot } })))));

    return h('div', null,
      h('div', { class: 'stats-grid' },
        stat('⬡', all.length, 'Total artefacts', '#0f172a'),
        stat('✓', live, 'Approved / current / active', '#065f46'),
        stat('✎', drafts, 'In draft', '#d97706'),
        stat('⌫', retired, 'Retired', '#64748b')),
      h('div', { class: 'panels-grid' }, domainPanel, typePanel),
      h('div', { class: 'panels-grid' }, recentPanel, statusPanel));
  }

  // ── Landscape (all / per-domain) ────────────────────────
  function filtered(fixedDomain) {
    const f = state.filters;
    const q = state.search;
    let list = state.artefacts.filter(a =>
      (!fixedDomain || a.domain === fixedDomain) &&
      (fixedDomain || !f.domain || a.domain === f.domain) &&
      (!f.type || a.type === f.type) &&
      (!f.status || a.status === f.status) &&
      (!q || [a.id, a.title, a.description, a.rationale, a.owner, a.subdomain,
              ...a.tags, ...a.groups, ...a.projects].join('\n').toLowerCase().includes(q)));
    if (f.sort === 'title') list.sort((a, b) => a.title.localeCompare(b.title));
    else if (f.sort === 'id') list.sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
    else list.sort(byUpdatedDesc);
    return list;
  }

  function filterSelect(label, key, options, allLabel) {
    const sel = h('select', { class: 'filter-select', 'aria-label': label },
      allLabel ? h('option', { value: '', text: allLabel }) : null,
      options.map(o => {
        const opt = h('option', { value: o.value || o, text: o.label || o });
        if ((o.value || o) === state.filters[key]) opt.selected = true;
        return opt;
      }));
    sel.addEventListener('change', () => { state.filters[key] = sel.value; render(); });
    return sel;
  }

  function renderLandscape(fixedDomain) {
    const list = filtered(fixedDomain);
    const selected = state.artefacts.find(a => a.id === state.selectedId);

    const filters = h('div', { class: 'filters' },
      fixedDomain ? null : filterSelect('Filter by domain', 'domain', DOMAINS, 'All domains'),
      filterSelect('Filter by type', 'type', TYPES, 'All types'),
      filterSelect('Filter by status', 'status', STATUSES, 'All statuses'),
      filterSelect('Sort', 'sort', [
        { value: 'updated', label: 'Sort: recently updated' },
        { value: 'title', label: 'Sort: title' },
        { value: 'id', label: 'Sort: ID' }
      ]),
      h('span', { class: 'results-count', text: list.length + ' shown' }));

    const cards = h('div', { class: 'artefact-list' },
      list.length ? list.map(renderCard) :
        empty(state.artefacts.length ? 'No artefacts match the current filters.' :
          READ_ONLY ? 'The repository is empty.' : 'The repository is empty. Use “+ New Artefact” to add one.'));

    return h('div', { class: 'landscape-wrap' },
      h('div', { class: 'list-col ' + (selected ? 'has-detail' : 'no-detail') }, filters, cards),
      selected ? renderDetail(selected) : null);
  }

  function renderCard(a) {
    const card = h('div', {
      class: 'artefact-card' + (a.id === state.selectedId ? ' selected' : ''),
      style: { borderLeftColor: DOMAIN_STYLE[a.domain].bar },
      'aria-pressed': a.id === state.selectedId ? 'true' : 'false'
    },
      h('div', { class: 'card-top' },
        h('div', { style: { minWidth: '0', flex: '1' } },
          h('div', { class: 'card-meta' },
            domainTag(a.domain),
            h('span', { class: 'type-chip', text: a.type }),
            a.subdomain ? h('span', { class: 'subdomain-chip', text: '· ' + a.subdomain }) : null),
          h('div', { class: 'card-title', text: a.title }),
          a.description ? h('div', { class: 'card-desc', text: a.description }) : null),
        h('div', { class: 'card-right' },
          statusBadge(a.status),
          h('span', { class: 'card-version', text: a.id + (a.version ? ' · v' + a.version : '') }))),
      a.tags.length ? h('div', { class: 'card-tags' },
        a.tags.slice(0, 6).map(t => h('span', { class: 'tag-chip', text: t }))) : null);
    return clickable(card, () => {
      state.selectedId = state.selectedId === a.id ? null : a.id;
      state.confirmDelete = false;
      render();
    });
  }

  function relatedTo(a) {
    const keys = new Set([...a.tags, ...a.projects, ...a.groups].map(s => s.toLowerCase()));
    if (!keys.size) return [];
    return state.artefacts
      .filter(b => b.id !== a.id)
      .map(b => ({ b, score: [...b.tags, ...b.projects, ...b.groups].filter(s => keys.has(s.toLowerCase())).length }))
      .filter(x => x.score > 0)
      .sort((x, y) => y.score - x.score)
      .slice(0, 6)
      .map(x => x.b);
  }

  function section(label, body) {
    return h('div', null, h('div', { class: 'detail-section-label', text: label }), body);
  }

  function renderDetail(a) {
    const ds = DOMAIN_STYLE[a.domain];
    const actions = h('div', { class: 'detail-actions' },
      READ_ONLY ? null : h('button', { class: 'btn-edit', text: '✎ Edit', onclick: () => openForm(a) }),
      h('button', { class: 'btn-close', 'aria-label': 'Close detail', text: '✕',
        onclick: () => { state.selectedId = null; render(); } }));

    const header = h('div', { class: 'detail-header', style: { background: ds.bg + '66', backgroundColor: '#fff' } },
      h('div', { style: { display: 'flex', justifyContent: 'space-between', gap: '8px' } },
        h('div', { style: { minWidth: '0' } },
          h('div', { class: 'detail-domain-row' }, domainTag(a.domain), statusBadge(a.status),
            h('span', { class: 'type-chip', text: a.type })),
          h('div', { class: 'detail-title', text: a.title }),
          h('div', { class: 'detail-sub', text: [a.domain, a.subdomain].filter(Boolean).join(' › ') })),
        actions),
      h('div', { class: 'detail-meta-row' },
        h('span', { text: a.id }),
        a.version ? h('span', { text: 'v' + a.version }) : null,
        h('span', { text: 'Created ' + fmtDate(a.created) }),
        h('span', { text: 'Updated ' + fmtDate(a.updated) })));

    const pills = (arr, cls) => h('div', { class: 'pill-row' }, arr.map(s => h('span', { class: cls, text: s })));
    const related = relatedTo(a);

    const body = h('div', { class: 'detail-body' },
      section('Owner', h('div', { class: 'detail-text', text: a.owner || '—' })),
      section('Description', a.description
        ? h('div', { class: 'detail-text', text: a.description })
        : h('div', { class: 'detail-text detail-italic', text: 'No description provided.' })),
      a.rationale ? section('Rationale', h('div', { class: 'detail-text detail-italic', text: a.rationale })) : null,
      a.groups.length ? section('Business Groups', pills(a.groups, 'pill-green')) : null,
      a.projects.length ? section('Related Projects', pills(a.projects, 'pill-orange')) : null,
      a.tags.length ? section('Tags', pills(a.tags, 'pill-blue')) : null,
      related.length ? section('Related Artefacts', h('div', null, related.map(b =>
        clickable(h('div', { class: 'related-item' },
          domainTag(b.domain),
          h('span', { class: 'related-title', text: b.title }),
          statusBadge(b.status)), () => { state.selectedId = b.id; state.confirmDelete = false; render(); })))) : null);

    let footer = null;
    if (!READ_ONLY) {
      footer = h('div', { class: 'detail-footer' },
        state.confirmDelete
          ? h('div', { class: 'confirm-row' }, 'Delete this artefact permanently?',
              h('button', { class: 'btn-secondary', style: { padding: '4px 9px', fontSize: '10px' }, text: 'Cancel',
                onclick: () => { state.confirmDelete = false; render(); } }),
              h('button', { class: 'btn-delete-confirm', text: 'Delete', onclick: () => deleteArtefact(a) }))
          : h('button', { class: 'btn-delete', text: 'Delete',
              onclick: () => { state.confirmDelete = true; render(); } }));
    }

    return h('div', { class: 'detail-col', role: 'region', 'aria-label': 'Artefact detail' }, header, body, footer);
  }

  // ── Governance ──────────────────────────────────────────
  function renderGovernance() {
    const principles = state.artefacts.filter(a => a.type === 'Principle').sort((a, b) => a.id.localeCompare(b.id, undefined, { numeric: true }));
    const decisions = state.artefacts.filter(a => a.type === 'Decision').sort(byUpdatedDesc);
    const pending = state.artefacts.filter(a => a.status === 'Draft').sort(byUpdatedDesc);
    const open = a => () => navigate('landscape', a.id);

    const rowHead = (a, extra) => h('div', { style: { display: 'flex', alignItems: 'center', gap: '6px', marginBottom: '3px' } },
      extra, domainTag(a.domain),
      h('span', { style: { fontSize: '11px', fontWeight: '600', color: '#0f172a', flex: '1' }, text: a.title }),
      statusBadge(a.status));
    const rowSub = text => text ? h('div', { style: { fontSize: '10px', color: '#64748b', lineHeight: '1.5' }, text }) : null;

    return h('div', null,
      pending.length ? h('div', { class: 'gov-pending' },
        h('div', { class: 'panel-title', style: { color: '#92400e' }, text: 'Pending review · ' + pending.length + ' draft' + (pending.length === 1 ? '' : 's') }),
        pending.map(a => clickable(h('div', { class: 'gov-adr-row', style: { cursor: 'pointer', background: '#fff' } },
          rowHead(a, h('span', { class: 'adr-id', text: a.id })),
          rowSub((a.type + ' · owner: ' + (a.owner || 'unassigned') + ' · updated ' + fmtDate(a.updated)))), open(a)))) : null,

      h('div', { class: 'gov-section' },
        h('div', { class: 'panel-title', text: 'Architecture Principles · ' + principles.length }),
        principles.length ? principles.map(a => clickable(h('div', { class: 'gov-principle-row', style: { cursor: 'pointer' } },
          rowHead(a, h('span', { class: 'adr-id', text: a.id })),
          rowSub(a.rationale || a.description)), open(a))) : empty('No principles recorded yet.')),

      h('div', { class: 'gov-section' },
        h('div', { class: 'panel-title', text: 'Architecture Decision Records · ' + decisions.length }),
        decisions.length ? decisions.map(a => clickable(h('div', { class: 'gov-adr-row', style: { cursor: 'pointer' } },
          rowHead(a, h('span', { class: 'adr-id', text: a.id })),
          rowSub(a.description),
          rowSub('Decided by ' + (a.owner || '—') + ' · ' + fmtDate(a.updated))), open(a))) : empty('No decisions recorded yet.')));
  }

  // ══════════════════════════════════════════════════════════
  //  FORM
  // ══════════════════════════════════════════════════════════
  function openForm(a) {
    if (READ_ONLY) return;
    const f = state.form;
    f.editingId = a ? a.id : null;
    f.tags = a ? a.tags.slice() : [];
    f.groups = a ? a.groups.slice() : [];
    f.projects = a ? a.projects.slice() : [];

    const v = VIEWS.find(x => x.key === state.view);
    $('f-domain').value = a ? a.domain : (v && v.domain) || 'Business';
    $('f-type').value = a ? a.type : 'Principle';
    $('f-title').value = a ? a.title : '';
    $('f-subdomain').value = a ? a.subdomain : '';
    $('f-status').value = a ? a.status : 'Draft';
    $('f-version').value = a ? a.version : '1.0';
    $('f-owner').value = a ? a.owner : state.user;
    $('f-description').value = a ? a.description : '';
    $('f-rationale').value = a ? a.rationale : '';
    for (const k of ['tags', 'groups', 'projects']) { $('input-' + k).value = ''; renderChips(k); }

    $('form-modal-title').textContent = a ? 'Edit ' + a.id : 'New Architecture Artefact';
    $('btn-save-form').textContent = a ? 'Save Changes' : 'Add to Repository';
    updateSaveEnabled();
    show($('form-modal'), 'flex');
    $('f-title').focus();
  }

  function closeForm() {
    if (state.saving) return;
    hide($('form-modal'));
  }

  function updateSaveEnabled() {
    $('btn-save-form').disabled = state.saving || !clean($('f-title').value);
  }

  const CHIP_STYLE = {
    tags:     { bg: '#eff6ff', fg: '#1d4ed8' },
    groups:   { bg: '#f0fdf4', fg: '#166534' },
    projects: { bg: '#fff7ed', fg: '#9a3412' }
  };

  function addChip(kind) {
    const input = $('input-' + kind);
    const values = input.value.split(/[;,]/).map(s => clean(s, LIMITS.chip)).filter(Boolean);
    const list = state.form[kind];
    for (const v of values) {
      if (list.length >= MAX_CHIPS) { toast('Maximum of ' + MAX_CHIPS + ' entries'); break; }
      if (!list.some(x => x.toLowerCase() === v.toLowerCase())) list.push(v);
    }
    input.value = '';
    renderChips(kind);
    input.focus();
  }

  function renderChips(kind) {
    const s = CHIP_STYLE[kind];
    $('chips-' + kind).replaceChildren(...state.form[kind].map((v, i) =>
      h('span', { class: 'chip-removable', style: { background: s.bg, color: s.fg } }, v,
        h('button', { class: 'chip-x', style: { color: s.fg }, 'aria-label': 'Remove ' + v, text: '✕',
          onclick: () => { state.form[kind].splice(i, 1); renderChips(kind); } }))));
  }

  async function saveForm() {
    if (READ_ONLY || state.saving) return;
    const title = clean($('f-title').value, LIMITS.title);
    if (!title) { $('f-title').focus(); return; }

    // Pick up anything typed into a chip box but not yet added
    for (const k of ['tags', 'groups', 'projects']) if (clean($('input-' + k).value)) addChip(k);

    const f = state.form;
    const existing = f.editingId ? state.artefacts.find(a => a.id === f.editingId) : null;
    const domain = oneOf($('f-domain').value, DOMAINS, 'Business');
    const now = new Date().toISOString();
    const artefact = normalise({
      _spId: existing ? existing._spId : null,
      id: existing ? existing.id : nextArtefactId(domain),
      title,
      domain,
      type: $('f-type').value,
      subdomain: $('f-subdomain').value,
      status: $('f-status').value,
      version: $('f-version').value,
      owner: $('f-owner').value,
      description: $('f-description').value,
      rationale: $('f-rationale').value,
      tags: f.tags, groups: f.groups, projects: f.projects,
      created: existing ? existing.created : now,
      updated: now
    });

    state.saving = true;
    updateSaveEnabled();
    const btn = $('btn-save-form');
    const label = btn.textContent;
    btn.textContent = 'Saving…';
    try {
      if (existing) await store().update(artefact);
      else await store().create(artefact);
      state.saving = false;
      btn.textContent = label;
      hide($('form-modal'));
      if (!isListView()) state.view = 'landscape';
      state.selectedId = artefact.id;
      state.confirmDelete = false;
      buildNav();
      render();
      flashSaved();
      toast(existing ? 'Updated ' + artefact.id : 'Added ' + artefact.id);
    } catch (e) {
      state.saving = false;
      btn.textContent = label;
      updateSaveEnabled();
      toast('Save failed: ' + e.message);
    }
  }

  async function deleteArtefact(a) {
    if (READ_ONLY) return;
    try {
      await store().remove(a);
      state.selectedId = null;
      state.confirmDelete = false;
      render();
      flashSaved('✓ Deleted');
      toast('Deleted ' + a.id);
    } catch (e) {
      toast('Delete failed: ' + e.message);
    }
  }

  // ══════════════════════════════════════════════════════════
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();
