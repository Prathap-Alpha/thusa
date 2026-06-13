'use strict';

/* ============================================================
   Thusa — chat-to-calendar assistant for Alpha Direct
   - LLM: Anthropic Claude (direct from browser, key in Setup)
   - Calendar: Microsoft Graph via MSAL.js (sign in once)
   All data stays in this browser's localStorage.
   ============================================================ */

const DEFAULT_MODEL = 'claude-haiku-4-5';
const TIMEZONE = 'Africa/Gaborone';
const GRAPH_SCOPES = ['User.Read', 'Calendars.ReadWrite'];
const S_KEY = 'thusa_settings';
const T_KEY = 'thusa_team';

/* ---------------- storage ---------------- */

function getSettings() {
  try { return JSON.parse(localStorage.getItem(S_KEY)) || {}; } catch { return {}; }
}
function saveSettings(patch) {
  localStorage.setItem(S_KEY, JSON.stringify({ ...getSettings(), ...patch }));
}
function getTeam() {
  try { return JSON.parse(localStorage.getItem(T_KEY)) || []; } catch { return []; }
}
function saveTeam(team) {
  localStorage.setItem(T_KEY, JSON.stringify(team));
}
function nameFor(email) {
  const hit = getTeam().find(p => p.email.toLowerCase() === String(email).toLowerCase());
  return hit ? hit.name : email;
}

/* ---------------- small helpers ---------------- */

const el = id => document.getElementById(id);

function toast(msg, ms = 3500) {
  const t = el('toast');
  t.textContent = msg;
  t.hidden = false;
  clearTimeout(t._timer);
  t._timer = setTimeout(() => { t.hidden = true; }, ms);
}

function toLocalISO(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
}

function fmtWhen(startStr, minutes) {
  const d = new Date(startStr);
  if (isNaN(d)) return startStr;
  const s = d.toLocaleString('en-GB', {
    weekday: 'short', day: 'numeric', month: 'short',
    hour: '2-digit', minute: '2-digit',
  });
  return `${s} · ${minutes || 30} min`;
}

/* ---------------- Microsoft auth (MSAL) ---------------- */

let _pca = null;
let _pcaReady = null;

function msalConfigured() {
  return !!getSettings().clientId;
}

async function getMsal() {
  const s = getSettings();
  if (!s.clientId) return null;
  if (!_pca) {
    _pca = new msal.PublicClientApplication({
      auth: {
        clientId: s.clientId,
        authority: `https://login.microsoftonline.com/${s.tenantId || 'organizations'}`,
        redirectUri: location.origin + location.pathname,
        postLogoutRedirectUri: location.origin + location.pathname,
      },
      cache: { cacheLocation: 'localStorage' },
    });
    _pcaReady = _pca.initialize();
  }
  await _pcaReady;
  return _pca;
}

function currentAccount(pca) {
  return pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
}

async function initAuth() {
  if (!msalConfigured()) { renderAuthStatus(); return; }
  try {
    const pca = await getMsal();
    const result = await pca.handleRedirectPromise();
    if (result && result.account) {
      pca.setActiveAccount(result.account);
      toast('Sharp sharp! Signed in as ' + result.account.username);
    }
    if (currentAccount(pca)) refreshMe().catch(() => {});
  } catch (e) {
    toast('Microsoft sign-in error: ' + (e.errorMessage || e.message));
  }
  renderAuthStatus();
}

async function signIn() {
  const s = getSettings();
  if (!s.clientId) { toast('Paste the Application (client) ID first, then Save settings.'); return; }
  const pca = await getMsal();
  await pca.loginRedirect({ scopes: GRAPH_SCOPES, prompt: 'select_account' });
}

async function signOut() {
  if (!msalConfigured()) return;
  const pca = await getMsal();
  const account = currentAccount(pca);
  saveSettings({ userName: '', userMail: '' });
  await pca.logoutRedirect({ account });
}

async function getGraphToken(interactive) {
  const pca = await getMsal();
  if (!pca) throw new Error('Microsoft is not set up yet — open Setup and fill in section 2.');
  const account = currentAccount(pca);
  if (!account) {
    if (interactive) { await signIn(); }
    throw new Error('Not signed in to Microsoft — open Setup and tap "Sign in with Microsoft".');
  }
  try {
    const res = await pca.acquireTokenSilent({ scopes: GRAPH_SCOPES, account });
    return res.accessToken;
  } catch (e) {
    if (interactive && e instanceof msal.InteractionRequiredAuthError) {
      // Round-trips through Microsoft's page; with "stay signed in" this
      // completes without asking for a password.
      await pca.acquireTokenRedirect({ scopes: GRAPH_SCOPES, account });
    }
    throw new Error('Microsoft session needs a refresh — try again in a moment.');
  }
}

async function refreshMe() {
  const token = await getGraphToken(false);
  const r = await fetch('https://graph.microsoft.com/v1.0/me', {
    headers: { Authorization: 'Bearer ' + token },
  });
  if (r.ok) {
    const me = await r.json();
    saveSettings({ userName: me.displayName || '', userMail: me.mail || me.userPrincipalName || '' });
    renderAuthStatus();
  }
}

function renderAuthStatus() {
  const badge = el('authBadge');
  const s = getSettings();
  (async () => {
    if (!msalConfigured()) { badge.textContent = 'Setup needed'; badge.className = 'brand-sub'; return; }
    const pca = await getMsal();
    const account = currentAccount(pca);
    if (account) {
      badge.textContent = s.userName ? `${s.userName} · connected` : account.username;
      badge.className = 'brand-sub ok';
    } else {
      badge.textContent = 'Not signed in';
      badge.className = 'brand-sub';
    }
  })().catch(() => {});
}

/* ---------------- Microsoft Graph: create the event ---------------- */

async function createEvent(m) {
  const start = new Date(m.start);
  if (isNaN(start)) throw new Error('Could not understand the start time "' + m.start + '"');
  const minutes = Number(m.duration_minutes) > 0 ? Number(m.duration_minutes) : 30;
  const end = new Date(start.getTime() + minutes * 60000);

  const body = {
    subject: m.subject,
    body: { contentType: 'text', content: m.notes || 'Scheduled via Thusa.' },
    start: { dateTime: toLocalISO(start), timeZone: TIMEZONE },
    end: { dateTime: toLocalISO(end), timeZone: TIMEZONE },
    attendees: (m.attendee_emails || []).map(a => ({
      emailAddress: { address: a, name: nameFor(a) },
      type: 'required',
    })),
  };
  if (m.location) body.location = { displayName: m.location };
  if (m.online !== false) {
    body.isOnlineMeeting = true;
    body.onlineMeetingProvider = 'teamsForBusiness';
  }

  const token = await getGraphToken(true);
  const res = await fetch('https://graph.microsoft.com/v1.0/me/events', {
    method: 'POST',
    headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!res.ok) {
    let msg = 'Microsoft Graph error ' + res.status;
    try { const j = await res.json(); msg = (j.error && j.error.message) || msg; } catch {}
    throw new Error(msg);
  }
  return res.json();
}

/* ---------------- Anthropic (Claude) ---------------- */

const TOOLS = [{
  name: 'create_meeting',
  description: 'Create a calendar event in the user\'s Microsoft 365 calendar and send invites to the attendees. Call this once per meeting the user asks for.',
  input_schema: {
    type: 'object',
    properties: {
      subject: { type: 'string', description: 'Short business-style meeting title' },
      start: { type: 'string', description: 'Start as local Gaborone wall-clock ISO datetime, e.g. 2026-06-15T17:00:00 — no timezone suffix, no Z' },
      duration_minutes: { type: 'integer', description: 'Length in minutes; default 30' },
      attendee_emails: { type: 'array', items: { type: 'string' }, description: 'Email addresses of the attendees, taken from the team directory' },
      notes: { type: 'string', description: 'Optional agenda / body text for the invite' },
      location: { type: 'string', description: 'Optional physical location, e.g. Boardroom' },
      online: { type: 'boolean', description: 'true = attach a Microsoft Teams link (the default)' },
    },
    required: ['subject', 'start', 'attendee_emails'],
  },
}];

function systemPrompt() {
  const s = getSettings();
  const team = getTeam();
  const dir = team.length
    ? team.map(p => `- ${p.name} <${p.email}>`).join('\n')
    : '(empty — ask the user to add people in the Team tab, or to type the full email address)';
  return `You are Thusa, the personal meeting-scheduling assistant for ${s.userName || 'the CFO'} at Alpha Direct Insurance, Gaborone, Botswana.
Current date and time: ${new Date().toString()} (timezone ${TIMEZONE}, UTC+2, no daylight saving).

Team directory — the only people you may invite by name:
${dir}

Rules:
- Turn scheduling requests into create_meeting tool calls. "start" must be local Gaborone wall-clock time, ISO format like 2026-06-15T08:00:00, never with a Z or offset.
- Match names against the directory case-insensitively and tolerate small misspellings. If a name has no reasonable match, ask — never invent an email address. If the user types a full email address you may use it directly.
- If no time is given, propose the next working day at 08:00 and say clearly that you assumed it.
- If no duration is given, use 30 minutes.
- Default to an online Microsoft Teams meeting unless a physical location is mentioned.
- One message can contain several meetings — make one tool call per meeting.
- Sensitive topics (dismissals, appraisals, disciplinary): keep the invite subject discreet and professional; put detail in notes only if the user gave it.
- Keep replies short, warm and businesslike. A light touch of Setswana ("Sharp sharp!", "Go siame.") is welcome.`;
}

async function callClaude(messages) {
  const s = getSettings();
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': s.apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: s.model || DEFAULT_MODEL,
      max_tokens: 2048,
      system: systemPrompt(),
      tools: TOOLS,
      messages,
    }),
  });
  if (!res.ok) {
    let msg = 'Anthropic API error ' + res.status;
    try { const j = await res.json(); msg = (j.error && j.error.message) || msg; } catch {}
    if (res.status === 401) msg = 'Your Anthropic API key was rejected — check it in Setup.';
    if (res.status === 429) msg = 'Rate limit reached — wait a few seconds and try again.';
    throw new Error(msg);
  }
  return res.json();
}

/* ---------------- chat engine ---------------- */

let history = [];      // Anthropic-format message history (this session only)
let awaiting = null;   // { ids: [], results: {}, total, cards: {} } while cards are pending
let busy = false;

function trimHistory() {
  // Drop oldest complete turn-groups (user-text → assistant → tool_results)
  // without ever draining the array to empty or leaving an assistant turn
  // at the head — both would 400 on the next request.
  while (history.length > 40) {
    history.shift();
    while (history.length && history[0].role === 'assistant') history.shift();
    while (history.length && history[0].role === 'user' &&
           Array.isArray(history[0].content) &&
           history[0].content[0] && history[0].content[0].type === 'tool_result') {
      history.shift();
    }
  }
  // Safety: the first message sent to the API must be a user turn.
  while (history.length && history[0].role !== 'user') history.shift();
}

function addBubble(role, text) {
  const w = el('welcome');
  if (w) w.remove();
  const div = document.createElement('div');
  div.className = 'bubble ' + role;
  div.textContent = text;
  el('msgs').appendChild(div);
  el('msgs').scrollTop = el('msgs').scrollHeight;
  return div;
}

let thinkingBubble = null;
function setThinking(on) {
  busy = on;
  el('sendBtn').disabled = on;
  if (on && !thinkingBubble) {
    thinkingBubble = addBubble('assistant thinking', 'Ke a akanya… (thinking)');
  } else if (!on && thinkingBubble) {
    thinkingBubble.remove();
    thinkingBubble = null;
  }
}

function addCard(toolUse) {
  const w = el('welcome');
  if (w) w.remove();
  const input = toolUse.input || {};
  const card = document.createElement('div');
  card.className = 'meeting-card';

  const subject = document.createElement('div');
  subject.className = 'mc-subject';
  subject.textContent = '📅 ' + (input.subject || 'Meeting');
  card.appendChild(subject);

  const when = document.createElement('div');
  when.className = 'mc-line';
  when.innerHTML = '<b>When:</b> ';
  when.appendChild(document.createTextNode(fmtWhen(input.start, input.duration_minutes)));
  card.appendChild(when);

  const who = document.createElement('div');
  who.className = 'mc-line';
  who.innerHTML = '<b>Invitees:</b> ';
  who.appendChild(document.createTextNode((input.attendee_emails || []).map(nameFor).join(', ') || '—'));
  card.appendChild(who);

  if (input.location) {
    const loc = document.createElement('div');
    loc.className = 'mc-line';
    loc.innerHTML = '<b>Where:</b> ';
    loc.appendChild(document.createTextNode(input.location));
    card.appendChild(loc);
  } else if (input.online !== false) {
    const loc = document.createElement('div');
    loc.className = 'mc-line';
    loc.innerHTML = '<b>Where:</b> Microsoft Teams';
    card.appendChild(loc);
  }

  if (input.notes) {
    const notes = document.createElement('div');
    notes.className = 'mc-line';
    notes.innerHTML = '<b>Notes:</b> ';
    notes.appendChild(document.createTextNode(input.notes));
    card.appendChild(notes);
  }

  const actions = document.createElement('div');
  actions.className = 'mc-actions';
  const sendBtn = document.createElement('button');
  sendBtn.className = 'primary';
  sendBtn.textContent = 'Send invite';
  const cancelBtn = document.createElement('button');
  cancelBtn.className = 'ghost-btn';
  cancelBtn.textContent = 'Cancel';
  actions.appendChild(sendBtn);
  actions.appendChild(cancelBtn);
  card.appendChild(actions);

  sendBtn.addEventListener('click', () => {
    actions.remove();
    executeMeeting(toolUse.id, input, card);
  });
  cancelBtn.addEventListener('click', () => {
    actions.remove();
    setCardStatus(card, 'Cancelled', 'err');
    finishTool(toolUse.id, 'The user cancelled this meeting — do not create it.');
  });

  card._actions = actions;
  el('msgs').appendChild(card);
  el('msgs').scrollTop = el('msgs').scrollHeight;
  return card;
}

function setCardStatus(card, text, cls) {
  if (card._actions) card._actions.remove();
  let status = card.querySelector('.mc-status');
  if (!status) {
    status = document.createElement('div');
    status.className = 'mc-status';
    card.appendChild(status);
  }
  status.textContent = text;
  status.className = 'mc-status' + (cls ? ' ' + cls : '');
}

async function executeMeeting(id, input, card) {
  setCardStatus(card, 'Sending…');
  try {
    const ev = await createEvent(input);
    setCardStatus(card, '✅ Invite sent', 'ok');
    finishTool(id, 'Success — the event was created and the invites were sent. Outlook link: ' + (ev.webLink || 'n/a'));
  } catch (e) {
    setCardStatus(card, '❌ ' + e.message, 'err');
    finishTool(id, 'ERROR — the calendar event could not be created: ' + e.message, true);
  }
}

function finishTool(id, content, isError) {
  if (!awaiting || awaiting.results[id]) return;
  awaiting.results[id] = {
    type: 'tool_result',
    tool_use_id: id,
    content,
    ...(isError ? { is_error: true } : {}),
  };
  if (Object.keys(awaiting.results).length === awaiting.total) {
    const ordered = awaiting.ids.map(i => awaiting.results[i]);
    awaiting = null;
    history.push({ role: 'user', content: ordered });
    runLLM();
  }
}

async function onSend() {
  if (busy) return;
  const input = el('chatInput');
  const text = input.value.trim();
  if (!text) return;
  const s = getSettings();
  if (!s.apiKey) {
    addBubble('user', text);
    addBubble('assistant', 'Ke kopa API key pele — open Setup and paste your Anthropic API key, then Save.');
    switchView('setup');
    input.value = '';
    return;
  }
  input.value = '';
  input.style.height = 'auto';
  addBubble('user', text);

  if (awaiting) {
    // The user moved on while cards were still pending: resolve the open
    // ones as cancelled and bundle everything into one user turn.
    const ordered = awaiting.ids.map(tid => {
      if (awaiting.results[tid]) return awaiting.results[tid];
      const card = awaiting.cards[tid];
      if (card) setCardStatus(card, 'Cancelled', 'err');
      return {
        type: 'tool_result',
        tool_use_id: tid,
        content: 'The user moved on without confirming — treat this meeting as cancelled.',
      };
    });
    awaiting = null;
    history.push({ role: 'user', content: [...ordered, { type: 'text', text }] });
  } else {
    history.push({ role: 'user', content: [{ type: 'text', text }] });
  }
  await runLLM();
}

async function runLLM() {
  trimHistory();
  setThinking(true);
  const histLen = history.length;  // snapshot so a failed call can roll back
  let resp;
  try {
    resp = await callClaude(history);
  } catch (e) {
    setThinking(false);
    // Roll back the user turn (text or tool_results) that got no assistant
    // reply — otherwise the next send produces two consecutive user turns,
    // which the Anthropic API rejects with a 400 and breaks the chat.
    history.length = histLen;
    addBubble('assistant', '⚠️ ' + e.message);
    return;
  }
  setThinking(false);

  history.push({ role: 'assistant', content: resp.content });

  for (const block of resp.content) {
    if (block.type === 'text' && block.text.trim()) addBubble('assistant', block.text);
  }

  const toolUses = resp.content.filter(b => b.type === 'tool_use');
  if (!toolUses.length) return;

  awaiting = { ids: toolUses.map(b => b.id), results: {}, total: toolUses.length, cards: {} };
  const autoSend = !!getSettings().autoSend;
  for (const b of toolUses) {
    if (b.name !== 'create_meeting') {
      finishTool(b.id, 'Unknown tool "' + b.name + '" — only create_meeting is available.', true);
      continue;
    }
    const card = addCard(b);
    awaiting.cards[b.id] = card;
    if (autoSend) {
      card._actions.remove();
      executeMeeting(b.id, b.input, card);
    }
  }
}

/* ---------------- team view ---------------- */

function renderTeam() {
  const list = el('teamList');
  list.textContent = '';
  for (const [i, p] of getTeam().entries()) {
    const li = document.createElement('li');
    const left = document.createElement('div');
    const name = document.createElement('span');
    name.textContent = p.name;
    const email = document.createElement('span');
    email.className = 't-email';
    email.textContent = p.email;
    left.appendChild(name);
    left.appendChild(email);
    const del = document.createElement('button');
    del.textContent = '✕';
    del.setAttribute('aria-label', 'Remove ' + p.name);
    del.addEventListener('click', () => {
      const team = getTeam();
      team.splice(i, 1);
      saveTeam(team);
      renderTeam();
    });
    li.appendChild(left);
    li.appendChild(del);
    list.appendChild(li);
  }
}

function addTeamMember() {
  const name = el('teamName').value.trim();
  const email = el('teamEmail').value.trim();
  if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    toast('Eish — I need a first name and a valid email address.');
    return;
  }
  const team = getTeam().filter(p => p.email.toLowerCase() !== email.toLowerCase());
  team.push({ name, email });
  saveTeam(team);
  el('teamName').value = '';
  el('teamEmail').value = '';
  renderTeam();
}

function bulkImport() {
  const text = el('bulkBox').value;
  const found = [];
  for (const line of text.split(/\r?\n/)) {
    const l = line.trim();
    if (!l) continue;
    const m = l.match(/([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})/);
    if (!m) continue;
    const email = m[1];
    let name = l.replace(email, '').replace(/[<>,;|"\t]/g, ' ').replace(/\s+/g, ' ').trim();
    if (!name) name = email.split('@')[0];
    found.push({ name, email });
  }
  if (!found.length) { toast('No email addresses found in that list.'); return; }
  const team = getTeam();
  for (const p of found) {
    const idx = team.findIndex(t => t.email.toLowerCase() === p.email.toLowerCase());
    if (idx >= 0) team[idx] = p; else team.push(p);
  }
  saveTeam(team);
  el('bulkBox').value = '';
  renderTeam();
  toast('Sharp sharp! Imported ' + found.length + ' people.');
}

/* ---------------- setup view ---------------- */

function loadSettingsForm() {
  const s = getSettings();
  el('setApiKey').value = s.apiKey || '';
  el('setModel').value = s.model || DEFAULT_MODEL;
  el('setClientId').value = s.clientId || '';
  el('setTenantId').value = s.tenantId || '';
  el('setAutoSend').checked = !!s.autoSend;
}

function saveSettingsForm() {
  const prevClient = getSettings().clientId;
  const prevTenant = getSettings().tenantId;
  saveSettings({
    apiKey: el('setApiKey').value.trim(),
    model: el('setModel').value.trim() || DEFAULT_MODEL,
    clientId: el('setClientId').value.trim(),
    tenantId: el('setTenantId').value.trim(),
    autoSend: el('setAutoSend').checked,
  });
  const now = getSettings();
  // Rebuild the MSAL client if either ID changed — the authority URL bakes in
  // the tenant, so a corrected tenantId must not keep targeting the old one.
  if (prevClient !== now.clientId || prevTenant !== now.tenantId) {
    _pca = null;
    _pcaReady = null;
  }
  renderAuthStatus();
  toast('Go siame — settings saved on this phone.');
}

/* ---------------- navigation & wiring ---------------- */

function switchView(name) {
  for (const v of document.querySelectorAll('.view')) v.classList.remove('active');
  el('view-' + name).classList.add('active');
  for (const b of document.querySelectorAll('#tabs button')) {
    b.classList.toggle('active', b.dataset.view === name);
  }
}

let deferredInstall = null;
window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault();
  deferredInstall = e;
  el('installBtn').hidden = false;
});

function wire() {
  for (const b of document.querySelectorAll('#tabs button')) {
    b.addEventListener('click', () => switchView(b.dataset.view));
  }
  el('sendBtn').addEventListener('click', onSend);
  const input = el('chatInput');
  input.addEventListener('keydown', e => {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); onSend(); }
  });
  input.addEventListener('input', () => {
    input.style.height = 'auto';
    input.style.height = Math.min(input.scrollHeight, 110) + 'px';
  });
  for (const chip of document.querySelectorAll('.chip')) {
    chip.addEventListener('click', () => {
      el('chatInput').value = chip.textContent;
      el('chatInput').focus();
    });
  }
  el('teamAddBtn').addEventListener('click', addTeamMember);
  el('bulkBtn').addEventListener('click', bulkImport);
  el('saveBtn').addEventListener('click', saveSettingsForm);
  el('signInBtn').addEventListener('click', () => { saveSettingsForm(); signIn().catch(e => toast(e.errorMessage || e.message)); });
  el('signOutBtn').addEventListener('click', () => signOut().catch(e => toast(e.message)));
  el('installBtn').addEventListener('click', async () => {
    if (!deferredInstall) return;
    deferredInstall.prompt();
    deferredInstall = null;
    el('installBtn').hidden = true;
  });
}

async function main() {
  wire();
  renderTeam();
  loadSettingsForm();
  await initAuth();
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
  // first-run guidance
  const s = getSettings();
  if (!s.apiKey || !s.clientId) switchView('setup');
}

main();
