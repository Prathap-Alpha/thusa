'use strict';

/* ============================================================
   Thusa — chat-to-calendar assistant for Alpha Direct
   - LLM: Google Gemini (direct from browser, key in Setup)
     [AI vendor = Gemini per CFO override of AD-POL-AI-GOV-001, 2026-06-13]
   - Calendar: Microsoft Graph via MSAL.js (sign in once)
   All data stays in this browser's localStorage.
   ============================================================ */

const DEFAULT_MODEL = 'gemini-2.5-flash';
const TIMEZONE = 'Africa/Gaborone';
const GRAPH_SCOPES = ['User.Read', 'Calendars.ReadWrite'];
const S_KEY = 'thusa_settings';
const T_KEY = 'thusa_team';
const P_KEY = 'thusa_prompts';

// The 10 standard tap-to-fill quick meetings, seeded on first run.
const DEFAULT_PROMPTS = [
  'Set up a meeting with Kago at 5pm today about 3D accounts',
  'Meeting with Arun tomorrow morning re performance appraisal',
  '30 min with Unami at 8am Monday — staff dismissals',
  'Catch-up with the finance team Friday 10am',
  'Board meeting next Tuesday 2pm for 2 hours',
  '1-on-1 with Lakshmi Thursday 3pm',
  'Claims review with Underwriting Wednesday 9am',
  'Budget review Monday 11am — 45 min',
  'Quick call with Bharath today 4pm, 15 min',
  'Month-end close kickoff first of next month 9am',
];

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
function getPrompts() {
  // No stored list yet → start everyone off with the 10 standard ones.
  const raw = localStorage.getItem(P_KEY);
  if (raw === null) return DEFAULT_PROMPTS.slice();
  try { const a = JSON.parse(raw); return Array.isArray(a) ? a : DEFAULT_PROMPTS.slice(); }
  catch { return DEFAULT_PROMPTS.slice(); }
}
function savePrompts(list) {
  localStorage.setItem(P_KEY, JSON.stringify(list));
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

/* ---------------- Google Gemini ---------------- */

const GEMINI_TOOLS = [{
  name: 'create_meeting',
  description: 'Create a calendar event in the user\'s Microsoft 365 calendar and send invites to the attendees. Call this once per meeting the user asks for.',
  parameters: {
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

async function callGemini(contents) {
  const s = getSettings();
  const model = s.model || DEFAULT_MODEL;
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/'
    + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(s.apiKey);
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'content-type': 'application/json' },
    body: JSON.stringify({
      systemInstruction: { parts: [{ text: systemPrompt() }] },
      contents,
      tools: [{ functionDeclarations: GEMINI_TOOLS }],
      toolConfig: { functionCallingConfig: { mode: 'AUTO' } },
      generationConfig: { maxOutputTokens: 2048, temperature: 0.2 },
    }),
  });
  if (!res.ok) {
    let msg = 'Gemini API error ' + res.status;
    try { const j = await res.json(); msg = (j.error && j.error.message) || msg; } catch {}
    if (res.status === 400 && /API key/i.test(msg)) msg = 'Your Gemini API key was rejected — check it in Setup.';
    if (res.status === 403) msg = 'Your Gemini API key was rejected or lacks access — check it in Setup.';
    if (res.status === 429) msg = 'Rate limit reached — wait a few seconds and try again.';
    throw new Error(msg);
  }
  return res.json();
}

/* ---------------- chat engine ---------------- */

let history = [];      // Gemini-format `contents` for this session only
let awaiting = null;   // { order:[{localId,name}], results:{}, total, cards:{} } while cards pending
let busy = false;

function isTextUserTurn(c) {
  return c && c.role === 'user' && c.parts && c.parts[0] && typeof c.parts[0].text === 'string';
}

function trimHistory() {
  // Drop oldest complete turn-groups (user-text → model → function-responses)
  // without leaving a model turn or a dangling functionResponse at the head —
  // any of those would make Gemini reject the next request.
  while (history.length > 40) {
    history.shift();
    while (history.length && history[0].role === 'model') history.shift();
    while (history.length && history[0].role === 'user' &&
           history[0].parts && history[0].parts[0] && history[0].parts[0].functionResponse) {
      history.shift();
    }
  }
  // Safety: the first content sent to Gemini must be a plain-text user turn.
  while (history.length && !isTextUserTurn(history[0])) history.shift();
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

function addCard(localId, args) {
  const w = el('welcome');
  if (w) w.remove();
  const input = args || {};
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
    executeMeeting(localId, input, card);
  });
  cancelBtn.addEventListener('click', () => {
    actions.remove();
    setCardStatus(card, 'Cancelled', 'err');
    finishTool(localId, 'create_meeting', { result: 'The user cancelled this meeting — do not create it.' });
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

async function executeMeeting(localId, args, card) {
  setCardStatus(card, 'Sending…');
  try {
    const ev = await createEvent(args);
    setCardStatus(card, '✅ Invite sent', 'ok');
    finishTool(localId, 'create_meeting', { result: 'Success — the event was created and invites were sent.', webLink: ev.webLink || '' });
  } catch (e) {
    setCardStatus(card, '❌ ' + e.message, 'err');
    finishTool(localId, 'create_meeting', { error: 'The calendar event could not be created: ' + e.message });
  }
}

// Record a tool's outcome as a Gemini functionResponse part. Once every pending
// call has a result, the responses are sent back in their ORIGINAL order (Gemini
// matches parallel calls of the same name positionally — there are no call ids).
function finishTool(localId, name, responseObj) {
  if (!awaiting || awaiting.results[localId]) return;
  awaiting.results[localId] = { functionResponse: { name, response: responseObj } };
  if (Object.keys(awaiting.results).length === awaiting.total) {
    const parts = awaiting.order.map(o => awaiting.results[o.localId]);
    awaiting = null;
    history.push({ role: 'user', parts });
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
    addBubble('assistant', 'Ke kopa API key pele — open Setup and paste your Gemini API key, then Save.');
    switchView('setup');
    input.value = '';
    return;
  }
  input.value = '';
  input.style.height = 'auto';
  addBubble('user', text);

  if (awaiting) {
    // The user moved on while cards were still pending: resolve the open ones
    // as cancelled, then append the new text as a part in the same user turn
    // (keeps the model → user alternation Gemini expects).
    const parts = awaiting.order.map(o => {
      if (awaiting.results[o.localId]) return awaiting.results[o.localId];
      const card = awaiting.cards[o.localId];
      if (card) setCardStatus(card, 'Cancelled', 'err');
      return { functionResponse: { name: o.name, response: { result: 'The user moved on without confirming — treat this meeting as cancelled.' } } };
    });
    awaiting = null;
    history.push({ role: 'user', parts: [...parts, { text }] });
  } else {
    history.push({ role: 'user', parts: [{ text }] });
  }
  await runLLM();
}

async function runLLM() {
  trimHistory();
  setThinking(true);
  let resp;
  try {
    resp = await callGemini(history);
  } catch (e) {
    setThinking(false);
    // Roll back the one user turn (text or function-responses) the caller just
    // pushed and that got no model reply — otherwise the next send leaves two
    // consecutive user turns, which Gemini rejects and the chat breaks.
    // pop() is correct because every runLLM caller pushes exactly one Content,
    // and trimHistory only removes from the front, so that turn is always last.
    history.pop();
    addBubble('assistant', '⚠️ ' + e.message);
    return;
  }
  setThinking(false);

  const cand = resp.candidates && resp.candidates[0];
  const parts = (cand && cand.content && cand.content.parts) || [];
  if (!parts.length) {
    // Blocked, empty, or safety-stopped: drop the unanswered user turn too.
    history.pop();
    const why = (cand && cand.finishReason)
      || (resp.promptFeedback && resp.promptFeedback.blockReason)
      || 'no response';
    addBubble('assistant', '⚠️ Gemini returned nothing (' + why + '). Try rephrasing.');
    return;
  }

  // Echo the model turn back verbatim (role normalised to "model").
  history.push({ role: 'model', parts });

  for (const p of parts) {
    if (p.text && p.text.trim()) addBubble('assistant', p.text);
  }

  const calls = parts.filter(p => p.functionCall).map(p => p.functionCall);
  if (!calls.length) return;

  awaiting = { order: [], results: {}, total: calls.length, cards: {} };
  const autoSend = !!getSettings().autoSend;
  calls.forEach((fc, i) => {
    const localId = 'fc' + i;
    awaiting.order.push({ localId, name: fc.name });
    if (fc.name !== 'create_meeting') {
      finishTool(localId, fc.name, { error: 'Unknown tool "' + fc.name + '" — only create_meeting is available.' });
      return;
    }
    const args = fc.args || {};
    const card = addCard(localId, args);
    awaiting.cards[localId] = card;
    if (autoSend) {
      card._actions.remove();
      executeMeeting(localId, args, card);
    }
  });
}

/* ---------------- team view ---------------- */

let editingEmail = null;  // email of the row currently being edited, or null

function renderTeam() {
  const list = el('teamList');
  list.textContent = '';
  const team = getTeam().slice().sort((a, b) =>
    a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  const count = el('teamCount');
  if (count) count.textContent = team.length ? String(team.length) : '';
  for (const p of team) {
    const li = document.createElement('li');

    const left = document.createElement('div');
    left.className = 't-person';
    const avatar = document.createElement('span');
    avatar.className = 't-avatar';
    avatar.textContent = ((p.name.trim()[0]) || '?').toUpperCase();
    const info = document.createElement('div');
    const name = document.createElement('span');
    name.className = 't-name';
    name.textContent = p.name;
    const email = document.createElement('span');
    email.className = 't-email';
    email.textContent = p.email;
    info.appendChild(name);
    info.appendChild(email);
    left.appendChild(avatar);
    left.appendChild(info);

    const actions = document.createElement('div');
    actions.className = 't-actions';

    const edit = document.createElement('button');
    edit.className = 't-edit';
    edit.textContent = '✎';
    edit.setAttribute('aria-label', 'Edit ' + p.name);
    edit.addEventListener('click', () => startEditMember(p));

    const del = document.createElement('button');
    del.className = 't-del';
    del.textContent = '✕';
    del.setAttribute('aria-label', 'Remove ' + p.name);
    // Delete/edit by email, not list index — the displayed order is sorted and
    // would not line up with the stored array's indices.
    del.addEventListener('click', () => {
      if (editingEmail && editingEmail.toLowerCase() === p.email.toLowerCase()) cancelEdit();
      saveTeam(getTeam().filter(t => t.email.toLowerCase() !== p.email.toLowerCase()));
      renderTeam();
    });

    actions.appendChild(edit);
    actions.appendChild(del);
    li.appendChild(left);
    li.appendChild(actions);
    if (editingEmail && editingEmail.toLowerCase() === p.email.toLowerCase()) {
      li.classList.add('editing');
    }
    list.appendChild(li);
  }
}

// Load a person into the top fields and switch the form into "edit" mode.
function startEditMember(p) {
  editingEmail = p.email;
  el('teamName').value = p.name;
  el('teamEmail').value = p.email;
  el('teamAddBtn').textContent = 'Save';
  el('teamCancelEdit').hidden = false;
  renderTeam();                 // highlight the row being edited
  el('view-team').scrollTop = 0;
  el('teamName').focus();
}

function cancelEdit() {
  editingEmail = null;
  el('teamName').value = '';
  el('teamEmail').value = '';
  el('teamAddBtn').textContent = 'Add';
  el('teamCancelEdit').hidden = true;
}

function addTeamMember() {
  const name = el('teamName').value.trim();
  const email = el('teamEmail').value.trim();
  if (!name || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    toast('Eish — I need a first name and a valid email address.');
    return;
  }
  // Drop the old row being edited (if any) and any existing row with this email,
  // then add the new/updated one — so editing and de-duping both just work.
  const wasEditing = !!editingEmail;
  let team = getTeam().filter(p =>
    p.email.toLowerCase() !== email.toLowerCase() &&
    (!editingEmail || p.email.toLowerCase() !== editingEmail.toLowerCase()));
  team.push({ name, email });
  saveTeam(team);
  cancelEdit();                 // resets fields + button, clears editingEmail
  renderTeam();
  if (wasEditing) {
    toast('Updated ' + name + '.');
  } else {
    // Keep the cursor in First name so several people can be added quickly.
    el('teamName').focus();
  }
}

// Dump the whole current directory into the bulk box so it can be amended at once.
function loadCurrentList() {
  const team = getTeam().slice().sort((a, b) =>
    a.name.localeCompare(b.name, undefined, { sensitivity: 'base' }));
  if (!team.length) { toast('Your list is empty — nothing to load yet.'); return; }
  el('bulkBox').value = team.map(p => `${p.name}, ${p.email}`).join('\n');
  el('bulkBox').focus();
  toast('Loaded ' + team.length + ' people — edit the lines, then Import to update.');
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

/* ---------------- quick meetings (chips + editor) ---------------- */

// Build the tap-to-fill chips on the Chat welcome screen from the saved list.
function renderChips() {
  const wrap = el('chips');
  if (!wrap) return;             // welcome already cleared by the first message
  wrap.textContent = '';
  for (const text of getPrompts()) {
    const b = document.createElement('button');
    b.className = 'chip';
    b.textContent = text;
    b.addEventListener('click', () => {
      // One tap = set it up. onSend routes it to the AI, which replies with a
      // meeting card you confirm — so no invite goes out without your tap.
      switchView('chat');
      el('chatInput').value = text;
      onSend();
    });
    wrap.appendChild(b);
  }
}

// The editable list in Setup. Each row edits one prompt in place (by index, so
// duplicates are fine); typing saves on blur and refreshes the chips only —
// the editor is NOT re-rendered on keystroke, so the field keeps focus.
function renderPromptEditor() {
  const list = el('promptList');
  if (!list) return;
  list.textContent = '';
  const prompts = getPrompts();
  prompts.forEach((text, i) => {
    const li = document.createElement('li');

    const input = document.createElement('input');
    input.className = 'p-edit';
    input.value = text;
    input.addEventListener('change', () => {
      const cur = getPrompts();
      const v = input.value.trim();
      if (!v) { cur.splice(i, 1); savePrompts(cur); renderPromptEditor(); }
      else { cur[i] = v; savePrompts(cur); }
      renderChips();
    });

    const del = document.createElement('button');
    del.className = 't-del';
    del.textContent = '✕';
    del.setAttribute('aria-label', 'Remove this quick meeting');
    del.addEventListener('click', () => {
      const cur = getPrompts();
      cur.splice(i, 1);
      savePrompts(cur);
      renderPromptEditor();
      renderChips();
    });

    li.appendChild(input);
    li.appendChild(del);
    list.appendChild(li);
  });
}

function addPrompt() {
  const field = el('promptNew');
  const v = field.value.trim();
  if (!v) { toast('Type the quick meeting first.'); return; }
  const cur = getPrompts();
  cur.push(v);
  savePrompts(cur);
  field.value = '';
  renderPromptEditor();
  renderChips();
  field.focus();
}

function resetPrompts() {
  savePrompts(DEFAULT_PROMPTS.slice());
  renderPromptEditor();
  renderChips();
  toast('Reset to the 10 standard quick meetings.');
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

/* ---------------- sharing / invite links ---------------- */

// An invite link may carry the Microsoft app IDs (?client=…&tenant=…&model=…).
// These are NOT secrets — an Azure SPA client id is public by design. We apply
// them on load so a recipient skips re-typing section 2, then strip them from
// the URL (so they don't linger and so the MSAL redirect URI stays clean).
function applyInviteParams() {
  const p = new URLSearchParams(location.search);
  const patch = {};
  const cid = p.get('client'); if (cid) patch.clientId = cid.trim();
  const tid = p.get('tenant'); if (tid) patch.tenantId = tid.trim();
  const mdl = p.get('model');  if (mdl) patch.model = mdl.trim();
  if (Object.keys(patch).length) {
    saveSettings(patch);              // only Microsoft IDs + model — never a key/team
    _pca = null; _pcaReady = null;    // rebuild MSAL with the invited authority
  }
  // Strip the params so they don't linger / re-apply. Never let this abort
  // startup — replaceState can throw in sandboxed/proxied contexts.
  if (location.search) {
    try { history.replaceState(null, '', location.origin + location.pathname); }
    catch (e) { /* non-fatal — the params are harmless if they remain */ }
  }
}

// Build a share link that carries ONLY the non-secret Microsoft app IDs. The
// Gemini API key and the team directory are deliberately left out.
function buildShareUrl() {
  const s = getSettings();
  const base = location.origin + location.pathname;
  const p = new URLSearchParams();
  if (s.clientId) p.set('client', s.clientId);
  if (s.tenantId) p.set('tenant', s.tenantId);
  if (s.model && s.model !== DEFAULT_MODEL) p.set('model', s.model);
  const q = p.toString();
  return q ? base + '?' + q : base;
}

async function shareApp() {
  const url = buildShareUrl();
  const data = {
    title: 'Thusa — Meeting Assistant',
    text: 'Install Thusa to schedule meetings by chat. Open the link in Chrome, tap “Add to Home screen”, then add your own AI key and sign in.',
    url,
  };
  if (navigator.share) {
    try { await navigator.share(data); return; }
    catch (e) { if (e && e.name === 'AbortError') return; }  // user dismissed the sheet
  }
  try {
    await navigator.clipboard.writeText(url);
    toast('Link copied — paste it to whoever you want to share with.');
  } catch {
    toast(url);
  }
}

/* ---------------- navigation & wiring ---------------- */

function switchView(name) {
  for (const v of document.querySelectorAll('.view')) v.classList.remove('active');
  el('view-' + name).classList.add('active');
  for (const b of document.querySelectorAll('#tabs button')) {
    b.classList.toggle('active', b.dataset.view === name);
  }
}

/* ---------------- install (platform-aware) ---------------- */

let deferredInstall = null;

function detectOS() {
  const ua = navigator.userAgent || '';
  // iPadOS 13+ reports as "MacIntel" but has a touch screen — catch it too.
  const isIOS = /iPad|iPhone|iPod/.test(ua) ||
    (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
  const isAndroid = /Android/.test(ua);
  return { isIOS, isAndroid };
}

function isStandalone() {
  return window.matchMedia('(display-mode: standalone)').matches ||
    navigator.standalone === true;  // iOS Safari's own flag
}

// Highlight the block for the phone we're on; hide the whole card once installed.
function setupInstallUI() {
  const card = el('installCard');
  if (!card) return;
  if (isStandalone()) { card.hidden = true; return; }  // already a home-screen app
  const { isIOS, isAndroid } = detectOS();
  if (isAndroid) el('osAndroid').classList.add('match');
  if (isIOS) el('osApple').classList.add('match');
}

// Fired only on Chromium (Android/desktop) — iOS Safari has no such event,
// which is exactly why the Apple block gives manual Share-sheet steps instead.
window.addEventListener('beforeinstallprompt', e => {
  e.preventDefault();
  deferredInstall = e;
  el('installBtn').hidden = false;
  const a = el('installBtnAndroid'); if (a) a.hidden = false;
});

window.addEventListener('appinstalled', () => {
  deferredInstall = null;
  el('installBtn').hidden = true;
  const a = el('installBtnAndroid'); if (a) a.hidden = true;
  const card = el('installCard'); if (card) card.hidden = true;
});

async function promptInstall() {
  if (!deferredInstall) {
    toast('Use the Chrome ⋮ menu → "Add to Home screen".');
    return;
  }
  deferredInstall.prompt();
  deferredInstall = null;
  el('installBtn').hidden = true;
  const a = el('installBtnAndroid'); if (a) a.hidden = true;
}

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
  el('promptAddBtn').addEventListener('click', addPrompt);
  el('promptResetBtn').addEventListener('click', resetPrompts);
  el('promptNew').addEventListener('keydown', e => {
    if (e.key === 'Enter') { e.preventDefault(); addPrompt(); }
  });
  el('teamAddBtn').addEventListener('click', addTeamMember);
  el('teamCancelEdit').addEventListener('click', cancelEdit);
  el('bulkBtn').addEventListener('click', bulkImport);
  el('bulkLoadBtn').addEventListener('click', loadCurrentList);
  // Enter in either field adds/saves — quick when entering several people.
  for (const fid of ['teamName', 'teamEmail']) {
    el(fid).addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); addTeamMember(); }
    });
  }
  el('saveBtn').addEventListener('click', saveSettingsForm);
  el('signInBtn').addEventListener('click', () => { saveSettingsForm(); signIn().catch(e => toast(e.errorMessage || e.message)); });
  el('signOutBtn').addEventListener('click', () => signOut().catch(e => toast(e.message)));
  el('installBtn').addEventListener('click', promptInstall);
  el('installBtnAndroid').addEventListener('click', promptInstall);
  el('shareBtn').addEventListener('click', () => shareApp());
}

async function main() {
  wire();
  applyInviteParams();   // honour ?client=&tenant= from an invite link first
  setupInstallUI();      // highlight Android/Apple steps for the current phone
  renderChips();         // tap-to-fill quick meetings on the chat screen
  renderPromptEditor();  // their editor in Setup
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
