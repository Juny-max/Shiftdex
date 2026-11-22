import XLSX from 'xlsx-js-style';
import { initialState, staff, pods, cellFor, setCell, monthName, roleFor, cycle } from '../shared/roster.js';
import seedRoster from '../shared/seed-roster.json';
import shiftDeskMark from '../assets/shift-desk-mark.svg';
import './style.css';
import './polish.css';
import './pages.css';

// Core State
for (const key of ['admin-session', 'admin-session-expires']) {
  const legacyValue = localStorage.getItem(key);
  if (legacyValue !== null) {
    if (sessionStorage.getItem(key) === null) sessionStorage.setItem(key, legacyValue);
    localStorage.removeItem(key);
  }
}
const saved = localStorage.getItem('shift-roster');
let state = saved ? JSON.parse(saved) : seedRoster;
if (!state.assignments || Object.keys(state.assignments).length < 200) {
  state = {
    ...seedRoster,
    assignments: { ...seedRoster.assignments, ...(state.assignments || {}) },
    overrides: { ...(state.overrides || {}) },
    staffMonthly: { ...seedRoster.staffMonthly, ...(state.staffMonthly || {}) },
    customStaff: [...(state.customStaff || [])],
    inactiveStaffIds: [...(state.inactiveStaffIds || [])],
    coverAssignments: { ...(state.coverAssignments || {}) },
    coverTombstones: [...(state.coverTombstones || [])]
  };
  localStorage.setItem('shift-roster', JSON.stringify(state));
}
state.coverAssignments ||= {};
state.coverTombstones ||= [];
state.customStaff ||= [];
state.inactiveStaffIds ||= [];
for (const person of state.customStaff) {
  const existing = staff.find(existing => existing.id === person.id);
  if (!existing) staff.push(person);
  else Object.assign(existing, person);
}
const markObj = staff.find(s => s.id === 'mark');
if (markObj && markObj.status !== 'active') markObj.status = 'pending';
let month = localStorage.getItem('shift-roster-month') || '2026-11';
if (month < '2026-11') month = '2026-11';
let selected = null;
let undo = [];
let redo = [];
let requests = [];
let availableStaff = [];
let desktopUpdateState = { status: 'idle', currentVersion: '', version: null, percent: 0, message: 'Update checks run automatically every six hours.' };
let manualCheckUntil = 0;
const API = 'https://shift-roster-api.maxgrip500.workers.dev';
let activePage = 'overview';
let rosterFilter = 'all';
let rosterSearch = '';
let rosterShowPast = false;
let peopleFilter = 'all';
let peopleSearch = '';

const app = document.querySelector('#app');

// Minimalist, Clean Stroke SVG Icons
const ICONS = {
  overview: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="3" width="7" height="7" rx="1.5"/><rect x="14" y="14" width="7" height="7" rx="1.5"/><rect x="3" y="14" width="7" height="7" rx="1.5"/></svg>`,
  roster: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
  people: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>`,
  requests: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>`,
  search: `<svg class="search-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>`,
  arrowUpRight: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="7" y1="17" x2="17" y2="7"/><polyline points="7 7 17 7 17 17"/></svg>`,
  chevronLeft: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="15 18 9 12 15 6"/></svg>`,
  chevronRight: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="9 18 15 12 9 6"/></svg>`,
  copy: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>`,
  download: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>`,
  upload: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/></svg>`,
  plus: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>`,
  clock: `<svg class="analog-clock-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9.4"/><path d="M12 4.2v1M19.8 12h-1M12 19.8v-1M4.2 12h1" opacity=".45"/><line class="clock-hour-hand" x1="12" y1="12" x2="15.2" y2="13.5"/><line class="clock-minute-hand" x1="12" y1="12" x2="12" y2="6.5"/><line class="clock-second-hand" x1="12" y1="13" x2="12" y2="5.2"/><circle cx="12" cy="12" r="1" fill="currentColor" stroke="none"/></svg>`,
  undo: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/></svg>`,
  redo: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>`,
  settings: `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/></svg>`
};

function formatMonth(key) {
  const [y, m] = key.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, 1)).toLocaleDateString('en-US', { month: 'long', year: 'numeric', timeZone: 'UTC' });
}

function dateKey(day) {
  return `${month}-${String(day).padStart(2, '0')}`;
}

function escapeHtml(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function personFor(s, key) {
  const current = state.staffMonthly?.[`${key}:${s.id}`];
  if (current) return { ...s, ...current };
  const previous = Object.entries(state.staffMonthly || {})
    .filter(([entry, value]) => entry.endsWith(`:${s.id}`) && entry.slice(0, 7) < key && value?.podId)
    .sort(([a], [b]) => a.localeCompare(b))
    .at(-1)?.[1];
  return previous ? { ...s, podId: previous.podId, rowOrder: previous.rowOrder } : { ...s };
}

function activeStaff() {
  const inactive = new Set(state.inactiveStaffIds || []);
  return staff.filter(person => !inactive.has(person.id) && person.status !== 'pending' && person.status !== 'archived' && person.status !== 'inactive');
}

function pendingStaff() {
  const inactive = new Set(state.inactiveStaffIds || []);
  return staff.filter(person => !inactive.has(person.id) && person.status === 'pending');
}

function archivedStaff() {
  const inactive = new Set(state.inactiveStaffIds || []);
  return staff.filter(person => inactive.has(person.id) || person.status === 'archived' || person.status === 'inactive');
}

async function syncDirectory() {
  const token = sessionStorage.getItem('admin-session') || '';
  if (!token) return;
  try {
    const res = await fetch(API + '/api/admin/directory', {
      headers: { authorization: 'Bearer ' + token }
    });
    if (!res.ok) return;
    const data = await res.json();
    if (!Array.isArray(data.staff)) return;
    let changed = false;
    for (const item of data.staff) {
      let existing = staff.find(s => s.id === item.id);
      const isArchived = (item.status === 'archived' || item.status === 'inactive');
      const isPending = !isArchived && (item.status === 'pending');
      const resolvedStatus = isArchived ? 'archived' : (isPending ? 'pending' : 'active');

      if (isArchived) {
        state.inactiveStaffIds ||= [];
        if (!state.inactiveStaffIds.includes(item.id)) {
          state.inactiveStaffIds.push(item.id);
          changed = true;
        }
      } else {
        if (state.inactiveStaffIds?.includes(item.id)) {
          state.inactiveStaffIds = state.inactiveStaffIds.filter(id => id !== item.id);
          changed = true;
        }
      }

      if (!existing) {
        existing = {
          id: item.id,
          name: item.full_name,
          displayName: item.full_name,
          role: item.role_tag,
          podId: item.pod_id || undefined,
          phone: item.phone || undefined,
          effectiveFrom: item.effective_from || '2026-11-01',
          email: item.email || undefined,
          pendingEmail: item.pending_email || undefined,
          mustChangePassword: item.must_change_password || 0,
          status: resolvedStatus
        };
        staff.push(existing);
        state.customStaff.push(existing);
        changed = true;
      } else {
        if (existing.status !== resolvedStatus) {
          existing.status = resolvedStatus;
          changed = true;
        }
        if (item.email && existing.email !== item.email) {
          existing.email = item.email;
          changed = true;
        }
        if (existing.pendingEmail !== (item.pending_email || undefined)) {
          existing.pendingEmail = item.pending_email || undefined;
          changed = true;
        }
        if (existing.mustChangePassword !== item.must_change_password) {
          existing.mustChangePassword = item.must_change_password;
          changed = true;
        }
        const customItem = state.customStaff.find(s => s.id === item.id);
        if (customItem) {
          customItem.status = resolvedStatus;
          if (item.email) customItem.email = item.email;
          customItem.pendingEmail = item.pending_email || undefined;
          customItem.mustChangePassword = item.must_change_password;
        }
      }
    }
    if (changed) {
      save();
      render();
    }
  } catch (err) {
    console.warn('Could not sync staff directory:', err);
  }
}

function orderedPeople(key) {
  const people = activeStaff().map(s => personFor(s, key));
  const podIndex = id => pods.findIndex(p => p.id === id);
  const core = people.filter(s => s.podId).sort((a, b) => podIndex(a.podId) - podIndex(b.podId) || (a.rowOrder ?? staff.findIndex(x => x.id === a.id)) - (b.rowOrder ?? staff.findIndex(x => x.id === b.id)));
  const support = people.filter(s => !s.podId);
  return [...core, ...support];
}

function save() {
  localStorage.setItem('shift-roster', JSON.stringify(state));
  render();
}

function validateMonth(st, m) {
  const [y, mo] = m.split('-').map(Number);
  const n = new Date(y, mo, 0).getDate();
  const people = activeStaff().map(s => personFor(s, m)).filter(s => s.podId);
  const issues = [];
  for (let d = 1; d <= n; d++) {
    const dKey = `${m}-${String(d).padStart(2, '0')}`;
    const vals = people.map(s => ({ s, v: cellFor(st, s.id, dKey) }));
    const a = vals.filter(x => x.v === 'M');
    const b = vals.filter(x => x.v === 'N');
    // Operational coverage: A true conflict is when a shift is completely uncovered (0 on duty).
    // NSPs (National Service Personnel) provide added rotation, but 1+ on duty maintains operations.
    if (a.length === 0 || b.length === 0) {
      issues.push(`${dKey}: Shift uncovered (Day: ${a.length}, Night: ${b.length})`);
    }
  }
  return [...new Set(issues)];
}

function getMonthStats(m) {
  const [y, mo] = m.split('-').map(Number);
  const days = new Date(y, mo, 0).getDate();
  const core = activeStaff().filter(s => s.podId);
  let mCount = 0, nCount = 0, offCount = 0;
  for (let d = 1; d <= days; d++) {
    const k = `${m}-${String(d).padStart(2, '0')}`;
    for (const s of core) {
      const val = cellFor(state, s.id, k);
      if (val === 'M') mCount++;
      else if (val === 'N') nCount++;
      else offCount++;
    }
  }
  const totalSlots = core.length * days;
  const issues = validateMonth(state, m);
  return {
    days,
    mCount,
    nCount,
    offCount,
    totalSlots,
    totalShifts: mCount + nCount,
    hours: (mCount + nCount) * 12,
    mPct: Math.round(mCount / totalSlots * 100),
    nPct: Math.round(nCount / totalSlots * 100),
    offPct: Math.round(offCount / totalSlots * 100),
    issues
  };
}

function getTodayInfo() {
  const now = new Date();
  const todayStr = now.toISOString().slice(0, 10);
  const [curY, curM, curD] = todayStr.split('-').map(Number);
  const isCurrentMonth = (month === `${curY}-${String(curM).padStart(2, '0')}`);
  const todayDay = isCurrentMonth ? curD : 20;
  const todayKeyStr = `${month}-${String(todayDay).padStart(2, '0')}`;
  const hour = now.getHours();
  const isMorningTime = (hour >= 6 && hour < 18);

  const core = activeStaff().filter(s => s.podId);
  const morningStaff = core.filter(s => cellFor(state, s.id, todayKeyStr) === 'M');
  const nightStaff = core.filter(s => cellFor(state, s.id, todayKeyStr) === 'N');
  const morningPod = pods.find(p => morningStaff.some(s => s.podId === p.id));
  const nightPod = pods.find(p => nightStaff.some(s => s.podId === p.id));

  const activeShiftStaff = isMorningTime ? morningStaff : nightStaff;
  const activePod = isMorningTime ? morningPod : nightPod;

  return {
    isCurrentMonth,
    todayDay,
    todayKeyStr,
    hour,
    isMorningTime,
    morningStaff,
    nightStaff,
    morningPod,
    nightPod,
    activeStaff: activeShiftStaff,
    activePod
  };
}

function toast(message) {
  const el = document.querySelector('#toast');
  if (!el) return;
  el.textContent = message;
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 3800);
}

function shiftDateOffset(date, amount) {
  const d = new Date(`${date}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + amount);
  return d.toISOString().slice(0, 10);
}

function previewShiftEdit(staffId, date, value, coverWorkerId = '') {
  const target = staff.find(s => s.id === staffId);
  const originalCover = state.coverAssignments?.[`${date}:${staffId}`];
  const next = JSON.parse(JSON.stringify(state));
  next.coverAssignments ||= {};
  next.coverTombstones ||= [];
  const affected = new Set([staffId]);
  const hard = [], warnings = [], info = [];
  const currentValue = originalCover?.shiftCode || cellFor(state, staffId, date);

  if (originalCover) {
    setCell(next, originalCover.coverWorkerId, date, 'OFF');
    delete next.coverAssignments[`${date}:${staffId}`];
    if (!coverWorkerId && originalCover.id && !next.coverTombstones.some(item => item.id === originalCover.id)) {
      next.coverTombstones.push({ id: originalCover.id, date });
    }
    affected.add(originalCover.coverWorkerId);
  }

  if (coverWorkerId) {
    if (!target?.podId || !['M', 'N'].includes(value)) {
      hard.push('A stand-in can only cover a Morning or Night shift.');
    } else {
      const available = cellFor(next, coverWorkerId, date);
      if (available !== 'OFF') {
        const coverName = staff.find(s => s.id === coverWorkerId)?.name || 'Selected worker';
        hard.push(`${coverName} is already assigned ${available} on ${date} and cannot cover this shift.`);
      }
      setCell(next, staffId, date, 'OFF');
      setCell(next, coverWorkerId, date, value);
      next.coverAssignments[`${date}:${staffId}`] = { id: originalCover?.id || crypto.randomUUID(), coverWorkerId, shiftCode: value };
      affected.add(coverWorkerId);
      const coverName = staff.find(s => s.id === coverWorkerId)?.name || 'Selected worker';
      info.push(`${coverName} will cover ${target.displayName || target.name}'s ${value === 'M' ? 'Morning' : 'Night'} shift on ${date}.`);
    }
  } else {
    setCell(next, staffId, date, value);
    info.push(`${target.displayName || target.name}: ${currentValue} → ${value} on ${date}.`);
  }

  for (const id of affected) {
    const person = staff.find(s => s.id === id);
    if (!person?.podId) continue;
    for (const day of [shiftDateOffset(date, -1), date]) {
      const tomorrow = shiftDateOffset(day, 1);
      const todayCode = cellFor(next, id, day), nextCode = cellFor(next, id, tomorrow);
      if ((todayCode === 'M' && nextCode === 'N') || (todayCode === 'N' && nextCode === 'M')) {
        hard.push(`${person.name} would work ${todayCode} on ${day} and ${nextCode} on ${tomorrow} without rest.`);
      }
    }
  }

  const core = activeStaff().map(s => personFor(s, month)).filter(s => s.podId);
  for (const code of ['M', 'N']) {
    const assigned = core.filter(s => cellFor(next, s.id, date) === code);
    if (assigned.length !== 3) {
      const names = assigned.length ? assigned.map(s => `${s.displayName || s.name} (${s.role})`).join(', ') : 'nobody';
      warnings.push(`${code === 'M' ? 'Morning' : 'Night'} coverage would be ${assigned.length}/3. Assigned: ${names}.`);
    }
    for (const role of ['Team Lead', 'NSP-SECURITY']) {
      if (!assigned.some(s => s.role === role)) {
        const names = assigned.map(s => s.displayName || s.name).join(', ') || 'nobody';
        warnings.push(`${code === 'M' ? 'Morning' : 'Night'} would have no ${role}; remaining on shift: ${names}.`);
      }
    }
  }

  return { next, info, hard: [...new Set(hard)], warnings: [...new Set(warnings)] };
}

function openShiftEditor(staffId, day, requestedValue = null) {
  const target = personFor(staff.find(s => s.id === staffId), month);
  if (!target) return;
  const date = dateKey(day), coverKey = `${date}:${staffId}`;
  const currentCover = state.coverAssignments?.[coverKey];
  const currentValue = currentCover?.shiftCode || cellFor(state, staffId, date);
  const editor = document.querySelector('#shift-edit-dialog');
  const coverOptions = activeStaff().filter(s => s.podId && s.id !== staffId && date >= (s.effectiveFrom || '0000-00-00')).map(s => {
    const working = cellFor(state, s.id, date);
    const assignmentHint = currentCover?.coverWorkerId === s.id ? 'current stand-in' : working === 'OFF' ? 'Rest day' : `already on ${working}`;
    const label = `${s.displayName || s.name} · ${s.role} · ${assignmentHint}`;
    return `<option value="${s.id}" ${currentCover?.coverWorkerId === s.id ? 'selected' : ''}>${escapeHtml(label)}</option>`;
  }).join('');

  editor.innerHTML = `<form method="dialog" class="shift-edit-form">
    <div class="shift-edit-head"><div><span class="shift-edit-eyebrow">ROSTER EDIT</span><h2>Edit shift</h2><p>${escapeHtml(target.displayName || target.name)} · ${formatMonth(month)} ${day}</p></div><button type="button" class="shift-edit-close" id="shift-edit-close" aria-label="Close">×</button></div>
    <label class="shift-edit-field">Assignment<select id="shift-edit-value"><option value="M">Morning · 06:00–18:00</option><option value="N">Night · 18:00–06:00</option><option value="OFF">Rest day</option><option value="LEAVE">Leave</option></select></label>
    ${target.podId ? `<label class="shift-edit-field">Stand-in<select id="shift-edit-cover"><option value="">No stand-in</option>${coverOptions}</select></label><p class="shift-edit-hint">Select a colleague to move this shift to them. Their own roster cell will be updated too.</p>` : ''}
    <section class="shift-conflict-preview" id="shift-conflict-preview" aria-live="polite"></section>
    <div class="shift-edit-actions"><button type="button" class="btn-clean" id="shift-edit-cancel">Cancel</button><button type="button" class="btn-clean primary" id="shift-edit-apply">Apply change</button></div>
  </form>`;
  const valueSelect = editor.querySelector('#shift-edit-value');
  valueSelect.value = ['M', 'N', 'OFF', 'LEAVE'].includes(requestedValue) ? requestedValue : currentValue;
  const coverSelect = editor.querySelector('#shift-edit-cover');
  const refreshPreview = () => {
    const result = previewShiftEdit(staffId, date, valueSelect.value, coverSelect?.value || '');
    const box = editor.querySelector('#shift-conflict-preview');
    const items = [
      ...result.hard.map(text => `<li class="conflict-hard">${escapeHtml(text)}</li>`),
      ...result.warnings.map(text => `<li class="conflict-warning">${escapeHtml(text)}</li>`)
    ];
    box.className = `shift-conflict-preview ${result.hard.length ? 'has-hard-conflict' : result.warnings.length ? 'has-warning' : 'is-clear'}`;
    box.innerHTML = `<b>${result.hard.length ? 'Cannot apply safely' : result.warnings.length ? 'Coverage preview' : 'No conflicts found'}</b><p class="shift-change-summary">${result.info.map(escapeHtml).join(' ')}</p>${items.length ? `<ul>${items.join('')}</ul>` : '<p>This change keeps coverage and rest rules clear.</p>'}`;
    editor.querySelector('#shift-edit-apply').disabled = result.hard.length > 0;
  };
  valueSelect.onchange = refreshPreview;
  if (coverSelect) coverSelect.onchange = refreshPreview;
  editor.querySelector('#shift-edit-close').onclick = () => editor.close();
  editor.querySelector('#shift-edit-cancel').onclick = () => editor.close();
  editor.querySelector('#shift-edit-apply').onclick = () => {
    const result = previewShiftEdit(staffId, date, valueSelect.value, coverSelect?.value || '');
    if (result.hard.length) { refreshPreview(); return; }
    undo.push(JSON.stringify(state));
    redo = [];
    state = result.next;
    editor.close();
    save();
    const coverName = coverSelect?.value ? staff.find(s => s.id === coverSelect.value)?.name : '';
    toast(coverName ? `${coverName} assigned as stand-in for ${target.displayName || target.name}.` : 'Shift updated.');
  };
  refreshPreview();
  editor.showModal();
}

function openLeaveDialog() {
  const dialog = document.querySelector('#leave-dialog');
  const people = orderedPeople(month);
  const today = new Date().toISOString().slice(0, 10);
  const firstDay = month === today.slice(0, 7) ? today : `${month}-01`;
  dialog.innerHTML = `<form method="dialog" class="shift-edit-form" id="leave-form">
    <div class="shift-edit-head"><div><span class="shift-edit-eyebrow">TIME OFF</span><h2>Schedule leave</h2><p>Mark a day or a date range as leave on the roster.</p></div><button type="button" class="shift-edit-close" id="leave-close" aria-label="Close">×</button></div>
    <label class="shift-edit-field">Team member<select id="leave-person">${people.map(person => `<option value="${person.id}">${escapeHtml(person.displayName || person.name)}</option>`).join('')}</select></label>
    <div class="leave-date-fields"><label class="shift-edit-field">First day<input id="leave-start" type="date" value="${firstDay}" required></label><label class="shift-edit-field">Last day<input id="leave-end" type="date" value="${firstDay}" required></label></div>
    <p class="shift-edit-hint">Each day in this range will show as Leave. You can undo the change from the schedule.</p>
    <div class="shift-edit-actions"><button type="button" class="btn-clean" id="leave-cancel">Cancel</button><button type="submit" class="btn-clean primary">Save leave</button></div>
  </form>`;
  const close = () => dialog.close();
  dialog.querySelector('#leave-close').onclick = close;
  dialog.querySelector('#leave-cancel').onclick = close;
  dialog.querySelector('#leave-form').onsubmit = event => {
    event.preventDefault();
    const personId = dialog.querySelector('#leave-person').value;
    const start = dialog.querySelector('#leave-start').value;
    const end = dialog.querySelector('#leave-end').value;
    if (!personId || !start || !end || end < start) {
      toast('Choose a valid leave date range.');
      return;
    }
    undo.push(JSON.stringify(state));
    redo = [];
    let count = 0;
    for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
      const date = day.toISOString().slice(0, 10);
      setCell(state, personId, date, 'LEAVE');
      count++;
    }
    close();
    save();
    toast(`Saved ${count} leave day${count === 1 ? '' : 's'} locally. Publish schedule to sync online.`);
  };
  dialog.showModal();
}

function openApplyRuleDialog() {
  const dialog = document.querySelector('#apply-rule-dialog');
  if (!dialog) return;

  const today = new Date().toISOString().slice(0, 10);
  const firstDay = month === today.slice(0, 7) ? today : `${month}-01`;
  const [y, mo] = month.split('-').map(Number);
  const lastDay = `${month}-${String(new Date(y, mo, 0).getDate()).padStart(2, '0')}`;

  const renderRuleForm = (ruleType) => {
    let specificFields = '';

    if (ruleType === 'holiday') {
      specificFields = `
        <label class="shift-edit-field">Target Personnel
          <select id="rule-target">
            <option value="all">All Personnel (Core & Support · 17 members)</option>
            <option value="support">Day Support Staff Only (5 members)</option>
            <option value="core">Shift Crew (All 4 Teams · 12 members)</option>
            <option value="orange">Team 1 Only</option>
            <option value="green">Team 2 Only</option>
            <option value="blue">Team 3 Only</option>
            <option value="red">Team 4 Only</option>
          </select>
        </label>
        <div class="leave-date-fields">
          <label class="shift-edit-field">Start Date
            <input id="rule-start" type="date" value="${firstDay}" required>
          </label>
          <label class="shift-edit-field">End Date
            <input id="rule-end" type="date" value="${firstDay}" required>
          </label>
        </div>
        <div class="leave-date-fields">
          <label class="shift-edit-field">Assignment Type
            <select id="rule-shift-code">
              <option value="OFF">Rest / Day Off (OFF)</option>
              <option value="LEAVE">Official Leave (LEAVE)</option>
            </select>
          </label>
          <label class="shift-edit-field">Reason / Event Name
            <input id="rule-note" type="text" placeholder="e.g. Public Holiday, Maintenance" value="Public Holiday">
          </label>
        </div>
      `;
    } else if (ruleType === 'bulk_shift') {
      specificFields = `
        <label class="shift-edit-field">Target Group / Role
          <select id="rule-target">
            <option value="all">All Personnel</option>
          <option value="leads">All Team Leads</option>
            <option value="nsps">All NSP Personnel (NSP & NSP-Security)</option>
            <option value="orange">Team 1 Only</option>
            <option value="green">Team 2 Only</option>
            <option value="blue">Team 3 Only</option>
            <option value="red">Team 4 Only</option>
            <option value="support">Day Support Staff</option>
          </select>
        </label>
        <div class="leave-date-fields">
          <label class="shift-edit-field">Start Date
            <input id="rule-start" type="date" value="${firstDay}" required>
          </label>
          <label class="shift-edit-field">End Date
            <input id="rule-end" type="date" value="${firstDay}" required>
          </label>
        </div>
        <label class="shift-edit-field">Shift To Assign
          <select id="rule-shift-code">
            <option value="M">Morning Shift · 06:00 – 18:00 (M)</option>
            <option value="N">Night Shift · 18:00 – 06:00 (N)</option>
            <option value="OFF">Rest Day · Off Duty (OFF)</option>
            <option value="LEAVE">Official Leave (LEAVE)</option>
          </select>
        </label>
      `;
    } else if (ruleType === 'swap_teams') {
      specificFields = `
        <div class="leave-date-fields">
          <label class="shift-edit-field">Primary Team
            <select id="rule-team-a">
              <option value="orange" selected>Team 1</option>
              <option value="green">Team 2</option>
              <option value="blue">Team 3</option>
              <option value="red">Team 4</option>
            </select>
          </label>
          <label class="shift-edit-field">Trade Schedules With
            <select id="rule-team-b">
              <option value="green" selected>Team 2</option>
              <option value="orange">Team 1</option>
              <option value="blue">Team 3</option>
              <option value="red">Team 4</option>
            </select>
          </label>
        </div>
        <div class="leave-date-fields">
          <label class="shift-edit-field">Effective From
            <input id="rule-start" type="date" value="${firstDay}" required>
          </label>
          <label class="shift-edit-field">Through
            <input id="rule-end" type="date" value="${lastDay}" required>
          </label>
        </div>
        <p class="shift-edit-hint" style="color:var(--text-light);">
          Both teams will exchange their 8-day rotation cycles across the selected range. Team 1 members will run Team 2's cycle and vice versa.
        </p>
      `;
    } else if (ruleType === 'cycle_offset') {
      specificFields = `
        <div class="leave-date-fields">
          <label class="shift-edit-field">Select Team
            <select id="rule-team-target">
              <option value="orange">Team 1</option>
              <option value="green">Team 2</option>
              <option value="blue">Team 3</option>
              <option value="red">Team 4</option>
            </select>
          </label>
          <label class="shift-edit-field">Rotation Adjustment
            <select id="rule-offset">
              <option value="1">+1 Day (Advance rotation by 1 day)</option>
              <option value="2">+2 Days (Advance rotation by 2 days)</option>
              <option value="4">+4 Days (Invert Day & Night cycles)</option>
              <option value="-1">-1 Day (Delay rotation by 1 day)</option>
              <option value="-2">-2 Days (Delay rotation by 2 days)</option>
            </select>
          </label>
        </div>
        <div class="leave-date-fields">
          <label class="shift-edit-field">Effective From
            <input id="rule-start" type="date" value="${firstDay}" required>
          </label>
          <label class="shift-edit-field">Through
            <input id="rule-end" type="date" value="${lastDay}" required>
          </label>
        </div>
      `;
    }

    return `
      <form method="dialog" class="shift-edit-form" id="rule-form">
        <div class="shift-edit-head">
          <div>
            <span class="shift-edit-eyebrow">SYSTEM SCHEDULING RULES</span>
            <h2>Apply Global / Batch Rule</h2>
            <p>Define an operational policy or shift rule to affect staff collectively.</p>
          </div>
          <button type="button" class="shift-edit-close" id="rule-close" aria-label="Close">✕</button>
        </div>

        <label class="shift-edit-field">Rule Type
          <select id="rule-type-select">
            <option value="holiday" ${ruleType === 'holiday' ? 'selected' : ''}>Company Holiday / Standdown (Set Day Off)</option>
            <option value="bulk_shift" ${ruleType === 'bulk_shift' ? 'selected' : ''}>Batch Shift Assignment (Role or Group)</option>
            <option value="swap_teams" ${ruleType === 'swap_teams' ? 'selected' : ''}>Swap Team Schedules (Trade 8-Day Cycles)</option>
            <option value="cycle_offset" ${ruleType === 'cycle_offset' ? 'selected' : ''}>Adjust Team Rotation Phase (Re-Anchor)</option>
          </select>
        </label>

        <div id="rule-specific-fields">
          ${specificFields}
        </div>

        <div class="shift-conflict-preview" style="background:#f8fafc;border-color:#e2e8f0;color:#334155;">
          <b>Impact &amp; Reversibility</b>
          <p style="margin:0;font-size:11.5px;color:#64748b;">
            This rule updates schedule cells instantly across the chosen group. You can revert this entire batch anytime using the <strong>Undo</strong> button on the schedule.
          </p>
        </div>

        <div class="shift-edit-actions">
          <button type="button" class="btn-clean" id="rule-cancel">Cancel</button>
          <button type="submit" class="btn-clean primary">Apply Rule to Schedule</button>
        </div>
      </form>
    `;
  };

  const setupDialog = (type) => {
    dialog.innerHTML = renderRuleForm(type);
    const close = () => dialog.close();
    dialog.querySelector('#rule-close').onclick = close;
    dialog.querySelector('#rule-cancel').onclick = close;

    const typeSelect = dialog.querySelector('#rule-type-select');
    typeSelect.onchange = () => setupDialog(typeSelect.value);

    dialog.querySelector('#rule-form').onsubmit = (event) => {
      event.preventDefault();
      const currentType = dialog.querySelector('#rule-type-select').value;
      const start = dialog.querySelector('#rule-start')?.value;
      const end = dialog.querySelector('#rule-end')?.value;

      if (start && end && end < start) {
        toast('Start date cannot be after end date.');
        return;
      }

      undo.push(JSON.stringify(state));
      redo = [];

      let count = 0;
      let membersAffected = 0;

      if (currentType === 'holiday') {
        const target = dialog.querySelector('#rule-target').value;
        const shiftCode = dialog.querySelector('#rule-shift-code').value || 'OFF';
        const members = activeStaff().filter(s => {
          if (target === 'all') return true;
          if (target === 'support') return !s.podId;
          if (target === 'core') return Boolean(s.podId);
          return s.podId === target;
        });

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          for (const m of members) {
            setCell(state, m.id, dKey, shiftCode);
            count++;
          }
        }
        membersAffected = members.length;
        toast(`Holiday rule applied: ${count} shift updates across ${membersAffected} staff.`);
      } else if (currentType === 'bulk_shift') {
        const target = dialog.querySelector('#rule-target').value;
        const shiftCode = dialog.querySelector('#rule-shift-code').value;
        const members = activeStaff().filter(s => {
          if (target === 'all') return true;
          if (target === 'leads') return s.role === 'Team Lead';
          if (target === 'nsps') return (s.role || '').includes('NSP');
          if (target === 'support') return !s.podId;
          return s.podId === target;
        });

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          for (const m of members) {
            setCell(state, m.id, dKey, shiftCode);
            count++;
          }
        }
        membersAffected = members.length;
        toast(`Bulk shift applied: ${count} shifts set to '${shiftCode}' for ${membersAffected} staff.`);
      } else if (currentType === 'swap_teams') {
        const teamAId = dialog.querySelector('#rule-team-a').value;
        const teamBId = dialog.querySelector('#rule-team-b').value;
        if (teamAId === teamBId) {
          toast('Please choose two different teams to swap.');
          return;
        }
        const podA = pods.find(p => p.id === teamAId);
        const podB = pods.find(p => p.id === teamBId);
        if (!podA || !podB) return;

        const membersA = activeStaff().filter(s => s.podId === teamAId);
        const membersB = activeStaff().filter(s => s.podId === teamBId);

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          // Team A receives Team B's cycle
          const shiftFromB = cycle(dKey, podB.anchor);
          for (const m of membersA) {
            setCell(state, m.id, dKey, shiftFromB);
            count++;
          }
          // Team B receives Team A's cycle
          const shiftFromA = cycle(dKey, podA.anchor);
          for (const m of membersB) {
            setCell(state, m.id, dKey, shiftFromA);
            count++;
          }
        }
        membersAffected = membersA.length + membersB.length;
        toast(`Teams swapped: ${podA.name} and ${podB.name} cycles traded (${count} shift updates).`);
      } else if (currentType === 'cycle_offset') {
        const teamId = dialog.querySelector('#rule-team-target').value;
        const offset = Number(dialog.querySelector('#rule-offset').value || 0);
        const pod = pods.find(p => p.id === teamId);
        if (!pod) return;

        const members = activeStaff().filter(s => s.podId === teamId);
        const anchorDate = new Date(`${pod.anchor}T00:00:00Z`);
        anchorDate.setUTCDate(anchorDate.getUTCDate() + offset);
        const shiftedAnchor = anchorDate.toISOString().slice(0, 10);

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          const newShift = cycle(dKey, shiftedAnchor);
          for (const m of members) {
            setCell(state, m.id, dKey, newShift);
            count++;
          }
        }
        membersAffected = members.length;
        toast(`Rotation adjusted: ${pod.name} shifted by ${offset > 0 ? '+' : ''}${offset} days (${count} shift updates).`);
      }

      close();
      save();
    };
  };

  setupDialog('holiday');
  dialog.showModal();
}

function getMonthAssignments(monthKey) {
  const [year, monthNumber] = monthKey.split('-').map(Number);
  const days = new Date(year, monthNumber, 0).getDate();
  const assignments = {};
  for (let day = 1; day <= days; day++) {
    const date = `${monthKey}-${String(day).padStart(2, '0')}`;
    for (const person of activeStaff()) assignments[`${date}:${person.id}`] = cellFor(state, person.id, date);
  }
  return assignments;
}

function currentAdmin() {
  const session = sessionStorage.getItem('admin-session') || '';
  const expires = Date.parse(sessionStorage.getItem('admin-session-expires') || '');
  const email = localStorage.getItem('admin-email') || '';
  const name = localStorage.getItem('admin-name') || '';
  const loggedIn = Boolean(session && expires > Date.now());
  const displayName = name || (email ? email.split('@')[0] : 'Administrator');
  const initial = (displayName.charAt(0) || 'A').toUpperCase();
  return { session, expires, email, name: displayName, initial, loggedIn };
}

function updateGhanaClock() {
  const svg = document.querySelector('.analog-clock-icon');
  if (!svg) return;
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Africa/Accra', hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(new Date()).map(part => [part.type, part.value]));
  const hour = Number(parts.hour) % 12;
  const minute = Number(parts.minute);
  const second = Number(parts.second);
  svg.querySelector('.clock-hour-hand')?.setAttribute('transform', `rotate(${hour * 30 + minute * 0.5} 12 12)`);
  svg.querySelector('.clock-minute-hand')?.setAttribute('transform', `rotate(${minute * 6 + second * 0.1} 12 12)`);
  svg.querySelector('.clock-second-hand')?.setAttribute('transform', `rotate(${second * 6} 12 12)`);
  const time = `${parts.hour}:${parts.minute}:${parts.second}`;
  svg.parentElement?.setAttribute('aria-label', `Ghana local time ${time}`);
  svg.parentElement?.setAttribute('title', `Ghana local time ${time}`);
}

function openAdminProfileDialog() {
  const admin = currentAdmin();
  const dialog = document.querySelector('#cloudflare-connection-dialog');
  if (!admin.loggedIn) {
    requireCloudflareConnection(() => {
      toast('Signed in as ' + currentAdmin().name);
      render();
    });
    return;
  }
  dialog.innerHTML = `<div class="shift-edit-form">
    <div class="shift-edit-head">
      <div>
        <span class="shift-edit-eyebrow">ADMIN PROFILE</span>
        <h2>${escapeHtml(admin.name)}</h2>
        <p>${escapeHtml(admin.email)}</p>
      </div>
      <button type="button" class="shift-edit-close" id="admin-profile-close" aria-label="Close">×</button>
    </div>
    <div style="font-size:13px;color:var(--text-muted);display:flex;flex-direction:column;gap:6px;margin:8px 0 16px;">
      <p style="margin:0;display:flex;align-items:center;gap:6px;"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:#10b981;"></span><strong style="color:var(--text-main);">Connected</strong></p>
      <p style="margin:0;">Session active until: <strong>${new Date(admin.expires).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}</strong></p>
    </div>
    <div class="shift-edit-actions">
      <button type="button" class="btn-clean" id="admin-profile-cancel">Close</button>
      <button type="button" class="btn-clean" id="admin-profile-signout" style="color:#ef4444;border-color:#fca5a5;">Sign out</button>
    </div>
  </div>`;
  const close = () => dialog.close();
  dialog.querySelector('#admin-profile-close').onclick = close;
  dialog.querySelector('#admin-profile-cancel').onclick = close;
  dialog.querySelector('#admin-profile-signout').onclick = () => {
    sessionStorage.removeItem('admin-session');
    sessionStorage.removeItem('admin-session-expires');
    localStorage.removeItem('admin-email');
    localStorage.removeItem('admin-name');
    close();
    toast('Signed out.');
    render();
  };
  dialog.showModal();
}

function requireCloudflareConnection(onConnected) {
  const session = sessionStorage.getItem('admin-session') || '';
  const expires = Date.parse(sessionStorage.getItem('admin-session-expires') || '');
  if (session && expires > Date.now()) { onConnected(); return; }
  const dialog = document.querySelector('#cloudflare-connection-dialog');
  dialog.innerHTML = `<form method="dialog" class="shift-edit-form" id="cloudflare-connection-form">
    <div class="shift-edit-head"><div><span class="shift-edit-eyebrow">ADMIN SIGN IN</span><h2>Sign in to ShiftDex</h2><p>Use the administrator account created for this roster.</p></div><button type="button" class="shift-edit-close" id="cloudflare-connection-close" aria-label="Close">×</button></div>
    <label class="shift-edit-field">Email<input id="cloudflare-admin-email" type="email" autocomplete="username" placeholder="admin@example.com" required></label>
    <label class="shift-edit-field">Password<input id="cloudflare-admin-password" type="password" autocomplete="current-password" placeholder="••••••••••••" required></label>
    <p class="shift-edit-hint" id="cloudflare-login-error" role="status"></p>
    <div class="shift-edit-actions"><button type="button" class="btn-clean" id="cloudflare-connection-cancel">Cancel</button><button type="submit" class="btn-clean primary">Sign in</button></div>
  </form>`;
  const close = () => dialog.close();
  dialog.querySelector('#cloudflare-connection-close').onclick = close;
  dialog.querySelector('#cloudflare-connection-cancel').onclick = close;
  dialog.querySelector('#cloudflare-connection-form').onsubmit = event => {
    event.preventDefault();
    const email = dialog.querySelector('#cloudflare-admin-email').value.trim();
    const password = dialog.querySelector('#cloudflare-admin-password').value;
    const submit = dialog.querySelector('button[type="submit"]');
    submit.disabled = true;
    fetch(API + '/api/admin/auth/login', {
      method: 'POST', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ email, password })
    }).then(async response => {
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'Sign in failed.');
      sessionStorage.setItem('admin-session', data.token);
      sessionStorage.setItem('admin-session-expires', data.expiresAt);
      localStorage.setItem('admin-email', data.email);
      localStorage.setItem('admin-name', data.fullName || data.email.split('@')[0]);
      close();
      render();
      onConnected();
    }).catch(error => {
      dialog.querySelector('#cloudflare-login-error').textContent = error.message || 'Could not reach Shift Desk.';
      submit.disabled = false;
    });
  };
  dialog.showModal();
}
async function publishRoster() {
  const tok = sessionStorage.getItem('admin-session') || '';
  const expires = Date.parse(sessionStorage.getItem('admin-session-expires') || '');
  if (!tok || expires <= Date.now()) {
    requireCloudflareConnection(publishRoster);
    return;
  }
  const publishButton = document.querySelector('#btn-top-publish');
  const previousMarkup = publishButton?.innerHTML;
  if (publishButton) {
    publishButton.disabled = true;
    publishButton.classList.add('is-loading');
    publishButton.innerHTML = '<span class="button-spinner" aria-hidden="true"></span><span>Publishing schedule...</span>';
  }
  try {
    const coverAssignments = Object.entries(state.coverAssignments || {}).flatMap(([key, cover]) => {
      const date = key.slice(0, 10), requesterId = key.slice(11);
      return date.startsWith(`${month}-`) ? [{ id: cover.id, date, requesterId, coverWorkerId: cover.coverWorkerId, shiftCode: cover.shiftCode }] : [];
    });
    const removedCoverIds = (state.coverTombstones || []).filter(item => item.date.startsWith(`${month}-`)).map(item => item.id);
    const res = await fetch(API + '/api/admin/publish', {
      method: 'POST',
      headers: { 'content-type': 'application/json', authorization: 'Bearer ' + tok },
      body: JSON.stringify({ month, assignments: getMonthAssignments(month), coverAssignments, removedCoverIds, inactiveStaffIds: state.inactiveStaffIds || [], podAssignments: activeStaff().map(person => personFor(person, month)).filter(person => person.podId).map(person => ({ staffId: person.id, podId: person.podId, rowOrder: person.rowOrder ?? staff.findIndex(item => item.id === person.id) })), staffRecords: (state.customStaff || []).filter(person => !(state.inactiveStaffIds || []).includes(person.id)).map(person => ({ id: person.id, fullName: person.name, roleTag: person.role, podId: person.podId || null, phone: person.phone || null, effectiveFrom: person.effectiveFrom || '0000-00-00' })) })
    });
    if (!res.ok) throw Error(`Publish failed (${res.status})`);
    state.coverTombstones = (state.coverTombstones || []).filter(item => !item.date.startsWith(`${month}-`));
    localStorage.setItem('shift-roster', JSON.stringify(state));
    toast('Published live successfully');
  } catch (error) {
    toast(error.message?.startsWith('Publish failed') ? error.message : 'Could not connect to Worker API');
  } finally {
    if (publishButton?.isConnected) {
      publishButton.innerHTML = previousMarkup;
      publishButton.disabled = false;
      publishButton.classList.remove('is-loading');
    }
  }
}

function copyHandoverBrief() {
  const todayInfo = getTodayInfo();
  const dKey = todayInfo.todayKeyStr;
  const core = activeStaff().filter(s => s.podId);
  const morning = core.filter(s => cellFor(state, s.id, dKey) === 'M');
  const night = core.filter(s => cellFor(state, s.id, dKey) === 'N');
  const morningPod = todayInfo.morningPod?.name || 'Morning Team';
  const nightPod = todayInfo.nightPod?.name || 'Night Team';
  const support = activeStaff().filter(person => !person.podId);
  const supportLine = (label, predicate) => {
    const names = support.filter(predicate).map(person => person.displayName || person.name);
    return names.length ? `• ${label}: ${names.join(', ')}` : '';
  };

  const brief = `SHIFT HANDOVER · ${dKey}
Cycle: 2M → 2N → 4 OFF

06:00 – 18:00 MORNING SHIFT (${morningPod}):
${morning.map(s => `• ${s.name} (${s.role})`).join('\n')}

18:00 – 06:00 NIGHT SHIFT (${nightPod}):
${night.map(s => `• ${s.name} (${s.role})`).join('\n')}

DAY & SUPPORT CREW:
${[
    supportLine('Fueling', person => /energy|fuel/i.test(person.role || '')),
    supportLine('Power NOC', person => /power|noc/i.test(person.role || '')),
    supportLine('Management', person => /manager|admin/i.test(person.role || '')),
    supportLine('Support', person => /support|leave/i.test(person.role || '')),
  ].filter(Boolean).join('\n')}

Status: All 6 shift roles covered.`;

  navigator.clipboard.writeText(brief).then(() => {
    toast(`Copied handover brief for ${dKey}`);
  }).catch(() => {
    toast(`Handover brief generated`);
  });
}

function exportExcel() {
  const [y, m] = month.split('-').map(Number);
  const n = new Date(y, m, 0).getDate();
  const monthAbbr = monthName(month).slice(0, 3).toUpperCase();
  const sheetTitle = `${monthAbbr} ${y} ROOSTER`;

  // Row 1: YYYY-MM-01, empty, empty, DATE, SUN, MON...
  const rowDate = [`${month}-01`, '', '', 'DATE', ...Array.from({ length: n }, (_, i) => new Date(y, m - 1, i + 1).toLocaleDateString('en', { weekday: 'short' }).toUpperCase())];

  // Row 2: empty, empty, empty, DAYS, 1, 2, 3...
  const rowDays = ['', '', '', 'DAYS', ...Array.from({ length: n }, (_, i) => i + 1)];

  // Row 3: empty spacer row
  const rowSpacer = [];

  const rows = [rowDate, rowDays, rowSpacer];

  // Rows 4 to 15: The 4 Pods
  const ordered = orderedPeople(month);
  const staffRowMeta = [];
  for (const p of pods) {
    const podStaff = ordered.filter(x => x.podId === p.id);
    for (const s of podStaff) {
      let displayName = s.displayName;
      if (!displayName) {
        if (s.role === 'Team Lead') displayName = `${s.name} (Team Lead)`;
        else if (s.role) displayName = `${s.shortName || s.name}(${s.role})`;
        else displayName = s.shortName || s.name;
      }
      const shifts = Array.from({ length: n }, (_, i) => {
        const v = cellFor(state, s.id, dateKey(i + 1));
        return v === 'OFF' ? '' : v;
      });
      rows.push([displayName, '', '', '', ...shifts]);
      staffRowMeta.push({ podId: p.id, staff: s });
    }
  }

  // Rows 16, 17, 18: empty spacer rows
  rows.push([], [], []);

  // Rows 19 to 23: Support staff
  const supportRows = activeStaff().filter(person => !person.podId).map(person => ({
    role: person.role === 'Energy' ? 'Energy( Fueling)' : person.role === 'Power NOC' ? 'Power NOC( PM)' : person.role,
    name: person.displayName || person.name,
    color: person.role === 'Energy' ? 'C55A11' : person.role === 'Power NOC' ? '4472C4' : person.role === 'Manager' ? '7030A0' : '0284C7',
  }));
  for (const sr of supportRows) {
    rows.push([sr.role, sr.name]);
  }

  const ws = XLSX.utils.aoa_to_sheet(rows);

  // Border presets
  const borderThin = {
    top: { style: 'thin', color: { rgb: 'D9D9D9' } },
    bottom: { style: 'thin', color: { rgb: 'D9D9D9' } },
    left: { style: 'thin', color: { rgb: 'D9D9D9' } },
    right: { style: 'thin', color: { rgb: 'D9D9D9' } }
  };
  const borderHeader = {
    top: { style: 'medium', color: { rgb: '94A3B8' } },
    bottom: { style: 'medium', color: { rgb: '94A3B8' } },
    left: { style: 'thin', color: { rgb: 'CBD5E1' } },
    right: { style: 'thin', color: { rgb: 'CBD5E1' } }
  };

  // Pod color styles matching the boss's spreadsheet
  const podColors = {
    orange: { rgb: 'FFC000', text: '000000' }, // Gold / Amber
    green: { rgb: '92D050', text: '000000' },  // Apple Green
    blue: { rgb: '5B9BD5', text: 'FFFFFF' },   // Corporate Blue
    red: { rgb: 'FF0000', text: 'FFFFFF' }     // Crimson Red
  };

  // 1. Style Title Cell A1
  if (ws['A1']) {
    ws['A1'].s = {
      font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '0F172A' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: borderHeader
    };
  }

  // 2. Style Header DATE & DAYS (D1 & D2)
  if (ws['D1']) {
    ws['D1'].s = {
      font: { name: 'Calibri', sz: 9, bold: true, color: { rgb: '475569' } },
      fill: { fgColor: { rgb: 'F1F5F9' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: borderThin
    };
  }
  if (ws['D2']) {
    ws['D2'].s = {
      font: { name: 'Calibri', sz: 9, bold: true, color: { rgb: '475569' } },
      fill: { fgColor: { rgb: 'F1F5F9' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: borderThin
    };
  }

  // 3. Style Day Columns (E1 to E2 + n)
  for (let d = 1; d <= n; d++) {
    const colIdx = 3 + d;
    const dt = new Date(y, m - 1, d);
    const isWeekend = (dt.getDay() === 0 || dt.getDay() === 6);
    const dayBg = isWeekend ? 'F1F5F9' : 'F8FAFC';

    // Row 1 (Day of week)
    const cellR1 = XLSX.utils.encode_cell({ r: 0, c: colIdx });
    if (ws[cellR1]) {
      ws[cellR1].s = {
        font: { name: 'Calibri', sz: 8.5, bold: true, color: { rgb: isWeekend ? '64748B' : '1E293B' } },
        fill: { fgColor: { rgb: dayBg } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: borderThin
      };
    }

    // Row 2 (Day number)
    const cellR2 = XLSX.utils.encode_cell({ r: 1, c: colIdx });
    if (ws[cellR2]) {
      ws[cellR2].s = {
        font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '0F172A' } },
        fill: { fgColor: { rgb: isWeekend ? 'F1F5F9' : 'FFFFFF' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: borderThin
      };
    }
  }

  // 4. Style Pod Staff Rows (Row 4 to 15, zero-indexed rows 3 to 14)
  staffRowMeta.forEach((meta, idx) => {
    const rowIdx = 3 + idx;
    const cellA = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
    const pColor = podColors[meta.podId] || { rgb: 'E2E8F0', text: '000000' };

    if (ws[cellA]) {
      ws[cellA].s = {
        font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: pColor.text } },
        fill: { fgColor: { rgb: pColor.rgb } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: borderThin
      };
    }

    // Style shift cells in this row
    for (let d = 1; d <= n; d++) {
      const colIdx = 3 + d;
      const cellRef = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
      const val = (ws[cellRef]?.v || '').toString().trim().toUpperCase();

      let cellStyle;
      if (val === 'M') {
        // Yellow fill (matching boss's spreadsheet)
        cellStyle = {
          fill: { fgColor: { rgb: 'FFFF00' } },
          font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '000000' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: borderThin
        };
      } else if (val === 'N') {
        // Grey fill with white text (matching boss's spreadsheet)
        cellStyle = {
          fill: { fgColor: { rgb: '7F7F7F' } },
          font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: borderThin
        };
      } else if (val === 'LEAVE' || val === 'LV') {
        // Soft Red for Leave
        cellStyle = {
          fill: { fgColor: { rgb: 'FEE2E2' } },
          font: { name: 'Calibri', sz: 9.5, bold: true, color: { rgb: 'DC2626' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          border: borderThin
        };
      } else {
        // Rest / Off
        cellStyle = {
          fill: { fgColor: { rgb: 'FFFFFF' } },
          border: borderThin
        };
      }

      if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
      ws[cellRef].s = cellStyle;
    }
  });

  // 5. Style Support Staff Rows (Rows 19 to 23, zero-indexed 18 to 22)
  supportRows.forEach((sr, idx) => {
    const rowIdx = 18 + idx;
    const cellA = XLSX.utils.encode_cell({ r: rowIdx, c: 0 });
    const cellB = XLSX.utils.encode_cell({ r: rowIdx, c: 1 });

    if (ws[cellA]) {
      ws[cellA].s = {
        font: { name: 'Calibri', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
        fill: { fgColor: { rgb: sr.color } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: borderThin
      };
    }
    if (ws[cellB]) {
      ws[cellB].s = {
        font: { name: 'Calibri', sz: 10, bold: true, color: { rgb: '1E293B' } },
        fill: { fgColor: { rgb: 'F8FAFC' } },
        alignment: { horizontal: 'left', vertical: 'center' },
        border: borderThin
      };
    }
  });

  // Column widths
  ws['!cols'] = [
    { wch: 26 }, // Column A (Member / Pod role)
    { wch: 18 }, // Column B (Support Name)
    { wch: 4 },  // Column C
    { wch: 7 },  // Column D (DATE / DAYS)
    ...Array(n).fill({ wch: 4.8 }) // Day columns E..
  ];

  // Row heights
  ws['!rows'] = [
    { hpt: 22 }, // Row 1 Date
    { hpt: 20 }, // Row 2 Days
    { hpt: 8 },  // Row 3 Spacer
    ...Array(staffRowMeta.length).fill({ hpt: 21 }), // Staff rows 4-15
    { hpt: 8 }, { hpt: 8 }, { hpt: 8 }, // Spacer rows 16-18
    ...Array(supportRows.length).fill({ hpt: 19 }) // Support rows 19-23
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetTitle);
  XLSX.writeFile(wb, `${sheetTitle}.xlsx`);
  toast(`Exported ${sheetTitle}.xlsx with full system team & shift colors!`);
}

function workbookMonth(sheetName, firstRow) {
  const match = sheetName.match(/(NOV|DEC|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT)[A-Z\s_-]*(20\d{2})/i);
  if (match) {
    const monthIndex = { JAN: 1, FEB: 2, MAR: 3, APR: 4, MAY: 5, JUN: 6, JUL: 7, AUG: 8, SEP: 9, OCT: 10, NOV: 11, DEC: 12 }[match[1].toUpperCase()];
    return `${match[2]}-${String(monthIndex).padStart(2, '0')}`;
  }
  const serial = Number(firstRow?.[0]);
  const parsed = Number.isFinite(serial) ? XLSX.SSF.parse_date_code(serial) : null;
  return parsed ? `${parsed.y}-${String(parsed.m).padStart(2, '0')}` : null;
}

function parseWorkbook(workbook) {
  const imported = initialState();
  const summary = [];
  const podByColor = { 'FFC000': 'orange', '92D050': 'green', '00B0F0': 'blue', '5B9BD5': 'blue', 'FF0000': 'red' };

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const key = workbookMonth(sheetName, rows[0]);
    if (!key) continue;
    const [y, m] = key.split('-').map(Number);
    const days = new Date(y, m, 0).getDate();
    let matches = 0, cells = 0, missing = [];

    for (let r = 3; r < 15; r++) {
      const source = String(rows[r]?.[0] || '').trim();
      if (!source) continue;
      const clean = source.replace(/\s*\([^)]*\)\s*$/, '').trim();
      const identity = clean.toLowerCase();
      const person = staff.find(s => [s.name, s.shortName].filter(Boolean).some(a => identity === a.toLowerCase()));
      if (!person) continue;

      const color = String(sheet[`A${r + 1}`]?.s?.fgColor?.rgb || '').toUpperCase().slice(-6);
      const podId = podByColor[color] || ['orange', 'green', 'blue', 'red'][Math.floor((r - 3) / 3)];
      const tag = source.match(/\(([^)]+)\)\s*$/);
      const role = tag ? (tag[1].trim().toUpperCase() === 'NSP-SECURITY' ? 'NSP-SECURITY' : tag[1].trim().toUpperCase() === 'NSP' ? 'NSP' : tag[1].trim()) : (roleFor(person, key) || 'NSP');
      const displayName = clean;

      imported.staffMonthly[`${key}:${person.id}`] = { podId, role, displayName, podColor: color ? `#${color}` : undefined, rowOrder: r + 1 };
      matches++;

      for (let d = 1; d <= days; d++) {
        const raw = String(rows[r][3 + d] || '').trim().toUpperCase();
        const value = ['M', 'N', 'LEAVE', 'OFF'].includes(raw) ? raw : 'OFF';
        const date = `${key}-${String(d).padStart(2, '0')}`;
        imported.assignments[`${date}:${person.id}`] = value;
        cells++;
      }
    }

    for (const s of staff.filter(x => x.podId)) {
      if (!imported.staffMonthly[`${key}:${s.id}`]) missing.push(s.name);
    }
    if (matches) summary.push({ month: key, matches, cells, missing });
  }

  if (!summary.length) throw new Error('No monthly roster sheets with recognizable staff rows were found.');
  return { state: imported, summary };
}

let adminLoginError = '';
let adminLoginEmail = '';

function renderAdminLogin() {
  app.innerHTML = `
    <div class="admin-login-screen">
      <div class="admin-login-card">
        <div class="admin-login-brand">
          <img class="brand-logo-mark" src="${shiftDeskMark}" alt="Shift Desk logo">
          <div class="brand-text">shift desk<span>.</span></div>
        </div>
        <div class="admin-login-header">
          <span class="shift-edit-eyebrow">ADMINISTRATOR PORTAL</span>
          <h2>Sign in to Shift Desk</h2>
          <p>Enter your administrator credentials to access the roster dashboard, manage schedules, and review requests.</p>
        </div>

        ${adminLoginError ? `<div class="admin-login-error" role="alert">${escapeHtml(adminLoginError)}</div>` : ''}

        <form id="admin-login-screen-form" class="admin-login-fields">
          <div class="admin-field-group">
            <label for="admin-screen-email">Admin Email</label>
            <input id="admin-screen-email" type="email" autocomplete="username" placeholder="admin@example.com" required value="${escapeHtml(adminLoginEmail || '')}" />
          </div>

          <div class="admin-field-group">
            <label for="admin-screen-password">Password</label>
            <input id="admin-screen-password" type="password" autocomplete="current-password" placeholder="••••••••••••" required />
          </div>

          <button type="submit" class="btn-clean primary admin-login-btn" id="btn-admin-screen-submit">
            <span>Sign in to Dashboard</span>
          </button>
        </form>

        <div class="admin-login-footer">
          <span>Shift Desk · 24/7 Operations Hub</span>
        </div>
      </div>
      <div class="minimal-toast" id="toast"></div>
    </div>
  `;

  const form = document.querySelector('#admin-login-screen-form');
  if (form) {
    form.onsubmit = async event => {
      event.preventDefault();
      const email = document.querySelector('#admin-screen-email').value.trim();
      const password = document.querySelector('#admin-screen-password').value;
      const submitBtn = document.querySelector('#btn-admin-screen-submit');
      submitBtn.disabled = true;
      submitBtn.innerHTML = '<span>Signing in…</span>';
      try {
        const response = await fetch(API + '/api/admin/auth/login', {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ email, password })
        });
        const data = await response.json();
        if (!response.ok) throw new Error(data.error || 'Sign in failed.');
        sessionStorage.setItem('admin-session', data.token);
        sessionStorage.setItem('admin-session-expires', data.expiresAt);
        localStorage.setItem('admin-email', data.email);
        localStorage.setItem('admin-name', data.fullName || data.email.split('@')[0]);
        adminLoginError = '';
        adminLoginEmail = '';
        render();
        syncDirectory();
        toast('Welcome back, ' + (data.fullName || data.email.split('@')[0]));
      } catch (err) {
        adminLoginError = err.message || 'Could not connect to Shift Desk.';
        adminLoginEmail = email;
        renderAdminLogin();
      }
    };
  }
}

// MAIN RENDER FUNCTION
function render() {
  const admin = currentAdmin();
  if (!admin.loggedIn) {
    renderAdminLogin();
    return;
  }
  const activeId = document.activeElement?.id;
  const selStart = document.activeElement?.selectionStart;
  const selEnd = document.activeElement?.selectionEnd;
  const canvasEl = document.querySelector('.content-canvas');
  const canvasScrollTop = canvasEl ? canvasEl.scrollTop : 0;
  const canvasScrollLeft = canvasEl ? canvasEl.scrollLeft : 0;
  const windowScrollX = window.scrollX;
  const windowScrollY = window.scrollY;

  const choices = [];
  const start = new Date(2026, 10, 1); // November 2026 onwards (Month index 10 = Nov)
  for (let i = 0; i < 36; i++) {
    const d = new Date(start.getFullYear(), start.getMonth() + i, 1);
    choices.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  if (month < '2026-11') month = '2026-11';
  if (!choices.includes(month)) choices.push(month);
  choices.sort();

  app.innerHTML = `
    <div class="app-shell">
      <!-- Minimalist Payday Sidebar -->
      <aside class="sidebar">
        <div class="brand-row">
          <img class="brand-logo-mark" src="${shiftDeskMark}" alt="">
          <span class="brand-text">shift desk<span>.</span></span>
        </div>

        <div class="sidebar-caption">Workspace</div>
        <nav class="side-nav-list">
          <button class="side-nav-button ${activePage === 'overview' ? 'active' : ''}" data-nav="overview">
            <span class="side-nav-icon">${ICONS.overview}</span>
            <span>Overview</span>
          </button>
          <button class="side-nav-button ${activePage === 'roster' ? 'active' : ''}" data-nav="roster">
            <span class="side-nav-icon">${ICONS.roster}</span>
            <span>Schedule</span>
          </button>
          <button class="side-nav-button ${activePage === 'people' ? 'active' : ''}" data-nav="people">
            <span class="side-nav-icon">${ICONS.people}</span>
            <span>People</span>
            <span class="side-nav-counter">${activeStaff().length}</span>
          </button>
          <button class="side-nav-button ${activePage === 'requests' ? 'active' : ''}" data-nav="requests">
            <span class="side-nav-icon">${ICONS.requests}</span>
            <span>Requests</span>
            <span class="side-nav-counter">${requests.length}</span>
          </button>
          <button class="side-nav-button ${activePage === 'settings' ? 'active' : ''}" data-nav="settings">
            <span class="side-nav-icon">${ICONS.settings}</span>
            <span>Settings</span>
            ${['available', 'downloading', 'downloaded'].includes(desktopUpdateState.status) ? '<span class="update-nav-badge">Update</span>' : ''}
          </button>
        </nav>

        <div class="sidebar-spacer"></div>

        <div class="sidebar-footer" id="sidebar-admin-profile" role="button" tabindex="0" title="${admin.loggedIn ? 'Signed in as ' + escapeHtml(admin.name) + ' (' + escapeHtml(admin.email) + ') — Click for options' : 'Click to sign in'}" style="cursor:pointer;">
          <div class="sidebar-user-avatar">${escapeHtml(admin.initial)}</div>
          <div class="sidebar-user-info">
            <span class="sidebar-user-name">${escapeHtml(admin.name)}</span>
            <span class="sidebar-user-role">${admin.loggedIn ? 'Administrator · Connected' : 'Sign in to sync'}</span>
          </div>
        </div>
      </aside>

      <!-- Main App Canvas -->
      <div class="main-wrapper">
        <!-- Payday Style Header -->
        <header class="topbar-header">
          <div class="topbar-greeting">
            <span>${admin.loggedIn ? `Hello, welcome back, ${escapeHtml(admin.name)}` : 'Welcome to Shift Desk'}</span>
          </div>

          <div class="topbar-controls">
            <!-- Minimal Month Pill -->
            <div class="month-picker-pill">
              <button class="month-nav-btn" id="btn-month-prev" title="Previous month" ${month <= '2026-11' ? 'disabled style="opacity:0.3;cursor:not-allowed;"' : ''}>${ICONS.chevronLeft}</button>
              <select class="month-nav-select" id="select-active-month">
                ${choices.map(m => `<option value="${m}" ${m === month ? 'selected' : ''}>${formatMonth(m)}</option>`).join('')}
              </select>
              <button class="month-nav-btn" id="btn-month-next" title="Next month">${ICONS.chevronRight}</button>
            </div>

            <button class="btn-clean" id="btn-top-export">
              <span>Export Excel</span>
            </button>
            <button class="btn-clean primary pill" id="btn-top-publish">
              <span>+ Publish schedule</span>
            </button>
          </div>
        </header>

        <!-- Canvas Body -->
        <main class="content-canvas ${activePage === 'roster' ? 'roster-canvas' : ''}">
          ${renderActiveView()}
        </main>
      </div>
    </div>

    <!-- Toast -->
    <div class="minimal-toast" id="toast"></div>
    <input type="file" id="file-import-input" accept=".xlsx" style="display:none">
    <dialog class="shift-edit-dialog" id="shift-edit-dialog"></dialog>
    <dialog class="shift-edit-dialog" id="leave-dialog"></dialog>
    <dialog class="shift-edit-dialog" id="cloudflare-connection-dialog"></dialog>
    <dialog class="shift-edit-dialog" id="add-person-dialog"></dialog>
    <dialog class="staff-drawer-dialog" id="staff-detail-dialog"></dialog>
    <dialog class="shift-edit-dialog" id="apply-rule-dialog"></dialog>
    <dialog class="shift-edit-dialog" id="resolve-request-dialog"></dialog>
  `;

  updateGhanaClock();
  bindEvents();
  bindUpdateSettings();

  const newCanvas = document.querySelector('.content-canvas');
  if (newCanvas) {
    newCanvas.scrollTop = canvasScrollTop;
    newCanvas.scrollLeft = canvasScrollLeft;
  }
  if (windowScrollX || windowScrollY) {
    window.scrollTo(windowScrollX, windowScrollY);
  }

  if (activeId) {
    const el = document.getElementById(activeId);
    if (el) {
      el.focus();
      if (typeof selStart === 'number' && typeof selEnd === 'number' && el.setSelectionRange) {
        el.setSelectionRange(selStart, selEnd);
      }
    }
  }
}

function bindUpdateSettings() {
  const desktop = window.shiftDexDesktop;
  const check = document.querySelector('#btn-update-check');
  if (check) check.onclick = async (event) => {
    if (event) event.preventDefault();
    if (desktopUpdateState.status === 'checking') return;

    manualCheckUntil = Date.now() + 2500;
    desktopUpdateState = { ...desktopUpdateState, status: 'checking', message: 'Checking GitHub repository for updates…' };
    render();

    const minDelay = new Promise(resolve => setTimeout(resolve, 2500));

    if (!desktop?.updates) {
      await minDelay;
      desktopUpdateState = {
        ...desktopUpdateState,
        status: 'current',
        message: 'ShiftDex is running in preview mode. Update checks run in the installed desktop app.',
      };
      toast('ShiftDex is up to date.');
      render();
      return;
    }

    try {
      const [res] = await Promise.all([
        desktop.updates.check(),
        minDelay,
      ]);
      desktopUpdateState = res || { ...desktopUpdateState, status: 'current', message: 'ShiftDex is up to date.' };
      render();
    } catch (error) {
      await minDelay;
      desktopUpdateState = {
        ...desktopUpdateState,
        status: 'error',
        message: error.message || 'Could not check for updates.',
      };
      render();
    }
  };
  const download = document.querySelector('#btn-update-download');
  if (download) download.onclick = async () => {
    download.disabled = true;
    try { await desktop.updates.download(); }
    catch (error) { toast(error.message || 'Update download failed.'); }
  };
  const install = document.querySelector('#btn-update-install');
  if (install) install.onclick = async () => {
    install.disabled = true;
    try { await desktop.updates.install(); }
    catch (error) { toast(error.message || 'Could not install the update.'); install.disabled = false; }
  };
  const uninstall = document.querySelector('#btn-app-uninstall');
  if (uninstall) uninstall.onclick = async () => {
    if (!desktop?.uninstall) { toast('Uninstall is available in the installed Windows app.'); return; }
    if (!window.confirm('Uninstall ShiftDex and remove its local app data?')) return;
    uninstall.disabled = true;
    try { await desktop.uninstall(); }
    catch (error) { toast(error.message || 'Could not open the uninstall wizard.'); uninstall.disabled = false; }
  };
}

function renderActiveView() {
  if (activePage === 'overview') return renderOverviewView();
  if (activePage === 'roster') return renderRosterView();
  if (activePage === 'people') return renderPeopleView();
  if (activePage === 'requests') return renderRequestsView();
  if (activePage === 'settings') return renderSettingsView();
  return '';
}

// 1. AESTHETIC PAYDAY OVERVIEW DASHBOARD
function renderOverviewView() {
  const stats = getMonthStats(month);
  const todayInfo = getTodayInfo();

  // SVG Donut Math (circumference = 2 * PI * 44 ≈ 276.46)
  const radius = 44;
  const circ = 2 * Math.PI * radius;
  const mLen = (stats.mPct / 100) * circ;
  const nLen = (stats.nPct / 100) * circ;
  const offLen = (stats.offPct / 100) * circ;

  return `
    <!-- Payday Hero Duo Grid -->
    <div class="hero-duo-grid">
      <!-- Payday Card 1 (White Card) -->
      <div class="payday-hero-white">
        <div>
          <div class="payday-label-sub">Total monthly shifts</div>
          <div class="payday-huge-number">${stats.totalShifts} Shifts</div>
          <div class="payday-date-line">
            <span>${formatMonth(month)}</span>
            <span>·</span>
            <span>${stats.hours.toLocaleString()} hours delivered</span>
          </div>
        </div>

        <div class="payday-hero-actions">
          <button class="btn-clean primary pill" data-nav="roster">
            <span>View schedule</span>
          </button>
          <button class="btn-clean" id="btn-brief-copy">
            <span>Handover brief ↗</span>
          </button>
        </div>

        <!-- Payday circular blue arrow button in bottom right -->
        <button class="payday-corner-circle" data-nav="roster" title="Open schedule">
          ${ICONS.arrowUpRight}
        </button>
      </div>

      <!-- Payday Card 2 (Solid Royal Blue Safety Net Card) -->
      <div class="payday-hero-blue">
        <div>
          <div class="blue-card-top-row">
            <span class="blue-tag-pill">CURRENT ON-DUTY SHIFT</span>
            <span class="blue-arrow-corner" data-nav="roster">${ICONS.arrowUpRight}</span>
          </div>
          <div class="blue-pod-lead">${todayInfo.isMorningTime ? 'Day Shift (06:00 – 18:00)' : 'Night Shift (18:00 – 06:00)'}</div>
          <div class="blue-huge-title">3 on duty</div>
        </div>

        <div class="blue-sub-metrics-row">
          <div class="blue-sub-metric-box">
            <small>Upcoming handover</small>
            <b>${todayInfo.isMorningTime ? '18:00 (Night Shift)' : '06:00 (Day Shift)'}</b>
          </div>
          <div class="blue-sub-metric-box">
            <small>Shift compliance</small>
            <b>All clear ✓</b>
          </div>
        </div>
      </div>
    </div>

    <!-- Payday Tri-Card Grid -->
    <div class="tri-card-grid">
      <!-- Card 1: Shift Allocation (Donut Ring) -->
      <div class="clean-card">
        <div class="clean-card-header">
          <div>
            <h3 class="clean-card-title">Shift allocation</h3>
            <p class="clean-card-subtitle">Monthly balance across 30 days:</p>
          </div>
          <button class="clean-card-arrow-link" data-nav="roster">${ICONS.arrowUpRight}</button>
        </div>

        <div class="donut-wrapper">
          <div class="donut-svg-box">
            <svg viewBox="0 0 110 110" style="transform:rotate(-90deg);width:100%;height:100%;">
              <circle cx="55" cy="55" r="${radius}" fill="none" stroke="#f1f4f8" stroke-width="12" />
              <!-- Morning -->
              <circle cx="55" cy="55" r="${radius}" fill="none" stroke="#f59e0b" stroke-width="12"
                stroke-dasharray="${mLen} ${circ}" stroke-dashoffset="0" />
              <!-- Night -->
              <circle cx="55" cy="55" r="${radius}" fill="none" stroke="#1a62ff" stroke-width="12"
                stroke-dasharray="${nLen} ${circ}" stroke-dashoffset="${-mLen}" />
              <!-- Rest -->
              <circle cx="55" cy="55" r="${radius}" fill="none" stroke="#cbd5e1" stroke-width="12"
                stroke-dasharray="${offLen} ${circ}" stroke-dashoffset="${-(mLen + nLen)}" />
            </svg>
            <div class="donut-center-info">
              <span class="donut-center-pct">${stats.totalShifts}</span>
              <span class="donut-center-text">Shifts</span>
            </div>
          </div>

          <div class="donut-legend-flat">
            <div class="donut-legend-cell">
              <span><i class="donut-dot" style="background:#f59e0b"></i> Morning</span>
              <b>${stats.mPct}%</b>
            </div>
            <div class="donut-legend-cell">
              <span><i class="donut-dot" style="background:#1a62ff"></i> Night</span>
              <b>${stats.nPct}%</b>
            </div>
            <div class="donut-legend-cell">
              <span><i class="donut-dot" style="background:#cbd5e1"></i> Rest</span>
              <b>${stats.offPct}%</b>
            </div>
            <div class="donut-legend-cell">
              <span><i class="donut-dot" style="background:#10b981"></i> Coverage</span>
              <b>100%</b>
            </div>
          </div>
        </div>
      </div>

      <!-- Card 2: Today's Shift Roster (Payday Account Breakdown style) -->
      <div class="clean-card">
        <div class="clean-card-header">
          <div>
            <h3 class="clean-card-title">Today's Shift Roster</h3>
          </div>
          <button class="clean-card-arrow-link" data-nav="roster">${ICONS.arrowUpRight}</button>
        </div>

        <div class="breakdown-list">
          <div class="breakdown-row">
            <div class="breakdown-icon-box" style="background:#fef9c3; color:#a16207; border: 1.5px solid #fde047; font-weight:800; font-size:14px;">
              M
            </div>
            <div class="breakdown-row-details">
              <span class="breakdown-row-title">Day Shift (06:00 – 18:00)</span>
              <span class="breakdown-row-sub">${todayInfo.morningStaff.map(s => s.displayName || s.name).join(', ') || '3 staff on duty'}</span>
            </div>
            <span class="breakdown-status-tag ${todayInfo.isMorningTime ? 'active' : ''}">3 on duty</span>
          </div>

          <div class="breakdown-row">
            <div class="breakdown-icon-box" style="background:#1e3a8a; color:#ffffff; font-weight:800; font-size:14px;">
              N
            </div>
            <div class="breakdown-row-details">
              <span class="breakdown-row-title">Night Shift (18:00 – 06:00)</span>
              <span class="breakdown-row-sub">${todayInfo.nightStaff.map(s => s.displayName || s.name).join(', ') || '3 staff on duty'}</span>
            </div>
            <span class="breakdown-status-tag ${!todayInfo.isMorningTime ? 'active' : ''}">3 on duty</span>
          </div>

          <div class="breakdown-row">
            <div class="breakdown-icon-box" style="background:#f1f5f9; color:#475569; border: 1.5px solid #cbd5e1; font-weight:800; font-size:14px;">
              S
            </div>
            <div class="breakdown-row-details">
              <span class="breakdown-row-title">Day Support (08:00 – 17:00)</span>
              <span class="breakdown-row-sub">${activeStaff().filter(person => !person.podId).map(person => person.displayName || person.name).join(', ') || 'Support team'}</span>
            </div>
            <span class="breakdown-status-tag resting">Support</span>
          </div>
        </div>
      </div>

      <!-- Card 3: Payday Next Upgrades / Handover Brief -->
      <div class="clean-card">
        <div class="clean-card-header">
          <div>
            <h3 class="clean-card-title">Handover operations</h3>
            <p class="clean-card-subtitle">Today's active cycle:</p>
          </div>
          <button class="clean-card-arrow-link" id="btn-brief-card">${ICONS.arrowUpRight}</button>
        </div>

        <div class="handover-preview-box">
          <div class="handover-slot-brief">
            <b>06:00–18:00 Morning:</b> ${todayInfo.morningStaff.map(s => s.name).join(', ') || 'Assigned'}<br>
            <span style="display:block;margin-top:6px;"><b>18:00–06:00 Night:</b> ${todayInfo.nightStaff.map(s => s.name).join(', ') || 'Assigned'}</span>
          </div>

          <!-- Payday soft blue promo card -->
          <div class="payday-promo-card" id="btn-promo-handover">
            <div class="promo-card-left">
              <span class="promo-card-title">Today's handover brief</span>
              <span class="promo-card-sub">Click to copy formatted text for WhatsApp</span>
            </div>
            <span class="promo-card-arrow">${ICONS.arrowUpRight}</span>
          </div>
        </div>
      </div>
    </div>
  `;
}



// 2. AESTHETIC ACME INC & HUBSTAFF ROSTER MATRIX
function renderRosterView() {
  const [yy, mm] = month.split('-').map(Number);
  const daysInMonth = new Date(yy, mm, 0).getDate();
  const allStaff = orderedPeople(month);
  const todayStr = new Date().toISOString().slice(0, 10);
  const issues = validateMonth(state, month);

  const todayObj = new Date();
  const currentMonthStr = todayObj.toISOString().slice(0, 7);
  const isCurrentMonth = (month === currentMonthStr);
  const todayDay = isCurrentMonth ? todayObj.getDate() : 1;

  // Show active day and upcoming days by default; past days are hidden unless toggled
  const hasPastDays = isCurrentMonth && todayDay > 1;
  const startDay = (hasPastDays && !rosterShowPast) ? todayDay : 1;
  const numDaysShown = daysInMonth - startDay + 1;

  const coreStaff = allStaff.filter(s => s.podId);
  const supportStaff = allStaff.filter(s => !s.podId);
  const shownPods = rosterFilter === 'all' || rosterFilter === 'support' ? pods : pods.filter(p => p.id === rosterFilter);
  const shownSupport = rosterFilter === 'all' || rosterFilter === 'support' ? supportStaff : [];

  // Header Days Cells
  const headerCells = Array.from({ length: numDaysShown }, (_, i) => {
    const d = startDay + i;
    const dt = new Date(yy, mm - 1, d);
    const dKey = dateKey(d);
    const isToday = (dKey === todayStr);
    const dayOfWeek = dt.getDay();
    const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

    return `
      <div class="hubstaff-day-col ${isToday ? 'today active-day-col' : isWeekend ? 'weekend' : ''} ${dayOfWeek === 0 ? 'week-end' : ''}" data-day="${d}">
        <small>${isToday ? 'TODAY' : dt.toLocaleDateString('en', { weekday: 'short' }).toUpperCase()}</small>
        <span>${d}</span>
        ${isToday ? '<em class="today-label">ACTIVE</em>' : ''}
      </div>
    `;
  }).join('');

  const podSections = shownPods.map(p => {
    const members = coreStaff.filter(s => s.podId === p.id);
    if (!members.length) return '';
    return `<section class="pod-schedule-section" data-pod-drop="${p.id}" style="--pod-color:${p.color}">
      <div class="pod-schedule-heading">
        <div class="pod-heading-sticky sticky-member-col">
          <span class="pod-team-indicator" style="background:${p.color}"></span>
          <span class="pod-team-title">${escapeHtml(p.name)}</span>
        </div>
        <div class="pod-heading-bar"></div>
      </div>
      ${members.map(s => renderHubstaffRow(s, startDay, daysInMonth, todayStr, yy, mm)).join('')}
    </section>`;
  }).join('');

  return `
    <div class="roster-page-container">
      <!-- Clean Top Controls Bar -->
      <div class="acme-subbar">
        <div class="acme-subbar-left">
          <span style="font-weight:700;font-size:14px;color:var(--text-main);letter-spacing:-0.2px;">
            ${isCurrentMonth ? `Active: Today (${formatMonth(month)} ${todayDay})` : `${formatMonth(month)} Schedule`}
          </span>
          <span class="acme-status-pill ${issues.length ? 'warning' : 'success'}">
            ${issues.length ? `${issues.length} Conflicts` : 'All Clear · 100% Coverage'}
          </span>
          ${hasPastDays ? `
            <button type="button" class="btn-clean small" id="btn-toggle-past-days" style="margin-left:4px;font-size:11.5px;color:var(--text-light);" title="${rosterShowPast ? 'Hide past days' : 'Show past days of this month'}">
              ${rosterShowPast ? 'Hide past days' : `Show past (${formatMonth(month).slice(0, 3)} 1–${todayDay - 1})`}
            </button>
          ` : ''}
        </div>

        <div class="acme-subbar-right">
          <div class="clean-search-box">
            ${ICONS.search}
            <input type="text" id="input-search-roster" placeholder="Search member..." value="${escapeHtml(rosterSearch)}" />
          </div>
          <button class="btn-clean small" id="btn-roster-undo" ${undo.length ? '' : 'disabled'}>Undo</button>
          <button class="btn-clean small" id="btn-roster-redo" ${redo.length ? '' : 'disabled'}>Redo</button>
          <button class="btn-clean small" id="btn-roster-import" title="Upload monthly Excel roster">Import</button>
          <button class="btn-clean small" id="btn-roster-leave">Schedule leave</button>
        </div>
      </div>

      <!-- Schedule Matrix Card -->
      <div class="hubstaff-schedule-card">
        <div class="matrix-viewport">
          <div class="hubstaff-table" style="--days: ${numDaysShown}; --table-min-width:${180 + numDaysShown * 48}px">
            <!-- Header Row -->
            <div class="hubstaff-header-row">
              <div class="sticky-member-col">MEMBERS</div>
              ${headerCells}
            </div>

            <!-- Continuous 12 Shift Crew Rows -->
            ${podSections}

            <!-- Day & Support Crew Section -->
            ${shownSupport.length ? `
              <div class="support-section-divider">
                <div class="pod-heading-sticky sticky-member-col">
                  <span class="pod-team-indicator" style="background:#64748b"></span>
                  <span class="pod-team-title">Support Crew</span>
                </div>
                <div class="pod-heading-bar"></div>
              </div>
              ${shownSupport.map(s => renderHubstaffRow(s, startDay, daysInMonth, todayStr, yy, mm)).join('')}
            ` : ''}

          </div>
        </div>
      </div>
    </div>
  `;
}



function renderHubstaffRow(s, startDay, daysInMonth, todayStr, yy, mm) {
  const pod = pods.find(p => p.id === s.podId);
  const podColor = s.podColor || pod?.color || '#cbd5e1';
  const numDaysShown = daysInMonth - startDay + 1;
  const cells = Array.from({ length: numDaysShown }, (_, i) => {
    const d = startDay + i;
    const dt = new Date(yy, mm - 1, d);
    const dKey = dateKey(d);
    const val = cellFor(state, s.id, dKey);
    const isToday = (dKey === todayStr);
    const isSelected = (selected?.staff === s.id && selected?.day === d);
    const dayOfWeek = dt.getDay();
    const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
    const previous = cellFor(state, s.id, shiftDateOffset(dKey, -1));
    const following = cellFor(state, s.id, shiftDateOffset(dKey, 1));
    const runClass = ['M', 'N'].includes(val)
      ? previous === val && following === val ? 'run-middle'
        : previous === val ? 'run-end'
          : following === val ? 'run-start' : 'run-single'
      : '';
    const coverFor = state.coverAssignments?.[`${dKey}:${s.id}`];
    const coveringFor = Object.entries(state.coverAssignments || {}).find(([key, cover]) => key.startsWith(`${dKey}:`) && cover.coverWorkerId === s.id);
    const coverTarget = coveringFor ? staff.find(person => person.id === coveringFor[0].slice(dKey.length + 1)) : null;
    const displayValue = coverFor ? 'CV' : val === 'M' ? 'M' : val === 'N' ? 'N' : val === 'LEAVE' ? 'LV' : '';

    const shiftTitle = coverFor ? `Shift covered by ${staff.find(person => person.id === coverFor.coverWorkerId)?.name || 'stand-in'}`
      : coveringFor ? `Covering ${coverTarget?.name || 'a colleague'} · ${coveringFor[1].shiftCode === 'M' ? 'Morning' : 'Night'}`
      : val === 'M' ? 'Day Shift (06:00 - 18:00)'
      : val === 'N' ? 'Night Shift (18:00 - 06:00)'
      : val === 'LEAVE' ? 'On Leave'
      : 'Rest Day (OFF)';

    return `
      <button class="hubstaff-shift-cell ${s.podId ? '' : 'support-row'} ${isToday ? 'today-col active-today-cell' : isWeekend ? 'weekend' : ''} ${dayOfWeek === 0 ? 'week-end' : ''} ${runClass} ${coverFor ? 'covered-shift' : ''} ${coveringFor ? 'covering-shift' : ''} ${isSelected ? 'selected' : ''}"
        data-staff="${s.id}" data-day="${d}" tabindex="0"
        title="${escapeHtml(s.displayName || s.name)} · Day ${d} · ${shiftTitle}">
        <span class="shift-pill-badge ${val === 'M' ? 'morning' : val === 'N' ? 'night' : val === 'LEAVE' ? 'leave' : coverFor ? 'covered-off' : 'off'}">
          ${displayValue}
        </span>
        ${coveringFor ? '<i class="cover-cell-mark" aria-label="Cover assignment">C</i>' : ''}
      </button>
    `;
  }).join('');

  const nameStr = s.displayName || s.name;
  const initials = nameStr.split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();

  return `
    <div class="hubstaff-member-row ${s.podId ? 'pod-draggable-row' : 'support-row'}" ${s.podId ? `draggable="true" data-person-row="${s.id}"` : ''}>
      <div class="hubstaff-member-info sticky-member-col" style="--member-pod-color:${podColor}">
        <div class="member-avatar-chip" style="--member-pod-color:${podColor}">${initials}</div>
        <div class="member-name-block">
          <span class="member-name-text">${escapeHtml(nameStr)}</span>
          <span class="member-role-text">${escapeHtml(s.role)}${s.group ? ' · ' + s.group : ''}</span>
        </div>
      </div>
      ${cells}
    </div>
  `;
}

// 3. PEOPLE DIRECTORY VIEW (Minimalist Table + Zendenta-inspired Detail Drawer)
function renderPeopleView() {
  const allStaff = orderedPeople(month);
  const allPending = pendingStaff();
  const allArchived = archivedStaff();
  let filtered = allStaff;
  let filteredPending = allPending;
  let filteredArchived = allArchived;

  if (peopleFilter === 'core') {
    filtered = filtered.filter(s => s.podId);
    filteredPending = filteredPending.filter(s => s.podId);
    filteredArchived = filteredArchived.filter(s => s.podId);
  } else if (peopleFilter === 'support') {
    filtered = filtered.filter(s => !s.podId);
    filteredPending = filteredPending.filter(s => !s.podId);
    filteredArchived = filteredArchived.filter(s => !s.podId);
  }

  if (peopleSearch.trim()) {
    const q = peopleSearch.toLowerCase();
    filtered = filtered.filter(s => (s.displayName || s.name || '').toLowerCase().includes(q) || (s.role || '').toLowerCase().includes(q) || (s.email || '').toLowerCase().includes(q));
    filteredPending = filteredPending.filter(s => (s.displayName || s.name || '').toLowerCase().includes(q) || (s.role || '').toLowerCase().includes(q) || (s.email || '').toLowerCase().includes(q));
    filteredArchived = filteredArchived.filter(s => (s.displayName || s.name || '').toLowerCase().includes(q) || (s.role || '').toLowerCase().includes(q) || (s.email || '').toLowerCase().includes(q));
  }

  const [yy, mm] = month.split('-').map(Number);
  const daysInMonth = new Date(yy, mm, 0).getDate();
  const todayStr = new Date().toISOString().slice(0, 10);

  const showPending = (peopleFilter === 'all' || peopleFilter === 'pending') && filteredPending.length > 0;
  const showActive = peopleFilter !== 'pending' && peopleFilter !== 'archived';
  const showArchived = (peopleFilter === 'archived');

  return `
    <div class="people-section-wrapper">
      <div style="display:flex;align-items:center;justify-content:space-between;">
        <div>
          <h2 style="font-size:22px;font-weight:700;letter-spacing:-0.5px;margin:0 0 4px;">People</h2>
          <p style="font-size:13px;color:var(--text-light);margin:0;">Team members, role allocations, and shift loads. Tap any person to view full details.</p>
        </div>
        <div style="display:flex;align-items:center;gap:10px;">
          <div class="clean-search-box" style="width:240px;">
            ${ICONS.search}
            <input type="text" id="input-people-search" placeholder="Search people..." value="${escapeHtml(peopleSearch)}" />
          </div>
          <button class="btn-clean primary" id="btn-add-person">+ Add person</button>
        </div>
      </div>

      <div class="filter-pills-row">
        <button class="clean-filter-pill ${peopleFilter === 'all' ? 'active' : ''}" data-people-filter="all">All (${allStaff.length + allPending.length})</button>
        <button class="clean-filter-pill ${peopleFilter === 'core' ? 'active' : ''}" data-people-filter="core">Shift Crew (${allStaff.filter(s => s.podId).length})</button>
        <button class="clean-filter-pill ${peopleFilter === 'support' ? 'active' : ''}" data-people-filter="support">Day Support (${allStaff.filter(s => !s.podId).length})</button>
        <button class="clean-filter-pill ${peopleFilter === 'pending' ? 'active' : ''}" data-people-filter="pending" style="${allPending.length ? 'color:#b45309;font-weight:700;' : ''}">Pending Invites (${allPending.length})</button>
        <button class="clean-filter-pill ${peopleFilter === 'archived' ? 'active' : ''}" data-people-filter="archived" style="${allArchived.length ? 'color:#64748b;font-weight:700;' : ''}">Archive (${allArchived.length})</button>
      </div>

      ${showPending ? `
        <!-- Minimalist Pending Invitations (Zendenta style) -->
        <div class="pending-minimal-card">
          <div class="pending-minimal-head">
            <div class="pending-minimal-title">
              <span class="pending-pulse-dot"></span>
              Pending Invitations (${filteredPending.length})
            </div>
            <span style="font-size:12px;color:#92400e;">Awaiting first sign-in &amp; permanent password creation</span>
          </div>

          <div>
            ${filteredPending.map(p => {
              const initials = (p.displayName || p.name).split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
              const podName = pods.find(pod => pod.id === p.podId)?.name || (p.podId ? p.podId : 'Day Support');
              return `
                <div class="pending-minimal-row" data-view-staff="${p.id}" role="button" tabindex="0">
                  <div class="table-member-cell">
                    <div class="table-avatar" style="border-color:#fde68a;background:#fef3c7;color:#b45309;">${initials}</div>
                    <div class="table-member-info">
                      <span class="table-member-name">${escapeHtml(p.displayName || p.name)}</span>
                      <span class="table-member-sub">${escapeHtml(p.email || 'No email')}</span>
                    </div>
                  </div>

                  <div style="display:flex;align-items:center;gap:8px;">
                    <span class="table-role-tag">${escapeHtml(p.role)}</span>
                    <span class="table-pod-badge" style="--pod-col:${pods.find(pod => pod.id === p.podId)?.color || '#64748b'};"><span class="table-pod-dot"></span>${escapeHtml(podName)}</span>
                  </div>

                  <div style="display:flex;align-items:center;gap:8px;">
                    <span class="table-status-pill morning">Pending Acceptance</span>
                    ${p.tempPassword ? `<span class="pending-temp-pw" style="font-size:11px;">Temp: <code>${escapeHtml(p.tempPassword)}</code></span>` : ''}
                  </div>

                  <div class="pending-minimal-actions" onclick="event.stopPropagation();">
                    <button type="button" class="btn-pending-resend" data-resend-invite="${p.id}" title="Resend credentials email">Resend</button>
                    <button type="button" class="btn-pending-cancel" data-cancel-invite="${p.id}" title="Cancel invitation">Cancel</button>
                    <button type="button" class="table-view-btn" data-view-staff="${p.id}">View info &rarr;</button>
                  </div>
                </div>
              `;
            }).join('')}
          </div>
        </div>
      ` : ''}

      ${showArchived ? `
        <!-- Minimalist Archived Staff Table -->
        <div class="people-table-card">
          <div style="padding:16px 20px;border-bottom:1px solid var(--border-color);display:flex;align-items:center;justify-content:space-between;background:#f8fafc;">
            <div>
              <div style="font-size:14px;font-weight:700;color:var(--text-main);">Archived Staff Directory (${filteredArchived.length})</div>
              <p style="font-size:12px;color:var(--text-muted);margin:2px 0 0;">These team members are inactive and removed from active work and schedules. You can restore them to active duty or delete them permanently.</p>
            </div>
          </div>
          <table class="people-table">
            <thead>
              <tr>
                <th>Member</th>
                <th>Role</th>
                <th>Previous Team</th>
                <th>Status</th>
                <th style="text-align:right;">Actions</th>
              </tr>
            </thead>
            <tbody>
              ${filteredArchived.length ? filteredArchived.map(s => {
                const initials = (s.displayName || s.name).split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
                const pObj = pods.find(p => p.id === s.podId);
                const podName = pObj?.name || (s.podId ? s.podId : s.group || 'Day Support');
                return `
                  <tr class="people-table-row" data-view-staff="${s.id}">
                    <td>
                      <div class="table-member-cell">
                        <div class="table-avatar" style="border-color:#cbd5e1;background:#f1f5f9;color:#64748b;">${initials}</div>
                        <div class="table-member-info">
                          <span class="table-member-name">${escapeHtml(s.displayName || s.name)}</span>
                          <span class="table-member-sub">${escapeHtml(s.email || s.id)}</span>
                        </div>
                      </div>
                    </td>
                    <td>
                      <span class="table-role-tag">${escapeHtml(s.role)}</span>
                    </td>
                    <td>
                      <span style="font-size:12px;color:var(--text-muted);">${escapeHtml(podName)}</span>
                    </td>
                    <td>
                      <span class="table-status-pill off" style="background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;">
                        Archived / Inactive
                      </span>
                    </td>
                    <td style="text-align:right;" onclick="event.stopPropagation();">
                      <div style="display:inline-flex;align-items:center;gap:6px;">
                        <button type="button" class="btn-clean primary" style="font-size:12px;padding:5px 12px;" data-restore-staff="${s.id}">Bring back</button>
                        <button type="button" class="btn-clean" style="font-size:12px;padding:5px 10px;color:#dc2626;border-color:#fecaca;" data-delete-permanent="${s.id}">Delete</button>
                        <button type="button" class="table-view-btn" data-view-staff="${s.id}">Info &rarr;</button>
                      </div>
                    </td>
                  </tr>
                `;
              }).join('') : `
                <tr>
                  <td colspan="5" style="padding:48px 20px;text-align:center;color:var(--text-muted);">
                    <div style="font-size:14px;font-weight:600;color:var(--text-main);margin-bottom:4px;">No archived staff</div>
                    <p style="font-size:12px;margin:0;">There are currently no staff members in the archive.</p>
                  </td>
                </tr>
              `}
            </tbody>
          </table>
        </div>
      ` : ''}

      ${showActive ? `
        <!-- Minimalist Active Staff Table (Zendenta style) -->
        <div class="people-table-card">
          <table class="people-table">
            <thead>
              <tr>
                <th>Member</th>
                <th>Role</th>
                <th>Team</th>
                <th>Current Status</th>
                <th>Monthly Load</th>
                <th style="text-align:right;">Details</th>
              </tr>
            </thead>
            <tbody>
              ${filtered.length ? filtered.map(s => {
                let mCount = 0, nCount = 0;
                for (let d = 1; d <= daysInMonth; d++) {
                  const v = cellFor(state, s.id, dateKey(d));
                  if (v === 'M') mCount++;
                  else if (v === 'N') nCount++;
                }
                const todayVal = cellFor(state, s.id, todayStr);
                const initials = (s.displayName || s.name).split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
                const pObj = pods.find(p => p.id === s.podId);
                const podName = pObj?.name || (s.podId ? s.podId : s.group || 'Day Support');
                const podCol = pObj?.color || '#64748b';
                const shiftStatusClass = todayVal === 'M' ? 'morning' : todayVal === 'N' ? 'night' : 'off';
                const shiftStatusText = todayVal === 'M' ? 'Morning (06:00 – 18:00)' : todayVal === 'N' ? 'Night (18:00 – 06:00)' : 'Off Duty';

                return `
                  <tr class="people-table-row" data-view-staff="${s.id}">
                    <td>
                      <div class="table-member-cell">
                        <div class="table-avatar" style="--pod-col:${podCol};border-color:${podCol}50;">${initials}</div>
                        <div class="table-member-info">
                          <span class="table-member-name">${escapeHtml(s.displayName || s.name)}</span>
                          <span class="table-member-sub">${escapeHtml(s.email || s.id)}</span>
                        </div>
                      </div>
                    </td>
                    <td>
                      <span class="table-role-tag">${escapeHtml(s.role)}</span>
                    </td>
                    <td>
                      <span class="table-pod-badge" style="--pod-col:${podCol};">
                        <span class="table-pod-dot"></span>
                        ${escapeHtml(podName)}
                      </span>
                    </td>
                    <td>
                      <span class="table-status-pill ${shiftStatusClass}">
                        ${shiftStatusText}
                      </span>
                    </td>
                    <td>
                      <span class="table-load-text">${mCount + nCount} shifts</span>
                      <span class="table-load-sub">${(mCount + nCount) * 12}h · ${mCount}M / ${nCount}N</span>
                    </td>
                    <td style="text-align:right;">
                      <button type="button" class="table-view-btn" data-view-staff="${s.id}">View info &rarr;</button>
                    </td>
                  </tr>
                `;
              }).join('') : `
                <tr>
                  <td colspan="6" style="padding:48px 20px;text-align:center;color:var(--text-muted);">
                    <div style="font-size:14px;font-weight:600;color:var(--text-main);margin-bottom:4px;">No members found</div>
                    <p style="font-size:12px;margin:0;">No team members match the search query or filter.</p>
                  </td>
                </tr>
              `}
            </tbody>
          </table>
        </div>
      ` : ''}
    </div>
  `;
}

// SLIDE-OVER DETAIL DRAWER (Zendenta inspired)
function openStaffDetailDrawer(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  const dialog = document.querySelector('#staff-detail-dialog');
  if (!dialog) return;

  const [yy, mm] = month.split('-').map(Number);
  const daysInMonth = new Date(yy, mm, 0).getDate();
  const todayStr = new Date().toISOString().slice(0, 10);
  const isArchived = (state.inactiveStaffIds?.includes(person.id) || person.status === 'archived' || person.status === 'inactive');
  const isPending = !isArchived && (person.status === 'pending');

  let mCount = 0, nCount = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const v = cellFor(state, person.id, dateKey(d));
    if (v === 'M') mCount++;
    else if (v === 'N') nCount++;
  }
  const todayVal = cellFor(state, person.id, todayStr);
  const initials = (person.displayName || person.name).split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
  const pObj = pods.find(p => p.id === person.podId);
  const podName = pObj?.name || (person.podId ? person.podId : person.group || 'Day Support');
  const podCol = pObj?.color || '#64748b';
  const shiftStatusClass = isArchived ? 'off' : (todayVal === 'M' ? 'morning' : todayVal === 'N' ? 'night' : 'off');
  const shiftStatusText = isArchived ? 'Archived from duty' : (todayVal === 'M' ? 'Morning Shift (06:00 – 18:00)' : todayVal === 'N' ? 'Night Shift (18:00 – 06:00)' : 'Off Duty');
  const statusBadge = isArchived
    ? '<span class="table-status-pill off" style="background:#f1f5f9;color:#64748b;border:1px solid #e2e8f0;">Archived</span>'
    : (isPending ? '<span class="table-status-pill morning">Pending Acceptance</span>' : '<span class="table-status-pill off">Active Crew</span>');

  dialog.innerHTML = `
    <div class="drawer-inner">
      <div class="drawer-header">
        <div>
          <div class="drawer-eyebrow">STAFF PROFILE</div>
          <h2 class="drawer-title">${escapeHtml(person.displayName || person.name)}</h2>
          <p class="drawer-sub">ID: #${escapeHtml(person.id)} · ${escapeHtml(podName)}</p>
        </div>
        <button type="button" class="drawer-close-btn" id="drawer-close-btn" aria-label="Close">✕</button>
      </div>

      <div class="drawer-body">
        <!-- Profile Card (Zendenta style) -->
        <div class="drawer-profile-card">
          <div class="drawer-avatar" style="--pod-col:${podCol};border-color:${podCol};">${initials}</div>
          <div class="drawer-profile-meta">
            <span class="drawer-profile-name">${escapeHtml(person.displayName || person.name)}</span>
            <div class="drawer-tags-row">
              <span class="table-role-tag">${escapeHtml(person.role)}</span>
              <span class="table-pod-badge" style="--pod-col:${podCol};"><span class="table-pod-dot"></span>${escapeHtml(podName)}</span>
              ${statusBadge}
            </div>
          </div>
        </div>

        <!-- Contact & Assignment Info (Zendenta Bill-To style) -->
        <div class="drawer-breakdown-box">
          <div class="drawer-box-heading">Contact &amp; Assignment</div>
          <div class="drawer-breakdown-row">
            <span style="color:var(--text-light);">Account Email</span>
            <div style="display:flex;align-items:center;gap:6px;">
              <strong style="font-family:var(--font-mono);font-size:12.5px;color:var(--text-main);">${escapeHtml(person.email || 'None')}</strong>
              ${person.email ? `<button type="button" id="btn-drawer-copy-email" style="border:none;background:none;color:var(--blue-primary);cursor:pointer;padding:2px;" title="Copy email">${ICONS.copy}</button>` : ''}
              ${!isArchived ? `<button type="button" class="btn-clean" id="btn-drawer-change-email" style="font-size:11px;padding:2px 8px;margin-left:4px;">Change Email</button>` : ''}
            </div>
          </div>
          ${person.pendingEmail ? `
            <div style="background:#fffbeb;border:1px solid #fde68a;border-radius:10px;padding:12px 14px;margin:8px 0 12px;">
              <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:4px;">
                <span style="font-size:11px;font-weight:700;color:#b45309;text-transform:uppercase;letter-spacing:0.5px;">Pending Verification</span>
                <span style="font-size:11px;color:#92400e;">48h expiry</span>
              </div>
              <div style="font-family:var(--font-mono);font-size:13px;font-weight:600;color:#92400e;margin-bottom:6px;word-break:break-all;">${escapeHtml(person.pendingEmail)}</div>
              <div style="font-size:11.5px;color:#78350f;line-height:1.4;margin-bottom:8px;">
                Verification link sent. Active sign-in email remains <strong>${escapeHtml(person.email)}</strong> until confirmed.
              </div>
              <div style="display:flex;gap:8px;">
                <button type="button" class="btn-clean" id="btn-drawer-resend-email-verify" style="font-size:11.5px;padding:4px 10px;">Resend Link</button>
                <button type="button" class="btn-clean" id="btn-drawer-cancel-email-verify" style="font-size:11.5px;padding:4px 10px;color:#dc2626;">Cancel</button>
              </div>
            </div>
          ` : ''}
          <div class="drawer-breakdown-row">
            <span style="color:var(--text-light);">Phone Number</span>
            <span style="color:var(--text-main);font-weight:500;">${escapeHtml(person.phone || 'Not provided')}</span>
          </div>
          <div class="drawer-breakdown-row">
            <span style="color:var(--text-light);">Effective Date</span>
            <span style="color:var(--text-main);font-weight:500;">${escapeHtml(person.effectiveFrom || '2026-11-01')}</span>
          </div>
          <div class="drawer-breakdown-row">
            <span style="color:var(--text-light);">Duty Status (Today)</span>
            <span class="table-status-pill ${shiftStatusClass}">${shiftStatusText}</span>
          </div>
          ${isPending && person.tempPassword ? `
            <div class="drawer-breakdown-row">
              <span style="color:var(--text-light);">Temporary Password</span>
              <span style="font-family:var(--font-mono);font-weight:700;color:#b45309;background:#fef3c7;padding:2px 6px;border-radius:4px;">${escapeHtml(person.tempPassword)}</span>
            </div>
          ` : ''}
        </div>

        ${isArchived ? `
          <div class="drawer-breakdown-box">
            <div class="drawer-box-heading">Monthly Shift Load · ${formatMonth(month)}</div>
            <div style="padding:8px 0;font-size:12.5px;color:var(--text-muted);line-height:1.5;">
              This team member is archived and removed from active schedule generation. Bring them back to active staff to assign a pod or standard shift schedule.
            </div>
          </div>
        ` : `
          <!-- Monthly Shift Load Breakdown (Zendenta Itemized Box style) -->
          <div class="drawer-breakdown-box">
            <div class="drawer-box-heading">Monthly Shift Load · ${formatMonth(month)}</div>
            <div class="drawer-breakdown-row">
              <span style="display:flex;align-items:center;gap:8px;">
                <span style="width:8px;height:8px;border-radius:2px;background:#d97706;"></span>
                Morning Shifts (06:00 – 18:00)
              </span>
              <strong style="color:var(--text-main);">${mCount} shifts (${mCount * 12}h)</strong>
            </div>
            <div class="drawer-breakdown-row">
              <span style="display:flex;align-items:center;gap:8px;">
                <span style="width:8px;height:8px;border-radius:2px;background:#2563eb;"></span>
                Night Shifts (18:00 – 06:00)
              </span>
              <strong style="color:var(--text-main);">${nCount} shifts (${nCount * 12}h)</strong>
            </div>
            <div class="drawer-breakdown-row">
              <span style="display:flex;align-items:center;gap:8px;">
                <span style="width:8px;height:8px;border-radius:2px;background:#94a3b8;"></span>
                Off / Rest Days
              </span>
              <strong style="color:var(--text-main);">${daysInMonth - mCount - nCount} days</strong>
            </div>
            <div class="drawer-breakdown-row total-row">
              <span>Total Duty Load</span>
              <span style="color:var(--blue-primary);">${mCount + nCount} shifts · ${(mCount + nCount) * 12} hrs</span>
            </div>
          </div>
        `}
      </div>

      <!-- Drawer Actions Footer -->
      <div class="drawer-footer">
        ${isArchived ? `
          <button type="button" class="btn-clean primary" id="drawer-restore-btn" style="width:100%;">Bring Back to Active Staff</button>
          <button type="button" class="btn-clean" id="drawer-delete-permanent-btn" style="width:100%;color:#dc2626;border-color:#fecaca;">Delete Permanently</button>
        ` : (isPending ? `
          <button type="button" class="btn-clean primary" id="drawer-resend-btn" style="width:100%;">Resend Invitation Email</button>
          <button type="button" class="btn-clean" id="drawer-cancel-btn" style="width:100%;color:#dc2626;">Cancel Invitation</button>
        ` : `
          ${person.email ? `<button type="button" class="btn-clean primary" id="drawer-reset-pw-btn" style="width:100%;">Reset Password</button>` : `<button type="button" class="btn-clean primary" id="drawer-invite-btn" style="width:100%;">Send Credentials Email</button>`}
          <button type="button" class="btn-clean" id="drawer-archive-btn" style="width:100%;color:#64748b;">Archive Staff</button>
        `)}
      </div>
    </div>
  `;

  dialog.querySelector('#drawer-close-btn').onclick = () => dialog.close();
  dialog.onclick = e => { if (e.target === dialog) dialog.close(); };

  const copyEmailBtn = dialog.querySelector('#btn-drawer-copy-email');
  if (copyEmailBtn && person.email) {
    copyEmailBtn.onclick = () => {
      navigator.clipboard.writeText(person.email);
      toast(`Copied ${person.email} to clipboard!`);
    };
  }

  const changeEmailBtn = dialog.querySelector('#btn-drawer-change-email');
  if (changeEmailBtn) {
    changeEmailBtn.onclick = () => {
      dialog.close();
      openChangeEmailDialog(person.id);
    };
  }

  const resendVerifyBtn = dialog.querySelector('#btn-drawer-resend-email-verify');
  if (resendVerifyBtn) {
    resendVerifyBtn.onclick = () => resendEmailChange(person.id);
  }

  const cancelVerifyBtn = dialog.querySelector('#btn-drawer-cancel-email-verify');
  if (cancelVerifyBtn) {
    cancelVerifyBtn.onclick = () => cancelEmailChange(person.id);
  }

  const resetPwBtn = dialog.querySelector('#drawer-reset-pw-btn');
  if (resetPwBtn) {
    resetPwBtn.onclick = () => {
      dialog.close();
      openResetPasswordDialog(person.id);
    };
  }

  const archiveBtn = dialog.querySelector('#drawer-archive-btn');
  if (archiveBtn) {
    archiveBtn.onclick = () => {
      dialog.close();
      archiveStaff(person.id);
    };
  }

  const restoreBtn = dialog.querySelector('#drawer-restore-btn');
  if (restoreBtn) {
    restoreBtn.onclick = () => {
      dialog.close();
      openRestoreStaffDialog(person.id);
    };
  }

  const deletePermanentBtn = dialog.querySelector('#drawer-delete-permanent-btn');
  if (deletePermanentBtn) {
    deletePermanentBtn.onclick = () => {
      dialog.close();
      openPermanentDeleteDialog(person.id);
    };
  }

  const resendBtn = dialog.querySelector('#drawer-resend-btn');
  if (resendBtn) {
    resendBtn.onclick = async () => {
      resendBtn.disabled = true;
      resendBtn.textContent = 'Sending…';
      try {
        const token = sessionStorage.getItem('admin-session') || '';
        const res = await fetch(API + '/api/admin/workers/resend', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ staffId: person.id })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to resend email');
        if (result.tempPassword) {
          person.tempPassword = result.tempPassword;
          save();
        }
        if (result.emailSent) {
          toast(`Invitation email resent to ${result.email}!`);
        } else {
          toast(`Email delivery notice: ${result.emailError || 'Brevo notice'}`);
        }
        dialog.close();
        render();
      } catch (err) {
        toast('Error: ' + err.message);
        resendBtn.disabled = false;
        resendBtn.textContent = 'Resend Invitation Email';
      }
    };
  }

  const inviteBtn = dialog.querySelector('#drawer-invite-btn');
  if (inviteBtn) {
    inviteBtn.onclick = () => {
      dialog.close();
      openInvitePersonDialog(person.id);
    };
  }

  const cancelBtn = dialog.querySelector('#drawer-cancel-btn');
  if (cancelBtn) {
    cancelBtn.onclick = async () => {
      if (!confirm(`Cancel invitation for ${person.displayName || person.name}?`)) return;
      try {
        const token = sessionStorage.getItem('admin-session') || '';
        const res = await fetch(API + '/api/admin/workers/cancel', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ staffId: person.id })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to cancel invite');
        dialog.close();
        const staffIdx = staff.findIndex(s => s.id === person.id);
        if (staffIdx !== -1) staff.splice(staffIdx, 1);
        const customIdx = state.customStaff.findIndex(s => s.id === person.id);
        if (customIdx !== -1) state.customStaff.splice(customIdx, 1);
        save();
        render();
        toast('Invitation cancelled.');
      } catch (err) {
        toast('Error: ' + err.message);
      }
    };
  }

  dialog.showModal();
}

async function archiveStaff(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  if (!confirm(`Archive ${person.displayName || person.name}?\n\nThey will be removed from active work and schedules and placed in the Archive list. You can restore them anytime or delete them permanently.`)) return;

  undo.push(JSON.stringify(state));
  redo = [];
  state.inactiveStaffIds ||= [];
  if (!state.inactiveStaffIds.includes(person.id)) state.inactiveStaffIds.push(person.id);
  person.status = 'archived';
  person.podId = undefined;
  const customItem = state.customStaff.find(s => s.id === person.id);
  if (customItem) {
    customItem.status = 'archived';
    customItem.podId = undefined;
  }

  for (const [key, cover] of Object.entries(state.coverAssignments || {})) {
    const date = key.slice(0, 10), requesterId = key.slice(11);
    if (requesterId !== person.id && cover.coverWorkerId !== person.id) continue;
    if (requesterId === person.id) setCell(state, cover.coverWorkerId, date, 'OFF');
    else setCell(state, requesterId, date, cover.shiftCode);
    if (cover.id) state.coverTombstones.push({ id: cover.id, date });
    delete state.coverAssignments[key];
  }

  save();
  render();
  toast(`${person.displayName || person.name} moved to archive.`);

  try {
    const token = sessionStorage.getItem('admin-session') || '';
    if (token) {
      await fetch(API + '/api/admin/workers/archive', {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
        body: JSON.stringify({ staffId: person.id })
      });
    }
  } catch (err) {
    console.warn('Backend archive sync error:', err);
  }
}

function openRestoreStaffDialog(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  const dialog = document.querySelector('#add-person-dialog');
  if (!dialog) return;

  dialog.innerHTML = `
    <form method="dialog" class="shift-edit-form" id="form-restore-staff">
      <div class="shift-edit-head">
        <div>
          <span class="shift-edit-eyebrow" style="color:#2563eb;">RESTORE TO ACTIVE</span>
          <h2>Restore ${escapeHtml(person.displayName || person.name)}</h2>
          <p>Bring this team member back to active duty and assign their shift schedule.</p>
        </div>
        <button type="button" class="shift-edit-close" id="restore-close" aria-label="Close">×</button>
      </div>

      <div class="shift-edit-field">
        <label style="font-weight:600;margin-bottom:8px;display:block;">Schedule &amp; Shift Assignment</label>
        
        <div style="display:flex;flex-direction:column;gap:10px;">
          <label style="display:flex;align-items:flex-start;gap:10px;padding:12px;border:1px solid #e2e8f0;border-radius:10px;cursor:pointer;background:#f8fafc;" class="assignment-choice-card">
            <input type="radio" name="restore-assignment-type" value="support" checked style="margin-top:3px;" />
            <div>
              <div style="font-weight:600;font-size:13.5px;color:var(--text-main);">Day Support (Normal 8:00 AM – 5:00 PM)</div>
              <div style="font-size:12px;color:var(--text-muted);margin-top:2px;">Standard weekday shift hours, unattached to 24/7 rotational pods.</div>
            </div>
          </label>

          <label style="display:flex;align-items:flex-start;gap:10px;padding:12px;border:1px solid #e2e8f0;border-radius:10px;cursor:pointer;background:#f8fafc;" class="assignment-choice-card">
            <input type="radio" name="restore-assignment-type" value="crew" style="margin-top:3px;" />
            <div style="flex:1;">
              <div style="font-weight:600;font-size:13.5px;color:var(--text-main);">Shift Crew (24/7 Rotational Pod)</div>
              <div style="font-size:12px;color:var(--text-muted);margin-top:2px;">Active member of a 12-hour 2-Day / 2-Night rotational shift team.</div>
              
              <div id="restore-pod-select-wrap" style="display:none;margin-top:10px;">
                <label style="font-size:12px;font-weight:600;color:var(--text-main);margin-bottom:4px;display:block;">Select Team Pod</label>
                <select id="restore-pod-id" style="width:100%;padding:8px;border-radius:8px;border:1px solid #cbd5e1;background:#fff;font-size:13px;">
                  ${pods.map(p => `<option value="${p.id}">${escapeHtml(p.name)}</option>`).join('')}
                </select>
              </div>
            </div>
          </label>
        </div>
      </div>

      <label class="shift-edit-field">
        Role Tag
        <input id="restore-role" type="text" maxlength="60" value="${escapeHtml(person.role || 'Support')}" required />
      </label>

      <div class="shift-edit-actions">
        <button type="button" class="btn-clean" id="restore-cancel">Cancel</button>
        <button type="submit" class="btn-clean primary" id="restore-submit-btn">Restore to Active</button>
      </div>
    </form>
  `;

  dialog.querySelector('#restore-close').onclick = () => dialog.close();
  dialog.querySelector('#restore-cancel').onclick = () => dialog.close();

  const form = dialog.querySelector('#form-restore-staff');
  const typeRadios = form.querySelectorAll('input[name="restore-assignment-type"]');
  const podSelectWrap = form.querySelector('#restore-pod-select-wrap');

  typeRadios.forEach(radio => {
    radio.onchange = () => {
      podSelectWrap.style.display = radio.value === 'crew' && radio.checked ? 'block' : 'none';
    };
  });

  form.onsubmit = async (e) => {
    e.preventDefault();
    const btn = form.querySelector('#restore-submit-btn');
    btn.disabled = true;
    btn.textContent = 'Restoring…';

    const isCrew = form.querySelector('input[name="restore-assignment-type"]:checked').value === 'crew';
    const chosenPod = isCrew ? form.querySelector('#restore-pod-id').value : null;
    const roleTag = form.querySelector('#restore-role').value.trim() || 'Support';

    try {
      const token = sessionStorage.getItem('admin-session') || '';
      if (token) {
        const res = await fetch(API + '/api/admin/workers/restore', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ staffId: person.id, podId: chosenPod, roleTag })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to restore staff.');
      }

      state.inactiveStaffIds = (state.inactiveStaffIds || []).filter(id => id !== person.id);
      person.status = 'active';
      person.podId = chosenPod || undefined;
      person.role = roleTag;
      const customItem = state.customStaff.find(s => s.id === person.id);
      if (customItem) {
        customItem.status = 'active';
        customItem.podId = chosenPod || undefined;
        customItem.role = roleTag;
      }
      save();
      dialog.close();
      render();
      toast(`${person.displayName || person.name} restored to active staff!`);
    } catch (err) {
      toast('Error: ' + err.message);
      btn.disabled = false;
      btn.textContent = 'Restore to Active';
    }
  };

  dialog.showModal();
}

function openPermanentDeleteDialog(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  const dialog = document.querySelector('#add-person-dialog');
  if (!dialog) return;

  dialog.innerHTML = `
    <form method="dialog" class="shift-edit-form" id="form-delete-permanent">
      <div class="shift-edit-head">
        <div>
          <span class="shift-edit-eyebrow" style="color:#dc2626;">PERMANENT DELETION</span>
          <h2>Delete ${escapeHtml(person.displayName || person.name)}?</h2>
          <p>This action is completely permanent and cannot be undone.</p>
        </div>
        <button type="button" class="shift-edit-close" id="delete-close" aria-label="Close">×</button>
      </div>

      <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:14px;color:#991b1b;font-size:12.5px;line-height:1.5;">
        <strong>Warning:</strong> You are about to permanently delete <strong>${escapeHtml(person.displayName || person.name)}</strong>.
        <ul style="margin:8px 0 0 16px;padding:0;">
          <li>All staff records and profiles will be erased.</li>
          <li>Linked worker login account and credentials will be deleted.</li>
          <li>All shift assignments, swap records, and unavailabilities will be deleted.</li>
        </ul>
      </div>

      <div class="shift-edit-actions">
        <button type="button" class="btn-clean" id="delete-cancel">Cancel</button>
        <button type="submit" class="btn-clean" id="delete-submit-btn" style="background:#dc2626;color:#fff;border-color:#dc2626;">Delete Permanently</button>
      </div>
    </form>
  `;

  dialog.querySelector('#delete-close').onclick = () => dialog.close();
  dialog.querySelector('#delete-cancel').onclick = () => dialog.close();

  const form = dialog.querySelector('#form-delete-permanent');
  form.onsubmit = async (e) => {
    e.preventDefault();
    const btn = form.querySelector('#delete-submit-btn');
    btn.disabled = true;
    btn.textContent = 'Deleting…';

    try {
      const token = sessionStorage.getItem('admin-session') || '';
      if (token) {
        const res = await fetch(API + '/api/admin/workers/delete', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ staffId: person.id })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to delete staff member.');
      }

      state.inactiveStaffIds = (state.inactiveStaffIds || []).filter(id => id !== person.id);
      const staffIdx = staff.findIndex(s => s.id === person.id);
      if (staffIdx !== -1) staff.splice(staffIdx, 1);
      const customIdx = state.customStaff.findIndex(s => s.id === person.id);
      if (customIdx !== -1) state.customStaff.splice(customIdx, 1);

      for (const [key, cover] of Object.entries(state.coverAssignments || {})) {
        const requesterId = key.slice(11);
        if (requesterId === person.id || cover.coverWorkerId === person.id) {
          delete state.coverAssignments[key];
        }
      }
      for (const key of Object.keys(state.assignments || {})) {
        if (key.endsWith(':' + person.id)) delete state.assignments[key];
      }

      save();
      dialog.close();
      render();
      toast(`${person.displayName || person.name} was permanently deleted.`);
    } catch (err) {
      toast('Error: ' + err.message);
      btn.disabled = false;
      btn.textContent = 'Delete Permanently';
    }
  };

  dialog.showModal();
}

function openResetPasswordDialog(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  const dialog = document.querySelector('#add-person-dialog');
  if (!dialog) return;

  dialog.innerHTML = `
    <form method="dialog" class="shift-edit-form" id="form-reset-password">
      <div class="shift-edit-head">
        <div>
          <span class="shift-edit-eyebrow" style="color:#2563eb;">SECURITY &amp; ACCESS</span>
          <h2>Reset Password</h2>
          <p>Generate a new temporary password for ${escapeHtml(person.displayName || person.name)}.</p>
        </div>
        <button type="button" class="shift-edit-close" id="reset-pw-close" aria-label="Close">×</button>
      </div>

      <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:14px;font-size:13px;color:var(--text-main);line-height:1.5;">
        <div><strong>Worker:</strong> ${escapeHtml(person.displayName || person.name)}</div>
        <div style="margin-top:4px;"><strong>Account Email:</strong> <code style="font-family:var(--font-mono);">${escapeHtml(person.email || 'None')}</code></div>
      </div>

      <div style="font-size:12px;color:var(--text-light);background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:10px 12px;line-height:1.4;">
        Resetting will invalidate all current active sessions. A new secure temporary password will be generated and emailed to the user, and they will be required to set a new permanent password on next sign-in.
      </div>

      <div class="shift-edit-actions">
        <button type="button" class="btn-clean" id="reset-pw-cancel">Cancel</button>
        <button type="submit" class="btn-clean primary" id="reset-pw-submit-btn">Reset Password Now</button>
      </div>
    </form>
  `;

  dialog.querySelector('#reset-pw-close').onclick = () => dialog.close();
  dialog.querySelector('#reset-pw-cancel').onclick = () => dialog.close();

  const form = dialog.querySelector('#form-reset-password');
  form.onsubmit = async (e) => {
    e.preventDefault();
    const btn = form.querySelector('#reset-pw-submit-btn');
    btn.disabled = true;
    btn.textContent = 'Resetting…';

    try {
      const token = sessionStorage.getItem('admin-session') || '';
      const res = await fetch(API + '/api/admin/workers/reset-password', {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
        body: JSON.stringify({ staffId: person.id })
      });
      const result = await res.json();
      if (!res.ok) throw new Error(result.error || 'Failed to reset password.');

      dialog.innerHTML = `
        <div class="shift-edit-form">
          <div class="shift-edit-head">
            <div>
              <span class="shift-edit-eyebrow" style="color:#059669;">PASSWORD RESET</span>
              <h2>New Password Generated</h2>
              <p>Credentials updated for <strong>${escapeHtml(person.displayName || person.name)}</strong>.</p>
            </div>
            <button type="button" class="shift-edit-close" id="reset-done-close" aria-label="Close">×</button>
          </div>

          <div style="background:#f8fafc;border:1.5px solid #cbd5e1;border-radius:12px;padding:16px;margin-bottom:14px;">
            <div style="font-size:11px;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;margin-bottom:6px;">Temporary Password</div>
            <div style="display:flex;align-items:center;justify-content:space-between;background:#fff;padding:10px 14px;border:1px solid #cbd5e1;border-radius:8px;">
              <code style="font-family:var(--font-mono);font-size:16px;font-weight:700;letter-spacing:1px;color:#0f172a;">${escapeHtml(result.tempPassword)}</code>
              <button type="button" class="btn-clean" id="btn-copy-temp-pw" style="font-size:12px;padding:4px 10px;">Copy</button>
            </div>
          </div>

          <div style="font-size:12px;color:${result.emailSent ? '#166534' : '#92400e'};background:${result.emailSent ? '#f0fdf4' : '#fef3c7'};border:1px solid ${result.emailSent ? '#bbf7d0' : '#fde68a'};border-radius:8px;padding:10px 12px;line-height:1.4;">
            ${result.emailSent ? `Email with these credentials has been sent to <strong>${escapeHtml(result.email)}</strong>.` : `Notice: Email delivery unavailable (${escapeHtml(result.emailError || 'Manual delivery required')}). Please share the temporary password with the user directly.`}
          </div>

          <div class="shift-edit-actions">
            <button type="button" class="btn-clean primary" id="reset-done-btn">Done</button>
          </div>
        </div>
      `;

      dialog.querySelector('#reset-done-close').onclick = () => dialog.close();
      dialog.querySelector('#reset-done-btn').onclick = () => dialog.close();
      dialog.querySelector('#btn-copy-temp-pw').onclick = () => {
        navigator.clipboard.writeText(result.tempPassword);
        toast('Temporary password copied to clipboard!');
      };
      toast('Password reset successfully.');
    } catch (err) {
      toast('Error: ' + err.message);
      btn.disabled = false;
      btn.textContent = 'Reset Password Now';
    }
  };

  dialog.showModal();
}

function openChangeEmailDialog(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  const dialog = document.querySelector('#add-person-dialog');
  if (!dialog) return;

  dialog.innerHTML = `
    <form method="dialog" class="shift-edit-form" id="form-change-email">
      <div class="shift-edit-head">
        <div>
          <span class="shift-edit-eyebrow" style="color:#2563eb;">ACCOUNT EMAIL</span>
          <h2>Change Email Address</h2>
          <p>Update login email for ${escapeHtml(person.displayName || person.name)}.</p>
        </div>
        <button type="button" class="shift-edit-close" id="change-email-close" aria-label="Close">×</button>
      </div>

      <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:12px 14px;font-size:12.5px;color:var(--text-main);line-height:1.5;">
        <span style="color:var(--text-muted);display:block;font-size:11px;text-transform:uppercase;font-weight:700;margin-bottom:2px;">Current Active Email</span>
        <code style="font-family:var(--font-mono);font-size:13.5px;font-weight:600;">${escapeHtml(person.email || 'None')}</code>
      </div>

      <label class="shift-edit-field">
        New Email Address
        <input id="input-new-email" type="email" maxlength="254" placeholder="new.email@company.com" required />
      </label>

      <div style="font-size:12px;color:#78350f;background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;line-height:1.4;">
        <strong>Mandatory Verification:</strong> A confirmation link will be sent to the new email address. Until verified, the current email remains active for sign-in. The change only takes effect once the worker verifies the new email.
      </div>

      <div class="shift-edit-actions">
        <button type="button" class="btn-clean" id="change-email-cancel">Cancel</button>
        <button type="submit" class="btn-clean primary" id="change-email-submit-btn">Send Verification Link</button>
      </div>
    </form>
  `;

  dialog.querySelector('#change-email-close').onclick = () => dialog.close();
  dialog.querySelector('#change-email-cancel').onclick = () => dialog.close();

  const form = dialog.querySelector('#form-change-email');
  form.onsubmit = async (e) => {
    e.preventDefault();
    const newEmail = form.querySelector('#input-new-email').value.trim().toLowerCase();
    const btn = form.querySelector('#change-email-submit-btn');
    btn.disabled = true;
    btn.textContent = 'Sending link…';

    try {
      const token = sessionStorage.getItem('admin-session') || '';
      const res = await fetch(API + '/api/admin/workers/change-email', {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
        body: JSON.stringify({ staffId: person.id, newEmail })
      });
      const result = await res.json();
      if (!res.ok) throw new Error(result.error || 'Failed to request email change.');

      person.pendingEmail = newEmail;
      const customItem = state.customStaff.find(s => s.id === person.id);
      if (customItem) customItem.pendingEmail = newEmail;
      save();
      dialog.close();
      openStaffDetailDrawer(person.id);
      toast(`Verification link sent to ${newEmail}!`);
    } catch (err) {
      toast('Error: ' + err.message);
      btn.disabled = false;
      btn.textContent = 'Send Verification Link';
    }
  };

  dialog.showModal();
}

async function cancelEmailChange(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  try {
    const token = sessionStorage.getItem('admin-session') || '';
    if (token) {
      await fetch(API + '/api/admin/workers/cancel-email-change', {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
        body: JSON.stringify({ staffId: person.id })
      });
    }
    person.pendingEmail = undefined;
    const customItem = state.customStaff.find(s => s.id === person.id);
    if (customItem) customItem.pendingEmail = undefined;
    save();
    openStaffDetailDrawer(person.id);
    toast('Email change cancelled.');
  } catch (err) {
    toast('Error: ' + err.message);
  }
}

async function resendEmailChange(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person || !person.pendingEmail) return;
  try {
    const token = sessionStorage.getItem('admin-session') || '';
    if (token) {
      const res = await fetch(API + '/api/admin/workers/resend-email-change', {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
        body: JSON.stringify({ staffId: person.id })
      });
      const result = await res.json();
      if (!res.ok) throw new Error(result.error || 'Failed to resend link');
    }
    toast(`Verification link resent to ${person.pendingEmail}!`);
  } catch (err) {
    toast('Error: ' + err.message);
  }
}

function openInvitePersonDialog(staffId) {
  const person = staff.find(p => p.id === staffId);
  if (!person) return;
  const dialog = document.querySelector('#add-person-dialog');
  if (!dialog) return;

  dialog.innerHTML = `
    <form method="dialog" class="shift-edit-form" id="form-invite-person">
      <div class="shift-edit-head">
        <div>
          <span class="shift-edit-eyebrow">SEND CREDENTIALS</span>
          <h2>Invite ${escapeHtml(person.displayName || person.name)}</h2>
          <p>Send an invitation email with a temporary password via Brevo.</p>
        </div>
        <button type="button" class="shift-edit-close" id="invite-close" aria-label="Close">×</button>
      </div>

      <label class="shift-edit-field">
        Recipient Email
        <input id="invite-email" type="email" maxlength="254" placeholder="enter email" value="${escapeHtml(person.email || '')}" required />
      </label>

      <div style="font-size:12px;color:var(--text-light);background:#f8fafc;padding:12px 14px;border:1px solid #e2e8f0;border-radius:8px;line-height:1.5;">
        A secure temporary password will be generated and emailed from the configured sender. Upon first sign in, they will be required to choose a new permanent password.
      </div>

      <div class="shift-edit-actions">
        <button type="button" class="btn-clean" id="invite-cancel">Cancel</button>
        <button type="submit" class="btn-clean primary" id="btn-submit-invite">Send Invitation Email</button>
      </div>
    </form>
  `;

  dialog.querySelector('#invite-close').onclick = () => dialog.close();
  dialog.querySelector('#invite-cancel').onclick = () => dialog.close();

  const form = dialog.querySelector('#form-invite-person');
  form.onsubmit = async event => {
    event.preventDefault();
    const email = dialog.querySelector('#invite-email').value.trim().toLowerCase();
    const btn = dialog.querySelector('#btn-submit-invite');
    btn.disabled = true;
    btn.textContent = 'Sending email…';

    try {
      const token = sessionStorage.getItem('admin-session') || '';
      const res = await fetch(API + '/api/admin/workers', {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
        body: JSON.stringify({
          fullName: person.displayName || person.name,
          email,
          roleTag: person.role,
          podId: person.podId || null,
          phone: person.phone || null,
          effectiveFrom: person.effectiveFrom || '2026-11-01'
        })
      });
      const result = await res.json();
      if (!res.ok) throw new Error(result.error || 'Failed to send invite.');
      person.email = email;
      save();

      dialog.innerHTML = `
        <div class="shift-edit-form">
          <div class="shift-edit-head">
            <div>
              <span class="shift-edit-eyebrow" style="color:#10b981;">EMAIL SENT</span>
              <h2>Invitation Sent!</h2>
              <p>Credentials emailed to <strong>${escapeHtml(email)}</strong> via Brevo.</p>
            </div>
            <button type="button" class="shift-edit-close" id="invite-done-close" aria-label="Close">×</button>
          </div>
          <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:14px;font-size:12.5px;color:#166534;">
            Temporary password generated: <code style="font-weight:700;background:#fff;padding:2px 6px;border-radius:4px;border:1px solid #86efac;">${escapeHtml(result.tempPassword)}</code>
          </div>
          <div class="shift-edit-actions">
            <button type="button" class="btn-clean primary" id="invite-done-btn">Done</button>
          </div>
        </div>
      `;
      dialog.querySelector('#invite-done-close').onclick = () => dialog.close();
      dialog.querySelector('#invite-done-btn').onclick = () => dialog.close();
      toast(`Invitation sent to ${email}`);
    } catch (err) {
      toast('Error: ' + err.message);
      btn.disabled = false;
      btn.textContent = 'Send Invitation Email';
    }
  };
  dialog.showModal();
}

// 4. REQUESTS VIEW
function openAddPersonDialog() {
  const dialog = document.querySelector('#add-person-dialog');
  if (!dialog) return;
  dialog.innerHTML = `<form method="dialog" class="shift-edit-form add-person-form" id="form-add-person">
    <div class="shift-edit-head">
      <div>
        <span class="shift-edit-eyebrow">TEAM DIRECTORY</span>
        <h2>Add person</h2>
        <p>Add a team member to the roster and email their login credentials.</p>
      </div>
      <button type="button" class="shift-edit-close" id="add-person-close" aria-label="Close">×</button>
    </div>

    <label class="shift-edit-field">
      Full name
      <input id="new-person-name" type="text" maxlength="100" placeholder="e.g. Ama Mensah" required />
    </label>

    <label class="shift-edit-field">
      Email
      <input id="new-person-email" type="email" maxlength="254" placeholder="enter email" required />
    </label>

    <div class="add-person-fields">
      <label class="shift-edit-field">Role
        <select id="new-person-role" required>
          <option value="" selected disabled>Select a role</option>
          <option>Team Lead</option>
          <option>NSP</option>
          <option>NSP-SECURITY</option>
          <option>Energy</option>
          <option>Power NOC</option>
          <option>Manager</option>
          <option>Support</option>
        </select>
      </label>
      <label class="shift-edit-field">Team
        <select id="new-person-pod">
          <option value="">Day / support staff</option>
          ${pods.map(p => `<option value="${p.id}">${escapeHtml(p.name)}</option>`).join('')}
        </select>
      </label>
    </div>

    <div class="add-person-fields">
      <label class="shift-edit-field">Effective from
        <input id="new-person-effective" type="date" value="${new Date().toISOString().slice(0, 10)}" required />
      </label>
      <label class="shift-edit-field">Phone <span class="field-optional">Optional</span>
        <input id="new-person-phone" type="tel" maxlength="40" placeholder="+233..." />
      </label>
    </div>

    <div class="shift-edit-actions">
      <button type="button" class="btn-clean" id="add-person-cancel">Cancel</button>
      <button type="submit" class="btn-clean primary">Add to roster</button>
    </div>
  </form>`;

  const form = dialog.querySelector('#form-add-person');
  dialog.querySelector('#add-person-close').onclick = () => dialog.close();
  dialog.querySelector('#add-person-cancel').onclick = () => dialog.close();

  form.onsubmit = event => {
    event.preventDefault();
    const name = dialog.querySelector('#new-person-name').value.trim();
    const email = dialog.querySelector('#new-person-email').value.trim().toLowerCase();
    const roleTag = dialog.querySelector('#new-person-role').value;
    const podId = dialog.querySelector('#new-person-pod').value || undefined;
    const effectiveFrom = dialog.querySelector('#new-person-effective').value;
    const phone = dialog.querySelector('#new-person-phone').value.trim() || undefined;

    if (!name || !email || !roleTag || !effectiveFrom) return;
    const podName = pods.find(p => p.id === podId)?.name || 'Day / Support';

    dialog.innerHTML = `
      <div class="shift-edit-form">
        <div class="shift-edit-head">
          <div>
            <span class="shift-edit-eyebrow">CONFIRM ADDITION</span>
            <h2>Confirm roster addition</h2>
            <p>Review the member details before sending the automated invitation email.</p>
          </div>
          <button type="button" class="shift-edit-close" id="confirm-add-close" aria-label="Close">×</button>
        </div>

        <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:16px 18px;display:flex;flex-direction:column;gap:10px;">
          <div style="display:flex;justify-content:space-between;font-size:13px;">
            <span style="color:var(--text-muted);font-weight:500;">Full Name</span>
            <strong style="color:var(--text-main);">${escapeHtml(name)}</strong>
          </div>
          <div style="display:flex;justify-content:space-between;font-size:13px;">
            <span style="color:var(--text-muted);font-weight:500;">Email Address</span>
            <strong style="color:var(--text-main);font-family:var(--font-mono);">${escapeHtml(email)}</strong>
          </div>
          <div style="display:flex;justify-content:space-between;font-size:13px;">
            <span style="color:var(--text-muted);font-weight:500;">Role &amp; Team</span>
            <span style="color:var(--text-main);font-weight:600;">${escapeHtml(roleTag)} · ${escapeHtml(podName)}</span>
          </div>
        </div>

        <p style="font-size:12px;color:var(--text-light);line-height:1.5;margin:4px 0 0;">
          A temporary password will be generated and an invitation email will be sent automatically to <strong>${escapeHtml(email)}</strong> via Brevo.
        </p>

        <div class="shift-edit-actions">
          <button type="button" class="btn-clean" id="confirm-add-back">Back</button>
          <button type="button" class="btn-clean primary" id="btn-execute-add-person">Confirm &amp; Send Email</button>
        </div>
      </div>
    `;

    dialog.querySelector('#confirm-add-close').onclick = () => dialog.close();
    dialog.querySelector('#confirm-add-back').onclick = () => openAddPersonDialog();

    dialog.querySelector('#btn-execute-add-person').onclick = async () => {
      const btn = dialog.querySelector('#btn-execute-add-person');
      btn.disabled = true;
      btn.textContent = 'Sending email…';

      try {
        const token = sessionStorage.getItem('admin-session') || '';
        const res = await fetch(API + '/api/admin/workers', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ fullName: name, email, roleTag, podId, phone, effectiveFrom })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to add person.');

        const person = {
          id: result.staffId,
          name,
          displayName: name,
          role: roleTag,
          podId,
          effectiveFrom,
          phone,
          email,
          status: 'pending',
          tempPassword: result.tempPassword
        };
        state.customStaff.push(person);
        staff.push(person);
        save();
        render();

        if (result.emailSent) {
          dialog.innerHTML = `
            <div class="shift-edit-form">
              <div class="shift-edit-head">
                <div>
                  <span class="shift-edit-eyebrow" style="color:#10b981;">INVITATION SENT</span>
                  <h2>Invitation Email Sent!</h2>
                  <p>Welcome email sent to <strong>${escapeHtml(email)}</strong> via Brevo.</p>
                </div>
                <button type="button" class="shift-edit-close" id="success-add-close" aria-label="Close">×</button>
              </div>

              <div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:16px 18px;font-size:13px;color:#166534;line-height:1.5;">
                <p style="margin:0 0 6px;"><strong>${escapeHtml(name)}</strong> is saved under <strong>Pending Invitations</strong>.</p>
                <p style="margin:0 0 6px;">Temporary Password: <code style="font-weight:700;background:#ffffff;padding:2px 6px;border-radius:4px;border:1px solid #86efac;">${escapeHtml(result.tempPassword)}</code></p>
                <p style="margin:0;font-size:11.5px;color:#15803d;">Worker will move to the active roster once they sign in and set their permanent password.</p>
              </div>

              <div class="shift-edit-actions">
                <button type="button" class="btn-clean primary" id="success-add-done">Done</button>
              </div>
            </div>
          `;
          toast(`${name} added! Invitation email sent.`);
        } else {
          dialog.innerHTML = `
            <div class="shift-edit-form">
              <div class="shift-edit-head">
                <div>
                  <span class="shift-edit-eyebrow" style="color:#f59e0b;">ATTENTION · EMAIL NOT SENT</span>
                  <h2>Worker Added to Pending</h2>
                  <p>Member was added to Pending Invitations, but Brevo could not deliver the email.</p>
                </div>
                <button type="button" class="shift-edit-close" id="success-add-close" aria-label="Close">×</button>
              </div>

              <div style="background:#fffbeb;border:1px solid #fde68a;border-radius:12px;padding:16px 18px;font-size:13px;color:#92400e;line-height:1.5;">
                <p style="margin:0 0 6px;font-weight:700;">Brevo Delivery Notice:</p>
                <p style="margin:0 0 10px;font-size:12px;color:#b45309;background:#fef3c7;padding:8px 10px;border-radius:6px;word-break:break-word;">${escapeHtml(result.emailError || 'Brevo unauthorized / Key not found')}</p>
                <p style="margin:0 0 6px;">Temporary Password: <code style="font-weight:700;background:#ffffff;padding:2px 6px;border-radius:4px;border:1px solid #fde68a;color:#1e293b;">${escapeHtml(result.tempPassword)}</code></p>
                <p style="margin:0;font-size:11.5px;color:#78350f;">You can copy this temporary password for the worker now, or update your Brevo API key and click "Resend Email" from the Pending Invitations section.</p>
              </div>

              <div class="shift-edit-actions">
                <button type="button" class="btn-clean primary" id="success-add-done">Done</button>
              </div>
            </div>
          `;
          toast(`Added ${name} to pending, but email delivery failed.`);
        }
        dialog.querySelector('#success-add-close').onclick = () => dialog.close();
        dialog.querySelector('#success-add-done').onclick = () => dialog.close();
      } catch (err) {
        toast('Error: ' + err.message);
        btn.disabled = false;
        btn.textContent = 'Confirm & Send Email';
      }
    };
  };
  dialog.showModal();
}

function renderRequestsView() {
  return `
    <div style="display:flex;flex-direction:column;gap:18px;">
      <div style="display:flex;align-items:center;justify-content:space-between;">
        <div>
          <h2 style="font-size:22px;font-weight:700;letter-spacing:-0.5px;margin:0 0 4px;">Staff Requests & Approvals</h2>
          <p style="font-size:13px;color:var(--text-light);margin:0;">Incoming worker unavailable dates and account approvals.</p>
        </div>
        <button class="btn-clean primary" id="btn-sync-d1">
          <span>Sync requests</span>
        </button>
      </div>

      <div class="requests-list-clean" style="display:flex;flex-direction:column;gap:12px;">
        ${requests.length ? requests.map(r => {
          const initials = (r.name || 'W').split(' ').map(w => w[0]).join('').slice(0, 2).toUpperCase();
          if (r.type === 'unavailability') {
            const affected = [];
            if (r.startDate && r.endDate) {
              let cur = new Date(r.startDate + 'T00:00:00Z');
              const end = new Date(r.endDate + 'T00:00:00Z');
              while (cur <= end) {
                const dStr = cur.toISOString().slice(0, 10);
                const val = cellFor(state, r.staffId, dStr);
                if (['M', 'N'].includes(val)) {
                  affected.push({ date: dStr, shift: val });
                }
                cur.setUTCDate(cur.getUTCDate() + 1);
              }
            }
            return `
              <div class="clean-request-card" style="background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.03);">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;">
                  <div style="display:flex;gap:14px;align-items:flex-start;">
                    <div class="member-avatar-chip" style="width:40px;height:40px;border-radius:10px;font-size:13px;">${initials}</div>
                    <div>
                      <div style="display:flex;align-items:center;gap:8px;">
                        <b style="font-size:14px;color:var(--text-main);">${escapeHtml(r.name)}</b>
                        <span class="req-tag-pill" style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:#fef3c7;color:#92400e;border:1px solid #fde68a;">UNAVAILABLE DATES</span>
                      </div>
                      <div style="font-size:12.5px;color:var(--text-main);margin-top:4px;font-weight:500;">
                        Requested Range: <strong>${r.startDate} &rarr; ${r.endDate}</strong>
                      </div>
                      ${r.reason ? `<div style="font-size:12px;color:var(--text-light);margin-top:2px;">Reason: <em>${escapeHtml(r.reason)}</em></div>` : ''}
                      <div style="margin-top:8px;">
                        ${affected.length ? `
                          <span style="font-size:11.5px;color:#dc2626;background:#fef2f2;border:1px solid #fecaca;padding:3px 8px;border-radius:6px;font-weight:600;display:inline-flex;align-items:center;gap:4px;">
                            ${affected.length} scheduled working shift(s) affected (${affected.map(a => `${a.date}: ${a.shift}`).join(', ')})
                          </span>
                        ` : `
                          <span style="font-size:11.5px;color:#16a34a;background:#f0fdf4;border:1px solid #bbf7d0;padding:3px 8px;border-radius:6px;font-weight:600;">
                            Worker is already scheduled OFF for these dates
                          </span>
                        `}
                      </div>
                    </div>
                  </div>
                  <div style="display:flex;gap:8px;align-items:center;margin-top:4px;">
                    <button class="btn-clean small primary" data-review-unavail-id="${r.id}">Review & Approve</button>
                    <button class="btn-clean small" style="color:#dc2626;" data-reject-unavail-id="${r.id}">Reject</button>
                  </div>
                </div>
              </div>
            `;
          }

          if (r.type === 'account') {
            return `
              <div class="clean-request-card" style="background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.03);">
                <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:16px;flex-wrap:wrap;">
                  <div style="display:flex;gap:14px;align-items:flex-start;">
                    <div class="member-avatar-chip" style="width:40px;height:40px;border-radius:10px;font-size:13px;">${initials}</div>
                    <div>
                      <div style="display:flex;align-items:center;gap:8px;">
                        <b style="font-size:14px;color:var(--text-main);">${escapeHtml(r.name)}</b>
                        <span class="req-tag-pill" style="font-size:10px;font-weight:700;padding:2px 8px;border-radius:999px;background:#eff6ff;color:#1d4ed8;border:1px solid #bfdbfe;">NEW ACCOUNT</span>
                      </div>
                      <div style="font-size:12px;color:var(--text-light);margin-top:2px;">Email: ${escapeHtml(r.email || '')}</div>
                      <div style="margin-top:8px;">
                        <select class="account-staff-link" id="account-staff-${r.id}" aria-label="Link to staff profile" style="min-width:240px;">
                          <option value="">Link to staff member...</option>
                          ${availableStaff.map(person => `<option value="${person.id}">${escapeHtml(person.full_name)} · ${escapeHtml(person.role_tag)}</option>`).join('')}
                        </select>
                      </div>
                    </div>
                  </div>
                  <div style="display:flex;gap:8px;align-items:center;margin-top:4px;">
                    <button class="btn-clean small primary" data-approve-account-id="${r.id}">Approve Account</button>
                    <button class="btn-clean small" style="color:#dc2626;" data-reject-account-id="${r.id}">Reject</button>
                  </div>
                </div>
              </div>
            `;
          }

          return `
            <div class="clean-request-card" style="background:#fff;border:1px solid #e2e8f0;border-radius:14px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.03);">
              <div style="display:flex;justify-content:space-between;align-items:center;gap:16px;">
                <div style="display:flex;gap:14px;align-items:center;">
                  <div class="member-avatar-chip" style="width:40px;height:40px;border-radius:10px;font-size:13px;">${initials}</div>
                  <div>
                    <b style="font-size:14px;color:var(--text-main);">${escapeHtml(r.name)}</b>
                    <div style="font-size:12px;color:var(--text-light);">${escapeHtml(r.text)}</div>
                  </div>
                </div>
                <div style="display:flex;gap:8px;">
                  <button class="btn-clean small primary" data-legacy-req-id="${r.id}" data-req-action="approved">Approve</button>
                  <button class="btn-clean small" style="color:#dc2626;" data-legacy-req-id="${r.id}" data-req-action="rejected">Reject</button>
                </div>
              </div>
            </div>
          `;
        }).join('') : `
          <div class="clean-card" style="padding:48px 20px;text-align:center;color:var(--text-muted);background:#fff;border:1px solid #e2e8f0;border-radius:14px;">
            <div style="font-size:14px;font-weight:600;color:var(--text-main);margin-bottom:4px;">No pending requests</div>
            <p style="font-size:12px;margin:0;">Worker registrations and unavailable date notifications will appear here.</p>
          </div>
        `}
      </div>
    </div>
  `;
}

function openResolveUnavailabilityDialog(r) {
  const dialog = document.getElementById('resolve-request-dialog');
  if (!dialog) return;

  const affected = [];
  if (r.startDate && r.endDate) {
    let cur = new Date(r.startDate + 'T00:00:00Z');
    const end = new Date(r.endDate + 'T00:00:00Z');
    while (cur <= end) {
      const dStr = cur.toISOString().slice(0, 10);
      const val = cellFor(state, r.staffId, dStr);
      if (['M', 'N'].includes(val)) {
        affected.push({ date: dStr, shift: val });
      }
      cur.setUTCDate(cur.getUTCDate() + 1);
    }
  }

  const allActive = activeStaff();

  dialog.innerHTML = `
    <form method="dialog" class="shift-edit-form" id="form-resolve-unavail" style="max-width:540px;width:100%;box-sizing:border-box;">
      <div class="shift-edit-head">
        <div>
          <span class="shift-edit-eyebrow">UNAVAILABILITY RESOLUTION</span>
          <h2 style="font-size:19px;margin:4px 0 2px;">Resolve Unavailability & Assign Stand-in</h2>
          <p style="margin:0;color:var(--text-light);font-size:12px;">
            ${escapeHtml(r.name)} &middot; ${r.startDate} to ${r.endDate}
            ${r.reason ? `<br><em>Note: ${escapeHtml(r.reason)}</em>` : ''}
          </p>
        </div>
        <button type="button" class="shift-edit-close" id="btn-close-resolve" aria-label="Close">&times;</button>
      </div>

      <div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px;font-size:12px;color:#1e40af;line-height:1.5;">
        Approving this request will set <strong>${escapeHtml(r.name)}</strong> to <strong>OFF (Rest Day)</strong> on the calendar for these dates. For any shift requiring coverage, select an available colleague below.
      </div>

      <div style="display:flex;flex-direction:column;gap:12px;max-height:280px;overflow-y:auto;padding-right:4px;">
        ${affected.length ? affected.map(item => {
          const availableColleagues = allActive.filter(s => {
            if (s.id === r.staffId) return false;
            return cellFor(state, s.id, item.date) === 'OFF';
          });
          return `
            <div style="border:1px solid #e2e8f0;border-radius:10px;padding:12px;background:#f8fafc;">
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">
                <strong style="font-size:13px;color:var(--text-main);">${item.date}</strong>
                <span class="shift-pill-badge ${item.shift === 'M' ? 'morning' : 'night'}">${item.shift}</span>
              </div>
              <label style="font-size:11px;color:var(--text-light);display:block;margin-bottom:4px;font-weight:600;">Assign Stand-in / Cover Colleague:</label>
              <select class="account-staff-link" id="standin-${item.date}" style="width:100%;">
                <option value="">(Leave shift uncovered / no stand-in)</option>
                ${availableColleagues.map(c => `<option value="${c.id}">${escapeHtml(c.name)} (${escapeHtml(c.role)}) &mdash; Available (OFF)</option>`).join('')}
              </select>
            </div>
          `;
        }).join('') : `
          <div style="padding:20px;text-align:center;color:var(--text-light);font-size:13px;">
            ${escapeHtml(r.name)} is already scheduled OFF on all requested dates. No working shifts need stand-in coverage.
          </div>
        `}
      </div>

      <div style="display:flex;justify-content:flex-end;gap:10px;margin-top:8px;padding-top:14px;border-top:1px solid #e2e8f0;">
        <button type="button" class="btn-clean" id="btn-cancel-resolve">Cancel</button>
        <button type="button" class="btn-clean primary" id="btn-confirm-resolve">Confirm & Apply to Calendar</button>
      </div>
    </form>
  `;

  dialog.showModal();

  document.getElementById('btn-close-resolve').onclick = () => dialog.close();
  document.getElementById('btn-cancel-resolve').onclick = () => dialog.close();

  document.getElementById('btn-confirm-resolve').onclick = async () => {
    if (!confirm(`Apply changes to calendar and approve unavailability for ${r.name}?`)) {
      return;
    }

    // 1. Apply changes to calendar state
    if (r.startDate && r.endDate) {
      let cur = new Date(r.startDate + 'T00:00:00Z');
      const end = new Date(r.endDate + 'T00:00:00Z');
      while (cur <= end) {
        const dStr = cur.toISOString().slice(0, 10);
        const shiftVal = cellFor(state, r.staffId, dStr);
        if (['M', 'N'].includes(shiftVal)) {
          setCell(state, r.staffId, dStr, 'OFF');
          const standInSelect = document.getElementById(`standin-${dStr}`);
          const standInId = standInSelect?.value;
          if (standInId) {
            setCell(state, standInId, dStr, shiftVal);
            state.coverAssignments = state.coverAssignments || {};
            state.coverAssignments[`${dStr}:${r.staffId}`] = { coverWorkerId: standInId, shiftCode: shiftVal };
          }
        }
        cur.setUTCDate(cur.getUTCDate() + 1);
      }
      save();
    }

    // 2. Call backend API
    try {
      await fetch(API + '/api/admin/requests/' + r.id, {
        method: 'POST',
        headers: { 'content-type': 'application/json', authorization: 'Bearer ' + (sessionStorage.getItem('admin-session') || '') },
        body: JSON.stringify({ status: 'approved' })
      });
    } catch (err) {
      console.warn('API update warning:', err);
    }

    // 3. Publish updated roster online
    if (admin.loggedIn) {
      try {
        await publishRoster({ silent: true });
      } catch (err) {
        console.warn('Publish warning:', err);
      }
    }

    dialog.close();
    requests = requests.filter(x => x.id !== r.id);
    render();
    toast(`Approved: ${r.name} marked OFF and calendar updated.`);
  };
}

let settingsRuleType = 'holiday';

function renderSettingsView() {
  const today = new Date().toISOString().slice(0, 10);
  const firstDay = month === today.slice(0, 7) ? today : `${month}-01`;
  const [y, mo] = month.split('-').map(Number);
  const lastDay = `${month}-${String(new Date(y, mo, 0).getDate()).padStart(2, '0')}`;
  const admin = currentAdmin();

  let specificFields = '';
  if (settingsRuleType === 'holiday') {
    specificFields = `
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Target Personnel
          <select id="settings-rule-target">
            <option value="all">All Personnel (Core & Support · 17 members)</option>
            <option value="support">Day Support Staff Only (5 members)</option>
            <option value="core">Shift Crew (All 4 Teams · 12 members)</option>
            <option value="orange">Team 1 Only</option>
            <option value="green">Team 2 Only</option>
            <option value="blue">Team 3 Only</option>
            <option value="red">Team 4 Only</option>
          </select>
        </label>
        <label class="shift-edit-field">Assignment Type
          <select id="settings-rule-shift-code">
            <option value="OFF">Rest / Day Off (OFF)</option>
            <option value="LEAVE">Official Leave (LEAVE)</option>
          </select>
        </label>
      </div>
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Start Date
          <input id="settings-rule-start" type="date" value="${firstDay}" required>
        </label>
        <label class="shift-edit-field">End Date
          <input id="settings-rule-end" type="date" value="${firstDay}" required>
        </label>
      </div>
      <label class="shift-edit-field">Reason / Event Name
        <input id="settings-rule-note" type="text" placeholder="e.g. Public Holiday, Maintenance" value="Public Holiday">
      </label>
    `;
  } else if (settingsRuleType === 'bulk_shift') {
    specificFields = `
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Target Group / Role
          <select id="settings-rule-target">
            <option value="all">All Personnel</option>
            <option value="leads">All Team Leads</option>
            <option value="nsps">All NSP Personnel (NSP & NSP-Security)</option>
            <option value="orange">Team 1 Only</option>
            <option value="green">Team 2 Only</option>
            <option value="blue">Team 3 Only</option>
            <option value="red">Team 4 Only</option>
            <option value="support">Day Support Staff</option>
          </select>
        </label>
        <label class="shift-edit-field">Shift To Assign
          <select id="settings-rule-shift-code">
            <option value="M">Morning Shift · 06:00 – 18:00 (M)</option>
            <option value="N">Night Shift · 18:00 – 06:00 (N)</option>
            <option value="OFF">Rest Day · Off Duty (OFF)</option>
            <option value="LEAVE">Official Leave (LEAVE)</option>
          </select>
        </label>
      </div>
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Start Date
          <input id="settings-rule-start" type="date" value="${firstDay}" required>
        </label>
        <label class="shift-edit-field">End Date
          <input id="settings-rule-end" type="date" value="${firstDay}" required>
        </label>
      </div>
    `;
  } else if (settingsRuleType === 'swap_teams') {
    specificFields = `
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Primary Team
          <select id="settings-rule-team-a">
            <option value="orange" selected>Team 1</option>
            <option value="green">Team 2</option>
            <option value="blue">Team 3</option>
            <option value="red">Team 4</option>
          </select>
        </label>
        <label class="shift-edit-field">Trade Schedules With
          <select id="settings-rule-team-b">
            <option value="green" selected>Team 2</option>
            <option value="orange">Team 1</option>
            <option value="blue">Team 3</option>
            <option value="red">Team 4</option>
          </select>
        </label>
      </div>
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Effective From
          <input id="settings-rule-start" type="date" value="${firstDay}" required>
        </label>
        <label class="shift-edit-field">Through
          <input id="settings-rule-end" type="date" value="${lastDay}" required>
        </label>
      </div>
      <p style="font-size:12px;color:var(--text-light);margin:0;">
        Both teams will exchange their 8-day rotation cycles across the selected range. Team 1 members will run Team 2's cycle and vice versa.
      </p>
    `;
  } else if (settingsRuleType === 'cycle_offset') {
    specificFields = `
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Select Team
          <select id="settings-rule-team-target">
            <option value="orange">Team 1</option>
            <option value="green">Team 2</option>
            <option value="blue">Team 3</option>
            <option value="red">Team 4</option>
          </select>
        </label>
        <label class="shift-edit-field">Rotation Adjustment
          <select id="settings-rule-offset">
            <option value="1">+1 Day (Advance rotation by 1 day)</option>
            <option value="2">+2 Days (Advance rotation by 2 days)</option>
            <option value="4">+4 Days (Invert Day & Night cycles)</option>
            <option value="-1">-1 Day (Delay rotation by 1 day)</option>
            <option value="-2">-2 Days (Delay rotation by 2 days)</option>
          </select>
        </label>
      </div>
      <div class="settings-fields-grid">
        <label class="shift-edit-field">Effective From
          <input id="settings-rule-start" type="date" value="${firstDay}" required>
        </label>
        <label class="shift-edit-field">Through
          <input id="settings-rule-end" type="date" value="${lastDay}" required>
        </label>
      </div>
    `;
  }

  return `
    <div class="settings-page-container">
      <!-- Settings Header -->
      <div class="settings-header">
        <div>
          <h2 class="settings-title">Settings &amp; Shift Rules</h2>
          <p class="settings-subtitle">Define scheduling rules that apply collectively across teams or all personnel without manual cell editing.</p>
        </div>
        <div class="settings-header-actions">
          <button type="button" class="btn-clean" data-nav="roster">
            <span>View Schedule &rarr;</span>
          </button>
        </div>
      </div>

      <div class="settings-layout-grid">
        <!-- Main Column: Global Shift Rules Engine -->
        <div class="clean-card settings-card">
          <div class="clean-card-header">
            <div>
              <h3 class="clean-card-title">Define Collective Shift Rule</h3>
              <p class="clean-card-subtitle">Set holidays, batch overrides, team trades, or cycle adjustments for multiple staff in one step.</p>
            </div>
          </div>

          <!-- Rule Type Tabs -->
          <div class="settings-rule-tabs">
            <button type="button" class="settings-tab-btn ${settingsRuleType === 'holiday' ? 'active' : ''}" data-settings-tab="holiday">
              Company Holiday
            </button>
            <button type="button" class="settings-tab-btn ${settingsRuleType === 'bulk_shift' ? 'active' : ''}" data-settings-tab="bulk_shift">
              Batch Shift
            </button>
            <button type="button" class="settings-tab-btn ${settingsRuleType === 'swap_teams' ? 'active' : ''}" data-settings-tab="swap_teams">
              Swap Teams
            </button>
            <button type="button" class="settings-tab-btn ${settingsRuleType === 'cycle_offset' ? 'active' : ''}" data-settings-tab="cycle_offset">
              Adjust Rotation
            </button>
          </div>

          <!-- Interactive Rule Form -->
          <form id="settings-rule-form" class="settings-form-body">
            ${specificFields}

            <div class="shift-conflict-preview" style="background:#f8fafc;border-color:#e2e8f0;color:#334155;margin-top:6px;">
              <b>Instant Effect &amp; Reversibility</b>
              <p style="margin:2px 0 0;font-size:12px;color:#64748b;">
                Applying this rule modifies the calendar immediately. You can revert this entire batch anytime using the <strong>Undo</strong> button.
              </p>
            </div>

            <div class="settings-form-actions">
              ${undo.length ? `
                <button type="button" class="btn-clean" id="btn-settings-undo">
                  ${ICONS.undo}
                  <span>Undo Last Rule</span>
                </button>
              ` : ''}
              <button type="submit" class="btn-clean primary">
                <span>Apply Rule to Schedule</span>
              </button>
            </div>
          </form>
        </div>

        <!-- Side Column: Active Operational Policies & Team Definitions -->
        <div style="display:flex;flex-direction:column;gap:18px;">
          <!-- Card: Active Shift Teams -->
          <div class="clean-card settings-card">
            <div class="clean-card-header">
              <div>
                <h3 class="clean-card-title">Shift Teams &amp; Anchors</h3>
                <p class="clean-card-subtitle">Active 4-team rotation roster.</p>
              </div>
            </div>

            <div style="display:flex;flex-direction:column;gap:8px;">
              ${pods.map(p => {
                const members = activeStaff().filter(s => s.podId === p.id);
                const lead = members.find(s => s.role === 'Team Lead');
                return `
                  <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 12px;background:#f8fafc;border-radius:10px;border:1px solid #edf2f7;">
                    <div style="display:flex;align-items:center;gap:10px;">
                      <span style="width:10px;height:10px;border-radius:50%;background:${p.color};"></span>
                      <div>
                        <strong style="font-size:13px;color:var(--text-main);">${escapeHtml(p.name)}</strong>
                        <div style="font-size:11px;color:var(--text-light);">${lead ? `Lead: ${escapeHtml(lead.displayName || lead.name)}` : 'No lead'} · ${members.length} members</div>
                      </div>
                    </div>
                    <span style="font-family:var(--font-mono);font-size:11px;color:var(--text-muted);">${p.anchor}</span>
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        </div>
      </div>

      <section class="clean-card settings-card update-settings-card">
        <div class="update-card-copy">
          <span class="update-card-icon">${ICONS.download}</span>
          <div>
            <h3 class="clean-card-title">ShiftDex updates</h3>
            <p class="clean-card-subtitle">Installed version ${escapeHtml(desktopUpdateState.currentVersion || '—')} · Checks automatically every 6 hours.</p>
            <p class="update-state-line" role="status">${escapeHtml(desktopUpdateState.message || 'Ready to check for updates.')}</p>
          </div>
        </div>
        <div class="update-actions">
          ${desktopUpdateState.status === 'available' ? `<button type="button" class="btn-clean primary" id="btn-update-download">Download ${escapeHtml(desktopUpdateState.version || 'update')}</button>` : ''}
          ${desktopUpdateState.status === 'downloaded' ? '<button type="button" class="btn-clean primary" id="btn-update-install">Install and restart</button>' : ''}
          ${desktopUpdateState.status === 'downloading' ? `<span class="update-progress">${Number(desktopUpdateState.percent) || 0}%</span>` : ''}
          <button type="button" class="btn-clean ${desktopUpdateState.status === 'checking' ? 'is-checking' : ''}" id="btn-update-check" ${desktopUpdateState.status === 'checking' || desktopUpdateState.status === 'downloading' ? 'disabled' : ''}>
            ${desktopUpdateState.status === 'checking'
              ? '<span class="update-spinner" aria-hidden="true"></span><span>Checking for updates…</span>'
              : '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg><span>Check for updates</span>'}
          </button>
        </div>
        <div class="update-uninstall-row">
          <div><strong>Uninstall ShiftDex</strong><span>Opens the Windows uninstall wizard and removes app data.</span></div>
          <button type="button" class="btn-clean danger" id="btn-app-uninstall">Uninstall</button>
        </div>
      </section>
    </div>
  `;
}



// INTERACTIVE EVENT BINDINGS
function bindEvents() {
  // Navigation
  document.querySelectorAll('[data-nav]').forEach(el => {
    el.onclick = () => {
      activePage = el.dataset.nav;
      render();
    };
  });

  // Month navigation (strictly from November 2026 onwards)
  const monthSelect = document.querySelector('#select-active-month');
  if (monthSelect) {
    monthSelect.onchange = e => {
      const selectedM = e.target.value;
      if (selectedM >= '2026-11') {
        month = selectedM;
        localStorage.setItem('shift-roster-month', month);
        render();
      }
    };
  }
  const prevBtn = document.querySelector('#btn-month-prev');
  if (prevBtn) {
    prevBtn.onclick = () => {
      if (month <= '2026-11') return;
      const d = new Date(month + '-01T00:00:00Z');
      d.setUTCMonth(d.getUTCMonth() - 1);
      const prevM = d.toISOString().slice(0, 7);
      if (prevM >= '2026-11') {
        month = prevM;
        localStorage.setItem('shift-roster-month', month);
        render();
      }
    };
  }
  const nextBtn = document.querySelector('#btn-month-next');
  if (nextBtn) {
    nextBtn.onclick = () => {
      const d = new Date(month + '-01T00:00:00Z');
      d.setUTCMonth(d.getUTCMonth() + 1);
      month = d.toISOString().slice(0, 7);
      localStorage.setItem('shift-roster-month', month);
      render();
    };
  }

  // Toggle past days visibility in current month
  const togglePastBtn = document.querySelector('#btn-toggle-past-days');
  if (togglePastBtn) {
    togglePastBtn.onclick = () => {
      rosterShowPast = !rosterShowPast;
      render();
    };
  }

  // Handover copy buttons
  document.querySelectorAll('#btn-brief-copy, #btn-brief-card, #btn-promo-handover').forEach(btn => {
    btn.onclick = () => copyHandoverBrief();
  });

  // Export and Import actions
  const topExportBtn = document.querySelector('#btn-top-export');
  if (topExportBtn) topExportBtn.onclick = exportExcel;

  const rosterImportBtn = document.querySelector('#btn-roster-import');
  const fileImportInput = document.querySelector('#file-import-input');
  if (rosterImportBtn && fileImportInput) {
    rosterImportBtn.onclick = () => fileImportInput.click();
    fileImportInput.onchange = async e => {
      const f = e.target.files[0];
      e.target.value = '';
      if (!f) return;
      try {
        const parsed = parseWorkbook(XLSX.read(await f.arrayBuffer(), { cellStyles: true }));
        state = parsed.state;
        month = parsed.summary.map(x => x.month).sort().at(-1);
        localStorage.setItem('shift-roster-month', month);
        save();
        toast(`Imported ${f.name} (${formatMonth(month)})`);
      } catch (err) {
        toast(err.message || 'Excel import failed. Check sheet names.');
      }
    };
  }

  // Publish Schedule
  const publishBtn = document.querySelector('#btn-top-publish');
  if (publishBtn) publishBtn.onclick = () => requireCloudflareConnection(publishRoster);

  const leaveButton = document.querySelector('#btn-roster-leave');
  if (leaveButton) leaveButton.onclick = openLeaveDialog;

  // Settings Tab Switching
  document.querySelectorAll('[data-settings-tab]').forEach(btn => {
    btn.onclick = () => {
      settingsRuleType = btn.dataset.settingsTab;
      render();
    };
  });

  // Settings Rule Form Submission
  const settingsForm = document.querySelector('#settings-rule-form');
  if (settingsForm) {
    settingsForm.onsubmit = (event) => {
      event.preventDefault();
      const currentType = settingsRuleType;
      const start = document.querySelector('#settings-rule-start')?.value;
      const end = document.querySelector('#settings-rule-end')?.value;

      if (start && end && end < start) {
        toast('Start date cannot be after end date.');
        return;
      }

      undo.push(JSON.stringify(state));
      redo = [];

      let count = 0;
      let membersAffected = 0;

      if (currentType === 'holiday') {
        const target = document.querySelector('#settings-rule-target').value;
        const shiftCode = document.querySelector('#settings-rule-shift-code').value || 'OFF';
        const members = activeStaff().filter(s => {
          if (target === 'all') return true;
          if (target === 'support') return !s.podId;
          if (target === 'core') return Boolean(s.podId);
          return s.podId === target;
        });

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          for (const m of members) {
            setCell(state, m.id, dKey, shiftCode);
            count++;
          }
        }
        membersAffected = members.length;
        toast(`Holiday rule applied: ${count} shift updates across ${membersAffected} staff.`);
      } else if (currentType === 'bulk_shift') {
        const target = document.querySelector('#settings-rule-target').value;
        const shiftCode = document.querySelector('#settings-rule-shift-code').value;
        const members = activeStaff().filter(s => {
          if (target === 'all') return true;
          if (target === 'leads') return s.role === 'Team Lead';
          if (target === 'nsps') return (s.role || '').includes('NSP');
          if (target === 'support') return !s.podId;
          return s.podId === target;
        });

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          for (const m of members) {
            setCell(state, m.id, dKey, shiftCode);
            count++;
          }
        }
        membersAffected = members.length;
        toast(`Bulk shift applied: ${count} shifts set to '${shiftCode}' for ${membersAffected} staff.`);
      } else if (currentType === 'swap_teams') {
        const teamAId = document.querySelector('#settings-rule-team-a').value;
        const teamBId = document.querySelector('#settings-rule-team-b').value;
        if (teamAId === teamBId) {
          toast('Please choose two different teams to swap.');
          return;
        }
        const podA = pods.find(p => p.id === teamAId);
        const podB = pods.find(p => p.id === teamBId);
        if (!podA || !podB) return;

        const membersA = activeStaff().filter(s => s.podId === teamAId);
        const membersB = activeStaff().filter(s => s.podId === teamBId);

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          const shiftFromB = cycle(dKey, podB.anchor);
          for (const m of membersA) {
            setCell(state, m.id, dKey, shiftFromB);
            count++;
          }
          const shiftFromA = cycle(dKey, podA.anchor);
          for (const m of membersB) {
            setCell(state, m.id, dKey, shiftFromA);
            count++;
          }
        }
        membersAffected = membersA.length + membersB.length;
        toast(`Teams swapped: ${podA.name} and ${podB.name} cycles traded (${count} shift updates).`);
      } else if (currentType === 'cycle_offset') {
        const teamId = document.querySelector('#settings-rule-team-target').value;
        const offset = Number(document.querySelector('#settings-rule-offset').value || 0);
        const pod = pods.find(p => p.id === teamId);
        if (!pod) return;

        const members = activeStaff().filter(s => s.podId === teamId);
        const anchorDate = new Date(`${pod.anchor}T00:00:00Z`);
        anchorDate.setUTCDate(anchorDate.getUTCDate() + offset);
        const shiftedAnchor = anchorDate.toISOString().slice(0, 10);

        for (let day = new Date(`${start}T00:00:00Z`); day <= new Date(`${end}T00:00:00Z`); day.setUTCDate(day.getUTCDate() + 1)) {
          const dKey = day.toISOString().slice(0, 10);
          const newShift = cycle(dKey, shiftedAnchor);
          for (const m of members) {
            setCell(state, m.id, dKey, newShift);
            count++;
          }
        }
        membersAffected = members.length;
        toast(`Rotation adjusted: ${pod.name} shifted by ${offset > 0 ? '+' : ''}${offset} days (${count} shift updates).`);
      }

      save();
    };
  }

  // Settings Undo button
  const settingsUndoBtn = document.querySelector('#btn-settings-undo');
  if (settingsUndoBtn) {
    settingsUndoBtn.onclick = () => {
      if (!undo.length) return;
      redo.push(JSON.stringify(state));
      state = JSON.parse(undo.pop());
      save();
      toast('Reverted last change.');
    };
  }

  // Drag a whole roster row onto another member to exchange pod placement, or onto a pod to move it there.
  let draggedStaffId = '';
  const rowPosition = person => person.rowOrder ?? staff.findIndex(item => item.id === person.id);
  const setMonthlyPerson = (person, podId, rowOrder) => {
    state.staffMonthly[`${month}:${person.id}`] = { ...person, podId, rowOrder };
  };
  document.querySelectorAll('[data-person-row]').forEach(row => {
    row.addEventListener('dragstart', event => {
      draggedStaffId = row.dataset.personRow;
      event.dataTransfer?.setData('text/plain', draggedStaffId);
      if (event.dataTransfer) event.dataTransfer.effectAllowed = 'move';
      requestAnimationFrame(() => row.classList.add('is-dragging'));
    });
    row.addEventListener('dragend', () => {
      draggedStaffId = '';
      document.querySelectorAll('.is-dragging, .drop-target').forEach(element => element.classList.remove('is-dragging', 'drop-target'));
    });
    row.addEventListener('dragover', event => {
      if (!draggedStaffId || draggedStaffId === row.dataset.personRow) return;
      event.preventDefault();
      row.classList.add('drop-target');
    });
    row.addEventListener('dragleave', event => {
      if (!row.contains(event.relatedTarget)) row.classList.remove('drop-target');
    });
    row.addEventListener('drop', event => {
      event.preventDefault();
      event.stopPropagation();
      const fromId = event.dataTransfer?.getData('text/plain') || draggedStaffId;
      const toId = row.dataset.personRow;
      if (!fromId || fromId === toId) return;
      const from = personFor(staff.find(item => item.id === fromId), month);
      const to = personFor(staff.find(item => item.id === toId), month);
      if (!from?.podId || !to?.podId) return;
      undo.push(JSON.stringify(state));
      redo = [];
      state.staffMonthly ||= {};
      setMonthlyPerson(from, to.podId, rowPosition(to));
      setMonthlyPerson(to, from.podId, rowPosition(from));
      save();
      const affected = new Set([from.podId, to.podId]);
      const needsRoleReview = [...affected].some(podId => {
        const roles = activeStaff().map(person => personFor(person, month)).filter(person => person.podId === podId).map(person => person.role);
        return roles.length !== 3 || ['Team Lead', 'NSP', 'NSP-SECURITY'].some(role => roles.filter(value => value === role).length !== 1);
      });
      toast(needsRoleReview ? 'People swapped locally. Check the team role mix, then publish to sync.' : 'People swapped and saved on this device. Publish to sync online.');
    });
  });
  document.querySelectorAll('[data-pod-drop]').forEach(section => {
    section.addEventListener('dragover', event => {
      if (!draggedStaffId || event.target.closest('[data-person-row]')) return;
      event.preventDefault();
      section.classList.add('drop-target');
    });
    section.addEventListener('dragleave', event => {
      if (!section.contains(event.relatedTarget)) section.classList.remove('drop-target');
    });
    section.addEventListener('drop', event => {
      if (event.target.closest('[data-person-row]')) return;
      event.preventDefault();
      event.stopPropagation();
      const fromId = event.dataTransfer?.getData('text/plain') || draggedStaffId;
      const person = personFor(staff.find(item => item.id === fromId), month);
      const targetPod = section.dataset.podDrop;
      if (!person?.podId || person.podId === targetPod) return;
      undo.push(JSON.stringify(state));
      redo = [];
      state.staffMonthly ||= {};
      const targetMembers = activeStaff().map(item => personFor(item, month)).filter(item => item.podId === targetPod && item.id !== person.id);
      const nextOrder = Math.max(-1, ...targetMembers.map(rowPosition)) + 1;
      setMonthlyPerson(person, targetPod, nextOrder);
      save();
      const roles = activeStaff().map(item => personFor(item, month)).filter(item => item.podId === targetPod).map(item => item.role);
      const needsRoleReview = roles.length !== 3 || ['Team Lead', 'NSP', 'NSP-SECURITY'].some(role => roles.filter(value => value === role).length !== 1);
      toast(needsRoleReview ? "Person moved locally. Check this team's role mix, then publish to sync." : 'Person moved and saved on this device. Publish to sync online.');
    });
  });

  // Undo / Redo
  const undoBtn = document.querySelector('#btn-roster-undo');
  if (undoBtn) {
    undoBtn.onclick = () => {
      if (undo.length) {
        redo.push(JSON.stringify(state));
        state = JSON.parse(undo.pop());
        save();
      }
    };
  }
  const redoBtn = document.querySelector('#btn-roster-redo');
  if (redoBtn) {
    redoBtn.onclick = () => {
      if (redo.length) {
        undo.push(JSON.stringify(state));
        state = JSON.parse(redo.pop());
        save();
      }
    };
  }

  const adminProfileBtn = document.querySelector('#sidebar-admin-profile');
  if (adminProfileBtn) {
    adminProfileBtn.onclick = () => openAdminProfileDialog();
  }

  const rosterSearchInput = document.querySelector('#input-search-roster');
  document.querySelectorAll('[data-roster-filter]').forEach(btn => {
    btn.onclick = () => { rosterFilter = btn.dataset.rosterFilter; render(); };
  });
  if (rosterSearchInput) {
    if (rosterSearch.trim()) {
      const q = rosterSearch.toLowerCase().trim();
      document.querySelectorAll('.hubstaff-member-row').forEach(row => {
        const text = row.querySelector('.hubstaff-member-info')?.innerText.toLowerCase() || '';
        row.style.display = (!q || text.includes(q)) ? 'grid' : 'none';
      });
      document.querySelectorAll('.pod-schedule-section').forEach(section => {
        section.style.display = [...section.querySelectorAll('.hubstaff-member-row')].some(row => row.style.display !== 'none') ? 'flex' : 'none';
      });
      const supportRows = [...document.querySelectorAll('.support-section-divider, .hubstaff-member-row.support-row')];
      const hasSupport = supportRows.some(row => row.classList.contains('support-row') && row.style.display !== 'none');
      supportRows.forEach(row => { if (row.classList.contains('support-section-divider')) row.style.display = hasSupport ? 'flex' : 'none'; });
    }
    rosterSearchInput.oninput = e => {
      rosterSearch = e.target.value;
      const q = rosterSearch.toLowerCase().trim();
      document.querySelectorAll('.hubstaff-member-row').forEach(row => {
        const text = row.querySelector('.hubstaff-member-info')?.innerText.toLowerCase() || '';
        row.style.display = (!q || text.includes(q)) ? 'grid' : 'none';
      });
      document.querySelectorAll('.pod-schedule-section').forEach(section => {
        section.style.display = [...section.querySelectorAll('.hubstaff-member-row')].some(row => row.style.display !== 'none') ? 'flex' : 'none';
      });
      const supportRows = [...document.querySelectorAll('.support-section-divider, .hubstaff-member-row.support-row')];
      const hasSupport = supportRows.some(row => row.classList.contains('support-row') && row.style.display !== 'none');
      supportRows.forEach(row => { if (row.classList.contains('support-section-divider')) row.style.display = hasSupport ? 'flex' : 'none'; });
    };
  }

  // People Filter Pills
  const addPersonButton = document.querySelector('#btn-add-person');
  if (addPersonButton) addPersonButton.onclick = openAddPersonDialog;
  document.querySelectorAll('[data-remove-person]').forEach(button => {
    button.onclick = () => {
      const id = button.dataset.removePerson;
      const person = staff.find(entry => entry.id === id);
      if (!person || !confirm(`Remove ${person.name} from the active roster? Past shifts will be kept.`)) return;
      undo.push(JSON.stringify(state));
      redo = [];
      state.inactiveStaffIds ||= [];
      if (!state.inactiveStaffIds.includes(id)) state.inactiveStaffIds.push(id);
      for (const [key, cover] of Object.entries(state.coverAssignments || {})) {
        const date = key.slice(0, 10), requesterId = key.slice(11);
        if (requesterId !== id && cover.coverWorkerId !== id) continue;
        if (requesterId === id) setCell(state, cover.coverWorkerId, date, 'OFF');
        else setCell(state, requesterId, date, cover.shiftCode);
        if (cover.id) state.coverTombstones.push({ id: cover.id, date });
        delete state.coverAssignments[key];
      }
      save();
      toast(`${person.name} removed from the active roster.`);
    };
  });
  document.querySelectorAll('[data-send-invite]').forEach(button => {
    button.onclick = () => {
      const id = button.dataset.sendInvite;
      openInvitePersonDialog(id);
    };
  });
  document.querySelectorAll('[data-resend-invite]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.resendInvite;
      btn.disabled = true;
      btn.textContent = 'Sending…';
      try {
        const token = sessionStorage.getItem('admin-session') || '';
        const res = await fetch(API + '/api/admin/workers/resend', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ staffId: id })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to resend email');
        const person = staff.find(s => s.id === id);
        if (person && result.tempPassword) {
          person.tempPassword = result.tempPassword;
          const customItem = state.customStaff.find(s => s.id === id);
          if (customItem) customItem.tempPassword = result.tempPassword;
          save();
        }
        if (result.emailSent) {
          toast(`Invitation email resent to ${result.email}!`);
        } else {
          toast(`Email delivery notice: ${result.emailError || 'Brevo unauthorized / Key not found'}`);
        }
        render();
      } catch (err) {
        toast('Error: ' + err.message);
        btn.disabled = false;
        btn.textContent = 'Resend Email';
      }
    };
  });
  document.querySelectorAll('[data-cancel-invite]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.cancelInvite;
      const person = staff.find(s => s.id === id);
      if (!person || !confirm(`Cancel invitation for ${person.displayName || person.name}?`)) return;
      try {
        const token = sessionStorage.getItem('admin-session') || '';
        const res = await fetch(API + '/api/admin/workers/cancel', {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + token },
          body: JSON.stringify({ staffId: id })
        });
        const result = await res.json();
        if (!res.ok) throw new Error(result.error || 'Failed to cancel invite');
        
        const staffIdx = staff.findIndex(s => s.id === id);
        if (staffIdx !== -1) staff.splice(staffIdx, 1);
        const customIdx = state.customStaff.findIndex(s => s.id === id);
        if (customIdx !== -1) state.customStaff.splice(customIdx, 1);
        save();
        render();
        toast('Invitation cancelled.');
      } catch (err) {
        toast('Error: ' + err.message);
      }
    };
  });
  document.querySelectorAll('[data-people-filter]').forEach(btn => {
    btn.onclick = () => {
      peopleFilter = btn.dataset.peopleFilter;
      render();
    };
  });

  document.querySelectorAll('[data-restore-staff]').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      openRestoreStaffDialog(btn.dataset.restoreStaff);
    };
  });

  document.querySelectorAll('[data-delete-permanent]').forEach(btn => {
    btn.onclick = (e) => {
      e.stopPropagation();
      openPermanentDeleteDialog(btn.dataset.deletePermanent);
    };
  });

  document.querySelectorAll('[data-view-staff]').forEach(el => {
    el.onclick = (e) => {
      // Don't trigger if user clicked an action button inside the row (e.g. resend, cancel)
      if (e.target.closest('button:not(.table-view-btn)')) return;
      const staffId = el.dataset.viewStaff;
      if (staffId) openStaffDetailDrawer(staffId);
    };
  });

  const peopleSearchInput = document.querySelector('#input-people-search');
  if (peopleSearchInput) {
    peopleSearchInput.oninput = e => {
      peopleSearch = e.target.value;
      const q = peopleSearch.toLowerCase().trim();
      document.querySelectorAll('.people-table-row').forEach(row => {
        const text = row.innerText.toLowerCase();
        row.style.display = (!q || text.includes(q)) ? '' : 'none';
      });
      document.querySelectorAll('.pending-minimal-row').forEach(row => {
        const text = row.innerText.toLowerCase();
        row.style.display = (!q || text.includes(q)) ? '' : 'none';
      });
    };
  }

  // Cell Interaction (Click & Keyboard)
  const [yy, mm] = month.split('-').map(Number);
  const daysInMonth = new Date(yy, mm, 0).getDate();

  document.querySelectorAll('.hubstaff-shift-cell').forEach(btn => {
    btn.onclick = () => {
      const staffId = btn.dataset.staff;
      const day = +btn.dataset.day;
      const date = dateKey(day);
      const coveringEntry = Object.entries(state.coverAssignments || {}).find(([key, cover]) => key.startsWith(`${date}:`) && cover.coverWorkerId === staffId);
      const editorStaffId = coveringEntry ? coveringEntry[0].slice(date.length + 1) : staffId;
      selected = { staff: editorStaffId, day };
      document.querySelectorAll('.hubstaff-shift-cell.selected').forEach(c => c.classList.remove('selected'));
      const selectedCell = document.querySelector(`.hubstaff-shift-cell[data-staff="${editorStaffId}"][data-day="${day}"]`);
      selectedCell?.classList.add('selected');
      openShiftEditor(editorStaffId, day);
    };

    btn.onkeydown = e => {
      const keyUpper = e.key.toUpperCase();
      if (['M', 'N', 'L', 'BACKSPACE', 'DELETE', 'SPACE'].includes(keyUpper) || e.key === ' ') {
        e.preventDefault();
        const newVal = (keyUpper === 'BACKSPACE' || keyUpper === 'DELETE') ? 'OFF'
          : keyUpper === 'L' ? 'LEAVE'
          : ['M', 'N'].includes(keyUpper) ? keyUpper : null;
        const day = +btn.dataset.day, date = dateKey(day);
        const coveringEntry = Object.entries(state.coverAssignments || {}).find(([key, cover]) => key.startsWith(`${date}:`) && cover.coverWorkerId === btn.dataset.staff);
        const editorStaffId = coveringEntry ? coveringEntry[0].slice(date.length + 1) : btn.dataset.staff;
        selected = { staff: editorStaffId, day };
        openShiftEditor(editorStaffId, day, newVal);
      }
      if (e.key.startsWith('Arrow')) {
        e.preventDefault();
        const allCells = [...document.querySelectorAll('.hubstaff-shift-cell')];
        const idx = allCells.indexOf(btn);
        const todayObj = new Date();
        const isCurrentMonth = (month === todayObj.toISOString().slice(0, 7));
        const todayDay = isCurrentMonth ? todayObj.getDate() : 1;
        const stride = (isCurrentMonth && todayDay > 1 && !rosterShowPast) ? (daysInMonth - todayDay + 1) : daysInMonth;
        const nextIdx = idx + (e.key === 'ArrowRight' ? 1 : e.key === 'ArrowLeft' ? -1 : e.key === 'ArrowDown' ? stride : -stride);
        allCells[Math.max(0, Math.min(allCells.length - 1, nextIdx))]?.focus();
      }
    };
  });

  // Sync D1 Requests
  const syncBtn = document.querySelector('#btn-sync-d1');
  if (syncBtn) {
    syncBtn.onclick = () => requireCloudflareConnection(async () => {
      const tok = sessionStorage.getItem('admin-session') || '';
      try {
        await syncDirectory();
        const res = await fetch(API + '/api/admin/requests', { headers: { authorization: 'Bearer ' + tok } });
        if (!res.ok) throw Error();
        const data = await res.json();
        availableStaff = data.availableStaff || [];
        requests = [
          ...data.covers.map(x => ({
            id: x.id,
            type: 'cover',
            name: x.requester,
            requesterId: x.requester_id,
            targetWorkerId: x.target_worker_id,
            originalDate: x.original_date,
            shiftCode: x.shift_code,
            reason: x.reason || '',
            createdAt: x.created_at,
            text: `Shift Cover · ${x.original_date} · Shift ${x.shift_code}`
          })),
          ...data.unavailability.map(x => ({
            id: x.id,
            type: 'unavailability',
            staffId: x.staff_id,
            name: x.full_name,
            startDate: x.start_date,
            endDate: x.end_date,
            reason: x.reason || '',
            createdAt: x.created_at,
            text: `Unavailable · ${x.start_date} to ${x.end_date}`
          })),
          ...(data.accounts || []).map(x => ({
            id: x.id,
            type: 'account',
            name: x.full_name,
            email: x.email,
            createdAt: x.created_at,
            text: `New account · ${x.email}`
          }))
        ];
        render();
        toast(`Synced ${requests.length} requests & updated staff`);
      } catch {
        toast('Could not reach Worker API');
      }
    });
  }

  // Unavailability Review & Approve
  document.querySelectorAll('[data-review-unavail-id]').forEach(btn => {
    btn.onclick = () => {
      const id = btn.dataset.reviewUnavailId;
      const r = requests.find(x => x.id === id);
      if (r) openResolveUnavailabilityDialog(r);
    };
  });

  // Unavailability Reject with confirmation
  document.querySelectorAll('[data-reject-unavail-id]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.rejectUnavailId;
      const r = requests.find(x => x.id === id);
      if (!r) return;
      if (!confirm(`Are you sure you want to reject the unavailability request from ${r.name} (${r.startDate} to ${r.endDate})?`)) {
        return;
      }
      try {
        const res = await fetch(API + '/api/admin/requests/' + id, {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + (sessionStorage.getItem('admin-session') || '') },
          body: JSON.stringify({ status: 'rejected' })
        });
        if (!res.ok) throw new Error();
        requests = requests.filter(x => x.id !== id);
        render();
        toast('Request rejected.');
      } catch {
        toast('Error rejecting request');
      }
    };
  });

  // Account Approve
  document.querySelectorAll('[data-approve-account-id]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.approveAccountId;
      const r = requests.find(x => x.id === id);
      if (!r) return;
      const staffSelect = document.getElementById(`account-staff-${CSS.escape(id)}`);
      const staffId = staffSelect?.value;
      if (!staffId) {
        toast('Choose a matching staff member first');
        return;
      }
      const matched = availableStaff.find(s => s.id === staffId);
      if (!confirm(`Approve account registration for ${r.name} (${r.email}) linked to ${matched?.full_name || staffId}?`)) {
        return;
      }
      try {
        const res = await fetch(API + '/api/admin/accounts/' + id, {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + (sessionStorage.getItem('admin-session') || '') },
          body: JSON.stringify({ status: 'approved', staffId })
        });
        if (!res.ok) throw new Error();
        requests = requests.filter(x => x.id !== id);
        render();
        toast(`Account approved and linked to ${matched?.full_name || staffId}.`);
      } catch {
        toast('Error approving account');
      }
    };
  });

  // Account Reject
  document.querySelectorAll('[data-reject-account-id]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.rejectAccountId;
      const r = requests.find(x => x.id === id);
      if (!r) return;
      if (!confirm(`Reject account registration for ${r.name} (${r.email})?`)) {
        return;
      }
      try {
        const res = await fetch(API + '/api/admin/accounts/' + id, {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + (sessionStorage.getItem('admin-session') || '') },
          body: JSON.stringify({ status: 'rejected' })
        });
        if (!res.ok) throw new Error();
        requests = requests.filter(x => x.id !== id);
        render();
        toast('Account registration rejected.');
      } catch {
        toast('Error rejecting account');
      }
    };
  });

  // Fallback / legacy request buttons
  document.querySelectorAll('[data-legacy-req-id]').forEach(btn => {
    btn.onclick = async () => {
      const id = btn.dataset.legacyReqId;
      const action = btn.dataset.reqAction;
      try {
        const res = await fetch(API + '/api/admin/requests/' + id, {
          method: 'POST',
          headers: { 'content-type': 'application/json', authorization: 'Bearer ' + (sessionStorage.getItem('admin-session') || '') },
          body: JSON.stringify({ status: action })
        });
        if (!res.ok) throw new Error();
        requests = requests.filter(x => x.id !== id);
        render();
        toast(`Request ${action}`);
      } catch {
        toast('Error updating request');
      }
    };
  });
}

// Initial Run
render();
syncDirectory();
if (window.shiftDexDesktop?.updates) {
  let pendingStateTimer = null;
  window.shiftDexDesktop.updates.onStatus(state => {
    const hadUpdate = ['available', 'downloading', 'downloaded'].includes(desktopUpdateState.status);
    const remaining = Math.max(0, manualCheckUntil - Date.now());
    if (remaining > 0 && state.status !== 'checking') {
      clearTimeout(pendingStateTimer);
      pendingStateTimer = setTimeout(() => {
        desktopUpdateState = state;
        const hasUpdate = ['available', 'downloaded'].includes(state.status);
        if (hasUpdate && !hadUpdate) toast(state.message || `ShiftDex ${state.version || ''} is available.`);
        render();
      }, remaining);
      return;
    }
    desktopUpdateState = state;
    const hasUpdate = ['available', 'downloaded'].includes(state.status);
    if (hasUpdate && !hadUpdate) toast(state.message || `ShiftDex ${state.version || ''} is available.`);
    render();
  });
  window.shiftDexDesktop.updates.getState().then(state => {
    desktopUpdateState = state;
    render();
  }).catch(() => {});
}
