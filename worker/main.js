import { staff as seedStaff, pods, cellFor, cycle, initialState } from './roster.js';
import { API_BASE_URL } from './api-config.js';
import shiftDeskMark from './shift-desk-mark.svg';
import { mountLogin } from './login.js';
import { startRegistration, browserSupportsWebAuthn } from '@simplewebauthn/browser';

const api = API_BASE_URL.replace(/\/$/, '');
const app = document.querySelector('#app');
for (const key of ['worker-token', 'worker-id', 'worker-profile', 'shift-roster']) {
  const legacyValue = localStorage.getItem(key);
  if (legacyValue !== null) {
    if (sessionStorage.getItem(key) === null) sessionStorage.setItem(key, legacyValue);
    localStorage.removeItem(key);
  }
}
let token = sessionStorage.getItem('worker-token') || '';
let staffId = sessionStorage.getItem('worker-id') || '';
let profile = JSON.parse(sessionStorage.getItem('worker-profile') || 'null');
let staff = seedStaff;
let state = JSON.parse(sessionStorage.getItem('shift-roster') || 'null') || initialState();

// View state
let activeWorkerTab = 'overview'; // 'overview' | 'calendar' | 'requests'
let currentMonth = new Date().toISOString().slice(0, 7);
let selectedDayDetail = null;
let isMobileNavOpen = false;
let workerRequests = [];

const ICONS = {
  home: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="m3 9 9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/></svg>`,
  calendar: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>`,
  requests: `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 8v4l3 3"/><circle cx="12" cy="12" r="9"/></svg>`,
  arrowUpRight: `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M7 17L17 7"/><path d="M7 7h10v10"/></svg>`,
  swap: `<svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M16 3l4 4-4 4"/><path d="M20 7H4"/><path d="M8 21l-4-4 4-4"/><path d="M4 17h16"/></svg>`,
};

const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, char => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
}[char]));

const isoDate = date => {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
};

const niceDate = value => new Date(`${value}T12:00:00Z`).toLocaleDateString('en-US', {
  weekday: 'short', day: 'numeric', month: 'short', timeZone: 'UTC',
});

const fullDate = value => new Date(`${value}T12:00:00Z`).toLocaleDateString('en-US', {
  weekday: 'long', day: 'numeric', month: 'long', year: 'numeric', timeZone: 'UTC',
});

const formatMonthHeader = monthKey => {
  const [y, m] = monthKey.split('-').map(Number);
  return new Date(Date.UTC(y, m - 1, 1)).toLocaleDateString('en-US', {
    month: 'long', year: 'numeric', timeZone: 'UTC',
  });
};

function shiftFor(personId, date) {
  const key = `${date}:${personId}`;
  if (Object.hasOwn(state.overrides || {}, key) || Object.hasOwn(state.assignments || {}, key)) return cellFor(state, personId, date);
  const person = staff.find(item => item.id === personId);
  if (seedStaff.some(item => item.id === personId)) return cellFor(state, personId, date);
  const monthly = state.staffMonthly?.[`${date.slice(0, 7)}:${personId}`];
  const podId = monthly?.podId || person?.podId;
  if (!podId || date < (person?.effectiveFrom || '2026-11-01')) return 'OFF';
  const pod = pods.find(item => item.id === podId);
  return pod ? cycle(date, pod.anchor) : 'OFF';
}

async function apiRequest(path, options = {}) {
  const response = await fetch(`${api}${path}`, {
    ...options,
    headers: {
      ...(options.body ? { 'content-type': 'application/json' } : {}),
      ...(token ? { authorization: `Bearer ${token}` } : {}),
      ...options.headers,
    },
  });
  const data = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(data.error || 'Could not connect to ShiftDex.');
  return data;
}

function renderLogin() {
  mountLogin(app, api, result => {
    token = result.token;
    staffId = result.staffId;
    profile = { fullName: result.fullName, email: result.email };
    sessionStorage.setItem('worker-token', token);
    sessionStorage.setItem('worker-id', staffId);
    sessionStorage.setItem('worker-profile', JSON.stringify(profile));
    boot();
  });
}

function showWorkerToast(message, type = 'success') {
  let toast = document.querySelector('#worker-toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.className = 'worker-toast';
    toast.id = 'worker-toast';
    document.body.appendChild(toast);
  }
  const isErr = (type === 'error');
  toast.className = `worker-toast ${isErr ? 'error' : 'success'}`;
  toast.innerHTML = `
    <span class="worker-toast-icon ${isErr ? 'error' : ''}">
      ${isErr 
        ? `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`
        : `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg>`}
    </span>
    <span class="worker-toast-text">${escapeHtml(message)}</span>
    <button type="button" class="worker-toast-close" aria-label="Dismiss">&times;</button>
  `;
  const closeBtn = toast.querySelector('.worker-toast-close');
  if (closeBtn) {
    closeBtn.onclick = () => {
      toast.classList.remove('show');
    };
  }
  void toast.offsetWidth;
  toast.classList.add('show');
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.classList.remove('show'), 3500);
}

function showMessage(message) {
  showWorkerToast(message, /error|invalid|fail|choose|earlier/i.test(message) ? 'error' : 'success');
}

function renderDashboard() {
  const today = new Date();
  const todayIso = isoDate(today);
  const me = staff.find(person => person.id === staffId) || {
    id: staffId, name: profile?.fullName || 'Staff Member', role: 'Staff',
  };

  const myPod = pods.find(p => p.id === me.podId);
  const teamName = myPod?.name || (me.podId ? me.podId : 'Day Support');
  const teamColor = myPod?.color || '#64748b';
  const initial = (me.name || 'W').charAt(0).toUpperCase();

  const todayShift = shiftFor(staffId, todayIso) || 'OFF';
  const currentHour = today.getHours();
  const currentDuty = currentHour >= 6 && currentHour < 18 ? 'M' : 'N';
  const onDesk = staff.filter(person => person.podId && shiftFor(person.id, todayIso) === currentDuty);

  // Next upcoming working shift
  const nextShiftDay = Array.from({ length: 45 }, (_, index) => {
    const d = new Date(today);
    d.setDate(d.getDate() + index + (todayShift === 'OFF' ? 0 : 1));
    return isoDate(d);
  }).find(day => ['M', 'N'].includes(shiftFor(staffId, day)));

  // Calendar calculations for currentMonth
  const [year, monthNum] = currentMonth.split('-').map(Number);
  const daysInMonth = new Date(year, monthNum, 0).getDate();
  const firstDayOfWeek = new Date(Date.UTC(year, monthNum - 1, 1)).getUTCDay();
  const daysInPrevMonth = new Date(year, monthNum - 1, 0).getDate();

  // Monthly stats for currentMonth
  let monthlyShiftsCount = 0;
  let morningShiftsCount = 0;
  let nightShiftsCount = 0;
  let restDaysCount = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dKey = `${currentMonth}-${String(d).padStart(2, '0')}`;
    const code = shiftFor(staffId, dKey);
    if (code === 'M') { monthlyShiftsCount++; morningShiftsCount++; }
    else if (code === 'N') { monthlyShiftsCount++; nightShiftsCount++; }
    else { restDaysCount++; }
  }

  // Calendar cells
  const calendarCells = [];

  const renderCalShift = (code) => {
    if (code === 'M') {
      return `
        <div class="cal-shift-pill m">
          <div class="shift-pill-header">
            <span class="shift-pill-badge">M</span>
            <span class="shift-pill-title">Day</span>
          </div>
          <div class="shift-pill-time">06:00-18:00</div>
        </div>
      `;
    }
    if (code === 'N') {
      return `
        <div class="cal-shift-pill n">
          <div class="shift-pill-header">
            <span class="shift-pill-badge">N</span>
            <span class="shift-pill-title">Night</span>
          </div>
          <div class="shift-pill-time">18:00-06:00</div>
        </div>
      `;
    }
    if (code === 'LEAVE') {
      return `
        <div class="cal-shift-pill leave">
          <div class="shift-pill-header">
            <span class="shift-pill-badge">LV</span>
            <span class="shift-pill-title">Leave</span>
          </div>
          <div class="shift-pill-time">Approved</div>
        </div>
      `;
    }
    return `
      <div class="cal-shift-pill off">
        <span class="shift-off-pill">Rest</span>
      </div>
    `;
  };

  // Prev month leading days
  const prevMonthKey = new Date(Date.UTC(year, monthNum - 2, 1)).toISOString().slice(0, 7);
  for (let i = firstDayOfWeek - 1; i >= 0; i--) {
    const dayNum = daysInPrevMonth - i;
    const dateStr = `${prevMonthKey}-${String(dayNum).padStart(2, '0')}`;
    const code = shiftFor(staffId, dateStr);
    calendarCells.push(`
      <div class="cal-cell other-month" data-date="${dateStr}">
        <div class="cal-cell-head">
          <span class="cal-day-num">${dayNum}</span>
        </div>
        ${renderCalShift(code)}
      </div>
    `);
  }

  // Current month days
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${currentMonth}-${String(d).padStart(2, '0')}`;
    const code = shiftFor(staffId, dateStr);
    const isToday = dateStr === todayIso;

    calendarCells.push(`
      <div class="cal-cell ${isToday ? 'is-today' : ''}" data-date="${dateStr}" role="button" tabindex="0">
        <div class="cal-cell-head">
          <span class="cal-day-num">${d}</span>
          ${isToday ? '<span class="cal-today-tag">TODAY</span>' : ''}
        </div>
        ${renderCalShift(code)}
      </div>
    `);
  }

  // Next month trailing days
  const remainingCells = (7 - (calendarCells.length % 7)) % 7;
  const nextMonthKey = new Date(Date.UTC(year, monthNum, 1)).toISOString().slice(0, 7);
  for (let d = 1; d <= remainingCells; d++) {
    const dateStr = `${nextMonthKey}-${String(d).padStart(2, '0')}`;
    const code = shiftFor(staffId, dateStr);
    calendarCells.push(`
      <div class="cal-cell other-month" data-date="${dateStr}">
        <div class="cal-cell-head">
          <span class="cal-day-num">${d}</span>
        </div>
        ${renderCalShift(code)}
      </div>
    `);
  }

  // Upcoming 7 Days timeline cards (uiux.aditya Image 3 style)
  const upcomingCards = Array.from({ length: 7 }, (_, i) => {
    const d = new Date(today);
    d.setDate(d.getDate() + i);
    const dStr = isoDate(d);
    const code = shiftFor(staffId, dStr);
    const isT = (i === 0);
    const weekdayStr = isT ? 'Today' : (i === 1 ? 'Tomorrow' : d.toLocaleDateString('en-US', { weekday: 'short' }));
    const dayNumStr = d.getDate();
    const monthShort = d.toLocaleDateString('en-US', { month: 'short' });

    const shiftTitle = code === 'M' ? 'Day Shift · 06:00 - 18:00'
      : code === 'N' ? 'Night Shift · 18:00 - 06:00'
      : code === 'LEAVE' ? 'Approved Annual Leave'
      : 'Rest Day (OFF)';

    const colleaguesCount = staff.filter(s => s.podId && shiftFor(s.id, dStr) === code).length;

    return `
      <div class="timeline-day-card ${isT ? 'is-today' : ''}" data-date="${dStr}" role="button" tabindex="0">
        <div class="timeline-day-left">
          <div class="timeline-date-block">
            <span class="timeline-weekday">${weekdayStr}</span>
            <span class="timeline-daynum">${dayNumStr}</span>
            <small style="font-size:9.5px;color:var(--text-muted);display:block;">${monthShort}</small>
          </div>
          <div class="timeline-info">
            <strong>${shiftTitle}</strong>
            <span>${code === 'M' || code === 'N' ? `${colleaguesCount} on desk with you` : 'Continuous 8-day rotation rest'}</span>
          </div>
        </div>
        <div>
          <span class="shift-code-box ${code.toLowerCase()}" style="width:36px;height:36px;font-size:13px;">${code}</span>
        </div>
      </div>
    `;
  }).join('');

  // 8-Day Continuous Rotation Cycle Calculation
  const anchorDate = myPod?.anchor || '2026-11-01';
  const dayDiffFromAnchor = Math.floor(Date.parse(todayIso + 'T00:00:00Z') / 86400000) - Math.floor(Date.parse(anchorDate + 'T00:00:00Z') / 86400000);
  const currentRotationPhase = ((dayDiffFromAnchor % 8) + 8) % 8; // 0 to 7

  const rotationSteps = [
    { phase: 0, code: 'M', title: 'Day 1', label: 'Day Shift' },
    { phase: 1, code: 'M', title: 'Day 2', label: 'Day Shift' },
    { phase: 2, code: 'N', title: 'Night 1', label: 'Night Shift' },
    { phase: 3, code: 'N', title: 'Night 2', label: 'Night Shift' },
    { phase: 4, code: 'OFF', title: 'Rest 1', label: 'Rest Day' },
    { phase: 5, code: 'OFF', title: 'Rest 2', label: 'Rest Day' },
    { phase: 6, code: 'OFF', title: 'Rest 3', label: 'Rest Day' },
    { phase: 7, code: 'OFF', title: 'Rest 4', label: 'Rest Day' },
  ];

  const currentStep = rotationSteps[currentRotationPhase];
  const morningPct = Math.round((morningShiftsCount / daysInMonth) * 100);
  const nightPct = Math.round((nightShiftsCount / daysInMonth) * 100);
  const restPct = Math.max(0, 100 - morningPct - nightPct);

  app.innerHTML = `
    <!-- Top Navigation Bar (Image 2 style with Mobile Avatar + Greeting) -->
    <header class="app-header">
      <div class="mobile-header-user">
        <div class="user-avatar-circle" style="background:${teamColor};">${initial}</div>
        <div class="mobile-header-text">
          <span class="mobile-header-greeting">Hey,</span>
          <strong class="mobile-header-name">${escapeHtml(me.name.split(' ')[0] || me.name)}</strong>
        </div>
      </div>

      <div class="brand-wrap">
        <img src="${shiftDeskMark}" alt="ShiftDex logo">
        <b>shiftdex<span style="color:var(--blue-primary);">.</span></b>
      </div>

      <!-- Desktop User Profile Controls -->
      <div class="header-user-side">
        <div class="user-profile-badge">
          <div class="user-avatar-circle" style="background:${teamColor};">${initial}</div>
          <span class="user-badge-name">${escapeHtml(me.name)}</span>
          <span class="user-team-chip" style="border-color:${teamColor};">${escapeHtml(teamName)}</span>
        </div>
        <button class="btn-signout" id="btn-worker-change-pw" type="button" style="border-color:var(--border-card);color:var(--text-main);background:#ffffff;">Security &amp; Passkeys</button>
        <button class="btn-signout" id="logout" type="button">Sign out</button>
      </div>

      <!-- Mobile Hamburger Button (Image 2 circular outline button) -->
      <button class="hamburger-btn" id="btn-hamburger" type="button" aria-label="Open menu">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><line x1="3" y1="12" x2="21" y2="12"/><line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="18" x2="21" y2="18"/></svg>
      </button>
    </header>

    <!-- Mobile Slide-out Drawer -->
    <div class="mobile-nav-backdrop ${isMobileNavOpen ? 'open' : ''}" id="mobile-nav-backdrop"></div>
    <div class="mobile-nav-drawer ${isMobileNavOpen ? 'open' : ''}" id="mobile-nav-drawer">
      <div>
        <div class="drawer-head">
          <div class="brand-wrap">
            <img src="${shiftDeskMark}" alt="ShiftDex logo">
            <b>shiftdex<span style="color:var(--blue-primary);">.</span></b>
          </div>
          <button type="button" id="btn-close-drawer" style="border:none;background:none;font-size:24px;cursor:pointer;color:var(--text-light);">&times;</button>
        </div>
        <div class="drawer-user-info">
          <div class="user-avatar-circle" style="background:${teamColor};width:42px;height:42px;font-size:16px;">${initial}</div>
          <div>
            <div style="font-weight:700;font-size:15px;color:var(--text-main);">${escapeHtml(me.name)}</div>
            <div style="display:flex;align-items:center;gap:6px;margin-top:4px;">
              <span class="user-team-chip" style="border-color:${teamColor};">${escapeHtml(teamName)}</span>
              <span style="font-size:12px;color:var(--text-light);">${escapeHtml(me.role)}</span>
            </div>
          </div>
        </div>
        <div style="margin:14px 0 4px;padding:12px 14px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:14px;display:flex;align-items:center;justify-content:space-between;gap:10px;">
          <div style="display:flex;align-items:center;gap:8px;">
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 10a4 4 0 1 0 4 4"/><circle cx="12" cy="12" r="9"/><path d="m14 10 5-5"/><path d="m17 7 2 2"/></svg>
            <span style="font-size:12px;font-weight:700;color:#1e40af;">Biometric Passkeys</span>
          </div>
          <button type="button" id="drawer-manage-passkeys" style="border:none;background:#2563eb;color:#fff;border-radius:9999px;padding:5px 12px;font-size:11px;font-weight:750;cursor:pointer;">Set Up</button>
        </div>
      </div>
      <div style="display:flex;flex-direction:column;gap:10px;">
        <button class="btn-signout" id="drawer-worker-change-pw" type="button" style="width:100%;height:46px;border-radius:14px;border-color:var(--border-card);color:var(--text-main);background:#ffffff;font-weight:600;">Security &amp; Passkeys</button>
        <button class="btn-signout" id="drawer-logout" type="button" style="width:100%;height:46px;border-radius:14px;background:#121316;color:#ffffff;font-weight:700;">Sign out</button>
      </div>
    </div>

    <main class="page-wrap">
      <!-- 1. OVERVIEW VIEW (Apple Clean White Aesthetic with Aesthetic Ticket Cuts) -->
      ${activeWorkerTab === 'overview' ? `
        <section class="worker-home neo-home">
          <!-- HERO CARD: Today's Shift & Next Duty with Ticket Cut Notches -->
          <div class="neo-ticket-card">
            <!-- Top Edge Cut Notch -->
            <span class="card-edge-cut cut-top"></span>

            <!-- Center Top Pill Badge -->
            <div class="ticket-badge-wrap">
              <div class="neo-hero-top-badge">
                <span class="neo-team-dot" style="background:${teamColor};"></span>
                <span>${escapeHtml(teamName)} &middot; ${escapeHtml(me.role)}</span>
              </div>
            </div>

            <!-- Two Metric Columns -->
            <div class="neo-hero-metrics-row">
              <div class="neo-hero-metric-col">
                <span class="neo-metric-label">CURRENT STATUS</span>
                <div class="neo-metric-value">${todayShift === 'M' ? 'Morning Shift' : todayShift === 'N' ? 'Night Shift' : todayShift === 'LEAVE' ? 'On Leave' : 'Rest Day'}</div>
                <span class="neo-metric-subtext">${todayShift === 'M' ? '06:00 – 18:00 · Day' : todayShift === 'N' ? '18:00 – 06:00 · Night' : todayShift === 'LEAVE' ? 'Approved Leave' : 'No shift scheduled today'}</span>
              </div>

              <div class="neo-hero-metric-col right">
                <span class="neo-metric-label">NEXT ON DUTY</span>
                <div class="neo-metric-value">${nextShiftDay ? niceDate(nextShiftDay) : 'OFF'}</div>
                <span class="neo-metric-subtext">${nextShiftDay ? (shiftFor(staffId, nextShiftDay) === 'M' ? 'Day (06:00)' : 'Night (18:00)') : 'Resting Cadence'}</span>
              </div>
            </div>

            <!-- Aesthetic Ticket Notch Cut Divider -->
            <div class="ticket-notch-divider">
              <span class="notch-left"></span>
              <span class="notch-line"></span>
              <span class="notch-right"></span>
            </div>

            <!-- Dual Pill Buttons (Signature ShiftDex Blue & Clean White) -->
            <div class="neo-hero-actions-row">
              <button type="button" class="neo-btn-primary-blue" data-tab="requests">
                <span class="neo-btn-icon-circle">${ICONS.swap}</span>
                <span>Request Off</span>
              </button>
              <button type="button" class="neo-btn-secondary" data-tab="calendar">
                <span class="neo-btn-icon-circle secondary">${ICONS.arrowUpRight}</span>
                <span>Full Calendar</span>
              </button>
            </div>

            <!-- Bottom Edge Cut Notch -->
            <span class="card-edge-cut cut-bottom"></span>
          </div>

          <!-- 2. ON-DUTY CREW CARD (Clean White Card with Ticket Cutout) -->
          <section class="neo-white-card neo-desk-card">
            <!-- Top Edge Cut Notch -->
            <span class="card-edge-cut cut-top"></span>

            <div class="desk-card-main-content">
              <h3>${currentDuty === 'M' ? 'Morning Shift Crew' : 'Night Shift Crew'}</h3>
              <p>${currentDuty === 'M' ? '06:00 – 18:00' : '18:00 – 06:00'} <span>&middot;</span> Handover at ${currentDuty === 'M' ? '18:00' : '06:00'}</p>
            </div>

            <!-- Aesthetic Ticket Notch Cut Divider -->
            <div class="ticket-notch-divider subtle">
              <span class="notch-left"></span>
              <span class="notch-line"></span>
              <span class="notch-right"></span>
            </div>

            <!-- Crew Member Rounded White Pills -->
            <div class="desk-crew-pills-row">
              ${onDesk.map(person => {
                const pPod = pods.find(p => p.id === person.podId);
                const pColor = pPod?.color || '#3b82f6';
                return `
                  <div class="desk-crew-chip">
                    <span class="crew-chip-dot" style="background:${pColor};"></span>
                    <span class="crew-chip-name">${escapeHtml(person.name)}</span>
                    <span class="crew-chip-role">${escapeHtml(person.role)}</span>
                  </div>
                `;
              }).join('') || '<div class="desk-crew-chip"><span>Shift handover in progress</span></div>'}
            </div>

            <!-- Bottom Edge Cut Notch -->
            <span class="card-edge-cut cut-bottom"></span>
          </section>

          <!-- 3. UPCOMING SHIFTS PANEL ("Your Week" Agenda) -->
          <section class="neo-white-panel">
            <!-- Top Edge Cut Notch -->
            <span class="card-edge-cut cut-top"></span>

            <div class="neo-panel-heading">
              <div>
                <span class="neo-panel-kicker">YOUR WEEK</span>
                <h2>Upcoming shifts</h2>
              </div>
              <button type="button" class="neo-panel-link" data-tab="calendar">
                <span>See calendar</span>
                <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"/><path d="m12 5 7 7-7 7"/></svg>
              </button>
            </div>
            <div class="home-agenda-list">${upcomingCards}</div>

            <!-- Bottom Edge Cut Notch -->
            <span class="card-edge-cut cut-bottom"></span>
          </section>
        </section>
      ` : ''}
      <!-- 2. CALENDAR VIEW (Image 2 Wylro & Image 4 Pocket Cal) -->
      ${activeWorkerTab === 'calendar' ? `
        <div class="calendar-wylro-layout">
          <!-- Main Calendar Grid Column -->
          <div class="calendar-main-col">
            <section class="calendar-card">
              <div class="calendar-toolbar">
                <div class="calendar-title-wrap">
                  <h2 class="calendar-title">Monthly Calendar</h2>
                  <div class="month-nav-pill">
                    <button type="button" class="month-nav-btn" id="btn-month-prev" title="Previous month">&larr;</button>
                    <span class="month-nav-label">${formatMonthHeader(currentMonth)}</span>
                    <button type="button" class="month-nav-btn" id="btn-month-next" title="Next month">&rarr;</button>
                  </div>
                </div>
                <div style="display:flex;align-items:center;gap:8px;">
                  <button type="button" class="btn-signout" id="btn-month-today" style="border-color:#cbd5e1;color:var(--text-main);height:30px;padding:0 12px;font-size:11.5px;">Today</button>
                </div>
              </div>

              <div class="calendar-grid-container">
                <div class="calendar-weekdays-row">
                  <div class="weekday-header">Sun</div>
                  <div class="weekday-header">Mon</div>
                  <div class="weekday-header">Tue</div>
                  <div class="weekday-header">Wed</div>
                  <div class="weekday-header">Thu</div>
                  <div class="weekday-header">Fri</div>
                  <div class="weekday-header">Sat</div>
                </div>
                <div class="calendar-days-grid">
                  ${calendarCells.join('')}
                </div>
              </div>
            </section>
          </div>

          <!-- Right Sidebar Column (Wylro desktop style) -->
          <div class="calendar-side-col">
            <!-- Shift Breakdown Card -->
            <div class="side-summary-card">
              <div class="side-card-title">
                <span>Monthly Total</span>
                <span style="font-size:12px;color:var(--blue-primary);">${monthlyShiftsCount * 12} hrs</span>
              </div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px;">
                <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px;text-align:center;">
                  <span style="font-size:11px;color:var(--text-light);display:block;font-weight:600;">Day Shifts</span>
                  <strong style="font-size:18px;color:#92400e;">${morningShiftsCount}</strong>
                </div>
                <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:10px;text-align:center;">
                  <span style="font-size:11px;color:var(--text-light);display:block;font-weight:600;">Night Shifts</span>
                  <strong style="font-size:18px;color:#1d4ed8;">${nightShiftsCount}</strong>
                </div>
              </div>
              <div class="shift-legend-list">
                <div class="legend-item">
                  <div class="legend-badge m">M</div>
                  <span>Day Shift (06:00 - 18:00)</span>
                </div>
                <div class="legend-item">
                  <div class="legend-badge n">N</div>
                  <span>Night Shift (18:00 - 06:00)</span>
                </div>
                <div class="legend-item">
                  <div class="legend-badge off">OFF</div>
                  <span>Rest Day</span>
                </div>
              </div>
            </div>
          </div>
        </div>
      ` : ''}

      <!-- 3. REQUESTS VIEW -->
      ${activeWorkerTab === 'requests' ? `
        <div style="display:flex;flex-direction:column;gap:18px;">
          <!-- Unavailable Dates Submission Form -->
          <section class="unavailability-card">
            <h2 class="card-heading">Mark Unavailable Dates</h2>
            <p class="card-desc">Notify operations of upcoming dates you cannot work before monthly shifts are finalized (up to 60 consecutive days per request).</p>

            <form id="unavailability-form" class="unavail-form">
              <div class="form-grid-2">
                <div class="form-field">
                  <label for="unavailable-start">First Unavailable Day</label>
                  <input id="unavailable-start" type="date" required min="${todayIso}" value="${todayIso}">
                </div>
                <div class="form-field">
                  <label for="unavailable-end">Last Unavailable Day</label>
                  <input id="unavailable-end" type="date" required min="${todayIso}" value="${todayIso}">
                </div>
              </div>

              <div class="form-field">
                <label for="unavailable-reason">Reason / Note (Optional)</label>
                <input id="unavailable-reason" type="text" placeholder="e.g. Medical appointment, university exam, personal travel">
              </div>

              <button type="submit" class="btn-submit-action" id="btn-submit-unavailability">
                Submit Unavailability to Operations
              </button>
            </form>
          </section>

          <!-- Worker's Submitted Requests History -->
          <section class="unavailability-card">
            <h2 class="card-heading">My Submitted Requests</h2>
            <p class="card-desc">Current status of your unavailable dates and leave requests.</p>
            <div class="requests-list-wrap">
              ${workerRequests.length ? workerRequests.map(req => {
                const isPending = req.status === 'pending';
                const isApproved = req.status === 'approved';
                return `
                  <div class="my-request-card">
                    <div>
                      <b>${req.start_date} &rarr; ${req.end_date}</b>
                      ${req.reason ? `<span class="meta">Note: ${escapeHtml(req.reason)}</span>` : '<span class="meta">No note provided</span>'}
                    </div>
                    <div>
                      <span class="req-status-pill ${isPending ? 'pending' : isApproved ? 'approved' : 'rejected'}">
                        ${isPending ? 'Pending Operations Review' : isApproved ? 'Approved · Off Scheduled' : 'Rejected'}
                      </span>
                    </div>
                  </div>
                `;
              }).join('') : `
                <div style="padding:28px 16px;text-align:center;color:var(--text-light);font-size:13px;">
                  You have not submitted any unavailable dates yet. Use the form above to submit dates you cannot work.
                </div>
              `}
            </div>
          </section>
        </div>
      ` : ''}

      <!-- Day Detail Modal -->
      ${selectedDayDetail ? `
        <div class="day-modal-backdrop" id="day-modal-backdrop">
          <div class="day-modal-card">
            <div class="day-modal-header">
              <h3>${fullDate(selectedDayDetail.date)}</h3>
              <button type="button" class="day-modal-close-btn" id="btn-close-day-modal">&times;</button>
            </div>
            <div class="day-modal-body">
              <div style="display:flex;align-items:center;gap:14px;">
                <div class="shift-code-box ${selectedDayDetail.code.toLowerCase()}">${selectedDayDetail.code}</div>
                <div>
                  <strong style="font-size:15px;color:var(--text-main);display:block;">
                    ${selectedDayDetail.code === 'M' ? 'Day Shift (06:00 - 18:00)' : selectedDayDetail.code === 'N' ? 'Night Shift (18:00 - 06:00)' : selectedDayDetail.code === 'LEAVE' ? 'Approved Leave' : 'Rest Day (Off Duty)'}
                  </strong>
                  <span style="font-size:12px;color:var(--text-light);">
                    ${selectedDayDetail.code === 'M' || selectedDayDetail.code === 'N' ? '12-Hour Continuous Operations Shift' : 'Rest period'}
                  </span>
                </div>
              </div>

              ${['M', 'N'].includes(selectedDayDetail.code) ? `
                <div style="margin-top:8px;">
                  <div style="font-size:11.5px;font-weight:700;color:var(--text-secondary);margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px;">
                    Colleagues on duty with you:
                  </div>
                  <div style="display:flex;flex-direction:column;gap:8px;">
                    ${staff.filter(person => person.podId && shiftFor(person.id, selectedDayDetail.date) === selectedDayDetail.code).map(person => {
                      const pPod = pods.find(p => p.id === person.podId);
                      const pColor = pPod?.color || '#94a3b8';
                      return `
                        <div style="display:flex;align-items:center;justify-content:space-between;padding:8px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;">
                          <div style="display:flex;align-items:center;gap:10px;">
                            <span style="width:8px;height:8px;border-radius:50%;background:${pColor};"></span>
                            <span style="font-weight:700;font-size:13px;color:var(--text-main);">${escapeHtml(person.name)}</span>
                          </div>
                          <span style="font-size:11px;color:var(--text-light);">${escapeHtml(person.role)}</span>
                        </div>
                      `;
                    }).join('') || '<div style="font-size:12px;color:var(--text-light);">No other crew assigned</div>'}
                  </div>
                </div>
              ` : `
                <div style="padding:16px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;text-align:center;font-size:12.5px;color:var(--text-light);">
                  You have no shift duty scheduled for this date.
                </div>
              `}
            </div>
          </div>
        </div>
      ` : ''}
    </main>

    <!-- Dark Matte Floating Bottom Navigation Bar (Image 2 style) -->
    <nav class="liquid-dock-wrap" role="navigation" aria-label="Primary Mobile Dock">
      <div class="liquid-glass-dock">
        <button type="button" class="liquid-dock-item ${activeWorkerTab === 'overview' ? 'active' : ''}" data-tab="overview" aria-label="Home Dashboard">
          <span class="dock-icon-wrap">${ICONS.home}</span>
          ${activeWorkerTab === 'overview' ? '<span class="dock-label">Home</span>' : ''}
        </button>
        <button type="button" class="liquid-dock-item ${activeWorkerTab === 'calendar' ? 'active' : ''}" data-tab="calendar" aria-label="Calendar Matrix">
          <span class="dock-icon-wrap">${ICONS.calendar}</span>
          ${activeWorkerTab === 'calendar' ? '<span class="dock-label">Calendar</span>' : ''}
        </button>
        <button type="button" class="liquid-dock-item ${activeWorkerTab === 'requests' ? 'active' : ''}" data-tab="requests" aria-label="Requests">
          <span class="dock-icon-wrap">${ICONS.requests}</span>
          ${activeWorkerTab === 'requests' ? '<span class="dock-label">Requests</span>' : ''}
        </button>
      </div>
    </nav>
  `;

  // Bind View Switcher events (Liquid Glass Dock items)
  document.querySelectorAll('[data-tab]').forEach(btn => {
    btn.onclick = () => {
      const targetTab = btn.dataset.tab;
      if (['overview', 'calendar', 'requests'].includes(targetTab)) {
        activeWorkerTab = targetTab;
        renderDashboard();
      }
    };
  });

  // Month navigation in calendar view
  const btnPrev = document.querySelector('#btn-month-prev');
  if (btnPrev) {
    btnPrev.onclick = () => {
      const [y, m] = currentMonth.split('-').map(Number);
      currentMonth = new Date(Date.UTC(y, m - 2, 1)).toISOString().slice(0, 7);
      renderDashboard();
    };
  }

  const btnNext = document.querySelector('#btn-month-next');
  if (btnNext) {
    btnNext.onclick = () => {
      const [y, m] = currentMonth.split('-').map(Number);
      currentMonth = new Date(Date.UTC(y, m, 1)).toISOString().slice(0, 7);
      renderDashboard();
    };
  }

  const btnToday = document.querySelector('#btn-month-today');
  if (btnToday) {
    btnToday.onclick = () => {
      currentMonth = new Date().toISOString().slice(0, 7);
      renderDashboard();
    };
  }

  const homePrevMonth = document.querySelector('#home-prev-month');
  if (homePrevMonth) homePrevMonth.onclick = () => {
    const [y, m] = currentMonth.split('-').map(Number);
    currentMonth = new Date(Date.UTC(y, m - 2, 1)).toISOString().slice(0, 7);
    renderDashboard();
  };
  const homeNextMonth = document.querySelector('#home-next-month');
  if (homeNextMonth) homeNextMonth.onclick = () => {
    const [y, m] = currentMonth.split('-').map(Number);
    currentMonth = new Date(Date.UTC(y, m, 1)).toISOString().slice(0, 7);
    renderDashboard();
  };

  // Day cell click -> open day detail modal
  document.querySelectorAll('.cal-cell:not(.other-month), .timeline-day-card').forEach(cell => {
    cell.onclick = () => {
      const date = cell.dataset.date;
      if (!date) return;
      const code = shiftFor(staffId, date);
      selectedDayDetail = { date, code };
      renderDashboard();
    };
  });

  // Close day detail modal
  const btnCloseDayModal = document.querySelector('#btn-close-day-modal');
  if (btnCloseDayModal) {
    btnCloseDayModal.onclick = () => {
      selectedDayDetail = null;
      renderDashboard();
    };
  }
  const dayModalBackdrop = document.querySelector('#day-modal-backdrop');
  if (dayModalBackdrop) {
    dayModalBackdrop.onclick = e => {
      if (e.target === dayModalBackdrop) {
        selectedDayDetail = null;
        renderDashboard();
      }
    };
  }

  // Mobile Hamburger Toggle
  const btnHamburger = document.querySelector('#btn-hamburger');
  const btnCloseDrawer = document.querySelector('#btn-close-drawer');
  const navBackdrop = document.querySelector('#mobile-nav-backdrop');

  if (btnHamburger) {
    btnHamburger.onclick = () => {
      isMobileNavOpen = true;
      renderDashboard();
    };
  }
  if (btnCloseDrawer) {
    btnCloseDrawer.onclick = () => {
      isMobileNavOpen = false;
      renderDashboard();
    };
  }
  if (navBackdrop) {
    navBackdrop.onclick = () => {
      isMobileNavOpen = false;
      renderDashboard();
    };
  }

  // Submit Unavailable Dates Handler (Reusable for in-page and modal forms)
  const submitUnavailable = async (startDate, endDate, reason, btnEl, formEl) => {
    if (!startDate || !endDate) {
      showWorkerToast('Please choose both start and end dates.', 'error');
      return;
    }
    if (endDate < startDate) {
      showWorkerToast('End date cannot be earlier than start date.', 'error');
      return;
    }

    const startObj = new Date(startDate + 'T00:00:00Z');
    const endObj = new Date(endDate + 'T00:00:00Z');
    const diffDays = Math.round((endObj - startObj) / (24 * 60 * 60 * 1000)) + 1;
    if (diffDays > 60) {
      showWorkerToast('Unavailability requests cannot exceed 60 consecutive days. For extended leaves, contact operations.', 'error');
      return;
    }

    const originalText = btnEl.textContent;
    btnEl.disabled = true;
    btnEl.innerHTML = '<span class="btn-spinner"></span> Submitting…';

    try {
      const res = await apiRequest('/api/unavailability', {
        method: 'POST',
        body: JSON.stringify({ startDate, endDate, reason }),
      });
      showWorkerToast('Unavailable dates submitted to operations.');
      btnEl.disabled = false;
      btnEl.textContent = originalText;
      if (formEl) formEl.reset();

      // Add to local requests list
      workerRequests.unshift({
        id: res.id || crypto.randomUUID(),
        start_date: startDate,
        end_date: endDate,
        reason,
        status: 'pending',
      });

      // Update the requests list in-place without reloading the entire dashboard!
      const requestsWrap = document.querySelector('.requests-list-wrap');
      if (requestsWrap) {
        requestsWrap.innerHTML = workerRequests.map(req => {
          const isPending = req.status === 'pending';
          const isApproved = req.status === 'approved';
          return `
            <div class="my-request-card" style="animation: fade-in-up 0.25s ease;">
              <div>
                <b>${escapeHtml(req.start_date)} &rarr; ${escapeHtml(req.end_date)}</b>
                ${req.reason ? `<span class="meta">Note: ${escapeHtml(req.reason)}</span>` : '<span class="meta">No note provided</span>'}
              </div>
              <div>
                <span class="req-status-pill ${isPending ? 'pending' : isApproved ? 'approved' : 'rejected'}">
                  ${isPending ? 'Pending Operations Review' : isApproved ? 'Approved · Off Scheduled' : 'Rejected'}
                </span>
              </div>
            </div>
          `;
        }).join('');
      } else {
        renderDashboard();
      }
    } catch (error) {
      showWorkerToast(error.message || 'Failed to submit request', 'error');
      btnEl.disabled = false;
      btnEl.textContent = originalText;
    }
  };

  // Submit form on page (Requests tab)
  const unavailForm = document.querySelector('#unavailability-form');
  if (unavailForm) {
    unavailForm.onsubmit = event => {
      event.preventDefault();
      const startDate = document.querySelector('#unavailable-start')?.value;
      const endDate = document.querySelector('#unavailable-end')?.value;
      const reason = document.querySelector('#unavailable-reason')?.value || '';
      const submitBtn = document.querySelector('#btn-submit-unavailability');
      submitUnavailable(startDate, endDate, reason, submitBtn, unavailForm);
    };
  }

  // Sign out handlers
  const handleLogout = async () => {
    try { await apiRequest('/api/auth/logout', { method: 'POST' }); } catch {}
    sessionStorage.removeItem('worker-token');
    sessionStorage.removeItem('worker-id');
    sessionStorage.removeItem('worker-profile');
    sessionStorage.removeItem('shift-roster');
    token = '';
    staffId = '';
    profile = null;
    renderLogin();
  };

  const logoutBtn = document.querySelector('#logout');
  if (logoutBtn) logoutBtn.onclick = handleLogout;
  const drawerLogoutBtn = document.querySelector('#drawer-logout');
  if (drawerLogoutBtn) drawerLogoutBtn.onclick = handleLogout;

  const workerChangePwBtn = document.querySelector('#btn-worker-change-pw');
  if (workerChangePwBtn) workerChangePwBtn.onclick = () => openWorkerSecurityModal('passkeys');
  const drawerWorkerChangePwBtn = document.querySelector('#drawer-worker-change-pw');
  if (drawerWorkerChangePwBtn) {
    drawerWorkerChangePwBtn.onclick = () => {
      isMobileNavOpen = false;
      renderDashboard();
      openWorkerSecurityModal('passkeys');
    };
  }
  const drawerManagePasskeysBtn = document.querySelector('#drawer-manage-passkeys');
  if (drawerManagePasskeysBtn) {
    drawerManagePasskeysBtn.onclick = () => {
      isMobileNavOpen = false;
      renderDashboard();
      openWorkerSecurityModal('passkeys');
    };
  }
}

async function openWorkerSecurityModal(initialTab = 'passkeys') {
  let modalWrap = document.querySelector('#worker-pw-modal-wrap');
  if (!modalWrap) {
    modalWrap = document.createElement('div');
    modalWrap.id = 'worker-pw-modal-wrap';
    document.body.appendChild(modalWrap);
  }

  let currentTab = initialTab; // 'passkeys' | 'password'
  let passkeys = [];
  let isLoadingPasskeys = true;
  let passkeyNotice = '';
  let passkeyError = '';

  const loadPasskeys = async () => {
    isLoadingPasskeys = true;
    try {
      const res = await apiRequest('/api/auth/passkey/list');
      passkeys = res.passkeys || [];
    } catch (e) {
      passkeys = [];
    } finally {
      isLoadingPasskeys = false;
      render();
    }
  };

  const close = () => { modalWrap.innerHTML = ''; };

  const render = () => {
    const isWebAuthnAvailable = browserSupportsWebAuthn();

    modalWrap.innerHTML = `
      <div class="day-modal-backdrop" id="worker-pw-modal-backdrop">
        <div class="day-modal-card" style="max-width:440px;animation:modalPop 0.18s cubic-bezier(0.16,1,0.3,1);overflow:hidden;">
          <div class="day-modal-header">
            <div>
              <div style="font-size:11px;font-weight:800;letter-spacing:0.8px;text-transform:uppercase;color:var(--blue-primary);margin-bottom:2px;">ACCOUNT SECURITY</div>
              <h3 style="margin:0;">Security &amp; Passkeys</h3>
            </div>
            <button type="button" class="day-modal-close-btn" id="btn-close-worker-pw-modal">&times;</button>
          </div>

          <!-- Tab Bar -->
          <div style="display:flex;border-bottom:1px solid #e2e8f0;padding:0 20px;background:#f8fafc;">
            <button type="button" id="tab-btn-passkeys" style="padding:11px 14px;border:none;background:none;font-weight:750;font-size:12.5px;cursor:pointer;color:${currentTab === 'passkeys' ? 'var(--blue-primary, #2563eb)' : '#64748b'};border-bottom:${currentTab === 'passkeys' ? '2.5px solid var(--blue-primary, #2563eb)' : '2.5px solid transparent'};margin-bottom:-1px;">
              Passkeys (Biometrics)
            </button>
            <button type="button" id="tab-btn-password" style="padding:11px 14px;border:none;background:none;font-weight:750;font-size:12.5px;cursor:pointer;color:${currentTab === 'password' ? 'var(--blue-primary, #2563eb)' : '#64748b'};border-bottom:${currentTab === 'password' ? '2.5px solid var(--blue-primary, #2563eb)' : '2.5px solid transparent'};margin-bottom:-1px;">
              Password
            </button>
          </div>

          <!-- Content Body -->
          <div style="padding:20px;">
            ${currentTab === 'passkeys' ? `
              ${passkeyNotice ? `<div style="padding:10px 12px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;color:#1d4ed8;font-size:12px;font-weight:600;margin-bottom:12px;">${escapeHtml(passkeyNotice)}</div>` : ''}
              ${passkeyError ? `<div style="padding:10px 12px;background:#fef2f2;border:1px solid #fecaca;border-radius:10px;color:#b91c1c;font-size:12px;font-weight:600;margin-bottom:12px;">${escapeHtml(passkeyError)}</div>` : ''}

              <div style="font-size:12px;color:#475569;line-height:1.5;margin-bottom:16px;">
                Passkeys let you sign in with one tap using <strong>Face ID</strong>, <strong>fingerprint</strong>, or device screen lock. They are stored securely in your device hardware and are 100% phishing-resistant.
              </div>

              ${isWebAuthnAvailable ? `
                <button type="button" class="btn-primary-action" id="btn-add-passkey" style="width:100%;height:44px;justify-content:center;gap:8px;font-size:13px;border-radius:12px;margin-bottom:18px;">
                  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 10a4 4 0 1 0 4 4"/><circle cx="12" cy="12" r="9"/><path d="m14 10 5-5"/><path d="m17 7 2 2"/></svg>
                  <span>Register this Device (Face ID / Passkey)</span>
                </button>
              ` : `
                <div style="padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;font-size:12px;color:#64748b;margin-bottom:16px;">
                  Your current browser does not support WebAuthn passkeys. Please use Safari (iOS/macOS), Chrome, or Edge to register a passkey.
                </div>
              `}

              <div style="font-size:11px;font-weight:800;color:#64748b;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:8px;">
                Enrolled Devices (${passkeys.length})
              </div>

              ${isLoadingPasskeys ? `
                <div style="padding:16px;text-align:center;font-size:12px;color:#94a3b8;">Loading passkeys…</div>
              ` : passkeys.length === 0 ? `
                <div style="padding:18px;border:1px dashed #cbd5e1;border-radius:12px;text-align:center;font-size:12px;color:#94a3b8;">
                  No passkeys registered yet. Tap above to enable 1-tap sign-in!
                </div>
              ` : `
                <div style="display:flex;flex-direction:column;gap:8px;">
                  ${passkeys.map(pk => `
                    <div style="display:flex;align-items:center;justify-content:space-between;padding:10px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;">
                      <div>
                        <div style="font-weight:750;font-size:13px;color:#0f172a;">${escapeHtml(pk.device_name || 'Passkey Device')}</div>
                        <div style="font-size:10.5px;color:#64748b;margin-top:2px;">Added ${new Date(pk.created_at).toLocaleDateString()}</div>
                      </div>
                      <button type="button" class="btn-delete-passkey" data-pk-id="${pk.id}" style="border:none;background:#fee2e2;color:#dc2626;border-radius:8px;padding:4px 9px;font-size:11px;font-weight:700;cursor:pointer;">
                        Remove
                      </button>
                    </div>
                  `).join('')}
                </div>
              `}
            ` : `
              <!-- Change Password Form -->
              <form id="worker-change-pw-form" style="display:flex;flex-direction:column;gap:14px;">
                <div id="worker-pw-error" style="display:none;padding:10px 12px;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;color:#b91c1c;font-size:12.5px;font-weight:600;line-height:1.4;"></div>
                <div class="action-field">
                  <label for="worker-cur-pw">Current Password</label>
                  <input id="worker-cur-pw" name="currentPassword" type="password" required autocomplete="current-password" placeholder="Enter current password">
                </div>
                <div class="action-field">
                  <label for="worker-new-pw">New Password</label>
                  <input id="worker-new-pw" name="newPassword" type="password" required minlength="10" autocomplete="new-password" placeholder="At least 10 characters (letters & numbers/symbols)">
                </div>
                <div class="action-field">
                  <label for="worker-conf-pw">Confirm New Password</label>
                  <input id="worker-conf-pw" name="confirmPassword" type="password" required minlength="10" autocomplete="new-password" placeholder="Repeat new password">
                </div>
                <div style="display:flex;gap:10px;margin-top:6px;">
                  <button type="button" class="btn-signout" id="btn-cancel-worker-pw" style="flex:1;height:42px;border-color:var(--border-card);color:var(--text-main);background:#fff;">Cancel</button>
                  <button type="submit" class="btn-primary-action" id="btn-save-worker-pw" style="flex:1;height:42px;">Save Password</button>
                </div>
              </form>
            `}
          </div>
        </div>
      </div>
    `;

    // Bind Close
    modalWrap.querySelector('#btn-close-worker-pw-modal').onclick = close;
    const cancelBtn = modalWrap.querySelector('#btn-cancel-worker-pw');
    if (cancelBtn) cancelBtn.onclick = close;
    modalWrap.querySelector('#worker-pw-modal-backdrop').onclick = (e) => {
      if (e.target.id === 'worker-pw-modal-backdrop') close();
    };

    // Bind Tab Switching
    modalWrap.querySelector('#tab-btn-passkeys').onclick = () => {
      currentTab = 'passkeys';
      passkeyNotice = '';
      passkeyError = '';
      render();
    };
    modalWrap.querySelector('#tab-btn-password').onclick = () => {
      currentTab = 'password';
      render();
    };

    // Bind Add Passkey
    const addPkBtn = modalWrap.querySelector('#btn-add-passkey');
    if (addPkBtn) {
      addPkBtn.onclick = async () => {
        passkeyNotice = '';
        passkeyError = '';
        addPkBtn.disabled = true;
        addPkBtn.innerHTML = '<span style="font-size:12.5px;">Registering biometric passkey…</span>';

        try {
          const chalRes = await apiRequest('/api/auth/passkey/register-challenge', { method: 'POST' });
          if (!chalRes.ok) throw new Error(chalRes.error || 'Failed to start registration.');

          const regResponse = await startRegistration({ optionsJSON: chalRes.options });

          let deviceName = 'My Device';
          const ua = navigator.userAgent;
          if (/iPhone/i.test(ua)) deviceName = 'iPhone (Face ID)';
          else if (/iPad/i.test(ua)) deviceName = 'iPad (Touch ID)';
          else if (/Macintosh/i.test(ua)) deviceName = 'Mac (Touch ID)';
          else if (/Android/i.test(ua)) deviceName = 'Android (Biometrics)';
          else if (/Windows/i.test(ua)) deviceName = 'Windows PC (Hello)';
          else deviceName = 'Device Passkey';

          const verifyRes = await apiRequest('/api/auth/passkey/register-verify', {
            method: 'POST',
            body: JSON.stringify({
              challengeId: chalRes.challengeId,
              response: regResponse,
              deviceName
            })
          });

          if (!verifyRes.ok) throw new Error(verifyRes.error || 'Verification failed.');

          passkeyNotice = `Passkey registered for ${verifyRes.deviceName}! You can now sign in with one tap.`;
          showWorkerToast('Passkey registered successfully!');
          await loadPasskeys();
        } catch (err) {
          if (err.name === 'NotAllowedError') {
            passkeyError = 'Registration was cancelled or timed out.';
          } else {
            passkeyError = err.message || 'Passkey registration failed.';
          }
          render();
        }
      };
    }

    // Bind Delete Passkey
    modalWrap.querySelectorAll('.btn-delete-passkey').forEach(btn => {
      btn.onclick = async () => {
        const pkId = btn.dataset.pkId;
        if (!confirm('Remove this passkey from your account?')) return;
        btn.disabled = true;
        btn.textContent = '…';
        try {
          await apiRequest('/api/auth/passkey/delete', {
            method: 'POST',
            body: JSON.stringify({ passkeyId: pkId })
          });
          showWorkerToast('Passkey removed.');
          await loadPasskeys();
        } catch (e) {
          alert(e.message || 'Could not remove passkey.');
          btn.disabled = false;
          btn.textContent = 'Remove';
        }
      };
    });

    // Bind Change Password Form
    const pwForm = modalWrap.querySelector('#worker-change-pw-form');
    if (pwForm) {
      const errorBox = modalWrap.querySelector('#worker-pw-error');
      const submitBtn = modalWrap.querySelector('#btn-save-worker-pw');

      pwForm.onsubmit = async (e) => {
        e.preventDefault();
        errorBox.style.display = 'none';
        const curPw = pwForm.currentPassword.value;
        const newPw = pwForm.newPassword.value;
        const confPw = pwForm.confirmPassword.value;

        if (newPw.length < 10) {
          errorBox.textContent = 'New password must be at least 10 characters.';
          errorBox.style.display = 'block';
          return;
        }
        if (!/[A-Za-z]/.test(newPw) || !/[0-9!@#$%^&*()_+\-=\[\]{};':"\\|,.<>\/?~`]/.test(newPw)) {
          errorBox.textContent = 'Password must include both letters and numbers/symbols.';
          errorBox.style.display = 'block';
          return;
        }
        if (newPw !== confPw) {
          errorBox.textContent = 'New passwords do not match.';
          errorBox.style.display = 'block';
          return;
        }

        submitBtn.disabled = true;
        submitBtn.textContent = 'Saving…';

        try {
          const res = await apiRequest('/api/worker/change-password', {
            method: 'POST',
            body: JSON.stringify({ currentPassword: curPw, newPassword: newPw })
          });
          if (res?.token) {
            token = res.token;
            sessionStorage.setItem('worker-token', res.token);
          }
          close();
          showWorkerToast('Password updated successfully!');
        } catch (err) {
          errorBox.textContent = err.message || 'Could not update password.';
          errorBox.style.display = 'block';
          submitBtn.disabled = false;
          submitBtn.textContent = 'Save Password';
        }
      };
    }
  };

  loadPasskeys();
}

async function boot() {
  if (!token) return renderLogin();
  try {
    const now = new Date();
    const nextMonthDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
    const selectedMonth = d => d.toISOString().slice(0, 7);

    const [me, currentRoster, nextRoster, unavailData] = await Promise.all([
      apiRequest('/api/me'),
      apiRequest(`/api/roster?month=${selectedMonth(now)}`),
      apiRequest(`/api/roster?month=${selectedMonth(nextMonthDate)}`),
      apiRequest('/api/unavailability').catch(() => ({ flags: [] })),
    ]);

    staffId = me.id;
    profile = { fullName: me.staff.full_name, email: me.email };
    workerRequests = unavailData.flags || [];

    staff = (currentRoster.staff || []).map(person => ({
      id: person.id, name: person.full_name, role: person.role_tag,
      podId: person.pod_id, effectiveFrom: person.effective_from,
    }));

    state = JSON.parse(sessionStorage.getItem('shift-roster') || 'null') || initialState();
    state.staffMonthly ||= {};
    for (const roster of [currentRoster, nextRoster]) {
      for (const assignment of roster.assignments || []) {
        state.assignments[`${assignment.shift_date}:${assignment.staff_id}`] = assignment.shift_code;
      }
      for (const assignment of roster.podAssignments || []) {
        state.staffMonthly[`${roster.month}:${assignment.staff_id}`] = {
          podId: assignment.pod_id,
          rowOrder: assignment.row_order,
        };
      }
    }
    sessionStorage.setItem('shift-roster', JSON.stringify(state));
    renderDashboard();
  } catch (error) {
    if (/sign in|required|session|invalid/i.test(error.message)) {
      sessionStorage.removeItem('worker-token');
      sessionStorage.removeItem('worker-id');
      sessionStorage.removeItem('worker-profile');
      sessionStorage.removeItem('shift-roster');
      token = '';
      staffId = '';
      profile = null;
      renderLogin();
    } else {
      app.innerHTML = `
        <div class="login-wrap">
          <div class="login-card" style="text-align:center;">
            <div class="login-brand-row" style="justify-content:center;">
              <img src="${shiftDeskMark}" alt="ShiftDex logo">
              <span>ShiftDex</span>
            </div>
            <h2 class="hero-title" style="font-size:20px;">Could not connect to roster</h2>
            <p class="hero-subtitle" style="margin-bottom:18px;">${escapeHtml(error.message)}</p>
            <button id="retry" class="btn-submit-action" style="width:100%;justify-content:center;" type="button">Try Again</button>
          </div>
        </div>
      `;
      document.querySelector('#retry').onclick = boot;
    }
  }
}

boot();
