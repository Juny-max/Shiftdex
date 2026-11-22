import shiftDeskMark from './shift-desk-mark.svg';

const esc = value => String(value || '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char]));

export function mountLogin(app, api, onLogin) {
  let mode = 'login'; // 'login' | 'change_password'
  let notice = '';
  let prefillEmail = '';
  let prefillTempPassword = '';

  const urlParams = new URLSearchParams(window.location.search);
  const verifyToken = urlParams.get('verify_email') || urlParams.get('verify_token');
  if (verifyToken && api) {
    fetch(`${api}/api/verify-email`, {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'accept': 'application/json' },
      body: JSON.stringify({ token: verifyToken })
    }).then(res => res.json()).then(data => {
      if (data.ok && data.email) {
        notice = `Email verified! Your account email is now ${data.email}. Please sign in.`;
        prefillEmail = data.email;
        window.history.replaceState({}, document.title, window.location.pathname);
        draw();
      }
    }).catch(() => {});
  }

  const draw = () => {
    if (mode === 'change_password') {
      app.innerHTML = `
        <div class="login-wrap">
          <div class="login-card">
            <div class="login-brand-row">
              <img src="${shiftDeskMark}" alt="Shift Desk logo">
              <span>shift desk<span style="color:var(--blue-primary);">.</span></span>
            </div>
            <div class="hero-eyebrow">FIRST-TIME SIGN IN</div>
            <h2 class="hero-title" style="font-size:22px;">Set permanent password</h2>
            <p class="hero-subtitle" style="margin-bottom:18px;">Enter your temporary credentials and choose a private password for your account.</p>
            ${notice ? `<div class="auth-notice" role="status">${esc(notice)}</div>` : ''}
            <form id="worker-change-password-form" style="display:flex;flex-direction:column;gap:13px;">
              <div class="action-field">
                <label for="auth-email">Account Email</label>
                <input id="auth-email" name="email" type="email" autocomplete="email" required maxlength="254" value="${esc(prefillEmail)}" readonly style="background:#f1f5f9;cursor:not-allowed;">
              </div>
              <div class="action-field">
                <label for="auth-temp-password">Temporary Password</label>
                <input id="auth-temp-password" name="tempPassword" type="password" autocomplete="current-password" required maxlength="256" value="${esc(prefillTempPassword)}" placeholder="From your invite email">
              </div>
              <div class="action-field">
                <label for="auth-new-password">New Password</label>
                <input id="auth-new-password" name="newPassword" type="password" autocomplete="new-password" required minlength="12" maxlength="256" placeholder="At least 12 characters">
              </div>
              <div class="action-field">
                <label for="auth-confirm-password">Confirm Password</label>
                <input id="auth-confirm-password" name="confirmPassword" type="password" autocomplete="new-password" required minlength="8" maxlength="256" placeholder="Repeat new password">
              </div>
              <button class="btn-primary-action" id="auth-submit" type="submit" style="width:100%;justify-content:center;height:44px;margin-top:6px;">Update Password &amp; Enter</button>
            </form>
            <p style="margin:16px 0 0;text-align:center;">
              <button id="auth-back" type="button" style="border:0;background:none;color:var(--blue-primary);font:inherit;font-weight:700;cursor:pointer;font-size:12.5px;">&larr; Back to sign in</button>
            </p>
          </div>
        </div>
      `;

      document.querySelector('#auth-back').onclick = () => {
        mode = 'login';
        notice = '';
        draw();
      };

      document.querySelector('#worker-change-password-form').onsubmit = async event => {
        event.preventDefault();
        const form = new FormData(event.currentTarget);
        const email = String(form.get('email') || '').trim();
        const tempPassword = String(form.get('tempPassword') || '');
        const newPassword = String(form.get('newPassword') || '');
        const confirmPassword = String(form.get('confirmPassword') || '');

        if (newPassword.length < 12) {
          notice = 'Your new password must be at least 12 characters.';
          draw();
          return;
        }
        if (newPassword !== confirmPassword) {
          notice = 'New passwords do not match.';
          draw();
          return;
        }

        const submit = document.querySelector('#auth-submit');
        submit.disabled = true;
        submit.textContent = 'Updating password…';

        try {
          const response = await fetch(`${api}/api/auth/change-password`, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ email, tempPassword, newPassword }),
          });
          const result = await response.json();
          if (!response.ok) throw new Error(result.error || 'Password update failed.');
          onLogin(result);
        } catch (error) {
          notice = error.message || 'Could not update password.';
          draw();
        }
      };
      return;
    }

    // Default: Clean Sign In
    app.innerHTML = `
      <div class="login-wrap">
        <div class="login-card">
          <div class="login-brand-row">
            <img src="${shiftDeskMark}" alt="Shift Desk logo">
            <span>shift desk<span style="color:var(--blue-primary);">.</span></span>
          </div>
          <div class="hero-eyebrow">WORKER PORTAL</div>
          <h2 class="hero-title" style="font-size:24px;">Sign in to your schedule</h2>
          <p class="hero-subtitle" style="margin-bottom:18px;">Enter your company email and password.</p>
          ${notice ? `<div class="auth-notice" role="status">${esc(notice)}</div>` : ''}
          <form id="worker-auth-form" style="display:flex;flex-direction:column;gap:14px;">
            <div class="action-field">
              <label for="auth-email">Company Email</label>
              <input id="auth-email" name="email" type="email" autocomplete="email" required maxlength="254" placeholder="you@company.com" value="${esc(prefillEmail)}">
            </div>
            <div class="action-field">
              <label for="auth-password">Password</label>
              <input id="auth-password" name="password" type="password" autocomplete="current-password" required maxlength="256" placeholder="Your password">
            </div>
            <button class="btn-primary-action" id="auth-submit" type="submit" style="width:100%;justify-content:center;height:44px;margin-top:4px;">Sign In</button>
          </form>
          <p style="margin:20px 0 0;font-size:12px;color:var(--text-light);text-align:center;line-height:1.45;">
            Your operations administrator provisions your account and emails your credentials.
          </p>
        </div>
      </div>
    `;

    document.querySelector('#worker-auth-form').onsubmit = async event => {
      event.preventDefault();
      if (!api) {
        notice = 'The portal is not connected yet. Please contact your roster administrator.';
        draw();
        return;
      }
      const form = new FormData(event.currentTarget);
      const email = String(form.get('email') || '').trim();
      const password = String(form.get('password') || '');
      const submit = document.querySelector('#auth-submit');
      submit.disabled = true;
      submit.textContent = 'Signing in…';

      try {
        const response = await fetch(`${api}/api/auth/login`, {
          method: 'POST',
          headers: { 'content-type': 'application/json' },
          body: JSON.stringify({ email, password }),
        });
        const result = await response.json();
        if (response.ok && result.mustChangePassword) {
          mode = 'change_password';
          prefillEmail = email;
          prefillTempPassword = password;
          notice = 'First time signing in: please set your permanent password to continue.';
          draw();
          return;
        }
        if (!response.ok) throw new Error(result.error || 'Invalid email or password.');
        onLogin(result);
      } catch (error) {
        notice = error.message || 'Could not sign in.';
        draw();
      }
    };
  };

  draw();
}
