import shiftDeskMark from './shift-desk-mark.svg';
import { startAuthentication, browserSupportsWebAuthn, browserSupportsWebAuthnAutofill } from '@simplewebauthn/browser';

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
              <img src="${shiftDeskMark}" alt="ShiftDex logo">
              <span>shiftdex<span style="color:var(--blue-primary);">.</span></span>
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

    // Default: Clean Sign In with Passkey Support
    const canUsePasskeys = browserSupportsWebAuthn();

    app.innerHTML = `
      <div class="login-wrap">
        <div class="login-card">
          <div class="login-brand-row">
            <img src="${shiftDeskMark}" alt="ShiftDex logo">
            <span>shiftdex<span style="color:var(--blue-primary);">.</span></span>
          </div>
          <div class="hero-eyebrow">WORKER PORTAL</div>
          <h2 class="hero-title" style="font-size:24px;">Sign in to your schedule</h2>
          <p class="hero-subtitle" style="margin-bottom:18px;">Sign in with your biometric passkey or company password.</p>
          ${notice ? `<div class="auth-notice" role="status">${esc(notice)}</div>` : ''}

          ${canUsePasskeys ? `
            <div style="margin-bottom:16px;">
              <button type="button" class="btn-passkey-hero" id="btn-passkey-login" style="width:100%;height:46px;display:flex;align-items:center;justify-content:center;gap:10px;border-radius:14px;border:1.5px solid var(--blue-primary, #2563eb);background:#eff6ff;color:#1d4ed8;font-weight:750;font-size:13.5px;cursor:pointer;transition:all 0.18s ease;box-shadow:0 2px 6px rgba(37,99,235,0.08);">
                <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><circle cx="9.2" cy="7.5" r="3.4" fill="currentColor" fill-opacity="0.24" stroke="currentColor" stroke-width="1.9"/><path d="M3 19.2c0-3.1 2.75-5.2 6.2-5.2 1.25 0 2.4.3 3.35.85" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/><circle cx="17.2" cy="14" r="2.6" fill="currentColor" fill-opacity="0.24" stroke="currentColor" stroke-width="1.9"/><path d="M17.2 16.6V21m0-2.1h2.1" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/></svg>
                <span>Sign in with Passkey (Face ID / Fingerprint)</span>
              </button>
              <div style="display:flex;align-items:center;margin:16px 0 4px;gap:10px;">
                <span style="flex:1;height:1px;background:#e2e8f0;"></span>
                <span style="font-size:11px;font-weight:700;color:#94a3b8;text-transform:uppercase;letter-spacing:0.04em;">or password</span>
                <span style="flex:1;height:1px;background:#e2e8f0;"></span>
              </div>
            </div>
          ` : ''}

          <form id="worker-auth-form" style="display:flex;flex-direction:column;gap:14px;">
            <div class="action-field">
              <label for="auth-email">Company Email</label>
              <input id="auth-email" name="email" type="email" autocomplete="username webauthn" required maxlength="254" placeholder="you@company.com" value="${esc(prefillEmail)}">
            </div>
            <div class="action-field">
              <label for="auth-password">Password</label>
              <input id="auth-password" name="password" type="password" autocomplete="current-password" required maxlength="256" placeholder="Your password">
            </div>
            <button class="btn-primary-action" id="auth-submit" type="submit" style="width:100%;justify-content:center;height:44px;margin-top:4px;">Sign In with Password</button>
          </form>
          <p style="margin:20px 0 0;font-size:12px;color:var(--text-light);text-align:center;line-height:1.45;">
            Your operations administrator provisions your account and emails your credentials.
          </p>
        </div>
      </div>
    `;

    // 1-Tap Passkey Authentication
    const passkeyBtn = document.querySelector('#btn-passkey-login');
    if (passkeyBtn) {
      passkeyBtn.onclick = async () => {
        if (!api) {
          notice = 'The portal is not connected yet. Please contact your roster administrator.';
          draw();
          return;
        }
        const emailInput = document.querySelector('#auth-email');
        const emailVal = emailInput ? String(emailInput.value || '').trim() : '';
        passkeyBtn.disabled = true;
        passkeyBtn.innerHTML = '<span style="font-size:13px;font-weight:700;color:#2563eb;">Verifying biometric passkey…</span>';

        try {
          const chalRes = await fetch(`${api}/api/auth/passkey/login-challenge`, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ email: emailVal })
          });
          const chalData = await chalRes.json();
          if (!chalRes.ok || !chalData.ok) {
            throw new Error(chalData.error || 'Could not initiate passkey sign in.');
          }

          const assertionResponse = await startAuthentication({ optionsJSON: chalData.options });

          const verifyRes = await fetch(`${api}/api/auth/passkey/login-verify`, {
            method: 'POST',
            headers: { 'content-type': 'application/json' },
            body: JSON.stringify({ challengeId: chalData.challengeId, response: assertionResponse })
          });
          const verifyData = await verifyRes.json();
          if (!verifyRes.ok || !verifyData.ok) {
            throw new Error(verifyData.error || 'Passkey verification failed.');
          }

          if (verifyData.passkeyId) {
            localStorage.setItem('shiftdex-device-passkey-id', verifyData.passkeyId);
          }
          onLogin(verifyData);
        } catch (err) {
          passkeyBtn.disabled = false;
          passkeyBtn.innerHTML = `
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none"><circle cx="9.2" cy="7.5" r="3.4" fill="currentColor" fill-opacity="0.24" stroke="currentColor" stroke-width="1.9"/><path d="M3 19.2c0-3.1 2.75-5.2 6.2-5.2 1.25 0 2.4.3 3.35.85" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"/><circle cx="17.2" cy="14" r="2.6" fill="currentColor" fill-opacity="0.24" stroke="currentColor" stroke-width="1.9"/><path d="M17.2 16.6V21m0-2.1h2.1" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"/></svg>
            <span>Sign in with Passkey (Face ID / Fingerprint)</span>
          `;
          if (err.name === 'NotAllowedError') {
            notice = 'Passkey sign in was cancelled or timed out.';
          } else {
            notice = err.message || 'Passkey authentication failed.';
          }
          draw();
        }
      };
    }

    // Passkey Autofill / Conditional UI
    if (mode === 'login' && canUsePasskeys) {
      (async () => {
        try {
          const isAutofillAvailable = await browserSupportsWebAuthnAutofill();
          if (isAutofillAvailable) {
            const chalRes = await fetch(`${api}/api/auth/passkey/login-challenge`, {
              method: 'POST',
              headers: { 'content-type': 'application/json' },
              body: JSON.stringify({})
            });
            const chalData = await chalRes.json();
            if (chalData.ok && mode === 'login') {
              const assertion = await startAuthentication({
                optionsJSON: chalData.options,
                useBrowserAutofill: true
              });
              if (assertion) {
                const verifyRes = await fetch(`${api}/api/auth/passkey/login-verify`, {
                  method: 'POST',
                  headers: { 'content-type': 'application/json' },
                  body: JSON.stringify({ challengeId: chalData.challengeId, response: assertion })
                });
                const verifyData = await verifyRes.json();
                if (verifyData.ok) {
                  if (verifyData.passkeyId) {
                    localStorage.setItem('shiftdex-device-passkey-id', verifyData.passkeyId);
                  }
                  onLogin(verifyData);
                }
              }
            }
          }
        } catch (e) {
          // Conditional autofill is passive; ignore when not used
        }
      })();
    }

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
