import { LOGO_PNG_BASE64 } from './logo-base64.js';

const cors = {
  'access-control-allow-origin': '*',
  'access-control-allow-methods': 'GET,POST,OPTIONS',
  'access-control-allow-headers': 'authorization,content-type',
};
const json = (data, status = 200) => new Response(JSON.stringify(data), {
  status,
  headers: { ...cors, 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', 'x-content-type-options': 'nosniff', 'referrer-policy': 'no-referrer', 'x-frame-options': 'DENY' },
});
const encoder = new TextEncoder();
// Cloudflare Workers rejects PBKDF2 iteration counts above 100,000.
// Compensate with long admin passwords, per-IP/email throttling, and short sessions.
const CURRENT_PASSWORD_ITERATIONS = 100000;
const MIN_PASSWORD_ITERATIONS = 100000;
const MAX_PASSWORD_ITERATIONS = 100000;
const hex = bytes => [...new Uint8Array(bytes)].map(byte => byte.toString(16).padStart(2, '0')).join('');
const tokenString = bytes => btoa(String.fromCharCode(...new Uint8Array(bytes))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/g, '');

async function hashText(value) {
  return hex(await crypto.subtle.digest('SHA-256', encoder.encode(value)));
}

function safePasswordIterations(iterations) {
  const count = Number(iterations);
  return Number.isInteger(count) ? Math.min(MAX_PASSWORD_ITERATIONS, Math.max(MIN_PASSWORD_ITERATIONS, count)) : CURRENT_PASSWORD_ITERATIONS;
}

async function passwordDigest(password, salt, iterations = CURRENT_PASSWORD_ITERATIONS) {
  const key = await crypto.subtle.importKey('raw', encoder.encode(password), 'PBKDF2', false, ['deriveBits']);
  const bits = await crypto.subtle.deriveBits({ name: 'PBKDF2', salt: Uint8Array.from(salt.match(/.{2}/g), byte => parseInt(byte, 16)), iterations: safePasswordIterations(iterations), hash: 'SHA-256' }, key, 256);
  return hex(bits);
}

function sameHash(left, right) {
  if (typeof left !== 'string' || typeof right !== 'string' || left.length !== right.length) return false;
  let mismatch = 0;
  for (let i = 0; i < left.length; i++) mismatch |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return mismatch === 0;
}

async function audit(db, actor, action, type, id, details = {}) {
  await db.prepare('INSERT INTO audit_logs(id,actor_id,action,entity_type,entity_id,details_json) VALUES(?,?,?,?,?,?)')
    .bind(crypto.randomUUID(), actor, action, type, id, JSON.stringify(details)).run();
}

async function isAuthBlocked(db, scope, email, ipHash, now) {
  const row = await db.prepare('SELECT blocked_until FROM worker_auth_limits WHERE scope=? AND email=? AND ip_hash=?')
    .bind(scope, email, ipHash).first();
  return Boolean(row?.blocked_until && row.blocked_until > now.toISOString());
}

async function recordAuthFailure(db, scope, email, ipHash, now) {
  const row = await db.prepare('SELECT failures,window_started_at FROM worker_auth_limits WHERE scope=? AND email=? AND ip_hash=?')
    .bind(scope, email, ipHash).first();
  const active = row && Date.parse(row.window_started_at) > now.getTime() - 15 * 60 * 1000;
  const failures = active ? Number(row.failures) + 1 : 1;
  const started = active ? row.window_started_at : now.toISOString();
  const blocked = failures >= 8 ? new Date(now.getTime() + 15 * 60 * 1000).toISOString() : null;
  await db.prepare(`INSERT INTO worker_auth_limits(scope,email,ip_hash,failures,window_started_at,blocked_until) VALUES(?,?,?,?,?,?)
    ON CONFLICT(scope,email,ip_hash) DO UPDATE SET failures=excluded.failures,window_started_at=excluded.window_started_at,blocked_until=excluded.blocked_until`)
    .bind(scope, email, ipHash, failures, started, blocked).run();
}

async function clearAuthFailures(db, scope, email, ipHash) {
  await db.prepare('DELETE FROM worker_auth_limits WHERE scope=? AND email=? AND ip_hash=?').bind(scope, email, ipHash).run();
}

async function auth(req, env, db, admin = false) {
  const token = (req.headers.get('authorization') || '').replace(/^Bearer\s+/i, '');
  if (!token) return null;
  const tokenHash = await hashText(token);
  if (admin) {
    const account = await db.prepare(`SELECT a.id,a.email,a.full_name FROM admin_sessions s JOIN admin_accounts a ON a.id=s.admin_id
      WHERE s.token_hash=? AND s.expires_at>? AND a.status='active'`)
      .bind(tokenHash, new Date().toISOString()).first();
    return account ? { id: account.id, email: account.email, name: account.full_name, admin: true } : null;
  }
  const account = await db.prepare(`SELECT a.id account_id,a.staff_id,a.email,a.full_name,s.role_tag,s.pod_id,s.effective_from
    FROM worker_sessions ws JOIN worker_accounts a ON a.id=ws.account_id
    JOIN staff s ON s.id=a.staff_id
    WHERE ws.token_hash=? AND ws.expires_at>? AND a.status='active' AND s.status='active'`)
    .bind(tokenHash, new Date().toISOString()).first();
  return account ? { id: account.staff_id, accountId: account.account_id, email: account.email, name: account.full_name, admin: false } : null;
}

async function createSession(db, accountId) {
  const token = tokenString(crypto.getRandomValues(new Uint8Array(32)));
  const tokenHash = await hashText(token);
  const expires = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
  await db.prepare('INSERT INTO worker_sessions(token_hash,account_id,expires_at) VALUES(?,?,?)').bind(tokenHash, accountId, expires).run();
  return { token, expiresAt: expires };
}

async function createAdminSession(db, adminId) {
  const token = tokenString(crypto.getRandomValues(new Uint8Array(32)));
  const tokenHash = await hashText(token);
  const expires = new Date(Date.now() + 12 * 60 * 60 * 1000).toISOString();
  await db.prepare('INSERT INTO admin_sessions(token_hash,admin_id,expires_at) VALUES(?,?,?)').bind(tokenHash, adminId, expires).run();
  return { token, expiresAt: expires };
}

async function sendRawBrevo(env, { toEmail, toName, subject, htmlContent, textContent }) {
  if (!env.BREVO_API_KEY) {
    console.warn('BREVO_API_KEY is not configured in environment.');
    return { success: false, error: 'BREVO_API_KEY is missing' };
  }
  const key = String(env.BREVO_API_KEY || '').trim();
  if (key.startsWith('xsmtpsib-')) {
    console.warn('BREVO_API_KEY appears to be an SMTP key, not a REST API key.');
    return {
      success: false,
      error: "Invalid Brevo Key: You entered an SMTP key ('xsmtpsib-...'). Brevo's REST API requires an API key starting with 'xkeysib-'. Please create an API Key under Brevo Settings > SMTP & API > API Keys tab."
    };
  }
  const senderEmail = env.BREVO_SENDER_EMAIL;
  if (!senderEmail) return { success: false, error: 'BREVO_SENDER_EMAIL is not configured' };
  const senderName = env.BREVO_SENDER_NAME || 'Shift Desk';

  try {
    const res = await fetch('https://api.brevo.com/v3/smtp/email', {
      method: 'POST',
      headers: {
        'api-key': env.BREVO_API_KEY,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
      },
      body: JSON.stringify({
        sender: { name: senderName, email: senderEmail },
        to: [{ email: toEmail, name: toName || toEmail }],
        subject,
        htmlContent,
        textContent,
      }),
    });
    if (res.ok) {
      const data = await res.json().catch(() => ({}));
      return { success: true, messageId: data.messageId };
    }
    const errText = await res.text();
    console.error('Brevo error:', res.status, errText);
    return { success: false, status: res.status, error: errText };
  } catch (error) {
    console.error('Brevo delivery failed:', error.message);
    return { success: false, error: error.message };
  }
}

async function sendBrevoEmail(env, { email, fullName, tempPassword }) {
  const portalUrl = env.PORTAL_URL || 'https://shiftdex-workers.pages.dev';
  const logoUrl = 'https://shiftdex-workers.pages.dev/shift-desk-mark.png';
  const esc = s => String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Welcome to Shift Desk</title>
</head>
<body style="margin: 0; padding: 32px 16px; background-color: #f6f8fb; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; color: #111827;">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center">
        <!-- Main Card Container (Apple / OpenAI style) -->
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width: 540px; background-color: #ffffff; border-radius: 18px; border: 1px solid #e5e7eb; box-shadow: 0 8px 30px rgba(0, 0, 0, 0.04); overflow: hidden;">
          
          <!-- Brand Header with Official System Logo -->
          <tr>
            <td style="padding: 36px 36px 20px;">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="vertical-align: middle; padding-right: 14px;">
                    <img src="${logoUrl}" width="44" height="44" alt="Shift Desk" style="display: block; width: 44px; height: 44px; border-radius: 11px; border: 1px solid rgba(0, 0, 0, 0.06); object-fit: contain;">
                  </td>
                  <td style="vertical-align: middle;">
                    <div style="font-size: 20px; font-weight: 800; letter-spacing: -0.5px; color: #0f172a; line-height: 1.1;">
                      shift desk<span style="color: #2563eb;">.</span>
                    </div>
                    <div style="font-size: 11px; font-weight: 700; letter-spacing: 0.8px; color: #2563eb; text-transform: uppercase; margin-top: 3px;">
                      24/7 Operations Hub
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Main Content Body -->
          <tr>
            <td style="padding: 0 36px 36px;">
              <h1 style="margin: 0 0 12px; font-size: 23px; font-weight: 700; letter-spacing: -0.5px; color: #0f172a; line-height: 1.3;">
                Welcome to the team, ${esc(fullName)}
              </h1>
              <p style="margin: 0 0 24px; font-size: 15px; line-height: 1.6; color: #4b5563;">
                An operations administrator has added you to the company shift roster on <strong>Shift Desk</strong>. Your account has been provisioned and is ready for activation below.
              </p>

              <!-- Credentials Card (OpenAI / Google Security style) -->
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 14px; margin-bottom: 26px;">
                <tr>
                  <td style="padding: 22px 24px;">
                    <div style="margin-bottom: 14px;">
                      <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 4px;">Account Email</div>
                      <div style="font-size: 15px; font-weight: 600; color: #0f172a; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;">${esc(email)}</div>
                    </div>

                    <div style="height: 1px; background-color: #e2e8f0; margin: 0 0 16px;"></div>

                    <div style="margin-bottom: 14px;">
                      <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px;">Temporary Password</div>
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td style="background-color: #ffffff; border: 1.5px solid #cbd5e1; border-radius: 8px; padding: 8px 16px; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);">
                            <span style="font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 19px; font-weight: 700; letter-spacing: 1.8px; color: #0f172a;">${esc(tempPassword)}</span>
                          </td>
                        </tr>
                      </table>
                    </div>

                    <div style="font-size: 12px; color: #64748b; line-height: 1.5; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 14px;">
                      <strong style="color: #334155;">Account Security:</strong> You will be prompted to set your private permanent password immediately upon your first sign in.
                    </div>
                  </td>
                </tr>
              </table>

              <!-- Apple / ChatGPT Style CTA Button -->
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom: 26px;">
                <tr>
                  <td align="center">
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="center" style="border-radius: 10px; background-color: #2563eb; box-shadow: 0 4px 14px rgba(37, 99, 235, 0.28);">
                          <a href="${portalUrl}" target="_blank" style="display: inline-block; padding: 14px 34px; font-size: 15px; font-weight: 600; color: #ffffff; text-decoration: none; letter-spacing: -0.2px; border-radius: 10px;">
                            Activate Account &amp; Set Password &rarr;
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <!-- 3-Step Setup Guide (Google Onboarding Style) -->
              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f8fafc; border: 1px solid #edf2f7; border-radius: 12px; margin-bottom: 24px;">
                <tr>
                  <td style="padding: 18px 20px;">
                    <div style="font-size: 11px; font-weight: 700; color: #475569; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 12px;">Quick 3-Step Setup</div>
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" style="font-size: 13px; color: #475569; line-height: 1.5;">
                      <tr>
                        <td width="24" style="vertical-align: top; font-weight: 700; color: #2563eb;">1.</td>
                        <td style="padding-bottom: 8px;">Click <strong>Activate Account</strong> above to open the worker portal.</td>
                      </tr>
                      <tr>
                        <td width="24" style="vertical-align: top; font-weight: 700; color: #2563eb;">2.</td>
                        <td style="padding-bottom: 8px;">Enter your login email and temporary password from this email.</td>
                      </tr>
                      <tr>
                        <td width="24" style="vertical-align: top; font-weight: 700; color: #2563eb;">3.</td>
                        <td>Create your new permanent password to access your live 24/7 schedule.</td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <p style="font-size: 12px; color: #94a3b8; margin: 0 0 24px; line-height: 1.5; word-break: break-all;">
                Direct portal link: <a href="${portalUrl}" style="color: #2563eb; text-decoration: underline;">${portalUrl}</a>
              </p>

              <!-- Footer (Apple Style) -->
              <div style="border-top: 1px solid #f1f5f9; padding-top: 20px; font-size: 12px; color: #94a3b8; line-height: 1.6; text-align: center;">
                <strong>Shift Desk</strong> · 24/7 Operations Hub<br>
                Automated operational notification sent to ${esc(email)}.<br>
                If you did not expect this invitation, please contact your shift supervisor.
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  const textContent = `Welcome to Shift Desk, ${fullName}!\n\nAn operations administrator has added you to the company shift roster.\n\nPortal: ${portalUrl}\nLogin Email: ${email}\nTemporary Password: ${tempPassword}\n\nPlease visit the portal to set your new permanent password and access your schedule.\n\nShift Desk Operations Hub`;

  return sendRawBrevo(env, {
    toEmail: email,
    toName: fullName,
    subject: 'Welcome to Shift Desk — Your Account Details',
    htmlContent,
    textContent
  });
}

async function sendPasswordResetEmail(env, { email, fullName, tempPassword }) {
  const portalUrl = env.PORTAL_URL || 'https://shiftdex-workers.pages.dev';
  const logoUrl = 'https://shiftdex-workers.pages.dev/shift-desk-mark.png';
  const esc = s => String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Shift Desk Password Reset</title>
</head>
<body style="margin: 0; padding: 32px 16px; background-color: #f6f8fb; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; color: #111827;">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width: 540px; background-color: #ffffff; border-radius: 18px; border: 1px solid #e5e7eb; box-shadow: 0 8px 30px rgba(0, 0, 0, 0.04); overflow: hidden;">
          <tr>
            <td style="padding: 36px 36px 20px;">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="vertical-align: middle; padding-right: 14px;">
                    <img src="${logoUrl}" width="44" height="44" alt="Shift Desk" style="display: block; width: 44px; height: 44px; border-radius: 11px; border: 1px solid rgba(0, 0, 0, 0.06); object-fit: contain;">
                  </td>
                  <td style="vertical-align: middle;">
                    <div style="font-size: 20px; font-weight: 800; letter-spacing: -0.5px; color: #0f172a; line-height: 1.1;">
                      shift desk<span style="color: #2563eb;">.</span>
                    </div>
                    <div style="font-size: 11px; font-weight: 700; letter-spacing: 0.8px; color: #2563eb; text-transform: uppercase; margin-top: 3px;">
                      24/7 Operations Hub
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="padding: 0 36px 36px;">
              <h1 style="margin: 0 0 12px; font-size: 23px; font-weight: 700; letter-spacing: -0.5px; color: #0f172a; line-height: 1.3;">
                Password Reset for ${esc(fullName)}
              </h1>
              <p style="margin: 0 0 24px; font-size: 15px; line-height: 1.6; color: #4b5563;">
                An operations administrator has reset the password for your <strong>Shift Desk</strong> account. Use the temporary credentials below to sign in.
              </p>

              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 14px; margin-bottom: 26px;">
                <tr>
                  <td style="padding: 22px 24px;">
                    <div style="margin-bottom: 14px;">
                      <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 4px;">Account Email</div>
                      <div style="font-size: 15px; font-weight: 600; color: #0f172a; font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace;">${esc(email)}</div>
                    </div>

                    <div style="height: 1px; background-color: #e2e8f0; margin: 0 0 16px;"></div>

                    <div style="margin-bottom: 14px;">
                      <div style="font-size: 11px; font-weight: 700; color: #64748b; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 6px;">New Temporary Password</div>
                      <table border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td style="background-color: #ffffff; border: 1.5px solid #cbd5e1; border-radius: 8px; padding: 8px 16px; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.04);">
                            <span style="font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 19px; font-weight: 700; letter-spacing: 1.8px; color: #0f172a;">${esc(tempPassword)}</span>
                          </td>
                        </tr>
                      </table>
                    </div>

                    <div style="font-size: 12px; color: #64748b; line-height: 1.5; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 10px 14px;">
                      <strong style="color: #334155;">Account Security:</strong> You will be required to set a new private permanent password immediately upon signing in.
                    </div>
                  </td>
                </tr>
              </table>

              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom: 26px;">
                <tr>
                  <td align="center">
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="center" style="border-radius: 10px; background-color: #2563eb; box-shadow: 0 4px 14px rgba(37, 99, 235, 0.28);">
                          <a href="${portalUrl}" target="_blank" style="display: inline-block; padding: 14px 34px; font-size: 15px; font-weight: 600; color: #ffffff; text-decoration: none; letter-spacing: -0.2px; border-radius: 10px;">
                            Sign In to Shift Desk &rarr;
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <div style="border-top: 1px solid #f1f5f9; padding-top: 20px; font-size: 12px; color: #94a3b8; line-height: 1.6; text-align: center;">
                <strong>Shift Desk</strong> · 24/7 Operations Hub<br>
                Automated security notification sent to ${esc(email)}.<br>
                If you did not expect this password reset, please contact your shift supervisor immediately.
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  const textContent = `Shift Desk Password Reset\n\nHello ${fullName},\n\nAn operations administrator has reset your password for Shift Desk.\n\nPortal: ${portalUrl}\nAccount Email: ${email}\nTemporary Password: ${tempPassword}\n\nPlease sign in with this temporary password to create your new permanent password.\n\nShift Desk Operations Hub`;

  return sendRawBrevo(env, {
    toEmail: email,
    toName: fullName,
    subject: 'Shift Desk — Your Password Has Been Reset',
    htmlContent,
    textContent
  });
}

async function sendEmailVerificationEmail(env, { newEmail, currentEmail, fullName, verifyUrl }) {
  const logoUrl = 'https://shiftdex-workers.pages.dev/shift-desk-mark.png';
  const esc = s => String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const htmlContent = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Confirm Your New Email Address</title>
</head>
<body style="margin: 0; padding: 32px 16px; background-color: #f6f8fb; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; -webkit-font-smoothing: antialiased; -moz-osx-font-smoothing: grayscale; color: #111827;">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center">
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="max-width: 540px; background-color: #ffffff; border-radius: 18px; border: 1px solid #e5e7eb; box-shadow: 0 8px 30px rgba(0, 0, 0, 0.04); overflow: hidden;">
          <tr>
            <td style="padding: 36px 36px 20px;">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td style="vertical-align: middle; padding-right: 14px;">
                    <img src="${logoUrl}" width="44" height="44" alt="Shift Desk" style="display: block; width: 44px; height: 44px; border-radius: 11px; border: 1px solid rgba(0, 0, 0, 0.06); object-fit: contain;">
                  </td>
                  <td style="vertical-align: middle;">
                    <div style="font-size: 20px; font-weight: 800; letter-spacing: -0.5px; color: #0f172a; line-height: 1.1;">
                      shift desk<span style="color: #2563eb;">.</span>
                    </div>
                    <div style="font-size: 11px; font-weight: 700; letter-spacing: 0.8px; color: #2563eb; text-transform: uppercase; margin-top: 3px;">
                      24/7 Operations Hub
                    </div>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="padding: 0 36px 36px;">
              <h1 style="margin: 0 0 12px; font-size: 23px; font-weight: 700; letter-spacing: -0.5px; color: #0f172a; line-height: 1.3;">
                Confirm Your New Email Address
              </h1>
              <p style="margin: 0 0 20px; font-size: 15px; line-height: 1.6; color: #4b5563;">
                Hello <strong>${esc(fullName)}</strong>, an operations administrator has requested to update your Shift Desk login email address to:
              </p>

              <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 12px; padding: 16px 20px; margin-bottom: 24px;">
                <div style="font-size: 11px; font-weight: 700; color: #166534; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 4px;">Requested New Email</div>
                <div style="font-size: 16px; font-weight: 700; color: #15803d; font-family: ui-monospace, SFMono-Regular, Menlo, monospace;">${esc(newEmail)}</div>
              </div>

              <p style="margin: 0 0 24px; font-size: 14px; line-height: 1.6; color: #64748b;">
                To activate this email address for your Shift Desk account and start using it to sign in, please confirm your email address by clicking the button below:
              </p>

              <table width="100%" border="0" cellspacing="0" cellpadding="0" style="margin-bottom: 26px;">
                <tr>
                  <td align="center">
                    <table border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td align="center" style="border-radius: 10px; background-color: #2563eb; box-shadow: 0 4px 14px rgba(37, 99, 235, 0.28);">
                          <a href="${verifyUrl}" target="_blank" style="display: inline-block; padding: 14px 34px; font-size: 15px; font-weight: 600; color: #ffffff; text-decoration: none; letter-spacing: -0.2px; border-radius: 10px;">
                            Confirm &amp; Verify Email Address &rarr;
                          </a>
                        </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>

              <div style="font-size: 12px; color: #64748b; line-height: 1.6; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px 16px; margin-bottom: 20px;">
                <strong style="color: #334155;">Security Notice:</strong> Until you verify this new address, your current email (<strong>${esc(currentEmail)}</strong>) remains active for signing in. This link is valid for 48 hours. If you did not request this update, please inform your operations administrator.
              </div>

              <div style="border-top: 1px solid #f1f5f9; padding-top: 20px; font-size: 12px; color: #94a3b8; line-height: 1.6; text-align: center;">
                <strong>Shift Desk</strong> · 24/7 Operations Hub<br>
                Automated security verification sent to ${esc(newEmail)}.<br>
                Direct link: <a href="${verifyUrl}" style="color: #2563eb; word-break: break-all;">${verifyUrl}</a>
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
</body>
</html>`;

  const textContent = `Confirm Your New Email Address — Shift Desk\n\nHello ${fullName},\n\nAn operations administrator has requested to update your Shift Desk email to: ${newEmail}.\n\nUntil you verify this new email, your current email (${currentEmail}) remains active for signing in.\n\nTo verify, open this link:\n${verifyUrl}\n\nThis link is valid for 48 hours.\n\nShift Desk Operations Hub`;

  return sendRawBrevo(env, {
    toEmail: newEmail,
    toName: fullName,
    subject: 'Confirm Your New Email Address — Shift Desk',
    htmlContent,
    textContent
  });
}

function renderVerificationHtml({ success, title, message, portalUrl, logoUrl, newEmail }) {
  const esc = s => String(s || '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${esc(title)} — Shift Desk</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background-color: #f6f8fb;
      font-family: -apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      color: #0f172a;
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 24px;
    }
    .card {
      background: #ffffff;
      border: 1px solid #e5e7eb;
      border-radius: 20px;
      box-shadow: 0 10px 30px rgba(0, 0, 0, 0.04), 0 1px 3px rgba(0, 0, 0, 0.02);
      max-width: 480px;
      width: 100%;
      padding: 40px 36px;
      text-align: center;
    }
    .logo-row {
      display: inline-flex;
      align-items: center;
      gap: 10px;
      margin-bottom: 28px;
    }
    .logo-img {
      width: 38px;
      height: 38px;
      border-radius: 9px;
      border: 1px solid rgba(0,0,0,0.06);
    }
    .logo-text {
      font-size: 19px;
      font-weight: 800;
      letter-spacing: -0.5px;
      color: #0f172a;
    }
    .logo-dot { color: #2563eb; }
    .badge {
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 56px;
      height: 56px;
      border-radius: 50%;
      margin: 0 auto 20px;
    }
    .badge-success {
      background: #ecfdf5;
      color: #059669;
      border: 1px solid #a7f3d0;
    }
    .badge-error {
      background: #fef2f2;
      color: #dc2626;
      border: 1px solid #fecaca;
    }
    .title {
      font-size: 22px;
      font-weight: 700;
      letter-spacing: -0.4px;
      color: #0f172a;
      margin-bottom: 12px;
    }
    .desc {
      font-size: 14.5px;
      line-height: 1.6;
      color: #4b5563;
      margin-bottom: 28px;
    }
    .email-box {
      background: #f8fafc;
      border: 1px solid #e2e8f0;
      border-radius: 12px;
      padding: 14px 18px;
      font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
      font-size: 15px;
      font-weight: 600;
      color: #0f172a;
      margin-bottom: 28px;
      word-break: break-all;
    }
    .cta-btn {
      display: inline-block;
      width: 100%;
      padding: 14px 20px;
      background: #2563eb;
      color: #ffffff;
      font-size: 15px;
      font-weight: 600;
      text-decoration: none;
      border-radius: 12px;
      box-shadow: 0 4px 14px rgba(37, 99, 235, 0.28);
      transition: background 0.15s ease;
    }
    .cta-btn:hover {
      background: #1d4ed8;
    }
    .footer-note {
      font-size: 12px;
      color: #94a3b8;
      margin-top: 24px;
      line-height: 1.5;
    }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo-row">
      <img src="${logoUrl}" alt="Shift Desk" class="logo-img" />
      <div class="logo-text">shift desk<span class="logo-dot">.</span></div>
    </div>

    <div class="badge ${success ? 'badge-success' : 'badge-error'}">
      ${success
        ? '<svg width="28" height="28" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><path stroke-linecap="round" stroke-linejoin="round" d="M5 13l4 4L19 7"/></svg>'
        : '<svg width="28" height="28" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2.5"><path stroke-linecap="round" stroke-linejoin="round" d="M6 18L18 6M6 6l12 12"/></svg>'
      }
    </div>

    <h1 class="title">${esc(title)}</h1>
    <div class="desc">${message}</div>

    ${newEmail ? `<div class="email-box">${esc(newEmail)}</div>` : ''}

    <a href="${portalUrl}" class="cta-btn">Proceed to Shift Desk Portal &rarr;</a>

    <div class="footer-note">
      Shift Desk · 24/7 Operations Hub<br>
      Automated account security verification
    </div>
  </div>
</body>
</html>`;
}

export default {
  async fetch(req, env) {
    const origin = req.headers.get('origin');
    const allowedOrigin = !origin || origin === 'null' || origin === 'https://shiftdex-workers.pages.dev' || /^https?:\/\/(localhost|127\.0\.0\.1):(5173|5174|5175|5176|5177|5178|5179|5180|5181|5182|5183)$/.test(origin);
    if (!allowedOrigin) return json({ error: 'Origin not allowed.' }, 403);
    if (req.method === 'OPTIONS') return new Response(null, { headers: cors });
    const contentLength = Number(req.headers.get('content-length') || 0);
    if (contentLength > 65536) return json({ error: 'Request is too large.' }, 413);
    const url = new URL(req.url), db = env.DB;
    try {
      if (url.pathname === '/api/health') return json({ ok: true, service: 'shift-roster-api' });

      if ((url.pathname === '/brand-logo.png' || url.pathname === '/shift-desk-mark.png' || url.pathname === '/logo.png') && req.method === 'GET') {
        const bin = Uint8Array.from(atob(LOGO_PNG_BASE64), c => c.charCodeAt(0));
        return new Response(bin, {
          headers: {
            'content-type': 'image/png',
            'cache-control': 'public, max-age=31536000, immutable',
            'access-control-allow-origin': '*',
          }
        });
      }

      if (url.pathname === '/api/admin/auth/login' && req.method === 'POST') {
        const body = await req.json();
        const email = String(body.email || '').trim().toLowerCase();
        const password = String(body.password || '');
        if (email.length > 254 || password.length > 256) return json({ error: 'Email or password is incorrect.' }, 401);
        const ipHash = await hashText(req.headers.get('cf-connecting-ip') || 'unknown');
        const now = new Date();
        const nowIso = now.toISOString();
        const limit = await db.prepare('SELECT failures,window_started_at,blocked_until FROM admin_login_limits WHERE email=? AND ip_hash=?').bind(email, ipHash).first();
        if (limit?.blocked_until && limit.blocked_until > nowIso) return json({ error: 'Too many sign-in attempts. Try again in 15 minutes.' }, 429);
        const windowActive = limit && Date.parse(limit.window_started_at) > now.getTime() - 15 * 60 * 1000;
        const account = await db.prepare("SELECT id,email,full_name,password_salt,password_hash,password_iterations FROM admin_accounts WHERE email=? AND status='active'").bind(email).first();
        if (account && Number(account.password_iterations) > MAX_PASSWORD_ITERATIONS) return json({ error: 'This administrator password must be reset by the Cloudflare account owner.' }, 401);
        const passwordIterations = safePasswordIterations(account?.password_iterations);
        const digest = await passwordDigest(password || 'invalid-password', account?.password_salt || '00000000000000000000000000000000', passwordIterations);
        if (!account || !password || !sameHash(digest, account.password_hash)) {
          const failures = windowActive ? Number(limit.failures) + 1 : 1;
          const windowStartedAt = windowActive ? limit.window_started_at : nowIso;
          const blockedUntil = failures >= 8 ? new Date(now.getTime() + 15 * 60 * 1000).toISOString() : null;
          await db.prepare(`INSERT INTO admin_login_limits(email,ip_hash,failures,window_started_at,blocked_until) VALUES(?,?,?,?,?)
            ON CONFLICT(email,ip_hash) DO UPDATE SET failures=excluded.failures,window_started_at=excluded.window_started_at,blocked_until=excluded.blocked_until`)
            .bind(email, ipHash, failures, windowStartedAt, blockedUntil).run();
          return json({ error: 'Email or password is incorrect.' }, 401);
        }
        await db.prepare('DELETE FROM admin_login_limits WHERE email=? AND ip_hash=?').bind(email, ipHash).run();
        if (passwordIterations < CURRENT_PASSWORD_ITERATIONS) {
          const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
          const upgradedHash = await passwordDigest(password, salt, CURRENT_PASSWORD_ITERATIONS);
          await db.prepare('UPDATE admin_accounts SET password_salt=?,password_hash=?,password_iterations=?,updated_at=CURRENT_TIMESTAMP WHERE id=?')
            .bind(salt, upgradedHash, CURRENT_PASSWORD_ITERATIONS, account.id).run();
        }
        const session = await createAdminSession(db, account.id);
        await audit(db, account.id, 'login', 'admin_account', account.id);
        return json({ token: session.token, expiresAt: session.expiresAt, email: account.email, fullName: account.full_name });
      }

      if (url.pathname === '/api/admin/auth/logout' && req.method === 'POST') {
        const token = (req.headers.get('authorization') || '').replace(/^Bearer\s+/i, '');
        if (token) await db.prepare('DELETE FROM admin_sessions WHERE token_hash=?').bind(await hashText(token)).run();
        return json({ ok: true });
      }

      if (url.pathname === '/api/auth/register' && req.method === 'POST') {
        return json({ error: 'Public registration is disabled. An operations administrator will provision your account and email you your temporary credentials.' }, 403);
      }

      if (url.pathname === '/api/auth/login' && req.method === 'POST') {
        const body = await req.json();
        const email = String(body.email || '').trim().toLowerCase();
        const password = String(body.password || '');
        if (email.length > 254 || password.length > 256) return json({ error: 'Email or password is incorrect.' }, 401);
        const ipHash = await hashText(req.headers.get('cf-connecting-ip') || 'unknown');
        const now = new Date();
        if (await isAuthBlocked(db, 'login', email, ipHash, now)) return json({ error: 'Too many sign-in attempts. Try again in 15 minutes.' }, 429);
        const account = await db.prepare('SELECT id,email,full_name,password_salt,password_hash,password_iterations,staff_id,status,must_change_password FROM worker_accounts WHERE email=?').bind(email).first();
        if (account && Number(account.password_iterations) > MAX_PASSWORD_ITERATIONS) return json({ error: 'This password needs to be reissued by your roster administrator.' }, 401);
        const iterations = safePasswordIterations(account?.password_iterations);
        const digest = await passwordDigest(password || 'invalid-password', account?.password_salt || '00000000000000000000000000000000', iterations);
        if (!account || !password || !sameHash(digest, account.password_hash)) {
          await recordAuthFailure(db, 'login', email, ipHash, now);
          return json({ error: 'Email or password is incorrect.' }, 401);
        }
        await clearAuthFailures(db, 'login', email, ipHash);
        if (account.status === 'pending') return json({ error: 'Your account is waiting for administrator approval.' }, 403);
        if (account.status !== 'active' || !account.staff_id) return json({ error: 'This account is not active. Contact your administrator.' }, 403);
        const staffMember = await db.prepare("SELECT id,full_name,role_tag,pod_id,effective_from,status FROM staff WHERE id=? AND status IN ('active','pending')").bind(account.staff_id).first();
        if (!staffMember) return json({ error: 'Your linked staff profile is inactive. Contact your administrator.' }, 403);
        if (account.must_change_password) {
          return json({ mustChangePassword: true, email: account.email, fullName: account.full_name });
        }
        if (iterations < CURRENT_PASSWORD_ITERATIONS) {
          const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
          const upgradedHash = await passwordDigest(password, salt, CURRENT_PASSWORD_ITERATIONS);
          await db.prepare('UPDATE worker_accounts SET password_salt=?,password_hash=?,password_iterations=?,updated_at=CURRENT_TIMESTAMP WHERE id=?')
            .bind(salt, upgradedHash, CURRENT_PASSWORD_ITERATIONS, account.id).run();
        }
        const session = await createSession(db, account.id);
        await audit(db, account.id, 'login', 'worker_account', account.id);
        return json({ token: session.token, expiresAt: session.expiresAt, staffId: account.staff_id, fullName: account.full_name, email: account.email, staff: staffMember });
      }

      if (url.pathname === '/api/auth/change-password' && req.method === 'POST') {
        const body = await req.json();
        const email = String(body.email || '').trim().toLowerCase();
        const tempPassword = String(body.tempPassword || '');
        const newPassword = String(body.newPassword || '');
        if (!email || !tempPassword || !newPassword || newPassword.length < 12 || newPassword.length > 256) {
          return json({ error: 'Choose a new password between 12 and 256 characters.' }, 400);
        }
        const ipHash = await hashText(req.headers.get('cf-connecting-ip') || 'unknown');
        const now = new Date();
        if (await isAuthBlocked(db, 'password-change', email, ipHash, now)) return json({ error: 'Too many attempts. Try again in 15 minutes.' }, 429);
        const account = await db.prepare('SELECT id,email,full_name,password_salt,password_hash,password_iterations,staff_id,status,must_change_password FROM worker_accounts WHERE email=?').bind(email).first();
        if (account && Number(account.password_iterations) > MAX_PASSWORD_ITERATIONS) return json({ error: 'Ask your roster administrator to resend your account invitation.' }, 401);
        const iterations = safePasswordIterations(account?.password_iterations);
        const digest = await passwordDigest(tempPassword || 'invalid-password', account?.password_salt || '00000000000000000000000000000000', iterations);
        if (!account || !account.must_change_password || !sameHash(digest, account.password_hash)) {
          await recordAuthFailure(db, 'password-change', email, ipHash, now);
          return json({ error: 'The temporary password entered is incorrect.' }, 401);
        }
        await clearAuthFailures(db, 'password-change', email, ipHash);
        if (account.status !== 'active' || !account.staff_id) return json({ error: 'This account is not active. Contact your administrator.' }, 403);
        const staffMember = await db.prepare("SELECT id,full_name,role_tag,pod_id,effective_from,status FROM staff WHERE id=? AND status IN ('active','pending')").bind(account.staff_id).first();
        if (!staffMember) return json({ error: 'Your linked staff profile is inactive. Contact your administrator.' }, 403);

        const newSalt = hex(crypto.getRandomValues(new Uint8Array(16)));
        const newDigest = await passwordDigest(newPassword, newSalt, CURRENT_PASSWORD_ITERATIONS);
        await db.prepare('UPDATE worker_accounts SET password_salt=?,password_hash=?,password_iterations=?,must_change_password=0,updated_at=CURRENT_TIMESTAMP WHERE id=?')
          .bind(newSalt, newDigest, CURRENT_PASSWORD_ITERATIONS, account.id).run();
        await db.prepare('DELETE FROM worker_sessions WHERE account_id=?').bind(account.id).run();
        await db.prepare("UPDATE staff SET status='active' WHERE id=?")
          .bind(account.staff_id).run();
        const session = await createSession(db, account.id);
        await audit(db, account.id, 'change_password', 'worker_account', account.id);
        return json({ ok: true, token: session.token, expiresAt: session.expiresAt, staffId: account.staff_id, fullName: account.full_name, email: account.email, staff: { ...staffMember, status: 'active' } });
      }

      if (url.pathname === '/api/auth/logout' && req.method === 'POST') {
        const token = (req.headers.get('authorization') || '').replace(/^Bearer\s+/i, '');
        if (token) await db.prepare('DELETE FROM worker_sessions WHERE token_hash=?').bind(await hashText(token)).run();
        return json({ ok: true });
      }

      if (url.pathname === '/api/me' && req.method === 'GET') {
        const actor = await auth(req, env, db);
        if (!actor) return json({ error: 'Sign in required.' }, 401);
        const person = await db.prepare('SELECT id,full_name,role_tag,pod_id,effective_from FROM staff WHERE id=? AND status=\'active\'').bind(actor.id).first();
        return person ? json({ ...actor, staff: person }) : json({ error: 'Staff profile is unavailable.' }, 404);
      }

      if (url.pathname === '/api/roster' && req.method === 'GET') {
        const actor = await auth(req, env, db);
        if (!actor) return json({ error: 'Sign in required.' }, 401);
        const month = url.searchParams.get('month');
        if (!/^\d{4}-\d{2}$/.test(month || '')) return json({ error: 'month must be YYYY-MM' }, 400);
        const [assignments, people, pods, covers, podAssignments] = await Promise.all([
          db.prepare("SELECT shift_date,staff_id,shift_code,is_override FROM shift_assignments WHERE shift_date LIKE ?||'%' ORDER BY shift_date").bind(month).all(),
          db.prepare("SELECT id,full_name,role_tag,pod_id,effective_from FROM staff WHERE status='active'").all(),
          db.prepare('SELECT id,name,color_hex,anchor_date,cycle_offset FROM pods').all(),
          db.prepare("SELECT id,original_date,shift_code,requester_id,cover_worker_id FROM shift_covers WHERE status='approved' AND original_date LIKE ?||'%' ORDER BY original_date").bind(month).all(),
          db.prepare('SELECT staff_id,pod_id,row_order FROM staff_pod_assignments WHERE month=?').bind(month).all(),
        ]);
        return json({ month, assignments: assignments.results, staff: people.results, pods: pods.results, covers: covers.results, podAssignments: podAssignments.results });
      }

      if (url.pathname === '/api/admin/publish' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        if (!/^\d{4}-\d{2}$/.test(body.month || '') || !body.assignments || typeof body.assignments !== 'object' ||
          (body.coverAssignments !== undefined && !Array.isArray(body.coverAssignments)) ||
          (body.removedCoverIds !== undefined && !Array.isArray(body.removedCoverIds)) ||
          (body.staffRecords !== undefined && !Array.isArray(body.staffRecords)) ||
          (body.inactiveStaffIds !== undefined && !Array.isArray(body.inactiveStaffIds)) ||
          (body.podAssignments !== undefined && !Array.isArray(body.podAssignments))) return json({ error: 'Invalid publish payload' }, 400);
        const batch = [];
        for (const person of body.staffRecords || []) {
          if (!/^[a-z0-9-]{1,100}$/.test(person.id || '') || !person.fullName || !person.roleTag || !/^\d{4}-\d{2}-\d{2}$/.test(person.effectiveFrom || '') || (person.podId && !/^(orange|green|blue|red)$/.test(person.podId))) return json({ error: 'Invalid staff record' }, 400);
          batch.push(db.prepare("INSERT INTO staff(id,full_name,role_tag,pod_id,phone,status,effective_from,is_admin) VALUES(?,?,?,?,?,'active',?,0) ON CONFLICT(id) DO UPDATE SET full_name=excluded.full_name,role_tag=excluded.role_tag,pod_id=excluded.pod_id,phone=excluded.phone,effective_from=excluded.effective_from")
            .bind(person.id, person.fullName, person.roleTag, person.podId || null, person.phone || null, person.effectiveFrom));
        }
        if (body.podAssignments) {
          batch.push(db.prepare('DELETE FROM staff_pod_assignments WHERE month=?').bind(body.month));
          for (const item of body.podAssignments) {
            if (!/^[a-z0-9-]{1,100}$/.test(item.staffId || '') || !['orange', 'green', 'blue', 'red'].includes(item.podId) || !Number.isInteger(item.rowOrder)) return json({ error: 'Invalid monthly pod assignment' }, 400);
            batch.push(db.prepare('INSERT INTO staff_pod_assignments(month,staff_id,pod_id,row_order) VALUES(?,?,?,?)').bind(body.month, item.staffId, item.podId, item.rowOrder));
          }
        }
        for (const id of body.inactiveStaffIds || []) {
          if (!/^[a-z0-9-]{1,100}$/.test(id || '')) return json({ error: 'Invalid staff id' }, 400);
          batch.push(db.prepare("UPDATE staff SET status='inactive' WHERE id=?").bind(id));
        }
        for (const [key, code] of Object.entries(body.assignments)) {
          const match = key.match(/^(\d{4}-\d{2}-\d{2}):(.+)$/);
          if (!match || !['M', 'N', 'OFF', 'LEAVE'].includes(code) || match[1].slice(0, 7) !== body.month) continue;
          batch.push(db.prepare('INSERT INTO shift_assignments(id,shift_date,staff_id,shift_code,is_override,updated_at) VALUES(?,?,?,?,1,CURRENT_TIMESTAMP) ON CONFLICT(shift_date,staff_id) DO UPDATE SET shift_code=excluded.shift_code,is_override=1,updated_at=CURRENT_TIMESTAMP')
            .bind(crypto.randomUUID(), match[1], match[2], code));
        }
        for (const cover of body.coverAssignments || []) {
          if (!cover.id || !/^\d{4}-\d{2}-\d{2}$/.test(cover.date || '') || cover.date.slice(0, 7) !== body.month || !['M', 'N'].includes(cover.shiftCode) || !cover.requesterId || !cover.coverWorkerId || cover.requesterId === cover.coverWorkerId) return json({ error: 'Invalid cover assignment' }, 400);
          batch.push(db.prepare("INSERT INTO shift_covers(id,original_date,shift_code,requester_id,cover_worker_id,status,approved_by) VALUES(?,?,?,?,?,'approved',?) ON CONFLICT(id) DO UPDATE SET original_date=excluded.original_date,shift_code=excluded.shift_code,requester_id=excluded.requester_id,cover_worker_id=excluded.cover_worker_id,status='approved',approved_by=excluded.approved_by,updated_at=CURRENT_TIMESTAMP")
            .bind(cover.id, cover.date, cover.shiftCode, cover.requesterId, cover.coverWorkerId, actor.name || actor.id));
        }
        for (const id of body.removedCoverIds || []) if (typeof id === 'string' && id.length <= 100) batch.push(db.prepare('DELETE FROM shift_covers WHERE id=?').bind(id));
        if (batch.length) await db.batch(batch);
        await audit(db, actor.id, 'publish', 'roster', body.month, { count: batch.length, covers: (body.coverAssignments || []).length });
        return json({ ok: true, published: Object.keys(body.assignments).length, covers: (body.coverAssignments || []).length });
      }

      if (url.pathname === '/api/covers' && req.method === 'POST') {
        const actor = await auth(req, env, db);
        if (!actor) return json({ error: 'Sign in required.' }, 401);
        const body = await req.json();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(body.date || '') || !['M', 'N'].includes(body.shiftCode)) return json({ error: 'Invalid shift' }, 400);
        const id = crypto.randomUUID();
        await db.prepare('INSERT INTO shift_covers(id,original_date,shift_code,requester_id,cover_worker_id,status,return_date) VALUES(?,?,?,?,?,?,?)')
          .bind(id, body.date, body.shiftCode, actor.id, body.coverWorkerId || null, 'pending', body.returnDate || null).run();
        await audit(db, actor.id, 'create', 'shift_cover', id);
        return json({ ok: true, id }, 201);
      }

      if (url.pathname === '/api/unavailability' && req.method === 'POST') {
        const actor = await auth(req, env, db);
        if (!actor) return json({ error: 'Sign in required.' }, 401);
        const body = await req.json();
        if (!/^\d{4}-\d{2}-\d{2}$/.test(body.startDate || '') || !/^\d{4}-\d{2}-\d{2}$/.test(body.endDate || '') || body.endDate < body.startDate) return json({ error: 'Invalid request' }, 400);
        const id = crypto.randomUUID();
        await db.prepare('INSERT INTO unavailability_flags(id,staff_id,start_date,end_date,reason) VALUES(?,?,?,?,?)').bind(id, actor.id, body.startDate, body.endDate, (body.reason || '').slice(0, 500)).run();
        await audit(db, actor.id, 'create', 'unavailability', id);
        return json({ ok: true, id }, 201);
      }

      if (url.pathname === '/api/unavailability' && req.method === 'GET') {
        const actor = await auth(req, env, db);
        if (!actor) return json({ error: 'Sign in required.' }, 401);
        const flags = await db.prepare("SELECT * FROM unavailability_flags WHERE staff_id=? ORDER BY created_at DESC").bind(actor.id).all();
        return json({ flags: flags.results });
      }

      if (url.pathname === '/api/admin/requests' && req.method === 'GET') {
        if (!await auth(req, env, db, true)) return json({ error: 'Admin token required' }, 401);
        const [covers, flags, accounts, availableStaff] = await Promise.all([
          db.prepare("SELECT c.*,s.full_name requester FROM shift_covers c JOIN staff s ON s.id=c.requester_id WHERE c.status='pending' ORDER BY c.created_at").all(),
          db.prepare("SELECT u.*,s.full_name FROM unavailability_flags u JOIN staff s ON s.id=u.staff_id WHERE u.status='pending' ORDER BY u.created_at").all(),
          db.prepare("SELECT id,email,full_name,created_at FROM worker_accounts WHERE status='pending' ORDER BY created_at").all(),
          db.prepare("SELECT s.id,s.full_name,s.role_tag FROM staff s LEFT JOIN worker_accounts a ON a.staff_id=s.id AND a.status='active' WHERE s.status='active' AND s.is_admin=0 AND a.id IS NULL ORDER BY s.full_name").all(),
        ]);
        return json({ covers: covers.results, unavailability: flags.results, accounts: accounts.results, availableStaff: availableStaff.results });
      }

      if (url.pathname.startsWith('/api/admin/accounts/') && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const accountId = url.pathname.split('/').pop(), body = await req.json();
        if (body.status === 'rejected') {
          const result = await db.prepare("UPDATE worker_accounts SET status='rejected',updated_at=CURRENT_TIMESTAMP WHERE id=? AND status='pending'").bind(accountId).run();
          if (!result.meta.changes) return json({ error: 'Pending registration not found' }, 404);
          await audit(db, actor.id, 'reject', 'worker_account', accountId);
          return json({ ok: true });
        }
        if (body.status !== 'approved' || !/^[a-z0-9-]{1,100}$/.test(body.staffId || '')) return json({ error: 'Choose a staff member to link to this account.' }, 400);
        const staffMember = await db.prepare("SELECT id FROM staff WHERE id=? AND status='active'").bind(body.staffId).first();
        if (!staffMember) return json({ error: 'Active staff member not found.' }, 404);
        const result = await db.prepare("UPDATE worker_accounts SET staff_id=?,status='active',updated_at=CURRENT_TIMESTAMP WHERE id=? AND status='pending'").bind(body.staffId, accountId).run();
        if (!result.meta.changes) return json({ error: 'This staff member may already have an account, or the registration is no longer pending.' }, 409);
        await audit(db, actor.id, 'approve', 'worker_account', accountId, { staffId: body.staffId });
        return json({ ok: true });
      }

      if (url.pathname.startsWith('/api/admin/requests/') && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const id = url.pathname.split('/').pop(), body = await req.json();
        if (!['approved', 'rejected'].includes(body.status)) return json({ error: 'Invalid status' }, 400);
        let result = await db.prepare("UPDATE unavailability_flags SET status=? WHERE id=? AND status='pending'").bind(body.status, id).run();
        if (!result.meta.changes) {
          result = await db.prepare("UPDATE shift_covers SET status=?,approved_by=NULL,updated_at=CURRENT_TIMESTAMP WHERE id=? AND status='pending'").bind(body.status, id).run();
        }
        if (!result.meta.changes) return json({ error: 'Pending request not found' }, 404);
        await audit(db, actor.id, body.status, 'request', id);
        return json({ ok: true });
      }

      if (url.pathname === '/api/admin/workers' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const fullName = String(body.fullName || '').trim();
        const email = String(body.email || '').trim().toLowerCase();
        const roleTag = String(body.roleTag || 'Support').trim();
        const podId = body.podId && ['orange', 'green', 'blue', 'red'].includes(body.podId) ? body.podId : null;
        const phone = body.phone ? String(body.phone).trim().slice(0, 40) : null;
        const effectiveFrom = body.effectiveFrom && /^\d{4}-\d{2}-\d{2}$/.test(body.effectiveFrom) ? body.effectiveFrom : new Date().toISOString().slice(0, 10);

        if (!fullName || fullName.length < 2 || fullName.length > 120) return json({ error: 'Enter a valid full name.' }, 400);
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || email.length > 254) return json({ error: 'Enter a valid email address.' }, 400);

        const baseId = fullName.normalize('NFKD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'staff';
        let staffId = baseId;
        const existingStaff = await db.prepare("SELECT id FROM staff WHERE id=?").bind(staffId).first();
        if (existingStaff) {
          const linkedAccount = await db.prepare("SELECT staff_id FROM worker_accounts WHERE email=?").bind(email).first();
          if (linkedAccount?.staff_id) staffId = linkedAccount.staff_id;
          else {
            let suffix = 2;
            while (await db.prepare("SELECT id FROM staff WHERE id=?").bind(`${baseId}-${suffix}`).first()) suffix++;
            staffId = `${baseId}-${suffix}`;
          }
        }

        await db.prepare("INSERT INTO staff(id,full_name,role_tag,pod_id,phone,status,effective_from,is_admin) VALUES(?,?,?,?,?,'pending',?,0) ON CONFLICT(id) DO UPDATE SET full_name=excluded.full_name,role_tag=excluded.role_tag,pod_id=excluded.pod_id,phone=excluded.phone")
          .bind(staffId, fullName, roleTag, podId, phone, effectiveFrom).run();

        const chars = '23456789abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ';
        const randomBytes = crypto.getRandomValues(new Uint8Array(12));
        let code = '';
        for (let i = 0; i < 12; i++) code += chars[randomBytes[i] % chars.length];
        const tempPassword = `Shift-${code}`;

        const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
        const digest = await passwordDigest(tempPassword, salt, CURRENT_PASSWORD_ITERATIONS);
        const accountId = crypto.randomUUID();

        await db.prepare(`INSERT INTO worker_accounts(id,email,full_name,password_salt,password_hash,password_iterations,staff_id,status,must_change_password)
          VALUES(?,?,?,?,?,?,?,'active',1)
          ON CONFLICT(email) DO UPDATE SET full_name=excluded.full_name,password_salt=excluded.password_salt,password_hash=excluded.password_hash,password_iterations=excluded.password_iterations,staff_id=excluded.staff_id,status='active',must_change_password=1,updated_at=CURRENT_TIMESTAMP`)
          .bind(accountId, email, fullName, salt, digest, CURRENT_PASSWORD_ITERATIONS, staffId).run();
        const provisionedAccount = await db.prepare('SELECT id FROM worker_accounts WHERE email=?').bind(email).first();
        if (provisionedAccount) await db.prepare('DELETE FROM worker_sessions WHERE account_id=?').bind(provisionedAccount.id).run();

        const emailResult = await sendBrevoEmail(env, { email, fullName, tempPassword });
        await audit(db, actor.id, 'provision_worker', 'worker_account', accountId, { email, staffId, emailSent: emailResult.success });

        return json({
          ok: true,
          staffId,
          email,
          fullName,
          roleTag,
          podId,
          effectiveFrom,
          tempPassword,
          status: 'pending',
          emailSent: emailResult.success,
          emailError: emailResult.error || undefined
        }, 201);
      }

      if (url.pathname === '/api/admin/workers/resend' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        const email = body.email ? String(body.email).trim().toLowerCase() : null;
        let account = null;
        if (staffId) {
          account = await db.prepare("SELECT a.id, a.email, a.full_name, a.staff_id, s.status staff_status FROM worker_accounts a JOIN staff s ON s.id=a.staff_id WHERE a.staff_id=?").bind(staffId).first();
        } else if (email) {
          account = await db.prepare("SELECT a.id, a.email, a.full_name, a.staff_id, s.status staff_status FROM worker_accounts a JOIN staff s ON s.id=a.staff_id WHERE a.email=?").bind(email).first();
        }
        if (!account) return json({ error: 'Worker account not found.' }, 404);

        const chars = '23456789abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ';
        const randomBytes = crypto.getRandomValues(new Uint8Array(12));
        let code = '';
        for (let i = 0; i < 12; i++) code += chars[randomBytes[i] % chars.length];
        const tempPassword = `Shift-${code}`;

        const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
        const digest = await passwordDigest(tempPassword, salt, CURRENT_PASSWORD_ITERATIONS);

        await db.prepare("UPDATE worker_accounts SET password_salt=?, password_hash=?, password_iterations=?, must_change_password=1, updated_at=CURRENT_TIMESTAMP WHERE id=?")
          .bind(salt, digest, CURRENT_PASSWORD_ITERATIONS, account.id).run();
        await db.prepare('DELETE FROM worker_sessions WHERE account_id=?').bind(account.id).run();

        const emailResult = await sendBrevoEmail(env, { email: account.email, fullName: account.full_name, tempPassword });
        await audit(db, actor.id, 'resend_invite', 'worker_account', account.id, { email: account.email, staffId: account.staff_id, emailSent: emailResult.success });

        return json({
          ok: true,
          staffId: account.staff_id,
          email: account.email,
          fullName: account.full_name,
          tempPassword,
          emailSent: emailResult.success,
          emailError: emailResult.error || undefined
        });
      }

      if (url.pathname === '/api/admin/workers/cancel' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        const person = await db.prepare("SELECT id, status FROM staff WHERE id=?").bind(staffId).first();
        if (!person) return json({ error: 'Staff record not found' }, 404);
        if (person.status !== 'pending') {
          return json({ error: 'Only pending invitations can be cancelled.' }, 400);
        }

        await db.prepare("DELETE FROM worker_accounts WHERE staff_id=?").bind(staffId).run();
        await db.prepare("DELETE FROM staff_pod_assignments WHERE staff_id=?").bind(staffId).run();
        await db.prepare("DELETE FROM shift_assignments WHERE staff_id=?").bind(staffId).run();
        await db.prepare("DELETE FROM staff WHERE id=? AND status='pending'").bind(staffId).run();
        await audit(db, actor.id, 'cancel_invite', 'staff', staffId);

        return json({ ok: true, staffId });
      }

      if (url.pathname === '/api/verify-email' && (req.method === 'GET' || req.method === 'POST')) {
        let token = url.searchParams.get('token');
        if (!token && req.method === 'POST') {
          const body = await req.json().catch(() => ({}));
          token = body.token;
        }
        token = String(token || '').trim();
        const wantsJson = (req.headers.get('accept') || '').includes('application/json') || req.method === 'POST';
        const portalUrl = env.PORTAL_URL || 'https://shiftdex-workers.pages.dev';
        const logoUrl = 'https://shiftdex-workers.pages.dev/shift-desk-mark.png';

        if (!token) {
          if (wantsJson) return json({ error: 'Verification token is required.' }, 400);
          return new Response(renderVerificationHtml({ success: false, title: 'Invalid Verification Link', message: 'The verification link is missing a valid token. Please contact your operations administrator for a new link.', portalUrl, logoUrl }), {
            status: 400,
            headers: { 'content-type': 'text/html; charset=utf-8' }
          });
        }

        const nowIso = new Date().toISOString();
        const account = await db.prepare("SELECT id, email, pending_email, full_name, email_verify_expires, staff_id FROM worker_accounts WHERE email_verify_token=?").bind(token).first();

        if (!account) {
          if (wantsJson) return json({ error: 'This verification link is invalid or has already been used.' }, 404);
          return new Response(renderVerificationHtml({ success: false, title: 'Link Invalid or Already Used', message: 'This verification link is no longer valid or has already been used. If your email was not updated, ask your administrator to resend the verification link.', portalUrl, logoUrl }), {
            status: 404,
            headers: { 'content-type': 'text/html; charset=utf-8' }
          });
        }

        if (account.email_verify_expires && account.email_verify_expires < nowIso) {
          if (wantsJson) return json({ error: 'This verification link has expired. Please ask your administrator to send a new one.' }, 410);
          return new Response(renderVerificationHtml({ success: false, title: 'Verification Link Expired', message: 'This email verification link expired. For your security, links are valid for 48 hours. Please ask your operations administrator to request the email change again.', portalUrl, logoUrl }), {
            status: 410,
            headers: { 'content-type': 'text/html; charset=utf-8' }
          });
        }

        const oldEmail = account.email;
        const newEmail = account.pending_email;
        if (!newEmail) {
          if (wantsJson) return json({ error: 'No pending email found for this token.' }, 400);
          return new Response(renderVerificationHtml({ success: false, title: 'No Pending Email', message: 'There is no pending email change associated with this account.', portalUrl, logoUrl }), {
            status: 400,
            headers: { 'content-type': 'text/html; charset=utf-8' }
          });
        }

        // Apply email update
        await db.prepare("UPDATE worker_accounts SET email=?, pending_email=NULL, email_verify_token=NULL, email_verify_expires=NULL, updated_at=CURRENT_TIMESTAMP WHERE id=?")
          .bind(newEmail, account.id).run();

        // Invalidate active worker sessions so user logs in cleanly with new credentials
        await db.prepare("DELETE FROM worker_sessions WHERE account_id=?").bind(account.id).run();

        await audit(db, account.id, 'verify_email_change', 'worker_account', account.id, { oldEmail, newEmail, staffId: account.staff_id });

        if (wantsJson) {
          return json({ ok: true, email: newEmail, previousEmail: oldEmail });
        }

        return new Response(renderVerificationHtml({
          success: true,
          title: 'Email Address Verified!',
          message: `Your Shift Desk account email has been successfully updated to <strong>${newEmail}</strong>. You can now use this email to sign in to your schedule and operational hub.`,
          portalUrl,
          logoUrl,
          newEmail
        }), {
          status: 200,
          headers: { 'content-type': 'text/html; charset=utf-8' }
        });
      }

      if (url.pathname === '/api/admin/workers/archive' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        const person = await db.prepare("SELECT id, status, full_name FROM staff WHERE id=?").bind(staffId).first();
        if (!person) return json({ error: 'Staff member not found' }, 404);

        // Mark staff as archived and clear active pod assignment
        await db.prepare("UPDATE staff SET status='archived', pod_id=NULL WHERE id=?").bind(staffId).run();
        await db.prepare("DELETE FROM staff_pod_assignments WHERE staff_id=?").bind(staffId).run();

        // Invalidate active sessions for linked worker account
        const acc = await db.prepare("SELECT id FROM worker_accounts WHERE staff_id=?").bind(staffId).first();
        if (acc) {
          await db.prepare("DELETE FROM worker_sessions WHERE account_id=?").bind(acc.id).run();
        }

        await audit(db, actor.id, 'archive_staff', 'staff', staffId, { name: person.full_name });
        return json({ ok: true, staffId, status: 'archived' });
      }

      if (url.pathname === '/api/admin/workers/restore' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        const podId = body.podId && ['orange', 'green', 'blue', 'red'].includes(body.podId) ? body.podId : null;
        const roleTag = body.roleTag ? String(body.roleTag).trim() : null;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        const person = await db.prepare("SELECT id, status, full_name, role_tag FROM staff WHERE id=?").bind(staffId).first();
        if (!person) return json({ error: 'Staff member not found' }, 404);

        await db.prepare("UPDATE staff SET status='active', pod_id=?, role_tag=COALESCE(?, role_tag) WHERE id=?")
          .bind(podId, roleTag, staffId).run();

        const acc = await db.prepare("SELECT id FROM worker_accounts WHERE staff_id=?").bind(staffId).first();
        if (acc) {
          await db.prepare("UPDATE worker_accounts SET status='active', updated_at=CURRENT_TIMESTAMP WHERE id=?").bind(acc.id).run();
        }

        await audit(db, actor.id, 'restore_staff', 'staff', staffId, { name: person.full_name, podId, roleTag: roleTag || person.role_tag });
        return json({ ok: true, staffId, status: 'active', podId, roleTag: roleTag || person.role_tag });
      }

      if ((url.pathname === '/api/admin/workers/delete' || url.pathname === '/api/admin/workers/permanent') && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        const person = await db.prepare("SELECT id, full_name FROM staff WHERE id=?").bind(staffId).first();
        if (!person) return json({ error: 'Staff member not found' }, 404);

        const acc = await db.prepare("SELECT id FROM worker_accounts WHERE staff_id=?").bind(staffId).first();
        if (acc) {
          await db.prepare("DELETE FROM worker_sessions WHERE account_id=?").bind(acc.id).run();
          await db.prepare("DELETE FROM worker_accounts WHERE id=?").bind(acc.id).run();
        }
        await db.prepare("DELETE FROM staff_pod_assignments WHERE staff_id=?").bind(staffId).run();
        await db.prepare("DELETE FROM shift_assignments WHERE staff_id=?").bind(staffId).run();
        await db.prepare("DELETE FROM shift_covers WHERE requester_id=? OR cover_worker_id=?").bind(staffId, staffId).run();
        await db.prepare("DELETE FROM unavailability_flags WHERE staff_id=?").bind(staffId).run();
        await db.prepare("DELETE FROM staff WHERE id=?").bind(staffId).run();

        await audit(db, actor.id, 'delete_staff_permanent', 'staff', staffId, { name: person.full_name });
        return json({ ok: true, staffId });
      }

      if (url.pathname === '/api/worker/change-password' && req.method === 'POST') {
        const actor = await auth(req, env, db);
        if (!actor) return json({ error: 'Sign in required.' }, 401);
        const body = await req.json();
        const currentPassword = String(body.currentPassword || '');
        const newPassword = String(body.newPassword || '');
        if (!currentPassword || !newPassword || newPassword.length < 8) {
          return json({ error: 'New password must be at least 8 characters.' }, 400);
        }
        const account = await db.prepare("SELECT id, password_salt, password_hash, password_iterations FROM worker_accounts WHERE id=?").bind(actor.accountId).first();
        if (!account) return json({ error: 'Account not found.' }, 404);
        const digest = await passwordDigest(currentPassword, account.password_salt, safePasswordIterations(account.password_iterations));
        if (!sameHash(digest, account.password_hash)) {
          return json({ error: 'Current password is incorrect.' }, 401);
        }
        const newSalt = hex(crypto.getRandomValues(new Uint8Array(16)));
        const newDigest = await passwordDigest(newPassword, newSalt, CURRENT_PASSWORD_ITERATIONS);
        await db.prepare("UPDATE worker_accounts SET password_salt=?, password_hash=?, password_iterations=?, must_change_password=0, updated_at=CURRENT_TIMESTAMP WHERE id=?")
          .bind(newSalt, newDigest, CURRENT_PASSWORD_ITERATIONS, account.id).run();
        return json({ ok: true });
      }

      if ((url.pathname === '/api/admin/workers/reset-password' || url.pathname === '/api/admin/workers/change-password') && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        const account = await db.prepare("SELECT a.id, a.email, a.full_name, a.staff_id, s.status staff_status FROM worker_accounts a JOIN staff s ON s.id=a.staff_id WHERE a.staff_id=?").bind(staffId).first();
        if (!account) return json({ error: 'Linked worker account not found for this staff member.' }, 404);

        let finalPassword = body.newPassword ? String(body.newPassword).trim() : '';
        let mustChange = 0;
        if (!finalPassword) {
          const chars = '23456789abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ';
          const randomBytes = crypto.getRandomValues(new Uint8Array(12));
          let code = '';
          for (let i = 0; i < 12; i++) code += chars[randomBytes[i] % chars.length];
          finalPassword = `Shift-${code}`;
          mustChange = 1;
        }

        const salt = hex(crypto.getRandomValues(new Uint8Array(16)));
        const digest = await passwordDigest(finalPassword, salt, CURRENT_PASSWORD_ITERATIONS);

        await db.prepare("UPDATE worker_accounts SET password_salt=?, password_hash=?, password_iterations=?, must_change_password=?, updated_at=CURRENT_TIMESTAMP WHERE id=?")
          .bind(salt, digest, CURRENT_PASSWORD_ITERATIONS, mustChange, account.id).run();
        await db.prepare('DELETE FROM worker_sessions WHERE account_id=?').bind(account.id).run();

        const emailResult = await sendPasswordResetEmail(env, { email: account.email, fullName: account.full_name, tempPassword: finalPassword });
        await audit(db, actor.id, 'change_worker_password', 'worker_account', account.id, { email: account.email, staffId: account.staff_id, emailSent: emailResult.success });

        return json({
          ok: true,
          staffId: account.staff_id,
          email: account.email,
          fullName: account.full_name,
          tempPassword: finalPassword,
          emailSent: emailResult.success,
          emailError: emailResult.error || undefined
        });
      }

      if (url.pathname === '/api/admin/workers/change-email' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        const newEmail = String(body.newEmail || '').trim().toLowerCase();
        if (!staffId) return json({ error: 'staffId required' }, 400);
        if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(newEmail) || newEmail.length > 254) {
          return json({ error: 'Enter a valid new email address.' }, 400);
        }

        const account = await db.prepare("SELECT a.id, a.email, a.full_name, a.staff_id FROM worker_accounts a WHERE a.staff_id=?").bind(staffId).first();
        if (!account) return json({ error: 'Worker account not found.' }, 404);

        if (account.email.toLowerCase() === newEmail) {
          return json({ error: 'The new email is the same as the current email.' }, 400);
        }

        const existingOther = await db.prepare("SELECT id FROM worker_accounts WHERE email=? AND id!=?").bind(newEmail, account.id).first();
        if (existingOther) {
          return json({ error: 'This email address is already in use by another account.' }, 409);
        }

        const verifyToken = tokenString(crypto.getRandomValues(new Uint8Array(32)));
        const verifyExpires = new Date(Date.now() + 48 * 3600 * 1000).toISOString();

        await db.prepare("UPDATE worker_accounts SET pending_email=?, email_verify_token=?, email_verify_expires=?, updated_at=CURRENT_TIMESTAMP WHERE id=?")
          .bind(newEmail, verifyToken, verifyExpires, account.id).run();

        const apiBase = new URL(req.url).origin;
        const verifyUrl = `${apiBase}/api/verify-email?token=${encodeURIComponent(verifyToken)}`;

        const emailResult = await sendEmailVerificationEmail(env, {
          newEmail,
          currentEmail: account.email,
          fullName: account.full_name,
          verifyUrl
        });

        await audit(db, actor.id, 'request_email_change', 'worker_account', account.id, {
          staffId,
          currentEmail: account.email,
          newEmail,
          emailSent: emailResult.success
        });

        return json({
          ok: true,
          staffId,
          currentEmail: account.email,
          pendingEmail: newEmail,
          verifyExpires,
          emailSent: emailResult.success,
          emailError: emailResult.error || undefined
        });
      }

      if (url.pathname === '/api/admin/workers/cancel-email-change' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        await db.prepare("UPDATE worker_accounts SET pending_email=NULL, email_verify_token=NULL, email_verify_expires=NULL, updated_at=CURRENT_TIMESTAMP WHERE staff_id=?")
          .bind(staffId).run();

        await audit(db, actor.id, 'cancel_email_change', 'worker_account', staffId);
        return json({ ok: true, staffId });
      }

      if (url.pathname === '/api/admin/workers/resend-email-change' && req.method === 'POST') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const body = await req.json();
        const staffId = body.staffId;
        if (!staffId) return json({ error: 'staffId required' }, 400);

        const account = await db.prepare("SELECT id, email, pending_email, full_name, staff_id FROM worker_accounts WHERE staff_id=?").bind(staffId).first();
        if (!account || !account.pending_email) return json({ error: 'No pending email change found.' }, 404);

        const verifyToken = tokenString(crypto.getRandomValues(new Uint8Array(32)));
        const verifyExpires = new Date(Date.now() + 48 * 3600 * 1000).toISOString();

        await db.prepare("UPDATE worker_accounts SET email_verify_token=?, email_verify_expires=?, updated_at=CURRENT_TIMESTAMP WHERE id=?")
          .bind(verifyToken, verifyExpires, account.id).run();

        const apiBase = new URL(req.url).origin;
        const verifyUrl = `${apiBase}/api/verify-email?token=${encodeURIComponent(verifyToken)}`;

        const emailResult = await sendEmailVerificationEmail(env, {
          newEmail: account.pending_email,
          currentEmail: account.email,
          fullName: account.full_name,
          verifyUrl
        });

        return json({
          ok: true,
          staffId,
          pendingEmail: account.pending_email,
          emailSent: emailResult.success,
          emailError: emailResult.error || undefined
        });
      }

      if (url.pathname === '/api/admin/directory' && req.method === 'GET') {
        const actor = await auth(req, env, db, true);
        if (!actor) return json({ error: 'Admin token required' }, 401);
        const staffList = await db.prepare(`
          SELECT s.id, s.full_name, s.role_tag, s.pod_id, s.phone, s.status, s.effective_from,
                 a.id account_id, a.email, a.must_change_password, a.pending_email, a.email_verify_expires
          FROM staff s
          LEFT JOIN worker_accounts a ON a.staff_id=s.id
          WHERE s.is_admin=0
          ORDER BY s.full_name ASC
        `).all();
        return json({ staff: staffList.results });
      }

      return json({ error: 'Not found' }, 404);
    } catch (error) {
      return json({ error: 'Internal server error', detail: env.ENVIRONMENT === 'development' ? error.message : undefined }, 500);
    }
  },
};
