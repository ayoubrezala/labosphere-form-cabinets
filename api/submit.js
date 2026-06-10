// POST /api/submit
// Body: { cabinetId, cabinetTitle, token, data, fieldStates, timestamp }
// Validates token, stores submission as a GitHub Issue, notifies Cyndie via Brevo.
import crypto from 'crypto';

const REPO = process.env.GH_REPO || 'ayoubrezala/labosphere-form-cabinets';
const TOKEN_SECRET = process.env.TOKEN_SECRET || 'labosphere-3.0-token-salt-2026';
const BREVO_API_KEY = process.env.BREVO_API_KEY;
const NOTIFY_FROM_EMAIL = process.env.BREVO_NOTIFY_FROM_EMAIL || 'hello@label-co-pilotes.com';
const NOTIFY_FROM_NAME = process.env.BREVO_NOTIFY_FROM_NAME || 'Labosphère — Formulaire Cabinets';
const NOTIFY_TO = (process.env.BREVO_NOTIFY_TO || 'c.brien@label-co-pilotes.com')
  .split(',').map(s => s.trim()).filter(Boolean);
const DASHBOARD_URL = process.env.FORM_URL_BASE || 'https://labosphere-form-cabinets.vercel.app';

function expectedToken(cabinetId, cabinetTitle) {
  const payload = `${TOKEN_SECRET}|${cabinetId}|${(cabinetTitle || '').toLowerCase().trim()}`;
  return crypto.createHash('sha256').update(payload).digest('hex').slice(0, 10);
}

function htmlEscape(s) {
  return (s || '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

async function notifyCyndie({ cabinetTitle, timestamp, fieldStates, issueUrl }) {
  if (!BREVO_API_KEY || !NOTIFY_TO.length) return { skipped: 'BREVO_API_KEY or NOTIFY_TO missing' };
  const states = fieldStates || {};
  const modified = Object.entries(states).filter(([, v]) => v === 'modified').map(([k]) => k);
  const confirmed = Object.entries(states).filter(([, v]) => v === 'confirmed').length;
  const dt = new Date(timestamp || Date.now());
  const dateStr = dt.toLocaleString('fr-FR', { timeZone: 'Europe/Paris', dateStyle: 'long', timeStyle: 'short' });

  const modifiedHtml = modified.length
    ? `<ul style="margin:8px 0 0;padding-left:20px;">${modified.map(k => `<li>${htmlEscape(k)}</li>`).join('')}</ul>`
    : '<p style="color:#666;font-style:italic;margin:8px 0 0;">Aucun champ modifié — tout a été confirmé tel quel.</p>';

  const html = `<!doctype html><html lang="fr"><body style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;line-height:1.5;margin:0;padding:24px;background:#f7f7f7;">
  <div style="max-width:600px;margin:0 auto;background:#fff;padding:24px;border-radius:8px;border:1px solid #eee;">
    <h2 style="margin:0 0 16px;color:#d44a3a;font-size:18px;">📩 Nouvelle soumission au formulaire</h2>
    <p><strong>Cabinet :</strong> ${htmlEscape(cabinetTitle)}</p>
    <p><strong>Reçu le :</strong> ${dateStr}</p>
    <p><strong>Champs confirmés :</strong> ${confirmed}</p>
    <p style="margin-bottom:0;"><strong>Champs modifiés :</strong> ${modified.length}</p>
    ${modifiedHtml}
    <p style="margin:24px 0 0;">
      <a href="${DASHBOARD_URL}" style="background:#d44a3a;color:#fff;padding:10px 20px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;">Voir et valider dans le dashboard →</a>
    </p>
    <p style="font-size:11px;color:#999;margin-top:24px;border-top:1px solid #eee;padding-top:12px;">Notification automatique du formulaire Labosphère.</p>
  </div>
</body></html>`;

  try {
    const r = await fetch('https://api.brevo.com/v3/smtp/email', {
      method: 'POST',
      headers: { 'api-key': BREVO_API_KEY, 'Content-Type': 'application/json', 'Accept': 'application/json' },
      body: JSON.stringify({
        sender: { name: NOTIFY_FROM_NAME, email: NOTIFY_FROM_EMAIL },
        to: NOTIFY_TO.map(email => ({ email })),
        subject: `📩 Nouvelle soumission · ${cabinetTitle}`,
        htmlContent: html,
        tags: ['labosphere-form-notify']
      })
    });
    if (!r.ok) return { error: `brevo ${r.status}: ${(await r.text()).slice(0, 200)}` };
    const j = await r.json();
    return { ok: true, messageId: j.messageId };
  } catch (e) {
    return { error: e.message || String(e) };
  }
}

export default async function handler(req, res) {
  // CORS for any origin (the form may run on Vercel preview / GitHub Pages)
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let payload;
  try {
    payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body;
  } catch (e) {
    return res.status(400).json({ error: 'Invalid JSON' });
  }

  const { cabinetId, cabinetTitle, token, data, fieldStates, timestamp } = payload || {};
  if (!cabinetId || !cabinetTitle || !token || !data) {
    return res.status(400).json({ error: 'Missing required fields' });
  }

  // Validate token
  const expected = expectedToken(cabinetId, cabinetTitle);
  if (expected !== token) {
    return res.status(403).json({ error: 'Invalid token' });
  }

  if (!process.env.GH_TOKEN) {
    return res.status(500).json({ error: 'Server not configured (GH_TOKEN missing)' });
  }

  const ts = timestamp || new Date().toISOString();
  const body = [
    `**Cabinet:** ${cabinetTitle}`,
    `**Backup ID:** ${cabinetId}`,
    `**Soumis le:** ${ts}`,
    ``,
    `### Données soumises`,
    '```json',
    JSON.stringify({ data, fieldStates }, null, 2),
    '```'
  ].join('\n');

  const ghRes = await fetch(`https://api.github.com/repos/${REPO}/issues`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${process.env.GH_TOKEN}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      title: `Soumission · ${cabinetTitle}`,
      body,
      labels: ['submission', `cabinet-${cabinetId}`]
    })
  });

  if (!ghRes.ok) {
    const errText = await ghRes.text();
    return res.status(502).json({ error: 'GitHub API failed', details: errText.slice(0, 300) });
  }

  const issue = await ghRes.json();

  // Fire-and-forget notification email (don't block the user response if it fails)
  const notif = await notifyCyndie({
    cabinetTitle,
    timestamp: ts,
    fieldStates,
    issueUrl: issue.html_url
  });

  return res.status(200).json({ ok: true, issueNumber: issue.number, url: issue.html_url, notify: notif });
}
