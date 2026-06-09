// POST /api/send-mail
// Header: Authorization: Bearer <ADMIN_SECRET>
// Body: { cabinets: [{ id, title, email, referent }, ...] }
//
// Envoie un email personnalisé à chaque cabinet via Brevo API v3, avec :
//  - lien formulaire unique (token déterministe basé sur id+title)
//  - signature image Cyndie BRIEN embarquée inline (CID attachment)
// Cabinets sans email sont skippés (rapportés dans la réponse).

import crypto from 'crypto';
import fs from 'node:fs';
import path from 'node:path';

const TOKEN_SECRET = process.env.TOKEN_SECRET || 'labosphere-3.0-token-salt-2026';
const BREVO_API_KEY = process.env.BREVO_API_KEY;
const SENDER_EMAIL = process.env.BREVO_SENDER_EMAIL || 'c.brien@label-co-pilotes.com';
const SENDER_NAME = process.env.BREVO_SENDER_NAME || 'Cyndie BRIEN';
const FORM_URL_BASE = process.env.FORM_URL_BASE || 'https://labosphere-form-cabinets.vercel.app';

// Load signature once and embed as inline attachment with CID
let SIGNATURE_BASE64 = null;
function loadSignature() {
  if (SIGNATURE_BASE64) return SIGNATURE_BASE64;
  try {
    const p = path.join(process.cwd(), 'assets', 'signature-cyndie.png');
    SIGNATURE_BASE64 = fs.readFileSync(p).toString('base64');
  } catch (e) {
    SIGNATURE_BASE64 = '';
  }
  return SIGNATURE_BASE64;
}


function expectedToken(cabinetId, cabinetTitle) {
  const payload = `${TOKEN_SECRET}|${cabinetId}|${(cabinetTitle || '').toLowerCase().trim()}`;
  return crypto.createHash('sha256').update(payload).digest('hex').slice(0, 10);
}

function splitEmails(raw) {
  if (!raw) return [];
  return raw.split(/[;,]/).map(s => s.trim()).filter(s => s && /@/.test(s));
}

function htmlEscape(s) {
  return (s || '').replace(/[&<>"']/g, c => ({ '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;' }[c]));
}

function buildHtml(cabinet, link) {
  const greeting = cabinet.referent
    ? `Bonjour ${htmlEscape(cabinet.referent)},`
    : `Bonjour à tous,`;
  return `<!doctype html>
<html lang="fr">
<body style="font-family:Arial,Helvetica,sans-serif;font-size:14px;color:#333;line-height:1.5;margin:0;padding:24px;background:#fff;">
  <div style="max-width:640px;margin:0 auto;">
    <p>${greeting}</p>
    <p>Nous souhaiterions pouvoir mettre à jour les informations concernant votre cabinet dans notre base de données afin de pouvoir réaliser un trombinoscope <strong>"Cabinets"</strong> partagé avec vous, nous vous remercions de prendre quelques minutes pour compléter ce formulaire :</p>
    <p style="margin:24px 0;">
      <a href="${link}" style="background:#d44a3a;color:#fff;padding:12px 24px;text-decoration:none;border-radius:6px;font-weight:bold;display:inline-block;">Compléter le formulaire</a>
    </p>
    <p style="font-size:12px;color:#666;">Si le bouton ne s'affiche pas, copiez-collez ce lien dans votre navigateur :<br><a href="${link}">${link}</a></p>
    <p><strong>Merci de compléter ce formulaire au plus tard le 19 juin.</strong></p>
    <p>Bonne journée,<br>Bien cordialement,</p>
    <div style="margin-top:24px;">
      <img src="cid:signature-cyndie.png" alt="Cyndie BRIEN - Label Co-Pilotes" style="display:block;max-width:600px;width:100%;height:auto;border:0;" />
    </div>
  </div>
</body>
</html>`;
}

async function sendOne(cabinet) {
  const tos = splitEmails(cabinet.email);
  if (tos.length === 0) {
    return { id: cabinet.id, title: cabinet.title, skipped: 'no email' };
  }
  const token = expectedToken(cabinet.id, cabinet.title);
  const link = `${FORM_URL_BASE}/?c=${cabinet.id}&t=${token}`;
  const html = buildHtml(cabinet, link);

  const sig = loadSignature();
  const payload = {
    sender: { name: SENDER_NAME, email: SENDER_EMAIL },
    to: tos.map(email => ({ email })),
    subject: `${cabinet.title} — Mise à jour de vos informations pour le trombinoscope Cabinets`,
    htmlContent: html,
    tags: ['labosphere-form', `cabinet-${cabinet.id}`]
  };
  if (sig) {
    payload.attachment = [{ name: 'signature-cyndie.png', content: sig }];
  }
  const r = await fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'POST',
    headers: {
      'api-key': BREVO_API_KEY,
      'Content-Type': 'application/json',
      'Accept': 'application/json'
    },
    body: JSON.stringify(payload)
  });
  if (!r.ok) {
    const txt = await r.text();
    return { id: cabinet.id, title: cabinet.title, error: `brevo ${r.status}: ${txt.slice(0, 200)}` };
  }
  const j = await r.json();
  return { id: cabinet.id, title: cabinet.title, recipients: tos, messageId: j.messageId };
}

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const auth = req.headers.authorization || '';
  const provided = auth.replace(/^Bearer\s+/i, '');
  if (!process.env.ADMIN_SECRET || provided !== process.env.ADMIN_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  if (!BREVO_API_KEY) return res.status(500).json({ error: 'BREVO_API_KEY missing' });

  let payload;
  try { payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch { return res.status(400).json({ error: 'Invalid JSON' }); }
  const { cabinets } = payload || {};
  if (!Array.isArray(cabinets) || cabinets.length === 0) {
    return res.status(400).json({ error: 'Missing cabinets array' });
  }

  const targets = cabinets
    .filter(c => c && c.id != null && c.title)
    .map(c => ({ id: Number(c.id), title: c.title, email: c.email || '', referent: c.referent || '' }));

  // Send sequentially to keep within Vercel function time + Brevo rate-limit friendly
  const results = [];
  for (const cab of targets) {
    try {
      results.push(await sendOne(cab));
    } catch (e) {
      results.push({ id: cab.id, title: cab.title, error: e.message || String(e) });
    }
  }

  const sent = results.filter(r => r.messageId).length;
  const skipped = results.filter(r => r.skipped).length;
  const failed = results.filter(r => r.error).length;
  return res.status(200).json({ ok: true, summary: { total: targets.length, sent, skipped, failed }, results });
}
