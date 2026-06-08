// POST /api/send-mail
// Header: Authorization: Bearer <ADMIN_SECRET>
// Body: { cabinetIds?: number[] }   // omit or [] = send to ALL cabinets
//
// Envoie un email personnalisé à chaque cabinet via Brevo API v3, avec :
//  - lien formulaire unique (token déterministe)
//  - signature image Cyndie BRIEN (hébergée sur le déploiement Vercel)
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

const REPO = process.env.GH_REPO || 'ayoubrezala/labosphere-form-cabinets';
const GRAPH_SITE_ID = process.env.GRAPH_SITE_ID
  || 'labelcopilotes.sharepoint.com,f7dcb637-5a00-45f3-bfe7-35d64898817b,074deece-cf07-4a81-806f-5d91b14a10ee';
const GRAPH_LIST_ID = process.env.GRAPH_LIST_ID
  || '0a85d98e-641d-48d8-90c5-3ba16ffbb6bb';
const TENANT_ID = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;

function expectedToken(cabinetId, cabinetTitle) {
  const payload = `${TOKEN_SECRET}|${cabinetId}|${(cabinetTitle || '').toLowerCase().trim()}`;
  return crypto.createHash('sha256').update(payload).digest('hex').slice(0, 10);
}

async function getGraphToken() {
  const body = new URLSearchParams({
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials',
  });
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });
  if (!r.ok) throw new Error(`graph token http ${r.status}`);
  return (await r.json()).access_token;
}

async function listCabinets(token) {
  // Pull all items with their reference fields
  let url = `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE_ID}/lists/${GRAPH_LIST_ID}/items?$top=100&$expand=fields($select=Title,Adresse_x0020_mail_x0020_r_x00e9,Pr_x00e9_nom_x0020_et_x0020_Nom_)`;
  const items = [];
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) throw new Error(`graph list http ${r.status}`);
    const j = await r.json();
    items.push(...(j.value || []));
    url = j['@odata.nextLink'] || null;
  }
  return items.map(it => ({
    id: parseInt(it.id, 10),
    title: (it.fields && it.fields.Title) || '',
    email: (it.fields && it.fields.Adresse_x0020_mail_x0020_r_x00e9) || '',
    referent: (it.fields && it.fields.Pr_x00e9_nom_x0020_et_x0020_Nom_) || '',
  }));
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
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) return res.status(500).json({ error: 'MS_* missing' });

  let payload;
  try { payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch { return res.status(400).json({ error: 'Invalid JSON' }); }
  const { cabinetIds } = payload || {};
  const filterIds = Array.isArray(cabinetIds) && cabinetIds.length ? new Set(cabinetIds.map(Number)) : null;

  let cabinets;
  try {
    const t = await getGraphToken();
    cabinets = await listCabinets(t);
  } catch (e) {
    return res.status(502).json({ error: 'Cabinets list fetch failed', details: e.message });
  }

  let targets = cabinets.filter(c => c.title);
  if (filterIds) targets = targets.filter(c => filterIds.has(c.id));

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
