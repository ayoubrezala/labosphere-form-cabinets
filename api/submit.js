// POST /api/submit
// Body: { cabinetId, cabinetTitle, token, data, fieldStates, timestamp }
// Validates token, stores submission as a GitHub Issue.
import crypto from 'crypto';

const REPO = process.env.GH_REPO || 'ayoubrezala/labosphere-form-cabinets';
const TOKEN_SECRET = process.env.TOKEN_SECRET || 'labosphere-3.0-token-salt-2026';

function expectedToken(cabinetId, cabinetTitle) {
  const payload = `${TOKEN_SECRET}|${cabinetId}|${(cabinetTitle || '').toLowerCase().trim()}`;
  return crypto.createHash('sha256').update(payload).digest('hex').slice(0, 10);
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
  return res.status(200).json({ ok: true, issueNumber: issue.number, url: issue.html_url });
}
