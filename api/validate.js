// POST /api/validate
// Header: Authorization: Bearer <ADMIN_SECRET>
// Body: { issueNumber }
// Closes the corresponding GitHub Issue → marks the submission as validated.

const REPO = process.env.GH_REPO || 'ayoubrezala/labosphere-form-cabinets';

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
  if (!process.env.GH_TOKEN) {
    return res.status(500).json({ error: 'Server not configured (GH_TOKEN missing)' });
  }

  let payload;
  try { payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch { return res.status(400).json({ error: 'Invalid JSON' }); }

  const { issueNumber } = payload || {};
  if (!issueNumber) return res.status(400).json({ error: 'Missing issueNumber' });

  // Add a "validated" label + close the issue
  // 1) Add label
  await fetch(`https://api.github.com/repos/${REPO}/issues/${issueNumber}/labels`, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${process.env.GH_TOKEN}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ labels: ['validated'] })
  });

  // 2) Close the issue
  const r = await fetch(`https://api.github.com/repos/${REPO}/issues/${issueNumber}`, {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${process.env.GH_TOKEN}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ state: 'closed', state_reason: 'completed' })
  });

  if (!r.ok) {
    const e = await r.text();
    return res.status(502).json({ error: 'GitHub API failed', details: e.slice(0, 200) });
  }

  return res.status(200).json({ ok: true });
}
