// GET /api/submissions
// Header: Authorization: Bearer <ADMIN_SECRET>
// Returns: { submissions: [{ cabinetId, cabinetTitle, submittedAt, url, data, fieldStates }, ...] }

const REPO = process.env.GH_REPO || 'ayoubrezala/labosphere-form-cabinets';

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  // Admin auth
  const auth = req.headers.authorization || '';
  const provided = auth.replace(/^Bearer\s+/i, '');
  if (!process.env.ADMIN_SECRET || provided !== process.env.ADMIN_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  if (!process.env.GH_TOKEN) {
    return res.status(500).json({ error: 'Server not configured (GH_TOKEN missing)' });
  }

  const ghRes = await fetch(`https://api.github.com/repos/${REPO}/issues?labels=submission&state=all&per_page=100`, {
    headers: {
      'Authorization': `Bearer ${process.env.GH_TOKEN}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    }
  });
  if (!ghRes.ok) {
    return res.status(502).json({ error: 'GitHub API failed' });
  }
  const issues = await ghRes.json();

  const submissions = issues.map(issue => {
    let parsed = null;
    const match = (issue.body || '').match(/```json\n([\s\S]+?)\n```/);
    if (match) {
      try { parsed = JSON.parse(match[1]); } catch (e) {}
    }
    const cabinetLabel = (issue.labels || []).find(l => (l.name || '').startsWith('cabinet-'));
    const cabinetId = cabinetLabel ? parseInt(cabinetLabel.name.replace('cabinet-', ''), 10) : null;
    const cabinetTitle = (issue.title || '').replace(/^Soumission · /, '');
    return {
      issueNumber: issue.number,
      cabinetId,
      cabinetTitle,
      submittedAt: issue.created_at,
      url: issue.html_url,
      state: issue.state,
      data: parsed?.data || null,
      fieldStates: parsed?.fieldStates || null
    };
  });

  return res.status(200).json({ submissions });
}
