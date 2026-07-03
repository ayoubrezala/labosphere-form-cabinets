// GET /api/cabinets-sync
// Header: Authorization: Bearer <ADMIN_SECRET>
//
// Returns every cabinet currently in the SharePoint Cabinets list with:
//   { Id, Title, _token, Adressemailr_x00e9_f_x00e9_rent, Pr_x00e9_nomdur_x00e9_f_x00e9_re }
// Id is the SharePoint list-item id. _token is expectedToken(Id, Title).
//
// The admin dashboard calls this on load to detect cabinets added in SP
// that are not yet in the static embed baked into index.html.
import crypto from 'crypto';

const TOKEN_SECRET = process.env.TOKEN_SECRET || 'labosphere-3.0-token-salt-2026';
const TENANT_ID = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const GRAPH_SITE_ID = process.env.GRAPH_SITE_ID
  || 'labelcopilotes.sharepoint.com,f7dcb637-5a00-45f3-bfe7-35d64898817b,074deece-cf07-4a81-806f-5d91b14a10ee';
const GRAPH_LIST_ID = process.env.GRAPH_LIST_ID
  || '0a85d98e-641d-48d8-90c5-3ba16ffbb6bb';

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

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Authorization, Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  const auth = req.headers.authorization || '';
  const provided = auth.replace(/^Bearer\s+/i, '');
  if (!process.env.ADMIN_SECRET || provided !== process.env.ADMIN_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    return res.status(500).json({ error: 'MS_* env vars missing' });
  }

  try {
    const token = await getGraphToken();
    let url = `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE_ID}/lists/${GRAPH_LIST_ID}/items?$top=200&$expand=fields($select=Title,Adresse_x0020_mail_x0020_r_x00e9,Pr_x00e9_nom_x0020_et_x0020_Nom_)`;
    const items = [];
    while (url) {
      const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!r.ok) throw new Error(`graph list http ${r.status}: ${(await r.text()).slice(0, 200)}`);
      const j = await r.json();
      items.push(...(j.value || []));
      url = j['@odata.nextLink'] || null;
    }
    const cabinets = items
      .filter(it => it.fields && it.fields.Title)
      .map(it => {
        const id = parseInt(it.id, 10);
        const title = it.fields.Title;
        return {
          Id: id,
          Title: title,
          Adressemailr_x00e9_f_x00e9_rent: it.fields.Adresse_x0020_mail_x0020_r_x00e9 || '',
          Pr_x00e9_nomdur_x00e9_f_x00e9_re: it.fields.Pr_x00e9_nom_x0020_et_x0020_Nom_ || '',
          _token: expectedToken(id, title),
        };
      });
    return res.status(200).json({ cabinets });
  } catch (e) {
    return res.status(502).json({ error: 'SP fetch failed', details: e.message || String(e) });
  }
}
