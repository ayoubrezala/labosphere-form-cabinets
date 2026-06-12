// POST /api/validate
// Header: Authorization: Bearer <ADMIN_SECRET>
// Body: { issueNumber }
//
// Pipeline:
//   1) Fetch the GitHub Issue → parse cabinetTitle + { data, fieldStates }
//   2) Acquire Microsoft Graph token (client_credentials)
//   3) Find Cabinets item by Title (lower/trim match)
//   4) PATCH /sites/{siteId}/lists/{listId}/items/{itemId}/fields with the modified fields only
//   5) Add 'validated' label + close the issue

const REPO = process.env.GH_REPO || 'ayoubrezala/labosphere-form-cabinets';

const TENANT_ID = process.env.MS_TENANT_ID;
const CLIENT_ID = process.env.MS_CLIENT_ID;
const CLIENT_SECRET = process.env.MS_CLIENT_SECRET;
const GRAPH_SITE_ID = process.env.GRAPH_SITE_ID
  || 'labelcopilotes.sharepoint.com,f7dcb637-5a00-45f3-bfe7-35d64898817b,074deece-cf07-4a81-806f-5d91b14a10ee';
const GRAPH_LIST_ID = process.env.GRAPH_LIST_ID
  || '0a85d98e-641d-48d8-90c5-3ba16ffbb6bb';

// Form field internal name → Cabinets list InternalName + type
const FIELD_MAP = {
  Title:                            { name: 'Title',                            type: 'text'   },
  OFFRE:                            { name: 'Offre_x0020_Label',                type: 'text'   },
  Adresse:                          { name: 'Adresse_x0020_Si_x00e8_ge',        type: 'text'   },
  Pr_x00e9_nomdur_x00e9_f_x00e9_re: { name: 'Pr_x00e9_nom_x0020_et_x0020_Nom_', type: 'text'   },
  Pr_x00e9_nomNomassoci_x00e9_:     { name: 'Pr_x00e9_nom_x0020_et_x0020_Nom_0',type: 'text'   },
  NBCollabs:                        { name: 'Effectif_x0020_total',             type: 'number' },
  NBassoci_x00e9_s:                 { name: 'Nb_x0020_Associ_x00e9_s',          type: 'number' },
  EChorsassoci_x00e9_s:             { name: 'Nb_x0020_EC_x0020_hors_x0020_ass', type: 'number' },
  NBSites:                          { name: 'Nb_x0020_Sites_x0020_ou_x0020_Ag', type: 'number' },
  CodespostauxSites:                { name: 'Adresses_x0020_Sites',             type: 'text'   },
  Outils_x0020_Prod_x0020_Compta:   { name: 'Outils_x0020_Prod_x0020_Compta',   type: 'multi'  },
  Outils_x0020_Prod_x0020_Social:   { name: 'Outils_x0020_Prod_x0020_Social',   type: 'multi'  },
  Outils_x0020_Prod_x0020_Juridiqu: { name: 'Outils_x0020_Prod_x0020_Juridiqu', type: 'multi'  },
  Outils_x0020_Conseil:             { name: 'Outils_x0020_Conseil',             type: 'multi'  },
  Outils_x0020_flux_x0020_d_x0027_: { name: 'Outils_x0020_flux_x0020_d_x0027_', type: 'multi'  },
  Outil_x0020_bancaire:             { name: 'Outils_x0020_bancaires',           type: 'multi'  },
  Outil_x0020_de_x0020_signature_x: { name: 'Outils_x0020_de_x0020_signature_', type: 'multi'  },
  Outils_x0020_client_x0020_PA:     { name: 'Outils_x0020_clients_x0020_PA',    type: 'multi'  },
  Outils_x0020_client_x0020_et_x00: { name: 'Outils_x0020_clients_x0020_Emiss', type: 'multi'  },
  Outils_x0020_client_x0020_r_x00e: { name: 'Outils_x0020_clients_x0020_R_x00', type: 'multi'  },
  Certification:                    { name: 'Certifications',                   type: 'multi'  },
  Sp_x00e9_cialit_x00e9__x0020_de_: { name: 'Sp_x00e9_cialit_x00e9__x0028_s_x', type: 'text-to-multi'  },
};

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
  if (!r.ok) throw new Error(`token http ${r.status}: ${(await r.text()).slice(0, 300)}`);
  const j = await r.json();
  return j.access_token;
}

async function findItemIdByTitle(token, title) {
  const needle = (title || '').trim().toLowerCase();
  let url = `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE_ID}/lists/${GRAPH_LIST_ID}/items?$top=100&$expand=fields($select=Title)`;
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!r.ok) throw new Error(`list items http ${r.status}: ${(await r.text()).slice(0, 300)}`);
    const j = await r.json();
    for (const it of (j.value || [])) {
      const t = ((it.fields && it.fields.Title) || '').trim().toLowerCase();
      if (t === needle) return parseInt(it.id, 10);
    }
    url = j['@odata.nextLink'] || null;
  }
  return null;
}

function buildPatchBody(data, fieldStates) {
  const out = {};
  for (const [formKey, m] of Object.entries(FIELD_MAP)) {
    if (formKey === 'Title') continue;
    const state = fieldStates && fieldStates[formKey];
    if (state !== 'modified') continue;
    let v = data[formKey];
    if (m.type === 'number') {
      v = (v === '' || v === null || v === undefined) ? null : Number(v);
      if (Number.isNaN(v)) continue;
      out[m.name] = v;
    } else if (m.type === 'multi') {
      // For "Autre" entries (prefix "__OTHER__:"): strip the marker, split on
      // common separators (", "/" et "/"&"), trim, uppercase each piece.
      // Existing predefined options are kept as-is.
      const splitOther = txt => txt
        .split(/\s*(?:,|\bet\b|&)\s*/i)
        .map(s => s.trim())
        .filter(Boolean)
        .map(s => s.toUpperCase());
      const arr = Array.isArray(v) ? v
        .flatMap(x => {
          if (typeof x !== 'string') return [x];
          if (!x.startsWith('__OTHER__:')) return [x];
          const t = x.slice(10).trim();
          return t ? splitOther(t) : [];
        })
        .filter(x => x != null && x !== '') : [];
      // Deduplicate while preserving order
      const seen = new Set();
      const dedup = arr.filter(x => {
        const k = typeof x === 'string' ? x : String(x);
        if (seen.has(k)) return false;
        seen.add(k);
        return true;
      });
      out[`${m.name}@odata.type`] = '#Collection(Edm.String)';
      out[m.name] = dedup;
    } else if (m.type === 'text-to-multi') {
      const s = (v === undefined || v === null) ? '' : String(v).trim();
      out[`${m.name}@odata.type`] = '#Collection(Edm.String)';
      out[m.name] = s ? [s] : [];
    } else {
      out[m.name] = (v === undefined || v === null) ? null : String(v);
    }
  }
  return out;
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
  if (!process.env.GH_TOKEN) {
    return res.status(500).json({ error: 'Server not configured (GH_TOKEN missing)' });
  }
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    return res.status(500).json({ error: 'Server not configured (MS_* env vars missing)' });
  }

  let payload;
  try { payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch { return res.status(400).json({ error: 'Invalid JSON' }); }

  const { issueNumber } = payload || {};
  if (!issueNumber) return res.status(400).json({ error: 'Missing issueNumber' });

  // 1) Fetch the issue
  const issueRes = await fetch(`https://api.github.com/repos/${REPO}/issues/${issueNumber}`, {
    headers: {
      'Authorization': `Bearer ${process.env.GH_TOKEN}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28'
    }
  });
  if (!issueRes.ok) {
    return res.status(502).json({ error: 'GitHub fetch issue failed' });
  }
  const issue = await issueRes.json();
  const cabinetTitle = (issue.title || '').replace(/^Soumission · /, '').trim();
  const match = (issue.body || '').match(/```json\n([\s\S]+?)\n```/);
  if (!match) return res.status(400).json({ error: 'No JSON block in issue body' });
  let parsed;
  try { parsed = JSON.parse(match[1]); }
  catch { return res.status(400).json({ error: 'Invalid JSON in issue body' }); }
  const { data, fieldStates } = parsed || {};
  if (!data || !fieldStates) return res.status(400).json({ error: 'Missing data/fieldStates' });

  // 2) Push to SharePoint via Graph
  let pushResult = { pushed: false };
  try {
    const token = await getGraphToken();
    const itemId = await findItemIdByTitle(token, cabinetTitle);
    if (!itemId) {
      pushResult = { pushed: false, error: `Cabinet "${cabinetTitle}" introuvable dans la liste Cabinets` };
    } else {
      const body = buildPatchBody(data, fieldStates);
      const fieldCount = Object.keys(body).filter(k => !k.endsWith('@odata.type')).length;
      if (fieldCount === 0) {
        pushResult = { pushed: false, info: 'Aucun champ marqué modifié — rien à pousser' };
      } else {
        const patchRes = await fetch(
          `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE_ID}/lists/${GRAPH_LIST_ID}/items/${itemId}/fields`,
          {
            method: 'PATCH',
            headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
            body: JSON.stringify(body)
          }
        );
        if (!patchRes.ok) {
          const txt = await patchRes.text();
          pushResult = { pushed: false, error: `Graph PATCH http ${patchRes.status}`, details: txt.slice(0, 400) };
        } else {
          pushResult = { pushed: true, itemId, fieldCount };
        }
      }
    }
  } catch (e) {
    pushResult = { pushed: false, error: e.message || String(e) };
  }

  // 3) If SP push failed, abort with details so the admin sees the issue
  if (!pushResult.pushed && pushResult.error) {
    return res.status(502).json({ error: 'SharePoint push failed', sharepoint: pushResult });
  }

  // 4) Label + close issue
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
  const closeRes = await fetch(`https://api.github.com/repos/${REPO}/issues/${issueNumber}`, {
    method: 'PATCH',
    headers: {
      'Authorization': `Bearer ${process.env.GH_TOKEN}`,
      'Accept': 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2022-11-28',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ state: 'closed', state_reason: 'completed' })
  });
  if (!closeRes.ok) {
    return res.status(502).json({ error: 'GitHub close issue failed', sharepoint: pushResult });
  }

  return res.status(200).json({ ok: true, sharepoint: pushResult });
}
