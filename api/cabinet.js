// POST /api/cabinet
// Body: { id, token }
//
// Returns the cabinet data for a specific id if the provided token matches
// expectedToken(id, title). Used by the end-user form to resolve cabinets
// that are not (yet) in the static embed baked into index.html — typically
// cabinets recently added in SharePoint.
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

// Cabinets InternalName → form-side key (mirror of validate.js FIELD_MAP)
const CAB_TO_FORM = {
  Title:                              'Title',
  Offre_x0020_Label:                  'OFFRE',
  Adresse_x0020_Si_x00e8_ge:          'Adresse',
  Pr_x00e9_nom_x0020_et_x0020_Nom_:   'Pr_x00e9_nomdur_x00e9_f_x00e9_re',
  Pr_x00e9_nom_x0020_et_x0020_Nom_0:  'Pr_x00e9_nomNomassoci_x00e9_',
  Adresse_x0020_mail_x0020_r_x00e9:   'Adressemailr_x00e9_f_x00e9_rent',
  Effectif_x0020_total:               'NBCollabs',
  Nb_x0020_Associ_x00e9_s:            'NBassoci_x00e9_s',
  Nb_x0020_EC_x0020_hors_x0020_ass:   'EChorsassoci_x00e9_s',
  Nb_x0020_Sites_x0020_ou_x0020_Ag:   'NBSites',
  Adresses_x0020_Sites:               'CodespostauxSites',
  Outils_x0020_Prod_x0020_Compta:     'Outils_x0020_Prod_x0020_Compta',
  Outils_x0020_Prod_x0020_Social:     'Outils_x0020_Prod_x0020_Social',
  Outils_x0020_Prod_x0020_Juridiqu:   'Outils_x0020_Prod_x0020_Juridiqu',
  Outils_x0020_Conseil:               'Outils_x0020_Conseil',
  Outils_x0020_flux_x0020_d_x0027_:   'Outils_x0020_flux_x0020_d_x0027_',
  Outils_x0020_bancaires:             'Outil_x0020_bancaire',
  Outils_x0020_de_x0020_signature_:   'Outil_x0020_de_x0020_signature_x',
  Outils_x0020_clients_x0020_PA:      'Outils_x0020_client_x0020_PA',
  Outils_x0020_clients_x0020_Emiss:   'Outils_x0020_client_x0020_et_x00',
  Outils_x0020_clients_x0020_R_x00:   'Outils_x0020_client_x0020_r_x00e',
  Certifications:                     'Certification',
  Sp_x00e9_cialit_x00e9__x0028_s_x:   'Sp_x00e9_cialit_x00e9__x0020_de_',
};

export default async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.status(204).end();
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  let payload;
  try { payload = typeof req.body === 'string' ? JSON.parse(req.body) : req.body; }
  catch { return res.status(400).json({ error: 'Invalid JSON' }); }

  const { id, token } = payload || {};
  if (!id || !token) return res.status(400).json({ error: 'Missing id or token' });
  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) return res.status(500).json({ error: 'MS_* env vars missing' });

  try {
    const gToken = await getGraphToken();
    const r = await fetch(
      `https://graph.microsoft.com/v1.0/sites/${GRAPH_SITE_ID}/lists/${GRAPH_LIST_ID}/items/${id}?$expand=fields`,
      { headers: { Authorization: `Bearer ${gToken}` } }
    );
    if (r.status === 404) return res.status(404).json({ error: 'Cabinet not found' });
    if (!r.ok) return res.status(502).json({ error: `SP fetch http ${r.status}` });
    const j = await r.json();
    const f = j.fields || {};
    const title = f.Title;
    if (!title) return res.status(404).json({ error: 'Cabinet has no title' });

    const expected = expectedToken(id, title);
    if (expected !== token) return res.status(403).json({ error: 'Invalid token' });

    // Map SP fields → form-side keys
    const cabinet = { Id: parseInt(id, 10), _token: token };
    for (const [spKey, formKey] of Object.entries(CAB_TO_FORM)) {
      cabinet[formKey] = f[spKey] !== undefined ? f[spKey] : null;
    }
    return res.status(200).json({ cabinet });
  } catch (e) {
    return res.status(502).json({ error: 'SP lookup failed', details: e.message || String(e) });
  }
}
