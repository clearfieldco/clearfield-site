// Runs automatically after every verified (non-spam) Netlify Forms submission.
// For the "demo-request" form it adds the person to Kit and tags them. The Kit
// automation on that tag sends the private demo link and starts the follow-up
// emails. Nothing secret lives in this file: the key is a Netlify environment
// variable.
//
// How it reaches Kit:
//   1. Always: subscribes the person through the public "Comps demo signup" Kit
//      form (no API key needed). The Kit automation starts on "joins form".
//      Note: Kit answers "success" but silently drops form sign-ups sent from
//      servers, so the demo forms also join the Kit form from the visitor's
//      browser (window.kitJoin in index.html and demo.html). That browser call
//      is what actually starts the emails; this step is a backup.
//   2. Optional: if KIT_API_KEY and KIT_TAG_ID are set in Netlify, it also saves
//      the business details as custom fields and applies the tag.
//
// Netlify environment variables (Project configuration > Environment variables):
//   KIT_FORM_ID  optional, Kit form ID (defaults to the Comps demo signup form)
//   KIT_API_KEY  optional, Kit v4 API key (Kit > Settings > Developer > API keys)
//   KIT_TAG_ID   optional, numeric ID of the "Comps demo" tag in Kit

const KIT = 'https://api.kit.com/v4';
const DEFAULT_FORM_ID = '10251951'; // Kit form "Comps demo signup"

async function joinForm(formId, email, firstName) {
  const body = new URLSearchParams({ email_address: email });
  if (firstName) body.set('fields[first_name]', firstName);
  const res = await fetch('https://app.kit.com/forms/' + encodeURIComponent(formId) + '/subscriptions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded', Accept: 'application/json' },
    body: body.toString(),
  });
  const text = await res.text();
  return { ok: res.ok, status: res.status, text };
}

async function kit(path, body, key) {
  const res = await fetch(KIT + path, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json', 'X-Kit-Api-Key': key },
    body: JSON.stringify(body),
  });
  const text = await res.text();
  return { ok: res.ok, status: res.status, text };
}

exports.handler = async (event) => {
  let payload;
  try {
    payload = JSON.parse(event.body || '{}').payload || {};
  } catch (e) {
    return { statusCode: 400, body: 'Bad payload' };
  }
  if (payload.form_name !== 'demo-request') return { statusCode: 200, body: 'Not a demo request' };

  const d = payload.data || {};
  const email = String(d.email || '').trim().toLowerCase();
  if (!email) return { statusCode: 200, body: 'No email' };
  const firstName = String(d.first_name || '').trim();

  // Joining the Kit form starts the automation that emails the demo link.
  const f = await joinForm(process.env.KIT_FORM_ID || DEFAULT_FORM_ID, email, firstName);
  if (!f.ok) {
    console.error('Kit form subscribe failed', f.status, f.text.slice(0, 300));
    return { statusCode: 502, body: 'Kit form error' };
  }

  const key = process.env.KIT_API_KEY;
  const tagId = process.env.KIT_TAG_ID;
  if (!key || !tagId) return { statusCode: 200, body: 'ok (form only)' };

  const fields = {
    company: d.company || '',
    website: d.website || '',
    niche: d.niche || '',
    monthly_revenue: d.monthly_revenue || '',
    role: d.role || '',
    goal: d.goal || '',
    lead_source: d.source || '',
  };

  // Create or update the subscriber with their business details.
  let r = await kit('/subscribers', { email_address: email, first_name: d.first_name || '', state: 'active', fields }, key);
  if (!r.ok && r.status === 422) {
    // Usually a custom field that doesn't exist in Kit yet. Save the person anyway.
    console.warn('Kit rejected custom fields, retrying without them:', r.text.slice(0, 300));
    r = await kit('/subscribers', { email_address: email, first_name: d.first_name || '', state: 'active' }, key);
  }
  if (!r.ok) {
    console.error('Kit subscriber create failed', r.status, r.text.slice(0, 300));
    return { statusCode: 502, body: 'Kit error' };
  }

  // Tag as well, for segmenting.
  const t = await kit('/tags/' + encodeURIComponent(tagId) + '/subscribers', { email_address: email }, key);
  if (!t.ok) {
    console.error('Kit tag failed', t.status, t.text.slice(0, 300));
    return { statusCode: 502, body: 'Kit tag error' };
  }
  return { statusCode: 200, body: 'ok' };
};
