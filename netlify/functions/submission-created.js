// Runs automatically after every verified (non-spam) Netlify Forms submission.
// For the "demo-request" form it adds the person to Kit and tags them. The Kit
// automation on that tag sends the private demo link and starts the follow-up
// emails. Nothing secret lives in this file: the key is a Netlify environment
// variable.
//
// Netlify environment variables (Project configuration > Environment variables):
//   KIT_API_KEY  Kit v4 API key (Kit > Settings > Developer > API keys)
//   KIT_TAG_ID   numeric ID of the "Terminal demo" tag in Kit

const KIT = 'https://api.kit.com/v4';

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

  const key = process.env.KIT_API_KEY;
  const tagId = process.env.KIT_TAG_ID;
  if (!key || !tagId) {
    console.warn('KIT_API_KEY or KIT_TAG_ID is not set. Lead saved in Netlify Forms only.');
    return { statusCode: 200, body: 'Kit not configured' };
  }

  const d = payload.data || {};
  const email = String(d.email || '').trim().toLowerCase();
  if (!email) return { statusCode: 200, body: 'No email' };

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

  // Tagging starts the Kit automation that emails the demo link.
  const t = await kit('/tags/' + encodeURIComponent(tagId) + '/subscribers', { email_address: email }, key);
  if (!t.ok) {
    console.error('Kit tag failed', t.status, t.text.slice(0, 300));
    return { statusCode: 502, body: 'Kit tag error' };
  }
  return { statusCode: 200, body: 'ok' };
};
