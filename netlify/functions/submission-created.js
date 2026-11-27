// Runs automatically after every verified (non-spam) Netlify Forms submission.
// For the "demo-request" form it emails the private demo link (email 1) from
// Gmail, then adds the person to Kit and tags them. The Kit automation starts the
// follow-up emails (2 to 5). Nothing secret lives in this file: keys are Netlify environment
// variables.
//
// How it reaches Kit:
//   1. Always: subscribes the person through the public "Comps demo signup" Kit
//      form (no API key needed). The Kit automation starts on "joins form".
//      Note: Kit quarantines form sign-ups sent from servers (it answers 200
//      with status "quarantined"), so the demo forms join the Kit form from the
//      visitor's browser (window.kitJoin in index.html and demo.html). That
//      browser call is what starts the emails; this step is only a backup.
//   2. Optional: if KIT_API_KEY and KIT_TAG_ID are set in Netlify, it also saves
//      the business details as custom fields and applies the tag.
//
// Email 1 (the private demo link) is sent from Google Workspace, not Kit, so it
// arrives as a normal one-to-one email from info@clearfieldcomps.com and is
// more likely to land in the Gmail Primary tab. A Google Apps Script web app
// ("Clearfield demo email 1 sender", owned by info@clearfieldcomps.com) does the
// sending. Kit still sends emails 2 to 5; email 1 is unpublished in Kit.
//
// Netlify environment variables (Project configuration > Environment variables):
//   DEMO_EMAIL1_URL     Apps Script web app URL (ends in /exec)
//   DEMO_EMAIL1_SECRET  shared secret that matches SHARED_SECRET in the script
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

async function sendEmail1(email, firstName) {
  const url = process.env.DEMO_EMAIL1_URL;
  const secret = process.env.DEMO_EMAIL1_SECRET;
  if (!url || !secret) {
    console.error('DEMO_EMAIL1_URL or DEMO_EMAIL1_SECRET is not set; demo email 1 was not sent.');
    return;
  }
  try {
    // Apps Script answers with a redirect to the result; fetch follows it.
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ secret, email, first_name: firstName }),
    });
    const text = await res.text();
    if (!res.ok || text.indexOf('"ok":true') === -1) {
      console.error('Demo email 1 send failed', res.status, text.slice(0, 300));
    }
  } catch (e) {
    console.error('Demo email 1 send error', e && e.message);
  }
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

  // Send the private demo link from Gmail first, so a Kit problem can't block it.
  await sendEmail1(email, firstName);

  // Joining the Kit form starts the automation that emails the demo link.
  const f = await joinForm(process.env.KIT_FORM_ID || DEFAULT_FORM_ID, email, firstName);
  if (f.ok && f.text.indexOf('"quarantined"') !== -1) {
    console.warn('Kit quarantined the server-side form sign-up; the browser sign-up is the one that counts.');
  }
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
