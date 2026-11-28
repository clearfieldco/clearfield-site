/**
 * Clearfield comps request handler (Google Apps Script web app)
 *
 * Owned by info@clearfieldcomps.com. The site's Netlify function posts every
 * comps request here. For each request it:
 *   1. puts a ready-to-send first reply (the match confirmation) in Drafts,
 *      with the recommended package and a prefilled payment link. Nothing is
 *      emailed to the buyer automatically: your reply is the first touch.
 *   2. logs one row in the "Comps requests" sheet.
 * An answer to the thanks-page question ("what do you most want to know")
 * is written into that person's row.
 *
 * Follow-ups (runFollowUps, every hour):
 *   When you send your offer email (any email to them from this account that
 *   contains a buy.stripe.com or invoice.stripe.com link), the row is marked
 *   "Confirmation sent" and the follow-ups start: 24 hours, 3 days and 7 days
 *   later, each sent as a reply in the same thread, only between 8 am and 9 pm.
 *   They stop for good when:
 *     - the person replies to you (you take it from there),
 *     - you send them another email yourself after the offer,
 *     - Stripe shows a completed checkout for their email (needs STRIPE_KEY),
 *     - you type anything in the "Paid" or "Outcome / objection" column.
 *   An offer without a Stripe link (e.g. a "can't match" email) starts nothing.
 *
 * Setup (about 10 minutes):
 *   1. Signed in as info@clearfieldcomps.com, go to script.google.com > New project.
 *      Name it "Clearfield comps request handler" and paste this whole file.
 *   2. Project Settings > Script properties > add SHARED_SECRET with the same
 *      value as DEMO_EMAIL1_SECRET in Netlify (or a new one, then add it to
 *      Netlify as COMPS_REQUEST_SECRET).
 *   3. Run setup() once from the editor and approve the permissions. It creates
 *      the "Comps requests" sheet in your Drive and saves its ID.
 *   4. Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 *      Copy the URL that ends in /exec.
 *   5. Netlify > Project configuration > Environment variables: add
 *      COMPS_REQUEST_URL = that URL. Then trigger a deploy.
 *   6. Send yourself a test request from the site.
 *
 * Follow-up setup (5 more minutes):
 *   7. In the editor: Services (+) > Gmail API > Add. Follow-ups use it so they
 *      land as replies in the buyer's thread.
 *   8. Stripe > Developers > API keys > Create restricted key. Name it
 *      "Comps follow-ups", give it Checkout Sessions: Read and nothing else.
 *      Add it here as the script property STRIPE_KEY. Without it, follow-ups
 *      still stop on a reply; mark payments by hand in the "Paid" column.
 *   9. Edit SIGNER below, then run installTrigger() once. To pause everything,
 *      run removeTrigger().
 */

const FROM_NAME = 'Clearfield Comps';
const SIGNATURE = '[Name]\nClearfield Comps analyst'; // replace [Name] with the first name you sign with

const SIGNER = 'The Clearfield Comps Team'; // signs the automatic follow-ups
const SEND_FROM_HOUR = 8, SEND_UNTIL_HOUR = 21;   // follow-ups only go out in this window (script time zone)
const FOLLOW_UP_HOURS = [24, 72, 168];            // after the offer email

const LINKS = {
  'One Comp': 'https://buy.stripe.com/aFafZj9ZLccKfzVedeeAg00',
  'Three Comps + Calculator': 'https://buy.stripe.com/bJebJ3b3PekS4Vhd9aeAg01',
};

const COLUMNS = ['Received', 'Name', 'Email', 'Website', 'Monthly revenue', 'What they sell',
  'Where they sell', 'Package chosen', 'Recommended', 'Wants to know', 'Confirmation sent',
  'Follow-up 1', 'Follow-up 2', 'Follow-up 3', 'Paid', 'Package paid', 'Amount', 'Outcome / objection'];

function setup() {
  const ss = SpreadsheetApp.create('Comps requests');
  const sh = ss.getSheets()[0];
  sh.setName('Requests');
  sh.appendRow(COLUMNS);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  Logger.log('Sheet created: ' + ss.getUrl());
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ ok: false, error: 'bad json' }); }
  const secret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
  if (!secret || d.secret !== secret) return json({ ok: false, error: 'forbidden' });
  const email = String(d.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json({ ok: false, error: 'bad email' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.type === 'detail') return json(saveDetail(email, clean(d.want_to_know, 600)));
    return json(handleRequest(email, d));
  } finally {
    lock.releaseLock();
  }
}

function handleRequest(email, d) {
  const r = {
    name: clean(d.name, 80), website: clean(d.website, 200), revenue: clean(d.monthly_revenue, 40),
    niche: clean(d.niche, 120), channel: clean(d.sales_channel, 80), pkg: clean(d.package, 40) || 'Not sure yet',
  };
  const first = (r.name.split(/\s+/)[0] || 'there');
  const rec = recommend(r.pkg, r.revenue);
  // 1. Draft your first reply. Nothing is sent to the buyer automatically.
  const subject = 'Your comps for ' + (r.website || 'your business');
  GmailApp.createDraft(email, subject, confirmation(first, r, rec, email), { name: FROM_NAME });

  // 2. Log it.
  sheet().appendRow([new Date(), r.name, email, r.website, r.revenue, r.niche, r.channel, r.pkg, rec,
    '', '', '', '', '', '', '', '', '']);
  return { ok: true };
}

function saveDetail(email, text) {
  if (!text) return { ok: true };
  const sh = sheet();
  const rows = sh.getDataRange().getValues();
  const emailCol = COLUMNS.indexOf('Email'), wantCol = COLUMNS.indexOf('Wants to know');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][emailCol]).toLowerCase() === email) {
      sh.getRange(i + 1, wantCol + 1).setValue(text);
      return { ok: true };
    }
  }
  return { ok: true, note: 'no matching row' };
}

// Which package to lead with: the one they picked, or by revenue band.
function recommend(pkg, revenue) {
  if (pkg && pkg !== 'Not sure yet') return pkg;
  if (/under/i.test(revenue)) return 'One Comp';
  return 'Three Comps + Calculator';
}

function payLink(pkg, email) {
  const base = LINKS[pkg];
  return base ? base + '?prefilled_email=' + encodeURIComponent(email) : '';
}

function confirmation(first, r, rec, email) {
  const big = /\$50K to \$100K|over/i.test(r.revenue);
  const lines = [
    'Hi ' + first + ',',
    '',
    'I looked at ' + (r.website || 'your site') + '. [One specific, true line about their business.]',
    '',
    "Good news: we have strong matches. Here's who I'd put next to you:",
    '',
    '- Comp 1: [model], [$X to $Y] a month, [N] years in, [sold / listed in 2026]',
  ];
  if (rec !== 'One Comp') lines.push('- Comp 2: [...]', '- Comp 3: [...]');
  lines.push('', '[One true hook from the comps, no exact figures.]', '');

  if (rec === 'Comps + Analyst') {
    lines.push(
      "For Comps + Analyst ($1,499), I'll send a Stripe invoice. You get the three comps and calculator within 24 hours of paying, then an analyst reviews your last 12 months, a 60-minute recorded call within 5 business days, your top 3 fixes ranked by dollars a month and a 90-day plan.",
      '', '[Invoice link]');
  } else if (rec === 'One Comp') {
    lines.push(
      "I'd start with One Comp ($199): your closest match, its full P&L in dollars a month and % of net sales.",
      '', 'Start here: ' + payLink('One Comp', email), '',
      'Pay by [time today] and your comp arrives by [same time tomorrow]. The 24 hours start at payment.', '',
      'If it\'s useful, the $199 counts toward Three Comps + Calculator within 14 days.');
  } else {
    lines.push(
      "I'd recommend Three Comps + Calculator ($399). You get all three P&Ls side by side, every line in dollars a month, and the calculator puts your own numbers next to theirs.",
      '', 'Start here: ' + payLink('Three Comps + Calculator', email), '',
      'Pay by [time today] and your comps arrive by [same time tomorrow]. The 24 hours start at payment.', '',
      'Want to start smaller? One Comp is $199 for your closest match, and the $199 counts toward Three Comps if you upgrade within 14 days: ' + payLink('One Comp', email));
    if (big) lines.push('', 'If you\'d rather have an analyst go through your own books with you and rank what to fix first, that\'s Comps + Analyst ($1,499). Happy to send details.');
  }
  lines.push('',
    "Every package has our match guarantee: if a comp doesn't match the niche, model, size or stage you gave us, tell us within 7 days and we'll replace it, or refund you if there's no closer match.",
    '', "I'll hold these matches for you for 7 days.", '', SIGNATURE);
  return lines.join('\n');
}

function describe(r) {
  const bits = [];
  if (r.niche) bits.push(r.niche.toLowerCase());
  if (r.channel) bits.push('selling mostly on ' + r.channel.replace(/ \(.*\)$/, '').replace(/^My own/, 'your own'));
  if (r.revenue) bits.push('at ' + r.revenue.replace(/^Under/, 'under').replace(/^Over/, 'over') + ' a month');
  return bits.length ? bits.join(', ') : 'your business';
}

// ---------- Follow-ups ----------

function installTrigger() {
  removeTrigger();
  ScriptApp.newTrigger('runFollowUps').timeBased().everyHours(1).create();
}

function removeTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'runFollowUps')
    .forEach(t => ScriptApp.deleteTrigger(t));
}

function runFollowUps() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const sh = sheet();
    const rows = sh.getDataRange().getValues();
    const c = name => COLUMNS.indexOf(name);
    const me = Session.getEffectiveUser().getEmail().toLowerCase();
    const now = new Date();
    const hour = Number(Utilities.formatDate(now, Session.getScriptTimeZone(), 'H'));
    const inWindow = hour >= SEND_FROM_HOUR && hour < SEND_UNTIL_HOUR;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const email = String(row[c('Email')]).toLowerCase();
      if (!email || row[c('Paid')] !== '' || row[c('Outcome / objection')] !== '') continue;
      if (row[c('Follow-up 3')] !== '') continue;
      const set = (name, v) => sh.getRange(i + 1, c(name) + 1).setValue(v);
      const received = new Date(row[c('Received')]);

      // 1. Has the offer gone out yet?
      if (!row[c('Confirmation sent')] && now - received > 30 * 864e5) continue; // no offer within 30 days: stop looking
      let offer = row[c('Confirmation sent')] ? null : findOffer(email, received);
      if (!row[c('Confirmation sent')]) {
        if (!offer) continue;
        set('Confirmation sent', offer.date);
        row[c('Confirmation sent')] = offer.date;
        PropertiesService.getScriptProperties().setProperty('offer:' + email, offer.threadId);
      }
      const sentAt = new Date(row[c('Confirmation sent')]);
      const threadId = PropertiesService.getScriptProperties().getProperty('offer:' + email);
      if (!threadId) continue;
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) continue;

      // 2. Stop if they paid, replied, or you wrote to them yourself.
      const paid = stripePaid(email, received);
      if (paid) { set('Paid', paid.date); set('Package paid', paid.pkg); set('Amount', paid.amount); continue; }
      const msgs = thread.getMessages().filter(m => !m.isDraft() && m.getDate() > sentAt);
      if (repliedSince(email, sentAt)) { set('Outcome / objection', 'Replied: follow-ups stopped'); continue; }
      const autoSent = ['Follow-up 1', 'Follow-up 2', 'Follow-up 3'].filter(n => row[c(n)] !== '').length;
      const mine = msgs.filter(m => m.getFrom().toLowerCase().indexOf(me) !== -1).length;
      if (mine > autoSent) { set('Outcome / objection', 'You wrote to them: follow-ups stopped'); continue; }

      // 3. Send the next follow-up if it's due.
      const next = autoSent; // 0, 1 or 2
      const due = new Date(sentAt.getTime() + FOLLOW_UP_HOURS[next] * 3600e3);
      if (now < due || !inWindow) continue;
      const first = String(row[c('Name')]).split(/\s+/)[0] || 'there';
      const offerMsg = thread.getMessages().find(m => !m.isDraft() && m.getDate().getTime() >= sentAt.getTime() - 60e3
        && /https:\/\/(buy|invoice)\.stripe\.com\//.test(m.getPlainBody())) || thread.getMessages().pop();
      const links = (offerMsg.getPlainBody().match(/https:\/\/(buy|invoice)\.stripe\.com\/[^\s)>\]]+/g) || []);
      const body = followUp(next, first, String(row[c('Website')]), links, email);
      sendInThread(thread, offerMsg, email, body);
      set(['Follow-up 1', 'Follow-up 2', 'Follow-up 3'][next], now);
      if (next === 2) set('Outcome / objection', 'No reply after 3 follow-ups');
    }
  } finally {
    lock.releaseLock();
  }
}

// Your offer: the first email you sent them after the request that has a Stripe link.
function findOffer(email, after) {
  const q = 'in:sent to:' + email + ' (buy.stripe.com OR invoice.stripe.com) after:' + Utilities.formatDate(after, 'UTC', 'yyyy/MM/dd');
  const threads = GmailApp.search(q, 0, 10);
  let best = null;
  threads.forEach(t => t.getMessages().forEach(m => {
    const to = (m.getTo() + ',' + m.getCc()).toLowerCase();
    if (m.getDate() < after || to.indexOf(email) === -1) return;
    if (!/https:\/\/(buy|invoice)\.stripe\.com\//.test(m.getPlainBody())) return;
    if (!best || m.getDate() < best.date) best = { date: m.getDate(), threadId: t.getId() };
  }));
  return best;
}

// Any email from them after your offer, in this thread or any other.
function repliedSince(email, since) {
  const q = 'from:' + email + ' after:' + Utilities.formatDate(since, 'UTC', 'yyyy/MM/dd');
  return GmailApp.search(q, 0, 10).some(t => t.getMessages().some(m =>
    !m.isDraft() && m.getDate() > since && m.getFrom().toLowerCase().indexOf(email) !== -1));
}

function stripePaid(email, after) {
  const key = PropertiesService.getScriptProperties().getProperty('STRIPE_KEY');
  if (!key) return null;
  const url = 'https://api.stripe.com/v1/checkout/sessions?status=complete&limit=5'
    + '&customer_details%5Bemail%5D=' + encodeURIComponent(email)
    + '&created%5Bgte%5D=' + Math.floor(after.getTime() / 1000);
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + key }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) { console.error('Stripe check failed', res.getResponseCode(), res.getContentText().slice(0, 200)); return null; }
  const s = (JSON.parse(res.getContentText()).data || [])[0];
  if (!s) return null;
  const sub = s.amount_subtotal || s.amount_total || 0;
  const pkg = sub === 19900 ? 'One Comp' : sub === 39900 ? 'Three Comps + Calculator' : sub === 149900 ? 'Comps + Analyst' : '';
  return { date: new Date(s.created * 1000), pkg: pkg, amount: (s.amount_total || 0) / 100 };
}

function followUp(n, first, website, links, email) {
  const main = links[0] || '';
  const one = links.find(l => l.indexOf(LINKS['One Comp']) === 0) || (LINKS['One Comp'] + '?prefilled_email=' + encodeURIComponent(email));
  const ledWithOne = main.indexOf(LINKS['One Comp']) === 0;
  if (n === 0) return [
    'Hi ' + first + ',', '',
    "Quick one: is there a specific number you're trying to figure out? Ad spend, product cost, what to pay yourself? Tell me and I'll make sure the comps answer it.",
    '', main ? 'Your link, if you\'re ready: ' + main : '', '', SIGNER].join('\n');
  if (n === 1) {
    const lines = ['Hi ' + first + ',', '',
      'The question I get most: "How do I know these are real?"', '',
      "Every comp is an actual business that has sold or been listed for sale, pulled from licensed data and partners who share with permission. Nothing is modeled or averaged. We can't show names (our agreements require that), but you get every line of their books. And if a comp doesn't match what you told us, we replace it.", ''];
    if (ledWithOne || !main || main.indexOf('invoice.stripe.com') !== -1) lines.push('Your link: ' + (main || one));
    else lines.push("If $399 feels like a lot before you've seen one, start with your single closest match for $199. If it's useful, the $199 counts toward the other two: " + one);
    lines.push('', SIGNER);
    return lines.join('\n');
  }
  return ['Hi ' + first + ',', '',
    "I haven't heard back, so I'll release the matches I set aside for " + (website || 'you') + ' today.', '',
    'If the timing\'s off, no problem. Reply "later" and I\'ll check back in a month with fresh comps. If something about the offer didn\'t fit, I\'d honestly like to know what.',
    '', SIGNER].join('\n');
}

// Sends as a real reply in the buyer's thread (Gmail API advanced service).
function sendInThread(thread, offerMsg, email, body) {
  const subject = /^re:/i.test(thread.getFirstMessageSubject()) ? thread.getFirstMessageSubject() : 'Re: ' + thread.getFirstMessageSubject();
  const msgId = offerMsg.getHeader('Message-ID');
  const mime = [
    'From: ' + FROM_NAME + ' <' + Session.getEffectiveUser().getEmail() + '>',
    'To: ' + email,
    'Subject: =?UTF-8?B?' + Utilities.base64Encode(subject, Utilities.Charset.UTF_8) + '?=',
    msgId ? 'In-Reply-To: ' + msgId : '',
    msgId ? 'References: ' + msgId : '',
    'Content-Type: text/plain; charset=UTF-8',
    '', body].filter((l, i) => l !== '' || i > 5).join('\r\n');
  const raw = Utilities.base64EncodeWebSafe(Utilities.newBlob(mime).getBytes());
  Gmail.Users.Messages.send({ raw: raw, threadId: thread.getId() }, 'me');
}

function sheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Run setup() first');
  return SpreadsheetApp.openById(id).getSheetByName('Requests');
}

function clean(v, max) { return String(v || '').replace(/[\r\n]+/g, ' ').trim().slice(0, max); }

function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
