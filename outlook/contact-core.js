/* Email to Contacts for Outlook: reading the open email and saving its sender to Microsoft contacts.
   No DOM and no Office.js here, so it can be tested. The signature reader is parser.js (parseEmail). */
(function (root) {
  'use strict';

  const GRAPH = 'https://graph.microsoft.com/v1.0';
  const PHONE_LABELS = { mobile: 'Mobile', work: 'Work phone', home: 'Home phone', fax: 'Fax', other: 'Phone' };

  // from: Outlook's { displayName, emailAddress } for the sender.
  function readDetails(from, body, me, parseEmail) {
    from = from || {};
    const name = String(from.displayName || '').replace(/["<>]/g, '').trim();
    const email = String(from.emailAddress || '').trim().toLowerCase();
    const fromLine = email ? (name ? '"' + name + '" ' : '') + '<' + email + '>' : name;
    const card = parseEmail('From: ' + fromLine + '\n\n' + (body || ''));
    return {
      name: card.name || name,
      email: card.email || email,
      phones: (card.phones || []).slice(0, 3),
      title: card.title || '',
      company: card.company || '',
      notes: noteText(card.warnings || [], !!email && email === String(me || '').toLowerCase()),
    };
  }

  // Only the reader's warnings that tell the user something to check before saving.
  function noteText(warnings, ownEmail) {
    if (ownEmail) return 'This email is from you.';
    if (warnings.some((w) => /^Signature name differs/.test(w))) return "The signature names someone other than the sender, so only the sender's name and email are filled in.";
    if (warnings.some((w) => /^No phone number found/.test(w))) return 'No phone number found in the signature. You can type one in.';
    if (warnings.some((w) => /shared or automated address/.test(w))) return 'This looks like a shared or automated address.';
    return '';
  }

  // A Microsoft Graph contact. Graph has one mobile number and lists of business and home numbers;
  // it has no fax field, so a fax number is left out.
  function toGraphContact(d) {
    const clean = (v) => String(v == null ? '' : v).trim();
    const words = clean(d.name).split(/\s+/).filter(Boolean);
    const contact = {};
    if (words.length) {
      contact.displayName = words.join(' ');
      contact.givenName = words.length > 1 ? words.slice(0, -1).join(' ') : words[0];
      if (words.length > 1) contact.surname = words[words.length - 1];
    }
    if (clean(d.email)) contact.emailAddresses = [{ address: clean(d.email), name: contact.displayName || clean(d.email) }];
    for (const p of d.phones || []) {
      const number = clean(p.number);
      const label = clean(p.label).toLowerCase();
      if (!number || label === 'fax') continue;
      if (label === 'mobile' && !contact.mobilePhone) contact.mobilePhone = number;
      else if (label === 'home') (contact.homePhones = contact.homePhones || []).push(number);
      else (contact.businessPhones = contact.businessPhones || []).push(number);
    }
    if (clean(d.title)) contact.jobTitle = clean(d.title);
    if (clean(d.company)) contact.companyName = clean(d.company);
    return contact;
  }

  // A small Microsoft Graph client. fetchImpl: window.fetch, or a fake in tests.
  function graphClient(token, fetchImpl) {
    async function call(method, url, body) {
      const res = await fetchImpl(url, {
        method: method,
        headers: Object.assign({ Authorization: 'Bearer ' + token }, body ? { 'Content-Type': 'application/json' } : {}),
        body: body ? JSON.stringify(body) : undefined,
      });
      const text = await res.text();
      const json = text ? JSON.parse(text) : {};
      if (!res.ok) {
        const err = new Error((json.error && json.error.message) || 'Microsoft Graph error ' + res.status);
        err.status = res.status;
        throw err;
      }
      return json;
    }
    return { get: (url) => call('GET', url), post: (url, body) => call('POST', url, body) };
  }

  // The existing contact with this email address, or null. Tries a direct search first; if the
  // mailbox doesn't support that search, reads through the contact list instead.
  async function findExisting(graph, email) {
    email = String(email || '').trim().toLowerCase();
    if (!email) return null;
    const has = (c) => (c.emailAddresses || []).some((e) => String(e.address).toLowerCase() === email);
    try {
      const filter = "emailAddresses/any(a:a/address eq '" + email.replace(/'/g, "''") + "')";
      const res = await graph.get(GRAPH + '/me/contacts?$filter=' + encodeURIComponent(filter) + '&$select=id,displayName,emailAddresses&$top=5');
      return (res.value || []).find(has) || null;
    } catch (err) {
      if (err.status !== 400 && err.status !== 501) throw err;
    }
    let url = GRAPH + '/me/contacts?$select=id,displayName,emailAddresses&$top=500';
    while (url) {
      const page = await graph.get(url);
      const hit = (page.value || []).find(has);
      if (hit) return hit;
      url = page['@odata.nextLink'] || null;
    }
    return null;
  }

  // Save the edited details. Returns { status: 'saved' | 'exists' | 'missing', contact }.
  async function saveContact(graph, d) {
    if (!String(d.name || '').trim() && !String(d.email || '').trim()) return { status: 'missing' };
    const existing = await findExisting(graph, d.email);
    if (existing) return { status: 'exists', contact: existing };
    return { status: 'saved', contact: await graph.post(GRAPH + '/me/contacts', toGraphContact(d)) };
  }

  const api = { readDetails, noteText, toGraphContact, graphClient, findExisting, saveContact, PHONE_LABELS };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.EmailToContacts = api;
})(typeof globalThis === 'object' ? globalThis : this);
