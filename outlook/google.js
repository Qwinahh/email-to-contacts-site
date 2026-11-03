/* Email to Contacts for Outlook: signing in to Google and saving to Google Contacts.
   Saving to Google rather than Microsoft needs no Microsoft app registration; the person's phone
   shows Google contacts (Android always, iPhone with Contacts on for the Google account).
   Sign-in happens in Outlook's dialog window (google-dialog.html), which hands the token back.
   The token is kept for the browser session only; the chosen account's address is remembered so
   later sign-ins don't ask which account to use. */
(function (root) {
  'use strict';

  const PEOPLE = 'https://people.googleapis.com/v1';
  const USERINFO = 'https://www.googleapis.com/oauth2/v3/userinfo';
  const TYPES = { mobile: 'mobile', work: 'work', home: 'home', fax: 'workFax', other: 'other' };

  // ----- Sign-in -----

  function createGoogleAuth(env) {
    // env: { Office, dialogUrl, session (sessionStorage-like), local (localStorage-like), fetchImpl }
    const read = (store, key) => { try { return JSON.parse(store.getItem(key) || 'null'); } catch (e) { return null; } };
    const write = (store, key, value) => { try { store.setItem(key, JSON.stringify(value)); } catch (e) { /* private mode */ } };

    function cached() {
      const t = read(env.session, 'googleToken');
      return t && t.expiresAt - Date.now() > 60000 ? t.token : null;
    }

    function fromDialog() {
      const hint = read(env.local, 'googleAccount');
      const url = env.dialogUrl + (hint ? '?hint=' + encodeURIComponent(hint) : '');
      return new Promise((resolve, reject) => {
        env.Office.context.ui.displayDialogAsync(url, { height: 70, width: 40, promptBeforeOpen: false }, (res) => {
          if (res.status !== 'succeeded') return reject(new Error('The Google sign-in window could not open: ' + ((res.error && res.error.message) || res.status)));
          const dialog = res.value;
          dialog.addEventHandler(env.Office.EventType.DialogMessageReceived, (arg) => {
            dialog.close();
            let msg = {};
            try { msg = JSON.parse(arg.message); } catch (e) { /* handled below */ }
            if (!msg.accessToken) return reject(new Error(msg.error || 'Google sign-in did not finish.'));
            write(env.session, 'googleToken', { token: msg.accessToken, expiresAt: Date.now() + (Number(msg.expiresIn) || 3600) * 1000 });
            resolve(msg.accessToken);
          });
          dialog.addEventHandler(env.Office.EventType.DialogEventReceived, () => reject(new Error('The Google sign-in window was closed. Click Save to try again.')));
        });
      });
    }

    async function rememberAccount(token) {
      if (read(env.local, 'googleAccount')) return;
      try {
        const res = await env.fetchImpl(USERINFO, { headers: { Authorization: 'Bearer ' + token } });
        const info = res.ok ? await res.json() : {};
        if (info.email) write(env.local, 'googleAccount', String(info.email));
      } catch (e) { /* only a convenience */ }
    }

    // interactive false: only a token already available (used when the pane opens).
    async function getToken(interactive) {
      const token = cached();
      if (token || !interactive) return token;
      const fresh = await fromDialog();
      await rememberAccount(fresh);
      return fresh;
    }

    function forget() {
      try { env.session.removeItem('googleToken'); } catch (e) { /* ignore */ }
    }

    return { getToken, forget };
  }

  // The URL the dialog sends people to. The dialog page calls this with its own address.
  function signInUrl(clientId, redirectUri, hint) {
    const params = {
      client_id: clientId,
      redirect_uri: redirectUri,
      response_type: 'token',
      scope: 'https://www.googleapis.com/auth/contacts email',
      include_granted_scopes: 'true',
    };
    if (hint) params.login_hint = hint;
    else params.prompt = 'select_account';
    return 'https://accounts.google.com/o/oauth2/v2/auth?' + new URLSearchParams(params);
  }

  // ----- Google Contacts -----

  function toGooglePerson(d) {
    const clean = (v) => String(v == null ? '' : v).trim();
    const words = clean(d.name).split(/\s+/).filter(Boolean);
    const person = {};
    if (words.length) person.names = [words.length > 1 ? { givenName: words.slice(0, -1).join(' '), familyName: words[words.length - 1] } : { givenName: words[0] }];
    if (clean(d.email)) person.emailAddresses = [{ value: clean(d.email), type: 'work' }];
    const phones = (d.phones || []).filter((p) => clean(p.number)).map((p) => ({ value: clean(p.number), type: TYPES[clean(p.label).toLowerCase()] || 'other' }));
    if (phones.length) person.phoneNumbers = phones;
    if (clean(d.company) || clean(d.title)) {
      const org = {};
      if (clean(d.company)) org.name = clean(d.company);
      if (clean(d.title)) org.title = clean(d.title);
      person.organizations = [org];
    }
    return person;
  }

  function contactLink(resourceName) {
    return 'https://contacts.google.com/person/' + String(resourceName || '').replace(/^people\//, '');
  }

  function googleClient(token, fetchImpl, onUnauthorised) {
    return async function call(method, path, body) {
      const res = await fetchImpl(PEOPLE + path, {
        method: method,
        headers: Object.assign({ Authorization: 'Bearer ' + token }, body ? { 'Content-Type': 'application/json' } : {}),
        body: body ? JSON.stringify(body) : undefined,
      });
      const json = await res.json().catch(() => ({}));
      if (!res.ok) {
        if (res.status === 401 && onUnauthorised) onUnauthorised();
        const err = new Error((json.error && json.error.message) || 'Google Contacts error ' + res.status);
        err.status = res.status;
        throw err;
      }
      return json;
    };
  }

  // The existing contact's resource name, or null. The quick search can miss contacts until its
  // cache warms up, so the full list is checked when it finds nothing.
  async function findContact(call, email) {
    email = String(email || '').trim().toLowerCase();
    if (!email) return null;
    const has = (p) => (p.emailAddresses || []).some((e) => String(e.value).toLowerCase() === email);
    await call('GET', '/people:searchContacts?query=&readMask=emailAddresses');
    const found = await call('GET', '/people:searchContacts?' + new URLSearchParams({ query: email, readMask: 'emailAddresses' }));
    const hit = (found.results || []).find((r) => has(r.person));
    if (hit) return hit.person.resourceName;
    let pageToken = '';
    do {
      const page = await call('GET', '/people/me/connections?' + new URLSearchParams({ personFields: 'emailAddresses', pageSize: '1000', pageToken: pageToken }));
      const match = (page.connections || []).find(has);
      if (match) return match.resourceName;
      pageToken = page.nextPageToken || '';
    } while (pageToken);
    return null;
  }

  // Save the edited details. Returns { status: 'saved' | 'exists' | 'missing', link }.
  async function saveContact(call, d) {
    if (!String(d.name || '').trim() && !String(d.email || '').trim()) return { status: 'missing' };
    const existing = await findContact(call, d.email);
    if (existing) return { status: 'exists', link: contactLink(existing) };
    const created = await call('POST', '/people:createContact', toGooglePerson(d));
    return { status: 'saved', link: contactLink(created.resourceName) };
  }

  const api = { createGoogleAuth, signInUrl, toGooglePerson, contactLink, googleClient, findContact, saveContact };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.EmailToContactsGoogle = api;
})(typeof globalThis === 'object' ? globalThis : this);
