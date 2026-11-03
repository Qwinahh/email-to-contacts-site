/* Email to Contacts for Outlook: the task pane beside the open email. Fills a form from the sender's
   signature; Save signs in to Google (once per session) and saves the contact to Google Contacts,
   which the person's phone shows. All text is written with textContent or input values, never as HTML. */
(function () {
  'use strict';

  const Core = window.EmailToContacts;
  const Google = window.EmailToContactsGoogle;
  const $ = (id) => document.getElementById(id);
  let auth = null;
  let phoneLabels = [];

  function setStatus(text, isError) {
    $('status').textContent = text;
    $('status').className = 'status' + (isError ? ' error' : '');
    $('status').hidden = !text;
  }

  function readBody(item) {
    return new Promise((resolve) => {
      item.body.getAsync(Office.CoercionType.Text, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value : ''));
    });
  }

  function render(d) {
    $('done').hidden = true;
    $('note').textContent = d.notes;
    $('note').hidden = !d.notes;
    $('name').value = d.name;
    $('email').value = d.email;
    $('title').value = d.title;
    $('company').value = d.company;
    const phones = d.phones.length ? d.phones : [{ label: 'mobile', number: '' }];
    phoneLabels = phones.map((p) => p.label || 'other');
    const box = $('phones');
    box.textContent = '';
    phones.forEach((p, i) => {
      const label = document.createElement('label');
      label.textContent = Core.PHONE_LABELS[phoneLabels[i]] || 'Phone';
      const input = document.createElement('input');
      input.id = 'phone' + i;
      input.type = 'tel';
      input.autocomplete = 'off';
      input.value = p.number;
      label.appendChild(input);
      box.appendChild(label);
    });
    $('save').disabled = false;
    $('details').hidden = false;
    setStatus('');
  }

  function formDetails() {
    const v = (id) => $(id).value.trim();
    return {
      name: v('name'),
      email: v('email'),
      phones: phoneLabels.map((label, i) => ({ label: label, number: v('phone' + i) })).filter((p) => p.number),
      title: v('title'),
      company: v('company'),
    };
  }

  function showDone(name, message, link) {
    $('details').hidden = true;
    $('done-name').textContent = name;
    $('done-message').textContent = message;
    $('done-link').hidden = !link;
    if (link) $('done-link').href = link;
    $('done').hidden = false;
    setStatus('');
  }

  function contacts(token) {
    return Google.googleClient(token, fetch.bind(window), () => auth.forget());
  }

  async function load() {
    const item = Office.context.mailbox.item;
    if (!item) return setStatus('Open an email to save its sender.');
    setStatus('Reading the email…');
    const me = Office.context.mailbox.userProfile.emailAddress;
    const details = Core.readDetails(item.from || item.sender, await readBody(item), me, window.parseEmail);
    render(details);
    // Say so if they're already a contact, but only if Google sign-in is already available.
    try {
      const token = await auth.getToken(false);
      if (token && details.email) {
        const existing = await Google.findContact(contacts(token), details.email);
        if (existing) showDone(details.name || details.email, 'Already in your contacts.', Google.contactLink(existing));
      }
    } catch (e) { /* the form still works; Save will sign in */ }
  }

  async function save(event) {
    event.preventDefault();
    const d = formDetails();
    $('save').disabled = true;
    setStatus('Saving…');
    try {
      if (!d.name && !d.email) {
        $('save').disabled = false;
        return setStatus('Add a name or an email first.', true);
      }
      const token = await auth.getToken(true);
      const result = await Google.saveContact(contacts(token), d);
      showDone(d.name || d.email, result.status === 'exists'
        ? 'Already in your contacts.'
        : 'Saved to your Google Contacts. It will appear on your phone shortly.', result.link);
    } catch (err) {
      $('save').disabled = false;
      setStatus('Could not save: ' + err.message, true);
    }
  }

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Outlook) return;
    auth = Google.createGoogleAuth({
      Office: Office,
      dialogUrl: new URL('google-dialog.html', location.href).href,
      session: window.sessionStorage,
      local: window.localStorage,
      fetchImpl: fetch.bind(window),
    });
    $('details').addEventListener('submit', save);
    // When the pane is pinned, follow the reader to the next email.
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => load());
    load();
  });
})();
