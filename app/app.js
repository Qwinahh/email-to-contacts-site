/* Email to Contacts web app: the page. Paste or share email text, check the card, save as a contact. */
(function () {
  'use strict';
  const $ = (id) => document.getElementById(id);
  const Card = window.ContactCard;
  const FIELDS = ['name', 'phone', 'phone2', 'email', 'title', 'company'];

  function setStatus(text, isError) {
    $('status').textContent = text;
    $('status').className = 'status' + (isError ? ' error' : '');
    $('status').hidden = !text;
  }

  function show(text) {
    fill(Card.readText(text, window.parseEmail));
  }

  function fill(c) {
    $('name').value = c.name;
    $('phone').value = (c.phones[0] || {}).number || '';
    $('phone2').value = (c.phones[1] || {}).number || '';
    $('email').value = c.email;
    $('title').value = c.title;
    $('company').value = c.company;
    $('phone').dataset.label = (c.phones[0] || {}).label || 'mobile';
    $('phone2').dataset.label = (c.phones[1] || {}).label || 'work';
    $('card').hidden = false;
    setStatus(c.name || c.email || c.phones.length ? '' : 'Nothing found to save. Check the text, or type the details in.');
  }

  function current() {
    const v = (id) => $(id).value.trim();
    return {
      name: v('name'), email: v('email'), title: v('title'), company: v('company'),
      phones: [{ label: $('phone').dataset.label || 'mobile', number: v('phone') }, { label: $('phone2').dataset.label || 'work', number: v('phone2') }].filter((p) => p.number),
    };
  }

  // iPhone and iPad open a vCard link as a "Create New Contact" screen; other phones download the
  // file, and opening it adds the contact.
  function save(card) {
    const vcard = Card.buildVCard(card);
    const url = URL.createObjectURL(new Blob([vcard], { type: 'text/vcard' }));
    const isApple = /iPhone|iPad|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    if (isApple) {
      window.location.href = url;
      return 'If nothing opened, tap Save to contacts again.';
    }
    const a = document.createElement('a');
    a.href = url;
    a.download = Card.fileName(card);
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
    return 'Saved as a contact file. Open it (from the notification or Downloads) to add the contact.';
  }

  $('paste').addEventListener('click', async () => {
    try {
      const text = await navigator.clipboard.readText();
      if (!text.trim()) return setStatus('Nothing copied yet. Copy the signature in your email app first.', true);
      $('text').value = text;
      show(text);
    } catch (e) {
      $('text').focus();
      setStatus('Your phone did not allow pasting here. Press and hold in the box above, then Paste.', true);
      $('card').hidden = false;
    }
  });
  $('text').addEventListener('input', () => { if ($('text').value.trim()) show($('text').value); });
  $('card').addEventListener('submit', (event) => {
    event.preventDefault();
    const card = current();
    const problem = Card.problem(card);
    if (problem) return setStatus(problem, true);
    setStatus(save(card));
  });
  $('again').addEventListener('click', () => {
    $('text').value = '';
    FIELDS.forEach((id) => { $(id).value = ''; });
    $('card').hidden = true;
    setStatus('');
    window.scrollTo(0, 0);
  });

  // Shared from another app: Android's Share menu (installed app) sends ?text=…; the iPhone Shortcut
  // opens #text=…, which stays on the phone (the part after # is never sent to the web server).
  // Either way the contact is saved straight away; the card stays on screen to save again if needed.
  const fromQuery = new URLSearchParams(location.search);
  const fromHash = new URLSearchParams(location.hash.slice(1));
  const shared = [fromQuery.get('title'), fromQuery.get('text'), fromQuery.get('url'), fromHash.get('text')].filter(Boolean).join('\n');
  // The Outlook button sends the details it has already read as #card={…}.
  let given = null;
  try { given = fromHash.get('card') ? JSON.parse(fromHash.get('card')) : null; } catch (e) { given = null; }
  if (given || shared.trim()) {
    if (given) {
      const text = (key) => String(given[key] || '').slice(0, 200);
      fill({ name: text('name') || text('email'), email: text('email'), title: text('title'), company: text('company'),
        phones: (Array.isArray(given.phones) ? given.phones : []).slice(0, 2).map((p) => ({ label: String(p.label || 'mobile'), number: String(p.number || '').slice(0, 40) })) });
    } else {
      $('text').value = shared;
      show(shared);
    }
    history.replaceState(null, '', location.pathname);
    const card = current();
    if (!Card.problem(card)) setStatus(save(card));
  }

  if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});

  window.EmailToContactsApp = { show, current };
})();
