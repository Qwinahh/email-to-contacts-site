/* Email to Contacts web app: turning email text into a contact card and a vCard file the phone opens
   as "Create New Contact". No DOM here, so it can be tested. Uses parseEmail from parser.js. */
(function (root) {
  'use strict';

  function clean(value) {
    return String(value == null ? '' : value).replace(/[\u0000-\u0008\u000B-\u001F\u007F]/g, '').trim();
  }

  // Read pasted or shared text. A copied signature usually has no "From:" line; a whole copied or
  // forwarded email may have one, which tells the reader who the sender is.
  function readText(text, parseEmail) {
    text = String(text || '').replace(/\r\n?/g, '\n').trim();
    if (!text) return { name: '', email: '', phones: [], title: '', company: '' };
    const hasFrom = /^\s*(?:From|Von|De)\s*:/im.test(text);
    const card = parseEmail(hasFrom ? text : 'From: \n\n' + text);
    if (!card.email) card.email = (text.match(/[A-Z0-9._%+'-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i) || [''])[0];
    if (!card.name) {
      // Fall back to the first line that looks like a person's name.
      const line = text.split('\n').map((l) => l.trim()).find((l) => /^[\p{L}][\p{L}'’.-]*(?:\s+[\p{L}][\p{L}'’.-]*){1,3}$/u.test(l) && !/^(?:thanks|thank you|regards|kind regards|best regards|cheers|best|hi|hello|dear)\b/i.test(l));
      card.name = line || '';
    }
    return {
      name: clean(card.name),
      email: clean(card.email),
      phones: (card.phones || []).map((p) => ({ label: clean(p.label) || 'mobile', number: clean(p.number) })).filter((p) => p.number).slice(0, 3),
      title: clean(card.title),
      company: clean(card.company),
    };
  }

  // vCard 3.0 text escaping.
  function esc(value) {
    return String(value).replace(/\\/g, '\\\\').replace(/\r\n?|\n/g, '\\n').replace(/;/g, '\\;').replace(/,/g, '\\,');
  }

  function telType(label) {
    const l = clean(label).toLowerCase();
    if (/mob|cell/.test(l)) return 'CELL';
    if (/fax/.test(l)) return 'WORK,FAX';
    if (/home/.test(l)) return 'HOME,VOICE';
    if (/work|office|direct|main/.test(l)) return 'WORK,VOICE';
    return 'VOICE';
  }

  function buildVCard(card) {
    const name = clean(card.name);
    const company = clean(card.company);
    const words = name.split(/\s+/).filter(Boolean);
    const family = words.length > 1 ? words[words.length - 1] : '';
    const given = words.length > 1 ? words.slice(0, -1).join(' ') : name;
    const lines = ['BEGIN:VCARD', 'VERSION:3.0'];
    lines.push('N;CHARSET=UTF-8:' + esc(family) + ';' + esc(given) + ';;;');
    lines.push('FN;CHARSET=UTF-8:' + esc(name || company));
    if (company) lines.push('ORG;CHARSET=UTF-8:' + esc(company));
    if (clean(card.title)) lines.push('TITLE;CHARSET=UTF-8:' + esc(clean(card.title)));
    (card.phones || []).forEach((p) => { if (clean(p.number)) lines.push('TEL;TYPE=' + telType(p.label) + ':' + esc(clean(p.number))); });
    if (clean(card.email)) lines.push('EMAIL;TYPE=INTERNET:' + esc(clean(card.email)));
    lines.push('END:VCARD');
    return lines.join('\r\n') + '\r\n';
  }

  function fileName(card) {
    const base = clean(card.name || card.company || 'contact').replace(/[^\p{L}\p{N} _-]/gu, '').trim().replace(/\s+/g, '-');
    return (base || 'contact') + '.vcf';
  }

  // What stops saving, if anything.
  function problem(card) {
    if (!clean(card.name) && !clean(card.company)) return 'Add a name first.';
    if (clean(card.email) && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(clean(card.email))) return 'The email address looks incomplete.';
    return '';
  }

  const api = { readText, buildVCard, fileName, problem };
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.ContactCard = api;
})(typeof globalThis === 'object' ? globalThis : this);
