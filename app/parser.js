/* Offline, conservative plain-text signature reader. No DOM, storage or network. */
(function (root) {
  'use strict';
  const EMAIL = /[A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9](?:[A-Z0-9.-]*[A-Z0-9])?\.[A-Z]{2,}/gi;
  const SIGNOFF = /^(?:kind regards|best regards|warm regards|regards|best wishes|thanks(?: and regards)?|thank you|cheers|sincerely|yours sincerely|best)[,!\s]*$/i;
  const TITLE = /\b(?:manager|director|engineer|consultant|officer|coordinator|designer|developer|analyst|architect|partner|founder|president|specialist|accountant|solicitor|counsel|professor|lecturer|recruiter|assistant|executive|CEO|CTO|CFO|COO|head of|sales lead|team lead)\b/i;
  const COMPANY = /\b(?:pty\.?\s+ltd\.?|ltd\.?|limited|inc\.?|llc|llp|plc|gmbh|company|corporation|university|studio|studios|group|solutions|services|consulting|technologies|labs)\b/i;
  const LEGAL_SUFFIX = /\b(?:pty\.?\s+ltd\.?|ltd\.?|limited|inc\.?|llc|llp|plc|gmbh)\s*$/i;
  const PHONE_START = /^(?:mobile|mob|cell|office|work|direct|telephone|phone|tel|home|fax|[mtpf])\s*[:.]?\s*[+\d(]/i;
  function emails(s) { return [...new Set((s.match(EMAIL) || []))]; }
  function cleanName(s) { return s.replace(/^\s*["']|["']\s*$/g, '').trim(); }
  function websiteText(s) {
    // Includes bare domains and domain-like brands: never infer an organization
    // from a website, even when it occupies the usual company slot or is labelled.
    return /(?:https?:\/\/|www\.)\S+|(?:^|[\s:(])(?:[\p{L}\d](?:[\p{L}\d-]*[\p{L}\d])?\.)+[\p{L}]{2,}(?=$|[\s/:?#)])/iu.test(s);
  }
  function partsOf(line) {
    // A spaced dash separates fields; a hyphen inside a name/company does not.
    const parts = line.split(/\s*[|;\u00b7\u2022]\s*|\s+[-\u2013\u2014]\s+/);
    if (parts.length > 1 || /^(?:title|job title|role|company|organisation|organization):/i.test(line)) return parts;
    const comma = line.match(/^(.+),\s+([^,]+)$/);
    // Preserve common role qualifiers (Director, Customer Experience).
    if (comma && TITLE.test(comma[1]) && (LEGAL_SUFFIX.test(comma[2]) || !/^(?:customer|operations|sales|marketing|engineering|finance|product|people|technology|research|business)\b/i.test(comma[2]))) return comma.slice(1);
    return parts;
  }
  function organizationText(s) {
    return s.length <= 100 && /^[\p{L}\p{M}\d][\p{L}\p{M}\d &.'’,()\-]*$/u.test(s) &&
      !websiteText(s) && !nonIdentityText(s) && !/^(?:please|call|contact|sent|we|your|our|he|she|they)\b/i.test(s);
  }
  function nonIdentityText(s) {
    return /^(?:making|helping|building|creating|delivering|connecting|empowering|inspiring)\b/i.test(s) ||
      /^(?:customer experience|customer success|sales|marketing|engineering|finance|operations|human resources|research and development)$/i.test(s);
  }
  function nameLike(s) {
    return s.length <= 80 && /^[\p{L}\p{M}][\p{L}\p{M} .’'\-]*$/u.test(s) &&
      !TITLE.test(s) && !COMPANY.test(s) && !websiteText(s) && !/^(?:hello|hi|dear|sent from|please|contact|email|phone|mobile|office|team|support|sales)\b/i.test(s);
  }
  function signatureIdentity(line, sender) {
    const first = partsOf(line)[0].replace(/^Name:\s*/i, '');
    // Only a complete matching Latin From name may anchor an appended rendering.
    // Do not transliterate, match initials, or accept another Latin name.
    if (sender && /^[\p{Script=Latin}\p{M} .’'\-]+$/u.test(sender) &&
        first.slice(0, sender.length).toLowerCase() === sender.toLowerCase() &&
        /\s/.test(first.charAt(sender.length))) {
      let suffix = first.slice(sender.length).trim();
      if (suffix.startsWith('(') && suffix.endsWith(')')) suffix = suffix.slice(1, -1).trim();
      if (/^[\p{L}\p{M} .’'\-]+$/u.test(suffix) && /\p{L}/u.test(suffix) && !/\p{Script=Latin}/u.test(suffix)) return sender;
    }
    // Retain a name-shaped mismatch for the identity conflict check, including
    // an unmatched Latin name with a parenthesized non-Latin rendering.
    return nameLike(first.replace(/[()]/g, '')) ? first : '';
  }
  function parseEmail(input) {
    const result = { name: '', email: '', phones: [], title: '', company: '', warnings: [] };
    const warn = s => { if (!result.warnings.includes(s)) result.warnings.push(s); };
    const finish = () => {
      if (!result.phones.length) warn('No phone number found.');
      if (!result.email) warn('No sender email found; enter it manually.');
      return result;
    };
    if (typeof input !== 'string' || !input.trim()) { warn('Paste a plain-text email to read its signature.'); return finish(); }
    if (input.length > 200000) { warn('Email is too long; paste only the latest message and signature.'); return finish(); }
    // Raw HTML and attached messages need a separate, reliable text conversion.
    if (/<\/?(?:html|body|div|p|br|table|script|img|a)\b[^>]*>/i.test(input)) {
      warn('HTML is not supported; paste the visible plain-text message and signature for review.'); return finish();
    }
    let lines = input.replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ').split('\n');
    let start = 0;
    while (start < lines.length && !lines[start].trim()) start++;
    const headers = {};
    if (/^(?:From|Sent|Date|To|Subject|Reply-To):/i.test(lines[start] || '')) {
      while (start < lines.length) {
        const m = lines[start].match(/^(From|Sent|Date|To|Cc|Bcc|Subject|Reply-To):\s*(.*)$/i);
        if (!m) break;
        const key = m[1].toLowerCase();
        headers[key] = headers[key] ? headers[key] + ', ' + m[2] : m[2];
        start++;
      }
    }
    const senderEmails = emails(headers.from || '');
    if (senderEmails.length === 1) {
      result.email = senderEmails[0];
      const display = cleanName((headers.from || '').replace(/<[^>]*>/g, '').replace(EMAIL, ''));
      if (nameLike(display)) result.name = display;
    } else if (senderEmails.length > 1) {
      warn('Multiple sender addresses; review and enter the intended contact manually.'); return finish();
    }
    if (headers['reply-to'] && emails(headers['reply-to']).some(e => e.toLowerCase() !== result.email.toLowerCase())) {
      warn('Reply-To differs from the sender address; review the email field.');
    }
    let body = [];
    for (let i = start; i < lines.length; i++) {
      const line = lines[i].trim();
      const foldedReply = /^On\s/i.test(line) && lines.slice(i, i + 3).some(s => /wrote:\s*$/i.test(s));
      const replyHeader = /^(?:Date|Sent|To|Subject):/i.test(line) && lines.slice(i + 1, i + 5).some(s => /^From:/i.test(s.trim()));
      if (foldedReply || replyHeader || /^>|^On\s.+wrote:\s*$|^-{2,}\s*(?:original message|forwarded message)|^Begin forwarded message:|^_{5,}|^From:\s*\S/i.test(line)) {
        warn('Quoted or forwarded content was excluded; review the latest sender only.'); break;
      }
      if (/^(?:confidential(?:ity)?(?: notice)?\b|disclaimer\b|this (?:email|e-mail|message) (?:and|is|contains)|the information contained in this|please consider the environment|if you (?:are not|have received)|sent from my\b)/i.test(line)) break;
      body.push(line);
    }
    while (body.length && !body[body.length - 1]) body.pop();
    let boundary = -1;
    const markers = [];
    body.forEach((s, i) => { if (SIGNOFF.test(s) || /^--\s*$/.test(s)) markers.push(i); });
    if (markers.length) {
      boundary = markers[0] + 1;
      // A short personal sign-off can precede the full automated signature.
      // Skip it only when both names match From and nothing but spacing or --
      // intervenes; never skip a phone, prose or another person's identity.
      const shortIndex = body.findIndex((s, i) => i >= boundary && s);
      let fullIndex = shortIndex + 1;
      while (fullIndex < body.length && (!body[fullIndex] || /^--\s*$/.test(body[fullIndex]))) fullIndex++;
      if (SIGNOFF.test(body[markers[0]]) && result.name.includes(' ') && shortIndex >= boundary &&
          body[shortIndex].toLowerCase() === result.name.split(/\s+/)[0].toLowerCase() &&
          signatureIdentity(body[fullIndex] || '', result.name).toLowerCase() === result.name.toLowerCase()) body[shortIndex] = '';
      // A sign-off followed by the standard -- separator is still one signature.
      const realMarkers = markers.filter((m, i) => i === 0 || body.slice(markers[i - 1] + 1, m).some(Boolean));
      if (realMarkers.length > 1) { warn('Multiple signature blocks; review the sender and enter details manually.'); return finish(); }
    } else {
      // Only a final compact block with clear contact evidence is eligible.
      let lastBlank = body.lastIndexOf('');
      let block = body.slice(lastBlank + 1);
      if (block.length <= 10 && block.some(s => emails(s).length || PHONE_START.test(s) || /^(?:mobile|phone|tel|office|work|fax|m|t|p|e|email)\s*[:|]/i.test(s))) boundary = lastBlank + 1;
    }
    if (boundary < 0) { warn('No clear signature found; review and enter missing details manually.'); return finish(); }
    const signature = body.slice(boundary).filter(s => s && s !== '--');
    if (signature.length > 16) { warn('Signature boundaries are uncertain; review and enter details manually.'); return finish(); }
    const sigEmails = emails(signature.join('\n'));
    if (sigEmails.length > 1 || (result.email && sigEmails.some(e => e.toLowerCase() !== result.email.toLowerCase()))) {
      warn('Signature email differs from the sender or lists multiple people; review details manually.'); return finish();
    }
    const signatureName = signatureIdentity(signature[0] || '', result.name);
    if (result.name && signatureName && result.name.toLocaleLowerCase() !== signatureName.toLocaleLowerCase()) {
      warn('Signature name differs from the sender name; review details manually.'); return finish();
    }
    const parts = signature.flatMap(partsOf);
    if (signatureName) parts[0] = signatureName;
    if (parts.some(s => /^(?:assistant|secretary|colleague|alternate contact|other contact|on behalf of|for enquiries|for inquiries)\s*[:\-]/i.test(s))) {
      warn('Signature lists another contact; review the sender and enter details manually.'); return finish();
    }
    // A company may occupy the slot immediately after a recognizable role,
    // before contact details. Do not treat text after a phone as a company.
    const companySlots = new Set();
    let contactSeen = false;
    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];
      if (emails(part).length || websiteText(part) || PHONE_START.test(part) || /^(?:\+?\d|\(\d)/.test(part)) contactSeen = true;
      if (!contactSeen && i > 0 && TITLE.test(parts[i - 1]) &&
          (!TITLE.test(part) || COMPANY.test(part)) && organizationText(part) && part !== signatureName) companySlots.add(i);
    }
    const unknownIdentity = parts.some((s, i) => s !== signatureName && !/^Name:/i.test(s) &&
      nameLike(s) && !companySlots.has(i) && !nonIdentityText(s));
    if (unknownIdentity) {
      warn('Signature boundaries or identity are uncertain; review details manually.'); return finish();
    }
    if (!result.name && signatureName) result.name = signatureName;
    if (!result.email && sigEmails.length === 1) {
      result.email = sigEmails[0];
      warn('No From header: review the signature email and sender identity.');
    }
    // A compact block without identity evidence is unsafe (e.g. a colleague's number).
    if (!signatureName && !sigEmails.length) {
      warn('Signature cannot be associated with the sender; review details manually.'); return finish();
    }
    const titleCandidates = [], companyCandidates = [];
    for (const [partIndex, part] of parts.entries()) {
      if (!part || part === signatureName || /^Name:/i.test(part)) continue;
      const title = part.match(/^(?:title|job title|role):\s*(.+)$/i);
      const company = part.match(/^(?:company|organisation|organization):\s*(.+)$/i);
      if (title) { titleCandidates.push(title[1]); continue; }
      if (websiteText(part)) continue;
      if (company) { companyCandidates.push(company[1]); continue; }
      if (companySlots.has(partIndex)) {
        companyCandidates.push(part);
        if (!COMPANY.test(part)) warn('Review company and phone ownership: text after the role is treated as a company, but could be another person or department.');
        continue;
      }
      if (nonIdentityText(part)) { warn('Uncertain signature text was left blank; review name, title and company.'); continue; }
      if (!emails(part).length && !/\d|https?:|www\./i.test(part)) {
        if (LEGAL_SUFFIX.test(part)) companyCandidates.push(part);
        else if (TITLE.test(part)) titleCandidates.push(part);
        else if (COMPANY.test(part)) companyCandidates.push(part);
        else if (nameLike(part)) warn('Uncertain signature text was left blank; review name, title and company.');
      }
      // Unlabelled local numbers need phone-like grouping; reject dates and bare IDs.
      const phone = part.match(/^(?:(mobile|mob|cell|m|office|work|direct|telephone|phone|tel|t|p|home|fax|f)\s*[:.]?\s+|(?:mobile|mob|cell|m|office|work|direct|telephone|phone|tel|t|p|home|fax|f)\s*:\s*|tel:)?(\+?[\d(][\d ()\-.]*\d(?:\s*(?:ext\.?|extension|x|#)\s*\d+)?)\s*$/i);
      if (!phone) {
        if (/^(?:mobile|mob|cell|office|work|direct|telephone|phone|tel|home|fax|[mtpf])\s*[:.]/i.test(part)) warn('Check this number: format is uncertain; review and enter it manually.');
        continue;
      }
      const number = phone[2].trim();
      const labelText = (phone[1] || (part.match(/^([a-z]+)\s*:/i) || [])[1] || '').toLowerCase();
      if (!labelText && !number.startsWith('+') &&
          (!/[ ()\-.]/.test(number) || /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$|^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}$/.test(number))) {
        warn('Unlabelled number was left blank; review it manually.'); continue;
      }
      const digits = number.replace(/\s*(?:ext\.?|extension|x|#)\s*\d+$/i, '').replace(/\D/g, '');
      if (!labelText && !/^(?:\+|00)/.test(number) && digits.length > 11) {
        warn('Check this number: an unlabelled long identifier was left blank; review it manually.'); continue;
      }
      if (digits.length < 7 || digits.length > 15) { warn('Check this number: unsupported length; enter it manually.'); continue; }
      const label = /^(mobile|mob|cell|m)$/.test(labelText) ? 'mobile' : /^(fax|f)$/.test(labelText) ? 'fax' : labelText === 'home' ? 'home' : labelText ? 'work' : 'other';
      if (!result.phones.some(p => p.number === number && p.label === label)) result.phones.push({ label, number });
      if (label === 'other') warn('Check this number: its phone label is unknown; review before saving.');
    }
    for (const [key, candidates] of [['title', titleCandidates], ['company', companyCandidates]]) {
      const unique = [...new Set(candidates)];
      if (unique.length === 1) result[key] = unique[0];
      else if (unique.length > 1) warn('Multiple ' + key + ' values; review and enter the correct one.');
    }
    if (result.phones.length > 1) warn('Several phone numbers found; review which ones to keep.');
    if (/^(?:info|support|sales|team|office|admin|no-?reply|notifications)@/i.test(result.email)) warn('Sender may be a shared or automated address; review the contact name.');
    return finish();
  }
  if (typeof module === 'object' && module.exports) module.exports = { parseEmail };
  else root.parseEmail = parseEmail;
})(typeof globalThis === 'object' ? globalThis : this);
