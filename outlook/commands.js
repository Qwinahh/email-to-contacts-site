/* Email to Contacts for Outlook: the one-click button. Reads the open email's sender and signature,
   saves them to Google Contacts (signing in to Google the first time), and says what happened in a
   line above the email. No pane opens. */
(function () {
  'use strict';

  const Core = window.EmailToContacts;
  const Google = window.EmailToContactsGoogle;
  let auth = null;

  function getBody(item, type) {
    return new Promise((resolve) => {
      item.body.getAsync(type, (res) => resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value || '' : ''));
    });
  }

  // Outlook's own plain-text version can run the signature lines together, so read the HTML.
  async function readBody(item) {
    const html = await getBody(item, Office.CoercionType.Html);
    return html ? Core.htmlToText(html) : getBody(item, Office.CoercionType.Text);
  }

  function tell(item, message, isError) {
    return new Promise((resolve) => {
      const types = Office.MailboxEnums.ItemNotificationMessageType;
      const note = isError
        ? { type: types.ErrorMessage, message: message.slice(0, 150) }
        : { type: types.InformationalMessage, message: message.slice(0, 150), icon: 'Icon.16', persistent: false };
      item.notificationMessages.replaceAsync('emailToContacts', note, () => resolve());
    });
  }

  async function saveSender(event) {
    const item = Office.context.mailbox.item;
    try {
      const me = Office.context.mailbox.userProfile.emailAddress;
      const d = Core.readDetails(item.from || item.sender, await readBody(item), me, window.parseEmail);
      if (/^This email is from you/.test(d.notes)) {
        await tell(item, 'This email is from you, so there is no one to save.', true);
      } else if (!d.name && !d.email) {
        await tell(item, 'No name or email address was found in this email.', true);
      } else {
        const token = await auth.getToken(true);
        const result = await Google.saveContact(Google.googleClient(token, fetch.bind(window), () => auth.forget()), d);
        const who = d.name || d.email;
        await tell(item, result.status === 'exists'
          ? who + ' is already in your contacts.'
          : 'Saved ' + who + ' to your contacts' + (d.phones.length ? ' with their phone number' : '') + '. It will appear on your phone shortly.');
      }
    } catch (err) {
      await tell(item, 'Could not save: ' + err.message, true);
    }
    event.completed();
  }

  Office.onReady(() => {
    auth = Google.createGoogleAuth({
      Office: Office,
      dialogUrl: new URL('google-dialog.html', location.href).href,
      session: window.sessionStorage,
      local: window.localStorage,
      fetchImpl: fetch.bind(window),
    });
    if (Office.actions && Office.actions.associate) Office.actions.associate('saveSender', saveSender);
  });
  // Outlook looks the function up by name.
  window.saveSender = saveSender;
})();
