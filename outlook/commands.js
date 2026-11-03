/* Email to Contacts for Outlook: the one-click button. Reads the open email's sender and signature
   and opens the Email to Contacts page in the browser with those details, which saves them as a
   contact card on the computer. No sign-in, and the details travel only in the part of the address
   after #, which browsers never send to the web server. */
(function () {
  'use strict';

  const Core = window.EmailToContacts;
  // The page that turns the details into a contact card: ../app/ next to this add-in.
  const CARD_PAGE = new URL('../app/', location.href).href;

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

  function cardUrl(d) {
    const card = { name: d.name, email: d.email, title: d.title, company: d.company, phones: d.phones };
    return CARD_PAGE + '#card=' + encodeURIComponent(JSON.stringify(card));
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
        const url = cardUrl(d);
        if (Office.context.ui.openBrowserWindow) Office.context.ui.openBrowserWindow(url);
        else window.open(url, '_blank', 'noopener');
        await tell(item, 'Contact card for ' + (d.name || d.email) + ' saved. Open it from your downloads to add them to your contacts.');
      }
    } catch (err) {
      await tell(item, 'Could not save: ' + err.message, true);
    }
    event.completed();
  }

  Office.onReady(() => {
    if (Office.actions && Office.actions.associate) Office.actions.associate('saveSender', saveSender);
  });
  // Outlook looks the function up by name.
  window.saveSender = saveSender;
})();
