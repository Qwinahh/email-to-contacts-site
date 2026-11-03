/* Email to Contacts for Outlook: the Google sign-in window. Sends the person to Google, and when
   Google sends them back here with a token, hands it to the task pane and closes. */
(function () {
  'use strict';
  const reply = new URLSearchParams(location.hash.slice(1));
  const send = (message) => Office.context.ui.messageParent(JSON.stringify(message));

  Office.onReady(() => {
    if (reply.get('access_token')) {
      history.replaceState(null, '', location.pathname); // don't leave the token in the address
      return send({ accessToken: reply.get('access_token'), expiresIn: Number(reply.get('expires_in') || 3600) });
    }
    if (reply.get('error')) return send({ error: 'Google sign-in did not finish: ' + reply.get('error') });
    const hint = new URLSearchParams(location.search).get('hint');
    location.replace(window.EmailToContactsGoogle.signInUrl(window.EMAIL_TO_CONTACTS_CONFIG.googleClientId, location.origin + location.pathname, hint));
  });
})();
