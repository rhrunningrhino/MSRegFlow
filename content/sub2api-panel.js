// content/sub2api-panel.js — Read auth context from Sub2API admin dashboard

function readFirstNonEmptyStorageValue(storageObj, keys) {
  for (const key of keys) {
    try {
      const value = String(storageObj.getItem(key) || '').trim();
      if (value) return value;
    } catch {}
  }
  return '';
}

function readSub2apiAuthToken() {
  const candidateKeys = [
    'auth_token',
    'access_token',
    'admin_token',
    'token',
  ];

  const localToken = readFirstNonEmptyStorageValue(window.localStorage, candidateKeys);
  if (localToken) return localToken;

  const sessionToken = readFirstNonEmptyStorageValue(window.sessionStorage, candidateKeys);
  if (sessionToken) return sessionToken;

  return '';
}

chrome.runtime.onMessage.addListener((message, _sender, sendResponse) => {
  if (message.type !== 'GET_SUB2API_AUTH_TOKEN') {
    return false;
  }

  try {
    const token = readSub2apiAuthToken();
    if (!token) {
      sendResponse({
        error: 'No Sub2API admin token found in this page storage. Please log in to Sub2API admin first.',
      });
      return true;
    }

    sendResponse({
      ok: true,
      token,
      pageUrl: location.href,
    });
  } catch (err) {
    sendResponse({ error: String(err?.message || err || 'Unknown error') });
  }

  return true;
});
