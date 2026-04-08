// background.js — Service Worker: orchestration, state, tab management, message routing

importScripts('data/names.js');

const LOG_PREFIX = '[MultiPage:bg]';
const ICLOUD_SETUP_URLS = [
  'https://setup.icloud.com.cn/setup/ws/1',
  'https://setup.icloud.com/setup/ws/1',
];
const ICLOUD_LOGIN_URLS = [
  'https://www.icloud.com.cn/',
  'https://www.icloud.com/',
];
const STOP_ERROR_MESSAGE = 'Flow stopped by user.';
const HUMAN_STEP_DELAY_MIN = 700;
const HUMAN_STEP_DELAY_MAX = 2200;
const TOTAL_STEPS = 10;

initializeSessionStorageAccess();

// ============================================================
// State Management (chrome.storage.session)
// ============================================================

const DEFAULT_STATE = {
  currentStep: 0,
  stepStatuses: {
    1: 'pending', 2: 'pending', 3: 'pending', 4: 'pending', 5: 'pending',
    6: 'pending', 7: 'pending', 8: 'pending', 9: 'pending', 10: 'pending',
  },
  autoRunning: false,
  autoRunCurrentRun: 0,
  autoRunTotalRuns: 1,
  autoRunPausedPhase: null,
  language: 'zh-CN',
  oauthUrl: null,
  autoDeleteUsedIcloudAlias: false,
  email: null,
  password: null,
  accounts: [], // Successfully completed accounts: { email, password, createdAt }
  manualAliasUsage: {},
  lastEmailTimestamp: null,
  localhostUrl: null,
  flowStartTime: null,
  tabRegistry: {},
  logs: [],
  oauthProvider: 'cpaauth', // 'cpaauth' or 'sub2api'
  vpsUrl: '',
  sub2apiBaseUrl: '',
  sub2apiAdminApiKey: '',
  sub2apiSessionId: null,
  sub2apiRuntimeCredential: '',
  customPassword: '',
  mailProvider: 'microsoft-manager',
  inbucketHost: '',
  inbucketMailbox: '',
  microsoftManagerUrl: '',
  microsoftManagerToken: '',
  microsoftManagerMode: 'graph',
  microsoftManagerKeyword: '',
};

async function getState() {
  const state = await chrome.storage.session.get(null);
  const merged = {
    ...DEFAULT_STATE,
    ...state,
    stepStatuses: {
      ...DEFAULT_STATE.stepStatuses,
      ...(state.stepStatuses || {}),
    },
  };

  return {
    ...merged,
    oauthProvider: normalizeOauthProvider(merged.oauthProvider),
    mailProvider: normalizeMailProvider(merged.mailProvider),
    microsoftManagerMode: normalizeMicrosoftManagerMode(merged.microsoftManagerMode),
  };
}

function normalizeOauthProvider(rawValue) {
  return String(rawValue || '').trim().toLowerCase() === 'sub2api' ? 'sub2api' : 'cpaauth';
}

function normalizeMailProvider(rawValue) {
  void rawValue;
  return 'microsoft-manager';
}

function isSub2apiOauthProvider(state) {
  return normalizeOauthProvider(state?.oauthProvider) === 'sub2api';
}

async function initializeSessionStorageAccess() {
  try {
    if (chrome.storage?.session?.setAccessLevel) {
      await chrome.storage.session.setAccessLevel({
        accessLevel: 'TRUSTED_AND_UNTRUSTED_CONTEXTS',
      });
      console.log(LOG_PREFIX, 'Enabled storage.session for content scripts');
    }
  } catch (err) {
    console.warn(LOG_PREFIX, 'Failed to enable storage.session for content scripts:', err?.message || err);
  }
}

async function setState(updates) {
  console.log(LOG_PREFIX, 'storage.set:', JSON.stringify(updates).slice(0, 200));
  await chrome.storage.session.set(updates);
}

function broadcastDataUpdate(payload) {
  chrome.runtime.sendMessage({
    type: 'DATA_UPDATED',
    payload,
  }).catch(() => {});
}

function broadcastIcloudAliasesChanged(payload = {}) {
  chrome.runtime.sendMessage({
    type: 'ICLOUD_ALIASES_CHANGED',
    payload,
  }).catch(() => {});
}

async function setEmailState(email) {
  await setState({ email });
  broadcastDataUpdate({ email });
}

async function setPasswordState(password) {
  await setState({ password });
  broadcastDataUpdate({ password });
}

async function recordCompletedAccount() {
  const state = await getState();
  const email = String(state.email || '').trim();
  const password = String(state.password || '').trim();

  if (!email) return;

  const accounts = Array.isArray(state.accounts) ? [...state.accounts] : [];
  const existingIndex = accounts.findIndex(account => String(account?.email || '').trim() === email);
  const record = {
    email,
    password,
    createdAt: new Date().toISOString(),
  };

  if (existingIndex >= 0) {
    accounts[existingIndex] = {
      ...accounts[existingIndex],
      ...record,
    };
  } else {
    accounts.push(record);
  }

  const manualAliasUsage = {
    ...getManualAliasUsageMap(state),
    [email]: true,
  };

  await setState({ accounts, manualAliasUsage });
  broadcastIcloudAliasesChanged({ reason: 'used-updated', email, used: true });
}

async function maybeAutoDeleteCompletedIcloudAlias() {
  const state = await getState();
  if (!state.autoDeleteUsedIcloudAlias) return;

  const email = String(state.email || '').trim();
  if (!email) return;

  if (state.mailProvider === 'microsoft-manager') {
    try {
      await deleteMicrosoftManagerAccountByEmail(state, email);
      await addLog(`Microsoft Manager: Auto-deleted account ${email} after successful completion.`, 'ok');
    } catch (err) {
      await addLog(`Microsoft Manager: Auto-delete failed for ${email}: ${getErrorMessage(err)}`, 'warn');
    }
    return;
  }

  try {
    const aliases = await listIcloudAliases();
    const alias = aliases.find(item => String(item?.email || '').trim() === email);

    if (!alias) {
      await addLog(`iCloud: Auto-delete skipped. ${email} was not found in your Hide My Email alias list.`, 'warn');
      return;
    }

    if (!alias.anonymousId) {
      await addLog(`iCloud: Auto-delete skipped. ${email} is missing anonymousId; refresh aliases and retry manually.`, 'warn');
      return;
    }

    await deleteIcloudAlias(alias);
    await addLog(`iCloud: Auto-deleted used alias ${email} after successful completion.`, 'ok');
  } catch (err) {
    await addLog(`iCloud: Auto-delete failed for ${email}: ${getErrorMessage(err)}`, 'warn');
  }
}

async function setManualEmailState(email) {
  const trimmedEmail = String(email || '').trim();
  await setState({ email: trimmedEmail });
  broadcastDataUpdate({ email: trimmedEmail });
}

function getManualAliasUsageMap(state) {
  return state?.manualAliasUsage && typeof state.manualAliasUsage === 'object'
    ? { ...state.manualAliasUsage }
    : {};
}

function getEffectiveUsedEmails(state) {
  const usedEmails = new Set((state.accounts || []).map(account => String(account?.email || '').trim()).filter(Boolean));
  const manualAliasUsage = getManualAliasUsageMap(state);

  for (const [email, used] of Object.entries(manualAliasUsage)) {
    const normalizedEmail = String(email || '').trim();
    if (!normalizedEmail) continue;
    if (used) usedEmails.add(normalizedEmail);
    else usedEmails.delete(normalizedEmail);
  }

  return usedEmails;
}

async function setIcloudAliasUsedState(payload = {}) {
  const email = String(payload.email || '').trim();
  if (!email) {
    throw new Error('No iCloud alias email was provided.');
  }

  const state = await getState();
  const manualAliasUsage = getManualAliasUsageMap(state);
  manualAliasUsage[email] = Boolean(payload.used);
  await setState({ manualAliasUsage });
  await addLog(`iCloud: Marked ${email} as ${payload.used ? 'used' : 'unused'}`, 'ok');
  broadcastIcloudAliasesChanged({ reason: 'used-updated', email, used: Boolean(payload.used) });
  return { email, used: Boolean(payload.used) };
}

function getErrorMessage(error) {
  if (typeof error === 'string') return error;
  return String(error?.message || error || 'Unknown error');
}

function isIcloudLoginRequiredError(error) {
  const message = getErrorMessage(error).toLowerCase();
  return message.includes('could not validate icloud session')
    || message.includes('hide my email service was unavailable')
    || /\bstatus (401|403|409|421)\b/.test(message);
}

function getPreferredIcloudLoginUrl(error) {
  const message = getErrorMessage(error).toLowerCase();
  if (message.includes('icloud.com/setup') && !message.includes('icloud.com.cn/setup')) {
    return 'https://www.icloud.com/';
  }
  return 'https://www.icloud.com.cn/';
}

let lastIcloudLoginPromptAt = 0;

async function openIcloudLoginPage(preferredUrl) {
  const urlPatterns = [
    'https://www.icloud.com/*',
    'https://www.icloud.com.cn/*',
  ];
  const tabs = await chrome.tabs.query({ url: urlPatterns });
  const preferredHost = new URL(preferredUrl).host;
  const existing = tabs.find(tab => {
    try {
      return new URL(tab.url).host === preferredHost;
    } catch {
      return false;
    }
  }) || tabs[0];

  if (existing?.id) {
    await chrome.tabs.update(existing.id, { active: true });
    if (existing.url !== preferredUrl) {
      await chrome.tabs.update(existing.id, { url: preferredUrl });
    }
    return existing.id;
  }

  const created = await chrome.tabs.create({ url: preferredUrl, active: true });
  return created.id;
}

async function promptIcloudLogin(error, actionLabel = 'iCloud action') {
  const now = Date.now();
  const preferredUrl = getPreferredIcloudLoginUrl(error);
  const originalError = getErrorMessage(error);

  chrome.runtime.sendMessage({
    type: 'ICLOUD_LOGIN_REQUIRED',
    payload: {
      actionLabel,
      loginUrl: preferredUrl,
      message: 'iCloud sign-in is required. A login page has been opened for you.',
      detail: originalError,
    },
  }).catch(() => {});

  if (now - lastIcloudLoginPromptAt < 15000) {
    return;
  }
  lastIcloudLoginPromptAt = now;

  await addLog(`iCloud login required during ${actionLabel}. Opening ${new URL(preferredUrl).host}...`, 'warn');

  try {
    await openIcloudLoginPage(preferredUrl);
  } catch (tabErr) {
    await addLog(`iCloud: Failed to open login page automatically: ${getErrorMessage(tabErr)}`, 'warn');
  }
}

async function withIcloudLoginHelp(actionLabel, action) {
  try {
    return await action();
  } catch (err) {
    if (isIcloudLoginRequiredError(err)) {
      await addLog(`iCloud login check failed during ${actionLabel}: ${getErrorMessage(err)}`, 'warn');
      await promptIcloudLogin(err, actionLabel);
      throw new Error('Please finish signing in on the opened iCloud page, then click "I\'ve Signed In".');
    }
    throw err;
  }
}

async function checkIcloudSession() {
  return withIcloudLoginHelp('checking iCloud session', async () => {
    const { setupUrl } = await resolveIcloudPremiumMailService();
    await addLog(`iCloud: Session check passed via ${new URL(setupUrl).host}`, 'ok');
    return { ok: true, setupUrl };
  });
}

async function icloudRequest(method, url, options = {}) {
  const { data } = options;
  let response;
  try {
    response = await fetch(url, {
      method,
      body: data !== undefined ? JSON.stringify(data) : undefined,
    });
  } catch (err) {
    throw new Error(`iCloud request failed for ${method} ${url}: ${err.message}`);
  }

  if (!response.ok) {
    throw new Error(`iCloud request failed for ${method} ${url} with status ${response.status}`);
  }

  try {
    return await response.json();
  } catch (err) {
    throw new Error(`iCloud returned invalid JSON for ${method} ${url}: ${err.message}`);
  }
}

async function validateIcloudSession(setupUrl) {
  const data = await icloudRequest('POST', `${setupUrl}/validate`);
  if (!data?.webservices?.premiummailsettings?.url) {
    throw new Error('iCloud session validated, but Hide My Email service was unavailable.');
  }
  return data;
}

async function resolveIcloudPremiumMailService() {
  const errors = [];

  for (const setupUrl of ICLOUD_SETUP_URLS) {
    try {
      const data = await validateIcloudSession(setupUrl);
      return {
        setupUrl,
        serviceUrl: String(data.webservices.premiummailsettings.url).replace(/\/$/, ''),
      };
    } catch (err) {
      errors.push(`${new URL(setupUrl).host}: ${err.message}`);
    }
  }

  throw new Error(errors.length
    ? `Could not validate iCloud session. ${errors.join(' | ')}`
    : 'Could not validate iCloud session. Log into icloud.com.cn or icloud.com in this browser first.');
}

function getIcloudAliasLabel() {
  const now = new Date();
  const dateStr = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  return `MultiPage ${dateStr}`;
}

function findIcloudAliasArray(node, depth = 0) {
  if (!node || depth > 4) return null;
  if (Array.isArray(node)) {
    return node.some(item => typeof item === 'object') ? node : null;
  }
  if (typeof node !== 'object') return null;

  const priorityKeys = ['hmeEmails', 'hmeEmailList', 'hmeList', 'hmes', 'aliases', 'items'];
  for (const key of priorityKeys) {
    if (Array.isArray(node[key])) return node[key];
  }

  for (const value of Object.values(node)) {
    const nested = findIcloudAliasArray(value, depth + 1);
    if (nested) return nested;
  }

  return null;
}

function normalizeIcloudAliasRecord(raw, usedEmails = new Set()) {
  const anonymousId = String(raw?.anonymousId || raw?.id || '').trim();
  const email = String(
    raw?.hme
      || raw?.email
      || raw?.alias
      || raw?.address
      || raw?.metaData?.hme
      || ''
  ).trim();

  if (!email || !email.includes('@')) return null;

  const label = String(raw?.label || raw?.metaData?.label || '').trim();
  const note = String(raw?.note || raw?.metaData?.note || '').trim();
  const state = String(raw?.state || raw?.status || '').trim().toLowerCase();
  const createdAt = raw?.createTimestamp
    || raw?.createTime
    || raw?.createdAt
    || raw?.createdDate
    || null;

  return {
    anonymousId,
    email,
    label,
    note,
    state,
    active: raw?.active !== false && raw?.isActive !== false && state !== 'inactive' && state !== 'deleted',
    used: usedEmails.has(email),
    createdAt,
  };
}

async function listIcloudAliases() {
  return withIcloudLoginHelp('loading iCloud aliases', async () => {
    const { serviceUrl } = await resolveIcloudPremiumMailService();
    const response = await icloudRequest('GET', `${serviceUrl}/v2/hme/list`);
    const aliases = findIcloudAliasArray(response);
    const state = await getState();
    const usedEmails = getEffectiveUsedEmails(state);

    if (!aliases) return [];

    return aliases
      .map(alias => normalizeIcloudAliasRecord(alias, usedEmails))
      .filter(Boolean)
      .sort((a, b) => {
        if (a.active !== b.active) return a.active ? -1 : 1;
        if (a.used !== b.used) return a.used ? 1 : -1;
        return String(a.email).localeCompare(String(b.email));
      });
  });
}

async function deleteIcloudAlias(email) {
  return withIcloudLoginHelp('deleting iCloud alias', async () => {
    const alias = typeof email === 'string'
      ? { email: String(email).trim(), anonymousId: '' }
      : {
          email: String(email?.email || '').trim(),
          anonymousId: String(email?.anonymousId || '').trim(),
        };

    if (!alias.email) {
      throw new Error('No iCloud alias email was provided.');
    }
    if (!alias.anonymousId) {
      throw new Error(`No anonymousId found for iCloud alias ${alias.email}. Refresh the alias list and retry.`);
    }

    const { serviceUrl } = await resolveIcloudPremiumMailService();

    try {
      const directDelete = await icloudRequest('POST', `${serviceUrl}/v1/hme/delete`, {
        data: { anonymousId: alias.anonymousId },
      });
      if (directDelete?.success === false) {
        throw new Error(directDelete?.error?.errorMessage || 'delete failed');
      }
    } catch (err) {
      await addLog(`iCloud: Direct delete failed for ${alias.email}, trying deactivate fallback...`, 'warn');

      const deactivated = await icloudRequest('POST', `${serviceUrl}/v1/hme/deactivate`, {
        data: { anonymousId: alias.anonymousId },
      });
      if (deactivated?.success === false) {
        throw new Error(deactivated?.error?.errorMessage || `Failed to deactivate ${alias.email}`);
      }

      const deleted = await icloudRequest('POST', `${serviceUrl}/v1/hme/delete`, {
        data: { anonymousId: alias.anonymousId },
      });
      if (deleted?.success === false) {
        throw new Error(deleted?.error?.errorMessage || `Failed to delete ${alias.email}`);
      }
    }

    const state = await getState();
    const manualAliasUsage = getManualAliasUsageMap(state);
    if (alias.email in manualAliasUsage) {
      delete manualAliasUsage[alias.email];
      await setState({ manualAliasUsage });
    }

    await addLog(`iCloud: Deleted alias ${alias.email}`, 'ok');
    broadcastIcloudAliasesChanged({ reason: 'deleted', email: alias.email });
    return { email: alias.email };
  });
}

async function deleteUsedIcloudAliases() {
  const aliases = await listIcloudAliases();
  const usedAliases = aliases.filter(alias => alias.used);

  if (usedAliases.length === 0) {
    return { deleted: [], skipped: [] };
  }

  const deleted = [];
  const skipped = [];

  for (const alias of usedAliases) {
    try {
      await deleteIcloudAlias(alias);
      deleted.push(alias.email);
    } catch (err) {
      skipped.push({ email: alias.email, error: err.message });
    }
  }

  return { deleted, skipped };
}

async function fetchIcloudHideMyEmail() {
  return withIcloudLoginHelp('generating iCloud Hide My Email alias', async () => {
    throwIfStopped();
    await addLog('iCloud: Validating browser session for Hide My Email...');

    const { serviceUrl, setupUrl } = await resolveIcloudPremiumMailService();
    await addLog(`iCloud: Session validated via ${new URL(setupUrl).host}`, 'ok');

    const existingAliasesResponse = await icloudRequest('GET', `${serviceUrl}/v2/hme/list`);
    const state = await getState();
    const usedEmails = getEffectiveUsedEmails(state);
    const existingAliases = (findIcloudAliasArray(existingAliasesResponse) || [])
      .map(alias => normalizeIcloudAliasRecord(alias, usedEmails))
      .filter(Boolean);

    const reusableAlias = existingAliases.find(alias => alias.active && !alias.used);
    if (reusableAlias) {
      await setEmailState(reusableAlias.email);
      await addLog(`iCloud: Reusing unused alias ${reusableAlias.email}`, 'ok');
      broadcastIcloudAliasesChanged({ reason: 'selected', email: reusableAlias.email });
      return reusableAlias.email;
    }

    await addLog('iCloud: No unused active alias available, generating a new one...');

    const generated = await icloudRequest('POST', `${serviceUrl}/v1/hme/generate`);
    if (!generated?.success || !generated?.result?.hme) {
      throw new Error(generated?.error?.errorMessage || 'iCloud Hide My Email generate failed.');
    }

    const reservePayload = {
      hme: generated.result.hme,
      label: getIcloudAliasLabel(),
      note: 'Generated through MultiPage Automation',
    };
    const reserved = await icloudRequest('POST', `${serviceUrl}/v1/hme/reserve`, {
      data: reservePayload,
    });

    if (!reserved?.success || !reserved?.result?.hme?.hme) {
      throw new Error(reserved?.error?.errorMessage || 'iCloud Hide My Email reserve failed.');
    }

    const alias = reserved.result.hme.hme;
    await setEmailState(alias);
    await addLog(`iCloud: Reserved Hide My Email alias ${alias}`, 'ok');
    broadcastIcloudAliasesChanged({ reason: 'created', email: alias });
    return alias;
  });
}

function normalizeMicrosoftManagerBaseUrl(rawValue) {
  const value = String(rawValue || '').trim();
  if (!value) return '';

  const candidate = /^[a-zA-Z][a-zA-Z\d+\-.]*:\/\//.test(value) ? value : `https://${value}`;

  try {
    const parsed = new URL(candidate);
    return parsed.origin;
  } catch {
    return '';
  }
}

function normalizeMicrosoftManagerMode(rawValue) {
  const mode = String(rawValue || '').trim().toLowerCase();
  return mode === 'imap' ? 'imap' : 'graph';
}

async function requestMicrosoftManagerApi(state, path, options = {}) {
  const baseUrl = normalizeMicrosoftManagerBaseUrl(state.microsoftManagerUrl);
  if (!baseUrl) {
    throw new Error('Microsoft Account Manager URL is empty or invalid.');
  }

  const token = String(state.microsoftManagerToken || '').trim();
  if (!token) {
    throw new Error('Microsoft Account Manager token is empty.');
  }

  const url = new URL(path, `${baseUrl}/`);
  if (options.query && typeof options.query === 'object') {
    for (const [key, value] of Object.entries(options.query)) {
      if (value === undefined || value === null) continue;
      const normalizedValue = String(value).trim();
      if (!normalizedValue) continue;
      url.searchParams.set(key, normalizedValue);
    }
  }

  const headers = {
    'x-mail-api-token': token,
    ...options.headers,
  };

  const init = {
    method: options.method || 'GET',
    headers,
  };

  if (options.body !== undefined) {
    headers['Content-Type'] = 'application/json';
    init.body = JSON.stringify(options.body);
  }

  let response;
  try {
    response = await fetch(url.toString(), init);
  } catch (err) {
    throw new Error(`Microsoft Account Manager request failed: ${getErrorMessage(err)}`);
  }

  let payload = {};
  try {
    payload = await response.json();
  } catch {
    payload = {};
  }

  if (!response.ok) {
    const message = String(payload?.message || payload?.error || `${response.status} ${response.statusText}`).trim();
    throw new Error(`Microsoft Account Manager request failed: ${message}`);
  }

  return payload;
}

function normalizeMicrosoftManagerAccount(raw) {
  const id = Number(raw?.id || 0);
  const account = String(raw?.account || '').trim();
  if (!account) return null;
  return {
    id: Number.isFinite(id) ? id : 0,
    account,
    remark: String(raw?.remark || '').trim(),
  };
}

async function listMicrosoftManagerAccounts(state, options = {}) {
  const keyword = String(options.keyword ?? state.microsoftManagerKeyword ?? '').trim();
  const payload = await requestMicrosoftManagerApi(state, '/api/open/accounts', {
    method: 'GET',
    query: { keyword },
  });

  const items = Array.isArray(payload?.items) ? payload.items : [];
  return items
    .map(item => normalizeMicrosoftManagerAccount(item))
    .filter(Boolean)
    .filter(item => item.account.includes('@'));
}

function findExactMicrosoftManagerAccounts(accounts, email) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  if (!normalizedEmail) return [];

  return accounts
    .filter(item => String(item?.account || '').trim().toLowerCase() === normalizedEmail)
    .sort((a, b) => Number(b?.id || 0) - Number(a?.id || 0));
}

async function deleteMicrosoftManagerAccountByEmail(state, email) {
  const accounts = await listMicrosoftManagerAccounts(state, { keyword: email });
  const matches = findExactMicrosoftManagerAccounts(accounts, email);

  if (!matches.length) {
    throw new Error(`${email} was not found in Microsoft Account Manager account list.`);
  }

  const target = matches[0];
  const id = Number(target?.id || 0);
  if (!Number.isInteger(id) || id <= 0) {
    throw new Error(`No valid account id found for ${email}.`);
  }

  await requestMicrosoftManagerApi(state, `/api/open/accounts/${id}`, {
    method: 'DELETE',
  });
}

function pickMicrosoftManagerAccount(accounts, state) {
  if (!accounts.length) return null;

  const usedEmails = new Set(
    (state.accounts || [])
      .map(account => String(account?.email || '').trim())
      .filter(Boolean)
  );

  return accounts.find(account => !usedEmails.has(account.account)) || accounts[0];
}

async function fetchMicrosoftManagerEmail(options = {}) {
  const state = await getState();
  const accounts = await listMicrosoftManagerAccounts(state, {
    keyword: options.keyword,
  });

  if (!accounts.length) {
    throw new Error('No account found in Microsoft Account Manager.');
  }

  const selected = pickMicrosoftManagerAccount(accounts, state);
  if (!selected?.account) {
    throw new Error('No valid account email found in Microsoft Account Manager.');
  }

  await setEmailState(selected.account);
  await addLog(`Microsoft Manager: Selected account ${selected.account}`, 'ok');
  return selected.account;
}

function stripHtmlTags(value) {
  return String(value || '')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractVerificationCodeFromText(text) {
  const normalized = String(text || '');

  const matchCn = normalized.match(/(?:代码为|验证码[^0-9]*?)[\s：:]*(\d{6})/);
  if (matchCn) return matchCn[1];

  const matchEn = normalized.match(/code[:\s]+is[:\s]+(\d{6})|code[:\s]+(\d{6})/i);
  if (matchEn) return matchEn[1] || matchEn[2];

  const match6 = normalized.match(/\b(\d{6})\b/);
  if (match6) return match6[1];

  return null;
}

function parseReceivedAtMs(value) {
  const timestamp = Date.parse(String(value || '').trim());
  return Number.isFinite(timestamp) ? timestamp : 0;
}

function findMicrosoftManagerCode(messages, payload = {}) {
  const senderFilters = Array.isArray(payload.senderFilters) ? payload.senderFilters : [];
  const subjectFilters = Array.isArray(payload.subjectFilters) ? payload.subjectFilters : [];
  const filterAfterTimestamp = Number(payload.filterAfterTimestamp || 0);

  for (const message of messages) {
    const receivedAt = parseReceivedAtMs(message?.receivedAt);
    if (filterAfterTimestamp && receivedAt && receivedAt <= filterAfterTimestamp) {
      continue;
    }

    const sender = String(message?.from || '').trim().toLowerCase();
    const subject = String(message?.subject || '').trim().toLowerCase();
    const preview = String(message?.preview || '').trim();
    const content = stripHtmlTags(String(message?.content || '').trim());
    const combinedText = [subject, sender, preview, content]
      .filter(Boolean)
      .join(' ');
    const combinedLower = combinedText.toLowerCase();

    const senderMatch = senderFilters.some(filter => {
      const keyword = String(filter || '').trim().toLowerCase();
      return keyword && (sender.includes(keyword) || combinedLower.includes(keyword));
    });
    const subjectMatch = subjectFilters.some(filter => {
      const keyword = String(filter || '').trim().toLowerCase();
      return keyword && (subject.includes(keyword) || combinedLower.includes(keyword));
    });
    const keywordMatch = /openai|chatgpt|verify|verification|confirm|login|验证码|代码/.test(combinedLower);

    const code = extractVerificationCodeFromText(combinedText);
    if (!code) continue;
    if (!senderMatch && !subjectMatch && !keywordMatch) continue;

    return {
      code,
      emailTimestamp: receivedAt || Date.now(),
      mailId: String(message?.id || '').trim(),
      subject: String(message?.subject || '').trim(),
    };
  }

  return null;
}

async function pollMicrosoftManagerCode(state, step, pollPayload = {}) {
  const targetEmail = String(pollPayload.targetEmail || state.email || '').trim();
  if (!targetEmail) {
    throw new Error('No email address provided for Microsoft Account Manager polling.');
  }

  const mode = normalizeMicrosoftManagerMode(state.microsoftManagerMode);
  const maxAttempts = Math.max(1, Number(pollPayload.maxAttempts || 20));
  const intervalMs = Math.max(500, Number(pollPayload.intervalMs || 3000));

  let lastErrorMessage = '';

  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    throwIfStopped();
    await addLog(`Step ${step}: Polling Microsoft Account Manager... attempt ${attempt}/${maxAttempts}`);

    try {
      const payload = await requestMicrosoftManagerApi(state, '/api/open/messages', {
        method: 'POST',
        body: {
          account: targetEmail,
          mode,
        },
      });

      const messages = Array.isArray(payload?.messages) ? payload.messages : [];
      const found = findMicrosoftManagerCode(messages, pollPayload);
      if (found?.code) {
        await addLog(`Step ${step}: Found code in Microsoft Account Manager (${found.subject || 'no subject'})`, 'ok');
        return {
          ok: true,
          code: found.code,
          emailTimestamp: found.emailTimestamp,
          mailId: found.mailId,
        };
      }
    } catch (err) {
      lastErrorMessage = getErrorMessage(err);
      await addLog(`Step ${step}: Microsoft Account Manager poll failed on attempt ${attempt}: ${lastErrorMessage}`, 'warn');
    }

    if (attempt < maxAttempts) {
      await sleepWithStop(intervalMs);
    }
  }

  const timeoutSec = (maxAttempts * intervalMs / 1000).toFixed(0);
  if (lastErrorMessage) {
    return {
      error: `No matching verification email found after ${timeoutSec}s. Last error: ${lastErrorMessage}`,
    };
  }

  return {
    error: `No matching verification email found in Microsoft Account Manager after ${timeoutSec}s.`,
  };
}

async function fetchConfiguredEmail(options = {}) {
  return fetchMicrosoftManagerEmail(options);
}

async function resetState() {
  console.log(LOG_PREFIX, 'Resetting all state');
  // Preserve settings and persistent data across resets
  const prev = await chrome.storage.session.get([
    'seenCodes',
    'seenInbucketMailIds',
    'accounts',
    'manualAliasUsage',
    'tabRegistry',
    'language',
    'oauthProvider',
    'vpsUrl',
    'sub2apiBaseUrl',
    'sub2apiAdminApiKey',
    'autoDeleteUsedIcloudAlias',
    'customPassword',
    'mailProvider',
    'inbucketHost',
    'inbucketMailbox',
    'microsoftManagerUrl',
    'microsoftManagerToken',
    'microsoftManagerMode',
    'microsoftManagerKeyword',
  ]);
  await chrome.storage.session.clear();
  await chrome.storage.session.set({
    ...DEFAULT_STATE,
    seenCodes: prev.seenCodes || [],
    seenInbucketMailIds: prev.seenInbucketMailIds || [],
    accounts: prev.accounts || [],
    manualAliasUsage: prev.manualAliasUsage && typeof prev.manualAliasUsage === 'object' ? prev.manualAliasUsage : {},
    tabRegistry: prev.tabRegistry || {},
    language: prev.language || 'zh-CN',
    oauthProvider: normalizeOauthProvider(prev.oauthProvider || 'cpaauth'),
    vpsUrl: prev.vpsUrl || '',
    sub2apiBaseUrl: prev.sub2apiBaseUrl || '',
    sub2apiAdminApiKey: prev.sub2apiAdminApiKey || '',
    autoDeleteUsedIcloudAlias: Boolean(prev.autoDeleteUsedIcloudAlias),
    customPassword: prev.customPassword || '',
    mailProvider: normalizeMailProvider(prev.mailProvider),
    inbucketHost: prev.inbucketHost || '',
    inbucketMailbox: prev.inbucketMailbox || '',
    microsoftManagerUrl: prev.microsoftManagerUrl || '',
    microsoftManagerToken: prev.microsoftManagerToken || '',
    microsoftManagerMode: normalizeMicrosoftManagerMode(prev.microsoftManagerMode || 'graph'),
    microsoftManagerKeyword: prev.microsoftManagerKeyword || '',
  });
}

/**
 * Generate a random password: 14 chars, mix of uppercase, lowercase, digits, symbols.
 */
function generatePassword() {
  const upper = 'ABCDEFGHJKLMNPQRSTUVWXYZ';
  const lower = 'abcdefghjkmnpqrstuvwxyz';
  const digits = '23456789';
  const symbols = '!@#$%&*?';
  const all = upper + lower + digits + symbols;

  // Ensure at least one of each type
  let pw = '';
  pw += upper[Math.floor(Math.random() * upper.length)];
  pw += lower[Math.floor(Math.random() * lower.length)];
  pw += digits[Math.floor(Math.random() * digits.length)];
  pw += symbols[Math.floor(Math.random() * symbols.length)];

  // Fill remaining 10 chars
  for (let i = 0; i < 10; i++) {
    pw += all[Math.floor(Math.random() * all.length)];
  }

  // Shuffle
  return pw.split('').sort(() => Math.random() - 0.5).join('');
}

// ============================================================
// Tab Registry
// ============================================================

async function getTabRegistry() {
  const state = await getState();
  return state.tabRegistry || {};
}

async function registerTab(source, tabId) {
  const registry = await getTabRegistry();
  registry[source] = { tabId, ready: true };
  await setState({ tabRegistry: registry });
  console.log(LOG_PREFIX, `Tab registered: ${source} -> ${tabId}`);
}

async function isTabAlive(source) {
  const registry = await getTabRegistry();
  const entry = registry[source];
  if (!entry) return false;
  try {
    await chrome.tabs.get(entry.tabId);
    return true;
  } catch {
    // Tab no longer exists — clean up registry
    registry[source] = null;
    await setState({ tabRegistry: registry });
    return false;
  }
}

async function getTabId(source) {
  const registry = await getTabRegistry();
  return registry[source]?.tabId || null;
}

// ============================================================
// Command Queue (for content scripts not yet ready)
// ============================================================

const pendingCommands = new Map(); // source -> { message, resolve, reject, timer }

function queueCommand(source, message, timeout = 15000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingCommands.delete(source);
      const err = `Content script on ${source} did not respond in ${timeout / 1000}s. Try refreshing the tab and retry.`;
      console.error(LOG_PREFIX, err);
      reject(new Error(err));
    }, timeout);
    pendingCommands.set(source, { message, resolve, reject, timer });
    console.log(LOG_PREFIX, `Command queued for ${source} (waiting for ready)`);
  });
}

function flushCommand(source, tabId) {
  const pending = pendingCommands.get(source);
  if (pending) {
    clearTimeout(pending.timer);
    pendingCommands.delete(source);
    chrome.tabs.sendMessage(tabId, pending.message).then(pending.resolve).catch(pending.reject);
    console.log(LOG_PREFIX, `Flushed queued command to ${source} (tab ${tabId})`);
  }
}

function cancelPendingCommands(reason = STOP_ERROR_MESSAGE) {
  for (const [source, pending] of pendingCommands.entries()) {
    clearTimeout(pending.timer);
    pending.reject(new Error(reason));
    pendingCommands.delete(source);
    console.log(LOG_PREFIX, `Cancelled queued command for ${source}`);
  }
}

// ============================================================
// Reuse or create tab
// ============================================================

async function reuseOrCreateTab(source, url, options = {}) {
  const alive = await isTabAlive(source);
  if (alive) {
    const tabId = await getTabId(source);
    const currentTab = await chrome.tabs.get(tabId);
    const sameUrl = currentTab.url === url;
    const shouldReloadOnReuse = sameUrl && options.reloadIfSameUrl;

    const registry = await getTabRegistry();
    if (sameUrl) {
      await chrome.tabs.update(tabId, { active: true });
      console.log(LOG_PREFIX, `Reused tab ${source} (${tabId}) on same URL`);

      if (shouldReloadOnReuse) {
        if (registry[source]) registry[source].ready = false;
        await setState({ tabRegistry: registry });
        await chrome.tabs.reload(tabId);

        await new Promise((resolve) => {
          const timer = setTimeout(() => { chrome.tabs.onUpdated.removeListener(listener); resolve(); }, 30000);
          const listener = (tid, info) => {
            if (tid === tabId && info.status === 'complete') {
              chrome.tabs.onUpdated.removeListener(listener);
              clearTimeout(timer);
              resolve();
            }
          };
          chrome.tabs.onUpdated.addListener(listener);
        });
      }

      // For dynamically injected pages like the CPA Auth panel, re-inject immediately.
      if (options.inject) {
        if (registry[source]) registry[source].ready = false;
        await setState({ tabRegistry: registry });
        if (options.injectSource) {
          await chrome.scripting.executeScript({
            target: { tabId },
            func: (injectedSource) => {
              window.__MULTIPAGE_SOURCE = injectedSource;
            },
            args: [options.injectSource],
          });
        }
        await chrome.scripting.executeScript({
          target: { tabId },
          files: options.inject,
        });
        await new Promise(r => setTimeout(r, 500));
      }

      return tabId;
    }

    // Mark as not ready BEFORE navigating — so READY signal from new page is captured correctly
    if (registry[source]) registry[source].ready = false;
    await setState({ tabRegistry: registry });

    // Navigate existing tab to new URL
    await chrome.tabs.update(tabId, { url, active: true });
    console.log(LOG_PREFIX, `Reused tab ${source} (${tabId}), navigated to ${url.slice(0, 60)}`);

    // Wait for page load complete (with 30s timeout)
    await new Promise((resolve) => {
      const timer = setTimeout(() => { chrome.tabs.onUpdated.removeListener(listener); resolve(); }, 30000);
      const listener = (tid, info) => {
        if (tid === tabId && info.status === 'complete') {
          chrome.tabs.onUpdated.removeListener(listener);
          clearTimeout(timer);
          resolve();
        }
      };
      chrome.tabs.onUpdated.addListener(listener);
    });

    // If dynamic injection needed (CPA Auth panel), re-inject after navigation
    if (options.inject) {
      if (options.injectSource) {
        await chrome.scripting.executeScript({
          target: { tabId },
          func: (injectedSource) => {
            window.__MULTIPAGE_SOURCE = injectedSource;
          },
          args: [options.injectSource],
        });
      }
      await chrome.scripting.executeScript({
        target: { tabId },
        files: options.inject,
      });
    }

    // Wait a bit for content script to inject and send READY
    await new Promise(r => setTimeout(r, 500));

    return tabId;
  }

  // Create new tab
  const tab = await chrome.tabs.create({ url, active: true });
  console.log(LOG_PREFIX, `Created new tab ${source} (${tab.id})`);

  // If dynamic injection needed (CPA Auth panel), inject scripts after load
  if (options.inject) {
    await new Promise((resolve) => {
      const timer = setTimeout(() => { chrome.tabs.onUpdated.removeListener(listener); resolve(); }, 30000);
      const listener = (tabId, info) => {
        if (tabId === tab.id && info.status === 'complete') {
          chrome.tabs.onUpdated.removeListener(listener);
          clearTimeout(timer);
          resolve();
        }
      };
      chrome.tabs.onUpdated.addListener(listener);
    });
    if (options.injectSource) {
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: (injectedSource) => {
          window.__MULTIPAGE_SOURCE = injectedSource;
        },
        args: [options.injectSource],
      });
    }
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: options.inject,
    });
  }

  return tab.id;
}

// ============================================================
// Send command to content script (with readiness check)
// ============================================================

async function sendToContentScript(source, message) {
  const registry = await getTabRegistry();
  const entry = registry[source];

  if (!entry || !entry.ready) {
    console.log(LOG_PREFIX, `${source} not ready, queuing command`);
    return queueCommand(source, message);
  }

  // Verify tab is still alive
  const alive = await isTabAlive(source);
  if (!alive) {
    // Tab was closed — queue the command, it will be sent when tab is reopened
    console.log(LOG_PREFIX, `${source} tab was closed, queuing command`);
    return queueCommand(source, message);
  }

  console.log(LOG_PREFIX, `Sending to ${source} (tab ${entry.tabId}):`, message.type);
  return chrome.tabs.sendMessage(entry.tabId, message);
}

// ============================================================
// Logging
// ============================================================

async function addLog(message, level = 'info') {
  const state = await getState();
  const logs = state.logs || [];
  const entry = { message, level, timestamp: Date.now() };
  logs.push(entry);
  // Keep last 500 logs
  if (logs.length > 500) logs.splice(0, logs.length - 500);
  await setState({ logs });
  // Broadcast to side panel
  chrome.runtime.sendMessage({ type: 'LOG_ENTRY', payload: entry }).catch(() => {});
}

// ============================================================
// Step Status Management
// ============================================================

async function setStepStatus(step, status) {
  const state = await getState();
  const statuses = { ...state.stepStatuses };
  statuses[step] = status;
  await setState({ stepStatuses: statuses, currentStep: step });
  // Broadcast to side panel
  chrome.runtime.sendMessage({
    type: 'STEP_STATUS_CHANGED',
    payload: { step, status },
  }).catch(() => {});
}

async function skipStep(step) {
  const stepNum = Number(step);
  if (!Number.isInteger(stepNum) || stepNum < 1 || stepNum > TOTAL_STEPS) {
    throw new Error(`Invalid step to skip: ${step}`);
  }

  const state = await getState();
  const currentStatus = state.stepStatuses?.[stepNum] || 'pending';
  if (currentStatus === 'running') {
    throw new Error(`Cannot skip step ${stepNum} while it is running.`);
  }

  await setStepStatus(stepNum, 'skipped');
  await addLog(`Step ${stepNum} skipped by user`, 'warn');
}

function isStopError(error) {
  const message = typeof error === 'string' ? error : error?.message;
  return message === STOP_ERROR_MESSAGE;
}

function clearStopRequest() {
  stopRequested = false;
}

function throwIfStopped() {
  if (stopRequested) {
    throw new Error(STOP_ERROR_MESSAGE);
  }
}

async function sleepWithStop(ms) {
  const start = Date.now();
  while (Date.now() - start < ms) {
    throwIfStopped();
    await new Promise(r => setTimeout(r, Math.min(100, ms - (Date.now() - start))));
  }
}

async function humanStepDelay(min = HUMAN_STEP_DELAY_MIN, max = HUMAN_STEP_DELAY_MAX) {
  const duration = Math.floor(Math.random() * (max - min + 1)) + min;
  await sleepWithStop(duration);
}

async function clickWithDebugger(tabId, rect) {
  if (!tabId) {
    throw new Error('No auth tab found for debugger click.');
  }
  if (!rect || !Number.isFinite(rect.centerX) || !Number.isFinite(rect.centerY)) {
    throw new Error('Step 8 debugger fallback needs a valid button position.');
  }

  const target = { tabId };
  try {
    await chrome.debugger.attach(target, '1.3');
  } catch (err) {
    throw new Error(
      `Debugger attach failed during step 8 fallback: ${err.message}. ` +
      'If DevTools is open on the auth tab, close it and retry.'
    );
  }

  try {
    const x = Math.round(rect.centerX);
    const y = Math.round(rect.centerY);

    await chrome.debugger.sendCommand(target, 'Page.bringToFront');
    await chrome.debugger.sendCommand(target, 'Input.dispatchMouseEvent', {
      type: 'mouseMoved',
      x,
      y,
      button: 'none',
      buttons: 0,
      clickCount: 0,
    });
    await chrome.debugger.sendCommand(target, 'Input.dispatchMouseEvent', {
      type: 'mousePressed',
      x,
      y,
      button: 'left',
      buttons: 1,
      clickCount: 1,
    });
    await chrome.debugger.sendCommand(target, 'Input.dispatchMouseEvent', {
      type: 'mouseReleased',
      x,
      y,
      button: 'left',
      buttons: 0,
      clickCount: 1,
    });
  } finally {
    await chrome.debugger.detach(target).catch(() => {});
  }
}

async function broadcastStopToContentScripts() {
  const registry = await getTabRegistry();
  for (const entry of Object.values(registry)) {
    if (!entry?.tabId) continue;
    try {
      await chrome.tabs.sendMessage(entry.tabId, {
        type: 'STOP_FLOW',
        source: 'background',
        payload: {},
      });
    } catch {}
  }
}

let stopRequested = false;

// ============================================================
// Message Handler (central router)
// ============================================================

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  console.log(LOG_PREFIX, `Received: ${message.type} from ${message.source || 'sidepanel'}`, message);

  handleMessage(message, sender).then(response => {
    sendResponse(response);
  }).catch(err => {
    console.error(LOG_PREFIX, 'Handler error:', err);
    sendResponse({ error: err.message });
  });

  return true; // async response
});

async function handleMessage(message, sender) {
  switch (message.type) {
    case 'CONTENT_SCRIPT_READY': {
      const tabId = sender.tab?.id;
      if (tabId && message.source) {
        await registerTab(message.source, tabId);
        flushCommand(message.source, tabId);
        await addLog(`Content script ready: ${message.source} (tab ${tabId})`);
      }
      return { ok: true };
    }

    case 'LOG': {
      const { message: msg, level } = message.payload;
      await addLog(`[${message.source}] ${msg}`, level);
      return { ok: true };
    }

    case 'STEP_COMPLETE': {
      if (stopRequested) {
        await setStepStatus(message.step, 'stopped');
        notifyStepError(message.step, STOP_ERROR_MESSAGE);
        return { ok: true };
      }
      await setStepStatus(message.step, 'completed');
      await addLog(`Step ${message.step} completed`, 'ok');
      await handleStepData(message.step, message.payload);
      notifyStepComplete(message.step, message.payload);
      return { ok: true };
    }

    case 'STEP_ERROR': {
      if (isStopError(message.error)) {
        await setStepStatus(message.step, 'stopped');
        await addLog(`Step ${message.step} stopped by user`, 'warn');
        notifyStepError(message.step, message.error);
      } else {
        await setStepStatus(message.step, 'failed');
        await addLog(`Step ${message.step} failed: ${message.error}`, 'error');
        notifyStepError(message.step, message.error);
      }
      return { ok: true };
    }

    case 'GET_STATE': {
      return await getState();
    }

    case 'RESET': {
      clearStopRequest();
      await resetState();
      await addLog('Flow reset', 'info');
      return { ok: true };
    }

    case 'EXECUTE_STEP': {
      clearStopRequest();
      const step = message.payload.step;
      // Save email if provided (from side panel step 3)
      if (message.payload.email) {
        await setManualEmailState(message.payload.email);
      }
      await executeStep(step);
      return { ok: true };
    }

    case 'SKIP_STEP': {
      const step = message.payload?.step;
      await skipStep(step);
      return { ok: true };
    }

    case 'AUTO_RUN': {
      clearStopRequest();
      const totalRuns = message.payload?.totalRuns || 1;
      autoRunLoop(totalRuns);  // fire-and-forget
      return { ok: true };
    }

    case 'RESUME_AUTO_RUN': {
      clearStopRequest();
      if (message.payload.email) {
        await setManualEmailState(message.payload.email);
      }
      await continueAutoRun();
      return { ok: true };
    }

    case 'CONTINUE_AUTO_RUN': {
      clearStopRequest();
      if (message.payload.email) {
        await setManualEmailState(message.payload.email);
      }
      await continueAutoRun();
      return { ok: true };
    }

    case 'SAVE_SETTING': {
      const updates = {};
      if (message.payload.language !== undefined) updates.language = message.payload.language;
      if (message.payload.oauthProvider !== undefined) updates.oauthProvider = normalizeOauthProvider(message.payload.oauthProvider);
      if (message.payload.vpsUrl !== undefined) updates.vpsUrl = message.payload.vpsUrl;
      if (message.payload.sub2apiBaseUrl !== undefined) {
        updates.sub2apiBaseUrl = message.payload.sub2apiBaseUrl;
        updates.sub2apiRuntimeCredential = '';
      }
      if (message.payload.sub2apiAdminApiKey !== undefined) {
        updates.sub2apiAdminApiKey = message.payload.sub2apiAdminApiKey;
        updates.sub2apiRuntimeCredential = '';
      }
      if (message.payload.autoDeleteUsedIcloudAlias !== undefined) updates.autoDeleteUsedIcloudAlias = Boolean(message.payload.autoDeleteUsedIcloudAlias);
      if (message.payload.customPassword !== undefined) updates.customPassword = message.payload.customPassword;
      if (message.payload.mailProvider !== undefined) updates.mailProvider = normalizeMailProvider(message.payload.mailProvider);
      if (message.payload.inbucketHost !== undefined) updates.inbucketHost = message.payload.inbucketHost;
      if (message.payload.inbucketMailbox !== undefined) updates.inbucketMailbox = message.payload.inbucketMailbox;
      if (message.payload.microsoftManagerUrl !== undefined) updates.microsoftManagerUrl = message.payload.microsoftManagerUrl;
      if (message.payload.microsoftManagerToken !== undefined) updates.microsoftManagerToken = message.payload.microsoftManagerToken;
      if (message.payload.microsoftManagerMode !== undefined) updates.microsoftManagerMode = normalizeMicrosoftManagerMode(message.payload.microsoftManagerMode);
      if (message.payload.microsoftManagerKeyword !== undefined) updates.microsoftManagerKeyword = message.payload.microsoftManagerKeyword;
      await setState(updates);
      return { ok: true };
    }

    // Side panel data updates
    case 'SAVE_EMAIL': {
      await setManualEmailState(message.payload.email);
      return { ok: true, email: message.payload.email };
    }

    case 'FETCH_AUTO_EMAIL': {
      clearStopRequest();
      const email = await fetchConfiguredEmail(message.payload || {});
      return { ok: true, email };
    }

    case 'CHECK_ICLOUD_SESSION': {
      clearStopRequest();
      return await checkIcloudSession();
    }

    case 'LIST_ICLOUD_ALIASES': {
      clearStopRequest();
      const aliases = await listIcloudAliases();
      return { ok: true, aliases };
    }

    case 'SET_ICLOUD_ALIAS_USED_STATE': {
      clearStopRequest();
      const result = await setIcloudAliasUsedState(message.payload || {});
      return { ok: true, ...result };
    }

    case 'DELETE_ICLOUD_ALIAS': {
      clearStopRequest();
      const result = await deleteIcloudAlias(message.payload);
      return { ok: true, ...result };
    }

    case 'DELETE_USED_ICLOUD_ALIASES': {
      clearStopRequest();
      const result = await deleteUsedIcloudAliases();
      return { ok: true, ...result };
    }

    case 'STOP_FLOW': {
      await requestStop();
      return { ok: true };
    }

    default:
      console.warn(LOG_PREFIX, `Unknown message type: ${message.type}`);
      return { error: `Unknown message type: ${message.type}` };
  }
}

// ============================================================
// Step Data Handlers
// ============================================================

async function handleStepData(step, payload) {
  switch (step) {
    case 1:
      if (payload.oauthUrl) {
        await setState({ oauthUrl: payload.oauthUrl });
        broadcastDataUpdate({ oauthUrl: payload.oauthUrl });
      }
      break;
    case 3:
      if (payload.email) await setEmailState(payload.email);
      break;
    case 4:
      if (payload.emailTimestamp) await setState({ lastEmailTimestamp: payload.emailTimestamp });
      break;
    case 8:
      if (payload.localhostUrl) {
        await setState({ localhostUrl: payload.localhostUrl });
        broadcastDataUpdate({ localhostUrl: payload.localhostUrl });
      }
      break;
    case 9:
      await recordCompletedAccount();
      break;
  }
}

// ============================================================
// Step Completion Waiting
// ============================================================

// Map of step -> { resolve, reject } for waiting on step completion
const stepWaiters = new Map();
let resumeWaiter = null;

function waitForStepComplete(step, timeoutMs = 120000) {
  return new Promise((resolve, reject) => {
    throwIfStopped();
    const timer = setTimeout(() => {
      stepWaiters.delete(step);
      reject(new Error(`Step ${step} timed out after ${timeoutMs / 1000}s`));
    }, timeoutMs);

    stepWaiters.set(step, {
      resolve: (data) => { clearTimeout(timer); stepWaiters.delete(step); resolve(data); },
      reject: (err) => { clearTimeout(timer); stepWaiters.delete(step); reject(err); },
    });
  });
}

function notifyStepComplete(step, payload) {
  const waiter = stepWaiters.get(step);
  if (waiter) waiter.resolve(payload);
}

function notifyStepError(step, error) {
  const waiter = stepWaiters.get(step);
  if (waiter) waiter.reject(new Error(error));
}

async function markRunningStepsStopped() {
  const state = await getState();
  const runningSteps = Object.entries(state.stepStatuses || {})
    .filter(([, status]) => status === 'running')
    .map(([step]) => Number(step));

  for (const step of runningSteps) {
    await setStepStatus(step, 'stopped');
  }
}

async function requestStop() {
  if (stopRequested) return;

  stopRequested = true;
  cancelPendingCommands();
  if (webNavListener) {
    chrome.webNavigation.onBeforeNavigate.removeListener(webNavListener);
    webNavListener = null;
  }

  await addLog('Stop requested. Cancelling current operations...', 'warn');
  await broadcastStopToContentScripts();

  for (const waiter of stepWaiters.values()) {
    waiter.reject(new Error(STOP_ERROR_MESSAGE));
  }
  stepWaiters.clear();

  if (resumeWaiter) {
    resumeWaiter.reject(new Error(STOP_ERROR_MESSAGE));
    resumeWaiter = null;
  }

  await markRunningStepsStopped();
  autoRunActive = false;
  autoRunPausedPhase = 'stopped';
  await syncAutoRunState({ autoRunning: false });
  chrome.runtime.sendMessage({
    type: 'AUTO_RUN_STATUS',
    payload: { phase: 'stopped', currentRun: autoRunCurrentRun, totalRuns: autoRunTotalRuns },
  }).catch(() => {});
}

// ============================================================
// Step Execution
// ============================================================

async function executeStep(step) {
  console.log(LOG_PREFIX, `Executing step ${step}`);
  throwIfStopped();
  await setStepStatus(step, 'running');
  await addLog(`Step ${step} started`);
  await humanStepDelay();

  const state = await getState();

  // Set flow start time on first step
  if (step === 1 && !state.flowStartTime) {
    const flowStartTime = Date.now();
    await setState({ flowStartTime });
    broadcastDataUpdate({ flowStartTime });
  }

  try {
    switch (step) {
      case 1: await executeStep1(state); break;
      case 2: await executeStep2(state); break;
      case 3: await executeStep3(state); break;
      case 4: await executeStep4(state); break;
      case 5: await executeStep5(state); break;
      case 6: await executeStep6(state); break;
      case 7: await executeStep7(state); break;
      case 8: await executeStep8(state); break;
      case 9: await executeStep9(state); break;
      case 10: await executeStep10(state); break;
      default:
        throw new Error(`Unknown step: ${step}`);
    }
  } catch (err) {
    if (isStopError(err)) {
      await setStepStatus(step, 'stopped');
      await addLog(`Step ${step} stopped by user`, 'warn');
      throw err;
    }
    await setStepStatus(step, 'failed');
    await addLog(`Step ${step} failed: ${err.message}`, 'error');
    throw err;
  }
}

/**
 * Execute a step and wait for it to complete before returning.
 * @param {number} step
 * @param {number} delayAfter - ms to wait after completion (for page transitions)
 */
async function executeStepAndWait(step, delayAfter = 2000) {
  throwIfStopped();
  const promise = waitForStepComplete(step, 120000);
  await executeStep(step);
  await promise;
  // Extra delay for page transitions / DOM updates
  if (delayAfter > 0) {
    await sleepWithStop(delayAfter + Math.floor(Math.random() * 1200));
  }
}

// ============================================================
// Auto Run Flow
// ============================================================

let autoRunActive = false;
let autoRunCurrentRun = 0;
let autoRunTotalRuns = 1;
let autoRunPausedPhase = null;

function getAutoRunStatusMessage(phase, currentRun, totalRuns) {
  return { type: 'AUTO_RUN_STATUS', payload: { phase, currentRun, totalRuns } };
}

async function syncAutoRunState(overrides = {}) {
  await setState({
    autoRunning: autoRunActive,
    autoRunCurrentRun,
    autoRunTotalRuns,
    autoRunPausedPhase,
    ...overrides,
  });
}

function getAutoStepDelay(step) {
  const delayMap = {
    1: 2000,
    2: 2000,
    3: 3000,
    4: 2000,
    5: 3000,
    6: 3000,
    7: 2000,
    8: 2000,
    9: 1000,
    10: 800,
  };
  return delayMap[step] || 0;
}

async function waitForSignupSurface(payload, timeout = 20000) {
  const startedAt = Date.now();
  let lastError = null;

  while (Date.now() - startedAt < timeout) {
    throwIfStopped();
    const tabId = await getTabId('signup-page');
    if (!tabId) {
      throw new Error('Signup/auth tab is not available.');
    }

    try {
      const response = await chrome.tabs.sendMessage(tabId, {
        type: 'WAIT_FOR_SURFACE',
        source: 'background',
        payload: {
          timeout: Math.min(5000, timeout),
          ...payload,
        },
      });
      if (response?.error) {
        throw new Error(response.error);
      }
      return response;
    } catch (err) {
      lastError = err;
      await sleepWithStop(250);
    }
  }

  throw new Error(`Signup page surface wait failed: ${getErrorMessage(lastError)}`);
}

function getAutoResumeStep(state) {
  const statuses = state?.stepStatuses || {};

  for (let step = 1; step <= TOTAL_STEPS; step++) {
    const status = statuses[step];
    if (status === 'failed' || status === 'stopped' || status === 'running') {
      return step;
    }
  }

  for (let step = 1; step <= TOTAL_STEPS; step++) {
    if (statuses[step] === 'pending') {
      return step;
    }
  }

  return null;
}

async function ensureAutoRunEmailReady(run, totalRuns) {
  const currentState = await getState();
  if (currentState.email) return;

  let emailReady = false;

  try {
    const email = await fetchConfiguredEmail({ generateNew: true });
    await addLog(`=== Run ${run}/${totalRuns} — Email ready: ${email} ===`, 'ok');
    emailReady = true;
  } catch (err) {
    await addLog(`Auto email fetch failed: ${err.message}`, 'warn');
  }

  if (!emailReady) {
    const pauseHint = 'Generate or paste an email address, then continue';
    await addLog(`=== Run ${run}/${totalRuns} PAUSED: ${pauseHint} ===`, 'warn');
    autoRunPausedPhase = 'waiting_email';
    await syncAutoRunState();
    chrome.runtime.sendMessage(getAutoRunStatusMessage('waiting_email', run, totalRuns)).catch(() => {});

    await waitForResume();

    autoRunPausedPhase = null;
    await syncAutoRunState();

    const resumedState = await getState();
    if (!resumedState.email) {
      throw new Error('Cannot resume: no email address.');
    }
  }
}

async function executeAutoRunSteps(run, totalRuns, options = {}) {
  const { resumeFromCurrentState = false } = options;
  const state = await getState();
  let startStep = resumeFromCurrentState ? getAutoResumeStep(state) : 1;

  if (!startStep) return;

  if (resumeFromCurrentState) {
    await addLog(`=== Run ${run}/${totalRuns} — Resuming from step ${startStep} ===`, 'warn');
  }

  if (startStep <= 2) {
    await addLog(`=== Auto Run ${run}/${totalRuns} — Phase 1: Get OAuth link & open signup ===`, 'info');
    chrome.runtime.sendMessage(getAutoRunStatusMessage('running', run, totalRuns)).catch(() => {});

    for (let step = startStep; step <= 2; step++) {
      const currentState = await getState();
      const stepStatus = currentState.stepStatuses?.[step];
      if (stepStatus === 'completed' || stepStatus === 'skipped') continue;
      await executeStepAndWait(step, getAutoStepDelay(step));
    }
    startStep = 3;
  }

  let needsPhase2 = false;
  for (let step = startStep; step <= TOTAL_STEPS; step++) {
    const stepStatus = (await getState()).stepStatuses?.[step];
    if (stepStatus !== 'completed' && stepStatus !== 'skipped') {
      needsPhase2 = true;
      break;
    }
  }

  if (!needsPhase2) return;

  await addLog(`=== Run ${run}/${totalRuns} — Phase 2: Register, verify, login, complete ===`, 'info');
  chrome.runtime.sendMessage(getAutoRunStatusMessage('running', run, totalRuns)).catch(() => {});

  await ensureAutoRunEmailReady(run, totalRuns);

  const signupTabId = await getTabId('signup-page');
  if (signupTabId) {
    await chrome.tabs.update(signupTabId, { active: true });
  }

  for (let step = Math.max(3, startStep); step <= TOTAL_STEPS; step++) {
    const currentState = await getState();
    const stepStatus = currentState.stepStatuses?.[step];
    if (stepStatus === 'completed' || stepStatus === 'skipped') continue;
    if (step === 3 && !currentState.email) {
      await ensureAutoRunEmailReady(run, totalRuns);
    }
    await executeStepAndWait(step, getAutoStepDelay(step));
  }
}

// Outer loop: runs the full flow N times
async function autoRunLoop(totalRuns, options = {}) {
  const { startRun = 1, preserveCurrentRunState = false } = options;
  if (autoRunActive) {
    await addLog('Auto run already in progress', 'warn');
    return;
  }

  clearStopRequest();
  autoRunActive = true;
  autoRunTotalRuns = totalRuns;
  autoRunPausedPhase = null;
  await syncAutoRunState({ autoRunning: true });

  for (let run = startRun; run <= totalRuns; run++) {
    autoRunCurrentRun = run;
    autoRunPausedPhase = null;
    await syncAutoRunState();

    const isResumingRun = preserveCurrentRunState && run === startRun;
    if (!isResumingRun) {
      // Reset everything at the start of each run (keep CPA Auth/mail settings)
      const prevState = await getState();
      const keepSettings = {
        oauthProvider: normalizeOauthProvider(prevState.oauthProvider),
        vpsUrl: prevState.vpsUrl,
        sub2apiBaseUrl: prevState.sub2apiBaseUrl,
        sub2apiAdminApiKey: prevState.sub2apiAdminApiKey,
        mailProvider: prevState.mailProvider,
        inbucketHost: prevState.inbucketHost,
        inbucketMailbox: prevState.inbucketMailbox,
        microsoftManagerUrl: prevState.microsoftManagerUrl,
        microsoftManagerToken: prevState.microsoftManagerToken,
        microsoftManagerMode: prevState.microsoftManagerMode,
        microsoftManagerKeyword: prevState.microsoftManagerKeyword,
        autoRunning: true,
        autoRunCurrentRun: run,
        autoRunTotalRuns: totalRuns,
        autoRunPausedPhase: null,
      };
      await resetState();
      await setState(keepSettings);
      // Tell side panel to reset all UI
      chrome.runtime.sendMessage({ type: 'AUTO_RUN_RESET' }).catch(() => {});
      await sleepWithStop(500);
    }

    try {
      throwIfStopped();
      chrome.runtime.sendMessage(getAutoRunStatusMessage('running', run, totalRuns)).catch(() => {});

      await executeAutoRunSteps(run, totalRuns, { resumeFromCurrentState: isResumingRun });

      await addLog(`=== Run ${run}/${totalRuns} COMPLETE! ===`, 'ok');

    } catch (err) {
      if (isStopError(err)) {
        await addLog(`Run ${run}/${totalRuns} stopped by user`, 'warn');
        chrome.runtime.sendMessage(getAutoRunStatusMessage('stopped', run, totalRuns)).catch(() => {});
      } else {
        autoRunPausedPhase = 'error';
        autoRunActive = false;
        await syncAutoRunState({ autoRunning: false });
        await addLog(`Run ${run}/${totalRuns} failed: ${err.message}`, 'error');
        chrome.runtime.sendMessage(getAutoRunStatusMessage('error', run, totalRuns)).catch(() => {});
        clearStopRequest();
        return;
      }
      break; // Stop on error
    }
  }

  const completedRuns = autoRunCurrentRun;
  if (stopRequested) {
    await addLog(`=== Stopped after ${Math.max(0, completedRuns - 1)}/${autoRunTotalRuns} runs ===`, 'warn');
    chrome.runtime.sendMessage(getAutoRunStatusMessage('stopped', completedRuns, autoRunTotalRuns)).catch(() => {});
  } else if (completedRuns >= autoRunTotalRuns) {
    await addLog(`=== All ${autoRunTotalRuns} runs completed successfully ===`, 'ok');
    chrome.runtime.sendMessage(getAutoRunStatusMessage('complete', completedRuns, autoRunTotalRuns)).catch(() => {});
  } else {
    await addLog(`=== Stopped after ${completedRuns}/${autoRunTotalRuns} runs ===`, 'warn');
    chrome.runtime.sendMessage(getAutoRunStatusMessage('stopped', completedRuns, autoRunTotalRuns)).catch(() => {});
  }
  autoRunActive = false;
  autoRunPausedPhase = null;
  await syncAutoRunState({ autoRunning: false, autoRunPausedPhase: null });
  clearStopRequest();
}

function waitForResume() {
  return new Promise((resolve, reject) => {
    throwIfStopped();
    resumeWaiter = { resolve, reject };
  });
}

async function resumeAutoRun() {
  throwIfStopped();
  const state = await getState();
  if (!state.email) {
    await addLog('Cannot resume: no email address. Paste email in Side Panel first.', 'error');
    return;
  }
  if (resumeWaiter) {
    autoRunPausedPhase = null;
    await syncAutoRunState();
    resumeWaiter.resolve();
    resumeWaiter = null;
  }
}

async function continueAutoRun() {
  const state = await getState();

  if (resumeWaiter) {
    await resumeAutoRun();
    return;
  }

  if (autoRunActive) {
    throw new Error('Auto run is already active.');
  }

  const currentRun = Number(state.autoRunCurrentRun || autoRunCurrentRun || 0);
  const totalRuns = Number(state.autoRunTotalRuns || autoRunTotalRuns || 1);
  if (!currentRun) {
    throw new Error('No interrupted auto run found.');
  }

  const resumeStep = getAutoResumeStep(state);
  const startRun = resumeStep ? currentRun : currentRun + 1;
  if (startRun > totalRuns) {
    throw new Error('There is no interrupted auto run to continue.');
  }

  autoRunLoop(totalRuns, {
    startRun,
    preserveCurrentRunState: Boolean(resumeStep),
  }); // fire-and-forget
}

// ============================================================
// Step 1: Get OAuth Link (via vps-panel.js)
// ============================================================

async function executeStep1(state) {
  if (isSub2apiOauthProvider(state)) {
    await executeStep1WithSub2api(state);
    return;
  }

  if (!state.vpsUrl) {
    throw new Error('No CPA Auth URL configured. Enter the CPA Auth address in Side Panel first.');
  }
  await addLog('Step 1: Opening CPA Auth panel...');
  await reuseOrCreateTab('vps-panel', state.vpsUrl, {
    inject: ['content/utils.js', 'content/vps-panel.js'],
    reloadIfSameUrl: true,
  });

  await sendToContentScript('vps-panel', {
    type: 'EXECUTE_STEP',
    step: 1,
    source: 'background',
    payload: {},
  });
}

async function executeStep1WithSub2api(state) {
  await addLog('Step 1: Requesting OAuth URL from Sub2API...');

  let authCredential = String(state.sub2apiAdminApiKey || state.sub2apiRuntimeCredential || '').trim();
  if (!authCredential) {
    authCredential = await resolveSub2apiCredentialFromDashboard(state);
  }

  let data;
  try {
    data = await requestSub2apiAdminApi(state, 'admin/openai/generate-auth-url', {
      method: 'POST',
      body: {},
      authCredential,
    });
  } catch (err) {
    const message = getErrorMessage(err);
    const isAuthError = /invalid admin api key|invalid token|authorization required|unauthorized|forbidden/i.test(message);
    if (!isAuthError) {
      throw err;
    }

    await addLog('Step 1: Sub2API auth failed, trying dashboard session token fallback...', 'warn');
    const fallbackCredential = await resolveSub2apiCredentialFromDashboard({
      ...state,
      sub2apiRuntimeCredential: '',
    });
    data = await requestSub2apiAdminApi(state, 'admin/openai/generate-auth-url', {
      method: 'POST',
      body: {},
      authCredential: fallbackCredential,
    });
  }

  const oauthUrl = String(data?.auth_url || '').trim();
  const sessionId = String(data?.session_id || '').trim();
  if (!oauthUrl || !sessionId) {
    throw new Error('Sub2API did not return auth_url/session_id.');
  }

  await setState({ sub2apiSessionId: sessionId });
  await addLog(`Step 1: Sub2API OAuth URL ready (${oauthUrl.slice(0, 80)}...)`, 'ok');
  await completeBackgroundStep(1, { oauthUrl });
}

// ============================================================
// Step 2: Open Signup Page (Background opens tab, signup-page.js clicks Register)
// ============================================================

async function executeStep2(state) {
  if (!state.oauthUrl) {
    throw new Error('No OAuth URL. Complete step 1 first.');
  }
  await addLog(`Step 2: Opening auth URL...`);
  await reuseOrCreateTab('signup-page', state.oauthUrl);

  await sendToContentScript('signup-page', {
    type: 'EXECUTE_STEP',
    step: 2,
    source: 'background',
    payload: {},
  });
}

// ============================================================
// Step 3: Fill Email & Password (via signup-page.js)
// ============================================================

async function executeStep3(state) {
  if (!state.email) {
    throw new Error('No email address. Paste email in Side Panel first.');
  }

  const password = state.customPassword || generatePassword();
  await setPasswordState(password);

  await addLog(
    `Step 3: Filling email ${state.email}, password ${state.customPassword ? 'customized' : 'generated'} (${password.length} chars)`
  );
  await sendToContentScript('signup-page', {
    type: 'EXECUTE_STEP',
    step: 3,
    source: 'background',
    payload: { email: state.email, password },
  });

  await waitForSignupSurface({
    step: 3,
    selectors: [
      'input[name="code"]',
      'input[name="otp"]',
      'input[maxlength="1"]',
      'input[name="name"]',
      'input[placeholder*="全名"]',
      '[role="spinbutton"][data-type="year"]',
      'input[name="age"]',
    ],
  });
}

// ============================================================
// Step 4: Get Signup Verification Code (Microsoft Account Manager API)
// ============================================================

function getMailConfig(state) {
  const provider = normalizeMailProvider(state.mailProvider);
  if (provider !== 'microsoft-manager') {
    return { error: 'Only Microsoft Account Manager API is supported in this build.' };
  }

  const managerUrl = normalizeMicrosoftManagerBaseUrl(state.microsoftManagerUrl);
  const managerToken = String(state.microsoftManagerToken || '').trim();
  const mode = normalizeMicrosoftManagerMode(state.microsoftManagerMode);

  if (!managerUrl) {
    return { error: 'Microsoft Account Manager URL is empty or invalid.' };
  }
  if (!managerToken) {
    return { error: 'Microsoft Account Manager token is empty.' };
  }

  return {
    source: 'microsoft-manager',
    label: `Microsoft Account Manager (${mode.toUpperCase()})`,
    mode,
    usesApi: true,
  };
}

function normalizeInbucketOrigin(rawValue) {
  const value = (rawValue || '').trim();
  if (!value) return '';

  const candidate = /^[a-zA-Z][a-zA-Z\d+\-.]*:\/\//.test(value) ? value : `https://${value}`;

  try {
    const parsed = new URL(candidate);
    return parsed.origin;
  } catch {
    return '';
  }
}

function normalizeSub2apiApiRoot(rawValue) {
  const value = String(rawValue || '').trim();
  if (!value) return '';

  const candidate = /^[a-zA-Z][a-zA-Z\d+\-.]*:\/\//.test(value) ? value : `https://${value}`;

  try {
    const parsed = new URL(candidate);
    let apiPath = parsed.pathname.replace(/\/+$/, '') || '';

    // Users may paste dashboard URLs like /admin or /admin/acc.
    // Convert them back to API root automatically.
    if (/^\/admin(?:\/|$)/i.test(apiPath)) {
      apiPath = '';
    }

    // If path already contains /api/v1, keep the prefix only.
    const apiV1Match = apiPath.match(/^(.*?\/api\/v1)(?:\/.*)?$/i);
    if (apiV1Match) {
      apiPath = apiV1Match[1];
    }

    if (!apiPath || apiPath === '/') {
      apiPath = '/api/v1';
    } else if (!/\/api\/v1$/i.test(apiPath)) {
      apiPath = `${apiPath}/api/v1`;
    }

    return `${parsed.origin}${apiPath}`;
  } catch {
    return '';
  }
}

function normalizeSub2apiDashboardUrl(rawValue) {
  const value = String(rawValue || '').trim();
  if (!value) return '';

  const candidate = /^[a-zA-Z][a-zA-Z\d+\-.]*:\/\//.test(value) ? value : `https://${value}`;

  try {
    const parsed = new URL(candidate);
    let path = parsed.pathname.replace(/\/+$/, '') || '';

    if (/\/api\/v1/i.test(path)) {
      path = '';
    }

    if (!/^\/admin(?:\/|$)/i.test(path)) {
      path = '/admin/accounts';
    }

    return `${parsed.origin}${path}`;
  } catch {
    return '';
  }
}

async function resolveSub2apiCredentialFromDashboard(state) {
  const dashboardUrl = normalizeSub2apiDashboardUrl(state.sub2apiBaseUrl);
  if (!dashboardUrl) {
    throw new Error('Sub2API URL is empty or invalid.');
  }

  await addLog('Sub2API: No API credential set, trying to reuse admin login session from dashboard...', 'warn');

  await reuseOrCreateTab('sub2api-panel', dashboardUrl, {
    inject: ['content/utils.js', 'content/sub2api-panel.js'],
    injectSource: 'sub2api-panel',
  });

  const response = await sendToContentScript('sub2api-panel', {
    type: 'GET_SUB2API_AUTH_TOKEN',
    source: 'background',
    payload: {},
  });

  if (response?.error) {
    throw new Error(response.error);
  }

  const token = String(response?.token || '').trim();
  if (!token) {
    throw new Error('No Sub2API admin token found. Please log in to Sub2API admin page first, then retry.');
  }

  const credential = /^Bearer\s+/i.test(token) ? token : `Bearer ${token}`;
  await setState({ sub2apiRuntimeCredential: credential });
  await addLog('Sub2API: Admin login token detected and will be used for API requests.', 'ok');
  return credential;
}

function buildSub2apiApiUrl(apiRoot, path) {
  const root = String(apiRoot || '').replace(/\/+$/, '');
  const normalizedPath = String(path || '').replace(/^\/+/, '');
  return `${root}/${normalizedPath}`;
}

async function requestSub2apiAdminApi(state, path, options = {}) {
  const apiRoot = normalizeSub2apiApiRoot(state.sub2apiBaseUrl);
  if (!apiRoot) {
    throw new Error('Sub2API URL is empty or invalid.');
  }

  const adminApiCredential = String(
    options.authCredential
      || state.sub2apiAdminApiKey
      || state.sub2apiRuntimeCredential
      || ''
  ).trim();

  const url = buildSub2apiApiUrl(apiRoot, path);
  const baseHeaders = {
    'Content-Type': 'application/json',
    ...(options.headers || {}),
  };

  async function sendSub2apiRequest(headers) {
    const init = {
      method: options.method || 'GET',
      headers,
    };

    if (options.body !== undefined) {
      init.body = JSON.stringify(options.body);
    }

    let response;
    try {
      response = await fetch(url, init);
    } catch (err) {
      throw new Error(`Sub2API request failed: ${getErrorMessage(err)}`);
    }

    let payload;
    try {
      payload = await response.json();
    } catch {
      payload = null;
    }

    const responseCode = Number(payload?.code);
    const responseMessage = String(payload?.message || payload?.detail || '').trim();
    const responseData = payload && typeof payload === 'object' && 'data' in payload ? payload.data : payload;

    if (!response.ok) {
      const fallbackMessage = `${response.status} ${response.statusText}`.trim();
      throw new Error(`Sub2API request failed: ${responseMessage || fallbackMessage}`);
    }

    if (Number.isFinite(responseCode) && responseCode !== 0) {
      throw new Error(`Sub2API request failed: ${responseMessage || `code ${responseCode}`}`);
    }

    return responseData;
  }

  if (!adminApiCredential) {
    return sendSub2apiRequest({ ...baseHeaders });
  }

  if (/^Bearer\s+/i.test(adminApiCredential)) {
    return sendSub2apiRequest({
      ...baseHeaders,
      Authorization: adminApiCredential,
    });
  }

  try {
    return await sendSub2apiRequest({
      ...baseHeaders,
      'x-api-key': adminApiCredential,
    });
  } catch (err) {
    const message = getErrorMessage(err);
    if (!/invalid admin api key/i.test(message)) {
      throw err;
    }

    await addLog('Sub2API auth fallback: x-api-key rejected, retrying as Bearer token...', 'warn');

    return sendSub2apiRequest({
      ...baseHeaders,
      Authorization: `Bearer ${adminApiCredential}`,
    });
  }
}

function parseOAuthCallbackParams(callbackUrl) {
  const value = String(callbackUrl || '').trim();
  if (!value) {
    throw new Error('No callback URL. Complete step 8 first.');
  }

  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    throw new Error('Callback URL is invalid.');
  }

  const code = String(parsed.searchParams.get('code') || '').trim();
  const state = String(parsed.searchParams.get('state') || '').trim();
  const oauthError = String(parsed.searchParams.get('error') || '').trim();
  const oauthErrorDescription = String(parsed.searchParams.get('error_description') || '').trim();

  if (oauthError) {
    const detail = oauthErrorDescription || oauthError;
    throw new Error(
      `OAuth callback returned error: ${detail}. Usually this means OAuth session/CSRF mismatch. Please rerun from step 1 to get a fresh OAuth URL, then continue.`
    );
  }

  if (!code || !state) {
    throw new Error('Callback URL is missing code/state query parameters. Please rerun from step 1 and retry.');
  }

  return {
    code,
    state,
    callbackUrl: parsed.toString(),
    redirectUri: `${parsed.origin}${parsed.pathname}`,
  };
}

async function completeBackgroundStep(step, payload = {}) {
  await setStepStatus(step, 'completed');
  await addLog(`Step ${step} completed`, 'ok');
  await handleStepData(step, payload);
  notifyStepComplete(step, payload);
}

async function pollVerificationCodeWithAutoResend(options) {
  const {
    step,
    mail,
    pollPayload,
    successSelectors,
    successMessage,
    resendRounds = 3,
    customPoll = null,
  } = options;

  let currentFilterAfter = pollPayload.filterAfterTimestamp || 0;

  for (let round = 1; round <= resendRounds; round++) {
    const currentPayload = {
      ...pollPayload,
      filterAfterTimestamp: currentFilterAfter,
    };

    const result = customPoll
      ? await customPoll(currentPayload)
      : await sendToContentScript(mail.source, {
          type: 'POLL_EMAIL',
          step,
          source: 'background',
          payload: currentPayload,
        });

    if (result?.code) {
      await setState({ lastEmailTimestamp: result.emailTimestamp || Date.now() });
      await addLog(`Step ${step}: ${successMessage}: ${result.code}`);

      const signupTabId = await getTabId('signup-page');
      if (!signupTabId) {
        throw new Error(`Signup page tab was closed. Cannot fill step ${step} verification code.`);
      }

      await chrome.tabs.update(signupTabId, { active: true });
      await sendToContentScript('signup-page', {
        type: 'FILL_CODE',
        step,
        source: 'background',
        payload: { code: result.code },
      });
      await waitForSignupSurface({
        step,
        selectors: successSelectors,
      });
      return;
    }

    const pollError = result?.error || `No verification code returned for step ${step}.`;
    if (round >= resendRounds) {
      throw new Error(pollError);
    }

    await addLog(`Step ${step}: No verification email received on round ${round}/${resendRounds}. Trying to resend code...`, 'warn');

    const signupTabId = await getTabId('signup-page');
    if (!signupTabId) {
      throw new Error(`Signup page tab was closed. Cannot resend step ${step} verification code.`);
    }

    await chrome.tabs.update(signupTabId, { active: true });
    const resendResponse = await sendToContentScript('signup-page', {
      type: 'RESEND_VERIFICATION_CODE',
      step,
      source: 'background',
      payload: {},
    });

    if (resendResponse?.error) {
      throw new Error(resendResponse.error);
    }

    currentFilterAfter = resendResponse?.resentAt || Date.now();
    await addLog(`Step ${step}: Resend triggered. Waiting for a fresh verification email...`, 'info');
  }
}

async function executeStep4(state) {
  const pollPayload = {
    filterAfterTimestamp: state.flowStartTime || 0,
    senderFilters: ['openai', 'noreply', 'verify', 'auth', 'forward'],
    subjectFilters: ['verify', 'verification', 'code', '验证', 'confirm'],
    targetEmail: state.email,
    maxAttempts: 20,
    intervalMs: 3000,
  };

  const mail = getMailConfig(state);
  if (mail.error) throw new Error(mail.error);

  if (mail.usesApi) {
    await addLog(`Step 4: Polling verification code via ${mail.label}...`);
    await pollVerificationCodeWithAutoResend({
      step: 4,
      mail,
      pollPayload,
      successMessage: 'Got verification code',
      successSelectors: [
        'input[name="name"]',
        'input[placeholder*="全名"]',
        '[role="spinbutton"][data-type="year"]',
        'input[name="birthday"]',
        'input[name="age"]',
      ],
      customPoll: (currentPayload) => pollMicrosoftManagerCode(state, 4, currentPayload),
    });
    return;
  }

  await addLog(`Step 4: Opening ${mail.label}...`);

  const alive = await isTabAlive(mail.source);
  if (alive) {
    if (mail.navigateOnReuse) {
      await reuseOrCreateTab(mail.source, mail.url, {
        inject: mail.inject,
        injectSource: mail.injectSource,
      });
    } else {
      const tabId = await getTabId(mail.source);
      await chrome.tabs.update(tabId, { active: true });
    }
  } else {
    await reuseOrCreateTab(mail.source, mail.url, {
      inject: mail.inject,
      injectSource: mail.injectSource,
    });
  }

  await pollVerificationCodeWithAutoResend({
    step: 4,
    mail,
    pollPayload,
    successMessage: 'Got verification code',
    successSelectors: [
      'input[name="name"]',
      'input[placeholder*="全名"]',
      '[role="spinbutton"][data-type="year"]',
      'input[name="birthday"]',
      'input[name="age"]',
    ],
  });
}

// ============================================================
// Step 5: Fill Name & Birthday (via signup-page.js)
// ============================================================

async function executeStep5(state) {
  const { firstName, lastName } = generateRandomName();
  const { year, month, day } = generateRandomBirthday();

  await addLog(`Step 5: Generated name: ${firstName} ${lastName}, Birthday: ${year}-${month}-${day}`);

  await sendToContentScript('signup-page', {
    type: 'EXECUTE_STEP',
    step: 5,
    source: 'background',
    payload: { firstName, lastName, year, month, day },
  });
}

async function refreshOAuthIfTimedOutBeforeStep6(state) {
  if (isSub2apiOauthProvider(state)) {
    return state;
  }

  if (!state.vpsUrl) {
    return state;
  }

  await addLog('Step 6: Checking CPA Auth status before login...');

  try {
    await reuseOrCreateTab('vps-panel', state.vpsUrl, {
      inject: ['content/utils.js', 'content/vps-panel.js'],
      injectSource: 'vps-panel',
    });

    const response = await sendToContentScript('vps-panel', {
      type: 'CHECK_OAUTH_TIMEOUT_STATUS',
      source: 'background',
    });

    if (response?.error) {
      throw new Error(response.error);
    }

    if (!response?.timedOut) {
      return state;
    }

    const timeoutText = response.statusText || '认证失败: Timeout waiting for OAuth callback';
    await addLog(`Step 6: CPA Auth reported OAuth timeout. Refreshing OAuth link... (${timeoutText})`, 'warn');

    await executeStepAndWait(1, 1500);

    const refreshedState = await getState();
    if (!refreshedState.oauthUrl) {
      throw new Error('Step 1 completed but no new OAuth URL was saved.');
    }

    await addLog('Step 6: New OAuth link obtained after timeout. Continuing login...', 'ok');
    return refreshedState;
  } catch (err) {
    await addLog(`Step 6: CPA Auth timeout check could not be completed, continuing with current OAuth URL. ${getErrorMessage(err)}`, 'warn');
    return state;
  }
}

// ============================================================
// Step 6: Login ChatGPT (Background opens tab, chatgpt.js handles login)
// ============================================================

async function executeStep6(state) {
  state = await refreshOAuthIfTimedOutBeforeStep6(state);

  if (!state.oauthUrl) {
    throw new Error('No OAuth URL. Complete step 1 first.');
  }
  if (!state.email) {
    throw new Error('No email. Complete step 3 first.');
  }

  await addLog(`Step 6: Opening OAuth URL for login...`);
  // Reuse the signup-page tab — navigate it to the OAuth URL
  await reuseOrCreateTab('signup-page', state.oauthUrl);

  // signup-page.js will inject (same auth.openai.com domain) and handle login
  await sendToContentScript('signup-page', {
    type: 'EXECUTE_STEP',
    step: 6,
    source: 'background',
    payload: { email: state.email, password: state.password },
  });

  await waitForSignupSurface({
    step: 6,
    selectors: [
      'input[name="code"]',
      'input[name="otp"]',
      'input[maxlength="1"]',
      'button[type="submit"][data-dd-action-name="Continue"]',
      'button[type="submit"]._primary_3rdp0_107',
    ],
  });
}

// ============================================================
// Step 7: Get Login Verification Code (Microsoft Account Manager API)
// ============================================================

async function executeStep7(state) {
  const pollPayload = {
    filterAfterTimestamp: state.lastEmailTimestamp || state.flowStartTime || 0,
    senderFilters: ['openai', 'noreply', 'verify', 'auth', 'chatgpt', 'forward'],
    subjectFilters: ['verify', 'verification', 'code', '验证', 'confirm', 'login'],
    targetEmail: state.email,
    maxAttempts: 20,
    intervalMs: 3000,
  };

  const mail = getMailConfig(state);
  if (mail.error) throw new Error(mail.error);

  if (mail.usesApi) {
    await addLog(`Step 7: Polling login code via ${mail.label}...`);
    await pollVerificationCodeWithAutoResend({
      step: 7,
      mail,
      pollPayload,
      successMessage: 'Got login verification code',
      successSelectors: [
        'button[type="submit"][data-dd-action-name="Continue"]',
        'button[type="submit"]._primary_3rdp0_107',
        'button[aria-label*="Continue"]',
      ],
      customPoll: (currentPayload) => pollMicrosoftManagerCode(state, 7, currentPayload),
    });
    return;
  }

  await addLog(`Step 7: Opening ${mail.label}...`);

  const alive = await isTabAlive(mail.source);
  if (alive) {
    if (mail.navigateOnReuse) {
      await reuseOrCreateTab(mail.source, mail.url, {
        inject: mail.inject,
        injectSource: mail.injectSource,
      });
    } else {
      const tabId = await getTabId(mail.source);
      await chrome.tabs.update(tabId, { active: true });
    }
  } else {
    await reuseOrCreateTab(mail.source, mail.url, {
      inject: mail.inject,
      injectSource: mail.injectSource,
    });
  }

  await pollVerificationCodeWithAutoResend({
    step: 7,
    mail,
    pollPayload,
    successMessage: 'Got login verification code',
    successSelectors: [
      'button[type="submit"][data-dd-action-name="Continue"]',
      'button[type="submit"]._primary_3rdp0_107',
      'button[aria-label*="Continue"]',
    ],
  });
}

// ============================================================
// Step 8: Complete OAuth (auto click + localhost listener)
// ============================================================

let webNavListener = null;

function isLocalCallbackUrl(rawUrl) {
  try {
    const parsed = new URL(rawUrl);
    return ['localhost', '127.0.0.1', '[::1]', '::1'].includes(parsed.hostname);
  } catch {
    return false;
  }
}

async function executeStep8(state) {
  if (!state.oauthUrl) {
    throw new Error('No OAuth URL. Complete step 1 first.');
  }

  await addLog('Step 8: Setting up localhost redirect listener...');

  // Register webNavigation listener (scoped to this step)
  return new Promise((resolve, reject) => {
    let resolved = false;
    let resolveCaptureWait = null;
    const captureWait = new Promise((resolveCapture) => {
      resolveCaptureWait = resolveCapture;
    });

    const cleanupListener = () => {
      if (webNavListener) {
        chrome.webNavigation.onBeforeNavigate.removeListener(webNavListener);
        webNavListener = null;
      }
    };

    const timeout = setTimeout(() => {
      cleanupListener();
      reject(new Error('Localhost redirect not captured after 120s. Step 8 click may have been blocked.'));
    }, 120000);

    webNavListener = (details) => {
      if (!isLocalCallbackUrl(details.url) || resolved) {
        return;
      }

      resolved = true;
      console.log(LOG_PREFIX, `Captured localhost redirect: ${details.url}`);
      cleanupListener();
      clearTimeout(timeout);
      if (resolveCaptureWait) resolveCaptureWait(details.url);

      (async () => {
        try {
          await setState({ localhostUrl: details.url });

          let callbackError = '';
          let callbackErrorDescription = '';
          try {
            const parsed = new URL(details.url);
            callbackError = String(parsed.searchParams.get('error') || '').trim();
            callbackErrorDescription = String(parsed.searchParams.get('error_description') || '').trim();
          } catch {}

          if (callbackError) {
            const detail = callbackErrorDescription || callbackError;
            throw new Error(
              `OAuth callback returned error: ${detail}. Usually this means OAuth session/CSRF mismatch. Please rerun from step 1 to get a fresh OAuth URL, then continue.`
            );
          }

          await addLog(`Step 8: Captured localhost URL: ${details.url}`, 'ok');
          await setStepStatus(8, 'completed');
          notifyStepComplete(8, { localhostUrl: details.url });
          broadcastDataUpdate({ localhostUrl: details.url });
          resolve();
        } catch (err) {
          reject(err);
        }
      })();
    };

    chrome.webNavigation.onBeforeNavigate.addListener(webNavListener);

    // After step 7, the auth page shows a consent screen ("使用 ChatGPT 登录到 Codex")
    // with a "继续" button. We locate the button in-page, then click it through
    // the debugger Input API directly.
    (async () => {
      try {
        let signupTabId = await getTabId('signup-page');
        if (signupTabId) {
          await chrome.tabs.update(signupTabId, { active: true });
          await addLog('Step 8: Switched to auth page. Preparing debugger click...');
        } else {
          signupTabId = await reuseOrCreateTab('signup-page', state.oauthUrl);
          await addLog('Step 8: Auth tab reopened. Preparing debugger click...');
        }

        const clickResult = await sendToContentScript('signup-page', {
          type: 'STEP8_FIND_AND_CLICK',
          source: 'background',
          payload: {},
        });

        if (clickResult?.error) {
          throw new Error(clickResult.error);
        }

        if (!resolved) {
          await clickWithDebugger(signupTabId, clickResult?.rect);
          await addLog('Step 8: Debugger click dispatched, waiting for redirect...');
        }
      } catch (err) {
        clearTimeout(timeout);
        cleanupListener();
        reject(err);
      }
    })();
  });
}

// ============================================================
// Step 9: CPA Auth Verify (via vps-panel.js)
// ============================================================

async function executeStep9(state) {
  if (isSub2apiOauthProvider(state)) {
    await executeStep9WithSub2api(state);
    return;
  }

  if (!state.localhostUrl) {
    throw new Error('No localhost URL. Complete step 8 first.');
  }
  if (!state.vpsUrl) {
    throw new Error('CPA Auth URL not set. Please enter the CPA Auth URL in the side panel.');
  }

  await addLog('Step 9: Opening CPA Auth panel...');

  let tabId = await getTabId('vps-panel');
  const alive = tabId && await isTabAlive('vps-panel');

  if (!alive) {
    // Create new tab
    const tab = await chrome.tabs.create({ url: state.vpsUrl, active: true });
    tabId = tab.id;
    await new Promise(resolve => {
      const listener = (tid, info) => {
        if (tid === tabId && info.status === 'complete') {
          chrome.tabs.onUpdated.removeListener(listener);
          resolve();
        }
      };
      chrome.tabs.onUpdated.addListener(listener);
    });
  } else {
    await chrome.tabs.update(tabId, { active: true });
  }

  // Inject scripts directly and wait for them to be ready
  await chrome.scripting.executeScript({
    target: { tabId },
    files: ['content/utils.js', 'content/vps-panel.js'],
  });
  await new Promise(r => setTimeout(r, 1000));

  // Send command directly — bypass queue/ready mechanism
  await addLog(`Step 9: Filling callback URL...`);
  await chrome.tabs.sendMessage(tabId, {
    type: 'EXECUTE_STEP',
    step: 9,
    source: 'background',
    payload: { localhostUrl: state.localhostUrl },
  });
}

async function executeStep9WithSub2api(state) {
  if (!state.localhostUrl) {
    throw new Error('No localhost URL. Complete step 8 first.');
  }

  const sessionId = String(state.sub2apiSessionId || '').trim();
  if (!sessionId) {
    throw new Error('No Sub2API session_id. Please rerun step 1 first.');
  }

  const callback = parseOAuthCallbackParams(state.localhostUrl);
  let authCredential = String(state.sub2apiAdminApiKey || state.sub2apiRuntimeCredential || '').trim();
  if (!authCredential) {
    authCredential = await resolveSub2apiCredentialFromDashboard(state);
  }

  await addLog('Step 9: Importing account to Sub2API...');
  let created;
  try {
    created = await requestSub2apiAdminApi(state, 'admin/openai/create-from-oauth', {
      method: 'POST',
      authCredential,
      body: {
        session_id: sessionId,
        code: callback.code,
        state: callback.state,
        redirect_uri: callback.redirectUri,
        name: String(state.email || '').trim() || undefined,
      },
    });
  } catch (err) {
    const message = getErrorMessage(err);
    const isAuthError = /invalid admin api key|invalid token|authorization required|unauthorized|forbidden/i.test(message);
    if (!isAuthError) {
      throw err;
    }

    await addLog('Step 9: Sub2API auth failed, trying dashboard session token fallback...', 'warn');
    const fallbackCredential = await resolveSub2apiCredentialFromDashboard({
      ...state,
      sub2apiRuntimeCredential: '',
    });
    created = await requestSub2apiAdminApi(state, 'admin/openai/create-from-oauth', {
      method: 'POST',
      authCredential: fallbackCredential,
      body: {
        session_id: sessionId,
        code: callback.code,
        state: callback.state,
        redirect_uri: callback.redirectUri,
        name: String(state.email || '').trim() || undefined,
      },
    });
  }

  await setState({ sub2apiSessionId: null });

  const accountId = Number(created?.id || 0);
  const accountName = String(created?.name || state.email || '').trim();
  if (accountId > 0) {
    await addLog(`Step 9: Sub2API account created #${accountId}${accountName ? ` (${accountName})` : ''}`, 'ok');
  } else {
    await addLog(`Step 9: Sub2API import completed${accountName ? ` (${accountName})` : ''}`, 'ok');
  }

  await completeBackgroundStep(9, {
    sub2apiAccountId: accountId || null,
    sub2apiAccountName: accountName || null,
  });
}

// ============================================================
// Step 10: Cleanup Source Email (iCloud alias / Microsoft account)
// ============================================================

async function executeStep10(state) {
  const email = String(state.email || '').trim();
  if (!email) {
    throw new Error('No email. Complete step 3 first.');
  }

  await addLog('Step 10: Cleaning up source email...');

  if (!state.autoDeleteUsedIcloudAlias) {
    await addLog('Step 10: Cleanup is disabled. Enable Cleanup if you want to auto-delete the source email.', 'warn');
    await completeBackgroundStep(10, {
      skipped: true,
      reason: 'cleanup_disabled',
      deleted: false,
      email,
    });
    return;
  }

  await maybeAutoDeleteCompletedIcloudAlias();
  await completeBackgroundStep(10, {
    deleted: true,
    email,
  });
}

// ============================================================
// Open Side Panel on extension icon click
// ============================================================

chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
