// sidepanel/sidepanel.js — Side Panel logic

const STATUS_ICONS = {
  pending: '',
  running: '',
  completed: '\u2713',  // ✓
  skipped: '\u00BB',    // »
  failed: '\u2717',     // ✗
  stopped: '\u25A0',    // ■
};
const TOTAL_STEPS = 10;

const logArea = document.getElementById('log-area');
const displayOauthUrl = document.getElementById('display-oauth-url');
const displayLocalhostUrl = document.getElementById('display-localhost-url');
const displayStatus = document.getElementById('display-status');
const statusBar = document.getElementById('status-bar');
const appTitle = document.getElementById('app-title');
const btnVersion = document.getElementById('btn-version');
const displayVersion = document.getElementById('display-version');
const displayElapsed = document.getElementById('display-elapsed');
const displayAverageDuration = document.getElementById('display-average-duration');
const displaySuccessRate = document.getElementById('display-success-rate');
const icloudSection = document.getElementById('icloud-section');
const icloudSummary = document.getElementById('icloud-summary');
const icloudList = document.getElementById('icloud-list');
const icloudLoginHelp = document.getElementById('icloud-login-help');
const icloudLoginHelpTitle = document.getElementById('icloud-login-help-title');
const icloudLoginHelpText = document.getElementById('icloud-login-help-text');
const btnIcloudLoginDone = document.getElementById('btn-icloud-login-done');
const btnIcloudRefresh = document.getElementById('btn-icloud-refresh');
const btnIcloudDeleteUsed = document.getElementById('btn-icloud-delete-used');
const checkboxAutoDeleteIcloud = document.getElementById('checkbox-auto-delete-icloud');
const aliasSourceValue = document.querySelector('.data-row .data-value[data-i18n="icloudAliasName"]');
const rowMailProvider = document.getElementById('row-mail-provider');
const inputEmail = document.getElementById('input-email');
const inputPassword = document.getElementById('input-password');
const btnFetchEmail = document.getElementById('btn-fetch-email');
const btnCopyEmail = document.getElementById('btn-copy-email');
const btnTogglePassword = document.getElementById('btn-toggle-password');
const btnCopyPassword = document.getElementById('btn-copy-password');
const btnStop = document.getElementById('btn-stop');
const btnReset = document.getElementById('btn-reset');
const stepsProgress = document.getElementById('steps-progress');
const btnAutoRun = document.getElementById('btn-auto-run');
const btnAutoContinue = document.getElementById('btn-auto-continue');
const autoContinueBar = document.getElementById('auto-continue-bar');
const btnClearLog = document.getElementById('btn-clear-log');
const selectLanguage = document.getElementById('select-language');
const selectOauthProvider = document.getElementById('select-oauth-provider');
const rowCpaAuthUrl = document.getElementById('row-cpa-auth-url');
const inputVpsUrl = document.getElementById('input-vps-url');
const btnPasteVpsUrl = document.getElementById('btn-paste-vps-url');
const rowSub2apiBaseUrl = document.getElementById('row-sub2api-base-url');
const inputSub2apiBaseUrl = document.getElementById('input-sub2api-base-url');
const rowSub2apiApiKey = document.getElementById('row-sub2api-api-key');
const inputSub2apiApiKey = document.getElementById('input-sub2api-api-key');
const selectMailProvider = document.getElementById('select-mail-provider');
const rowInbucketHost = document.getElementById('row-inbucket-host');
const inputInbucketHost = document.getElementById('input-inbucket-host');
const rowInbucketMailbox = document.getElementById('row-inbucket-mailbox');
const inputInbucketMailbox = document.getElementById('input-inbucket-mailbox');
const rowMicrosoftManagerUrl = document.getElementById('row-microsoft-manager-url');
const inputMicrosoftManagerUrl = document.getElementById('input-microsoft-manager-url');
const rowMicrosoftManagerToken = document.getElementById('row-microsoft-manager-token');
const inputMicrosoftManagerToken = document.getElementById('input-microsoft-manager-token');
const rowMicrosoftManagerMode = document.getElementById('row-microsoft-manager-mode');
const selectMicrosoftManagerMode = document.getElementById('select-microsoft-manager-mode');
const rowMicrosoftManagerKeyword = document.getElementById('row-microsoft-manager-keyword');
const inputMicrosoftManagerKeyword = document.getElementById('input-microsoft-manager-keyword');
const inputRunCount = document.getElementById('input-run-count');
const autoHint = document.getElementById('auto-hint');
let icloudRefreshQueued = false;
let currentLanguage = localStorage.getItem('multipage-language') || 'zh-CN';
let lastKnownState = null;
let lastRenderedIcloudAliases = [];
let runMetricsTicker = null;
let activeRunStartMs = 0;
let activeRunKey = '';
const completedRunDurationsMs = [];
const finishedRunKeys = new Set();
const successfulRunKeys = new Set();
const manifestInfo = chrome.runtime.getManifest();
const releaseRepo = 'Msg-Lbo/MSRegFlow';
const currentManifestVersion = normalizeVersionValue(manifestInfo.version || '0.0.0');
const currentManifestVersionLabel = formatVersionLabel(currentManifestVersion);
let latestReleaseVersion = '';
let latestReleaseUrl = `https://github.com/${releaseRepo}/releases`;
let hasNewRelease = false;
let isVersionCheckFinished = false;
let versionCheckInFlight = false;
let hasShownNewReleaseToast = false;

function normalizeMailProviderValue(rawValue) {
  void rawValue;
  return 'microsoft-manager';
}

// ============================================================
// Toast Notifications
// ============================================================

const toastContainer = document.getElementById('toast-container');

const TOAST_ICONS = {
  error: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>',
  warn: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>',
  success: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>',
  info: '<svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/></svg>',
};

const AUTO_BUTTON_ICON = '<svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><polygon points="5 3 19 12 5 21 5 3"/></svg>';

const I18N = {
  'zh-CN': {
    titleRunCount: '运行次数',
    titleAutoRun: '自动执行全部步骤',
    titleFetchEmail: '自动获取 iCloud 别名',
    titleFetchEmailMicrosoftManager: '自动获取 Microsoft 账号',
    titleStop: '停止当前流程',
    titleReset: '重置全部步骤',
    titleTheme: '切换主题',
    titleVersionBadge: '点击查看版本更新',
    titleSkipStep: '跳过这一步',
    titleClearLog: '清空日志',
    labelCpaAuth: 'CPA Auth',
    labelOauthTarget: 'OAuth',
    labelLanguage: '语言',
    labelAlias: '别名',
    labelCleanup: '清理',
    labelVerify: '验证',
    labelInbucket: 'Inbucket',
    labelMailbox: '邮箱名',
    labelMicrosoftManager: 'MS 管理',
    labelToken: '令牌',
    labelMode: '模式',
    labelKeyword: '筛选',
    labelSub2api: 'Sub2API',
    labelSub2apiApiKey: 'API Key',
    labelEmail: '邮箱',
    labelPassword: '密码',
    labelOauth: 'OAuth',
    labelCallback: '回调',
    labelElapsed: '计时',
    labelAverageDuration: '平均用时',
    labelSuccessRate: '成功率',
    icloudAliasName: 'iCloud Hide My Email',
    microsoftManagerEmailName: 'Microsoft 账号',
    cleanupAutoDelete: '成功使用后自动删除来源邮箱',
    mailProvider163: '163 邮箱 (mail.163.com)',
    mailProviderQq: 'QQ 邮箱 (wx.mail.qq.com)',
    mailProviderInbucket: 'Inbucket（自定义主机）',
    mailProviderMicrosoftManager: 'Microsoft Account Manager API',
    microsoftManagerModeGraph: 'Graph',
    microsoftManagerModeImap: 'IMAP',
    oauthProviderCpaAuth: 'CPA Auth',
    oauthProviderSub2api: 'Sub2API',
    placeholderCpaAuth: 'http://ip:port/management.html#/oauth',
    placeholderSub2apiBaseUrl: 'https://你的-sub2api域名',
    placeholderSub2apiApiKey: '可留空；或填写 x-api-key / Bearer token',
    placeholderInbucketHost: '你的 inbucket 主机或 https://你的主机',
    placeholderInbucketMailbox: '例如 zju2001',
    placeholderMicrosoftManagerUrl: 'https://你的-manager域名',
    placeholderMicrosoftManagerToken: '填写 MAIL_API_TOKEN',
    placeholderMicrosoftManagerKeyword: '可选关键词，用于筛选账号',
    placeholderEmail: '使用 Auto 生成 iCloud 别名，或手动粘贴',
    placeholderEmailMicrosoftManager: '使用 Auto 获取 Microsoft 账号，或手动粘贴',
    placeholderPassword: '留空则自动生成',
    waiting: '等待中...',
    btnAuto: '自动',
    btnStop: '停止',
    btnContinue: '继续',
    btnCopy: '复制',
    btnPaste: '粘贴',
    btnRefresh: '刷新',
    btnDeleteUsed: '删除已用',
    btnDelete: '删除',
    btnMarkUsed: '标记已用',
    btnMarkUnused: '标记未用',
    btnIcloudLoginDone: '我已登录',
    btnClear: '清空',
    btnSkip: '跳过',
    btnShow: '显示',
    btnHide: '隐藏',
    sectionIcloud: 'iCloud',
    sectionWorkflow: '流程',
    sectionConsole: '控制台',
    step1: '获取 OAuth 链接',
    step2: '打开注册页',
    step3: '填写邮箱 / 密码',
    step4: '获取注册验证码',
    step5: '填写姓名 / 生日',
    step6: '通过 OAuth 登录',
    step7: '获取登录验证码',
    step8: 'OAuth 自动确认',
    step9: '回调验证 / 导入',
    step10: '清理来源邮箱',
    statusRunning: ({ step }) => `第 ${step} 步执行中...`,
    statusFailed: ({ step }) => `第 ${step} 步失败`,
    statusStopped: ({ step }) => `第 ${step} 步已停止`,
    statusAllFinished: '全部步骤已完成',
    statusSkipped: ({ step }) => `第 ${step} 步已跳过`,
    statusDone: ({ step }) => `第 ${step} 步完成`,
    statusReady: '就绪',
    autoHintEmail: '使用 Auto 生成 iCloud 别名，或手动粘贴后继续',
    autoHintEmailMicrosoftManager: '使用 Auto 获取 Microsoft 账号邮箱，或手动粘贴后继续',
    autoHintError: '自动运行被错误中断。修复问题或跳过失败步骤后继续',
    fetchedEmail: ({ email }) => `已获取 ${email}`,
    autoFetchFailed: ({ message }) => `自动获取失败：${message}`,
    icloudSummaryInitial: '加载你的 Hide My Email 别名以便在这里管理。',
    icloudEmpty: '未找到 iCloud Hide My Email 别名。',
    icloudAliasesLoaded: ({ count, usedCount }) => `已加载 ${count} 个别名，其中 ${usedCount} 个已在插件中标记为 used。`,
    icloudLoading: '正在加载 iCloud 别名...',
    icloudLoadFailed: ({ message }) => `iCloud 加载失败：${message}`,
    deletingAlias: ({ email }) => `正在删除 ${email}...`,
    deletedAlias: ({ email }) => `已删除 ${email}`,
    deleteFailed: ({ message }) => `删除失败：${message}`,
    updatingAliasUsed: ({ email, used }) => `正在将 ${email} 标记为${used ? '已用' : '未用'}...`,
    updatedAliasUsed: ({ email, used }) => `${email} 已标记为${used ? '已用' : '未用'}`,
    updateAliasUsedFailed: ({ message }) => `标记失败：${message}`,
    deletingUsedAliases: '正在删除已使用的 iCloud 别名...',
    deletedUsedAliases: ({ deleted, skipped }) => skipped ? `已删除 ${deleted} 个已用别名，跳过 ${skipped} 个。` : `已删除 ${deleted} 个已用别名。`,
    bulkDeleteFailed: ({ message }) => `批量删除失败：${message}`,
    icloudLoginRequiredToast: '需要登录 iCloud，我已经为你打开登录页。',
    icloudLoginHelpTitle: '需要登录 iCloud',
    icloudLoginHelpText: ({ host }) => `我已经为你打开 ${host}。请在那个页面完成登录，然后回到这里点击“我已登录”。`,
    icloudSessionReady: 'iCloud 会话已恢复，别名列表已刷新。',
    icloudStillNotSignedIn: ({ message }) => `看起来还没有登录完成：${message}`,
    pleaseEnterEmailFirst: '请先粘贴邮箱地址或点击 Auto',
    skipFailed: ({ message }) => `跳过失败：${message}`,
    stepSkippedToast: ({ step }) => `第 ${step} 步已跳过`,
    stoppingFlow: '正在停止当前流程...',
    continueNeedEmail: '请先获取或粘贴邮箱地址',
    continueFailed: ({ message }) => `继续失败：${message}`,
    confirmReset: '要重置全部步骤和数据吗？',
    copiedValue: ({ label }) => `已复制${label}`,
    copiedValueFallback: ({ label }) => `已复制 ${label}`,
    copyFailed: ({ label, message }) => `${label}复制失败：${message}`,
    nothingToCopy: ({ label }) => `${label}为空，无法复制`,
    pastedCpaAuth: '已从剪贴板粘贴 CPA Auth 地址',
    pasteFailed: ({ message }) => `粘贴失败：${message}`,
    clipboardEmpty: '剪贴板为空',
    clipboardNoUsefulText: '剪贴板中没有可用内容',
    autoRunRunning: ({ runLabel }) => `运行中${runLabel}`,
    autoRunPaused: ({ runLabel }) => `已暂停${runLabel}`,
    autoRunInterrupted: ({ runLabel }) => `已中断${runLabel}`,
    versionChecking: '版本检查中...',
    versionTooltipLatest: ({ version }) => `当前已是最新版本 ${version}`,
    versionTooltipUpdateAvailable: ({ current, latest }) => `发现新版本 ${latest}（当前 ${current}），点击查看`,
    versionTooltipCheckFailed: '版本检查失败，点击查看 Releases',
    newVersionFound: ({ latest }) => `发现新版本 ${latest}，点击标题旁版本号查看`,
  },
  'en-US': {
    titleRunCount: 'Number of runs',
    titleAutoRun: 'Run all steps automatically',
    titleFetchEmail: 'Fetch an iCloud alias automatically',
    titleFetchEmailMicrosoftManager: 'Fetch a Microsoft account automatically',
    titleStop: 'Stop current flow',
    titleReset: 'Reset all steps',
    titleTheme: 'Toggle theme',
    titleVersionBadge: 'Click to view version updates',
    titleSkipStep: 'Skip this step',
    titleClearLog: 'Clear log',
    labelCpaAuth: 'CPA Auth',
    labelOauthTarget: 'OAuth',
    labelLanguage: 'Language',
    labelAlias: 'Alias',
    labelCleanup: 'Cleanup',
    labelVerify: 'Verify',
    labelInbucket: 'Inbucket',
    labelMailbox: 'Mailbox',
    labelMicrosoftManager: 'MSMgr',
    labelToken: 'Token',
    labelMode: 'Mode',
    labelKeyword: 'Filter',
    labelSub2api: 'Sub2API',
    labelSub2apiApiKey: 'API Key',
    labelEmail: 'Email',
    labelPassword: 'Password',
    labelOauth: 'OAuth',
    labelCallback: 'Callback',
    labelElapsed: 'Elapsed',
    labelAverageDuration: 'Avg Time',
    labelSuccessRate: 'Success Rate',
    icloudAliasName: 'iCloud Hide My Email',
    microsoftManagerEmailName: 'Microsoft account',
    cleanupAutoDelete: 'Delete source email after successful use',
    mailProvider163: '163 Mail (mail.163.com)',
    mailProviderQq: 'QQ Mail (wx.mail.qq.com)',
    mailProviderInbucket: 'Inbucket (custom host)',
    mailProviderMicrosoftManager: 'Microsoft Account Manager API',
    microsoftManagerModeGraph: 'Graph',
    microsoftManagerModeImap: 'IMAP',
    oauthProviderCpaAuth: 'CPA Auth',
    oauthProviderSub2api: 'Sub2API',
    placeholderCpaAuth: 'http://ip:port/management.html#/oauth',
    placeholderSub2apiBaseUrl: 'https://your-sub2api-host',
    placeholderSub2apiApiKey: 'Optional; use x-api-key or Bearer token',
    placeholderInbucketHost: 'your inbucket host or https://your-host',
    placeholderInbucketMailbox: 'e.g. zju2001',
    placeholderMicrosoftManagerUrl: 'https://your-manager-domain',
    placeholderMicrosoftManagerToken: 'Use MAIL_API_TOKEN',
    placeholderMicrosoftManagerKeyword: 'Optional keyword for account filter',
    placeholderEmail: 'Use Auto to generate an iCloud alias, or paste manually',
    placeholderEmailMicrosoftManager: 'Use Auto to fetch a Microsoft account, or paste manually',
    placeholderPassword: 'Leave blank to auto-generate',
    waiting: 'Waiting...',
    btnAuto: 'Auto',
    btnStop: 'Stop',
    btnContinue: 'Continue',
    btnCopy: 'Copy',
    btnPaste: 'Paste',
    btnRefresh: 'Refresh',
    btnDeleteUsed: 'Delete Used',
    btnDelete: 'Delete',
    btnMarkUsed: 'Mark Used',
    btnMarkUnused: 'Mark Unused',
    btnIcloudLoginDone: "I've Signed In",
    btnClear: 'Clear',
    btnSkip: 'Skip',
    btnShow: 'Show',
    btnHide: 'Hide',
    sectionIcloud: 'iCloud',
    sectionWorkflow: 'Workflow',
    sectionConsole: 'Console',
    step1: 'Get OAuth Link',
    step2: 'Open Signup',
    step3: 'Fill Email / Password',
    step4: 'Get Signup Code',
    step5: 'Fill Name / Birthday',
    step6: 'Login via OAuth',
    step7: 'Get Login Code',
    step8: 'OAuth Auto Confirm',
    step9: 'Callback Verify / Import',
    step10: 'Cleanup Source Email',
    statusRunning: ({ step }) => `Step ${step} running...`,
    statusFailed: ({ step }) => `Step ${step} failed`,
    statusStopped: ({ step }) => `Step ${step} stopped`,
    statusAllFinished: 'All steps finished',
    statusSkipped: ({ step }) => `Step ${step} skipped`,
    statusDone: ({ step }) => `Step ${step} done`,
    statusReady: 'Ready',
    autoHintEmail: 'Use Auto to generate an iCloud alias, or paste manually, then continue',
    autoHintEmailMicrosoftManager: 'Use Auto to fetch a Microsoft account email, or paste manually, then continue',
    autoHintError: 'Auto run was interrupted by an error. Fix it or skip the failed step, then continue',
    fetchedEmail: ({ email }) => `Fetched ${email}`,
    autoFetchFailed: ({ message }) => `Auto fetch failed: ${message}`,
    icloudSummaryInitial: 'Load your Hide My Email aliases to manage them here.',
    icloudEmpty: 'No iCloud Hide My Email aliases found.',
    icloudAliasesLoaded: ({ count, usedCount }) => `${count} aliases loaded. ${usedCount} marked as used in this plugin.`,
    icloudLoading: 'Loading iCloud aliases...',
    icloudLoadFailed: ({ message }) => `iCloud load failed: ${message}`,
    deletingAlias: ({ email }) => `Deleting ${email}...`,
    deletedAlias: ({ email }) => `Deleted ${email}`,
    deleteFailed: ({ message }) => `Delete failed: ${message}`,
    updatingAliasUsed: ({ email, used }) => `Marking ${email} as ${used ? 'used' : 'unused'}...`,
    updatedAliasUsed: ({ email, used }) => `${email} marked as ${used ? 'used' : 'unused'}`,
    updateAliasUsedFailed: ({ message }) => `Failed to update used state: ${message}`,
    deletingUsedAliases: 'Deleting used iCloud aliases...',
    deletedUsedAliases: ({ deleted, skipped }) => skipped ? `Deleted ${deleted} used aliases, ${skipped} skipped.` : `Deleted ${deleted} used aliases.`,
    bulkDeleteFailed: ({ message }) => `Bulk delete failed: ${message}`,
    icloudLoginRequiredToast: 'iCloud sign-in is required. A login page has been opened for you.',
    icloudLoginHelpTitle: 'iCloud sign-in required',
    icloudLoginHelpText: ({ host }) => `We opened ${host} for you. Please finish sign-in there, then return here and click "I've Signed In".`,
    icloudSessionReady: 'iCloud session is ready. Alias list refreshed.',
    icloudStillNotSignedIn: ({ message }) => `Still not signed in: ${message}`,
    pleaseEnterEmailFirst: 'Please paste email address or use Auto first',
    skipFailed: ({ message }) => `Skip failed: ${message}`,
    stepSkippedToast: ({ step }) => `Step ${step} skipped`,
    stoppingFlow: 'Stopping current flow...',
    continueNeedEmail: 'Please fetch or paste an email address first!',
    continueFailed: ({ message }) => `Continue failed: ${message}`,
    confirmReset: 'Reset all steps and data?',
    copiedValue: ({ label }) => `Copied ${label}`,
    copiedValueFallback: ({ label }) => `${label} copied`,
    copyFailed: ({ label, message }) => `Failed to copy ${label}: ${message}`,
    nothingToCopy: ({ label }) => `${label} is empty`,
    pastedCpaAuth: 'Pasted CPA Auth URL from clipboard',
    pasteFailed: ({ message }) => `Paste failed: ${message}`,
    clipboardEmpty: 'Clipboard is empty',
    clipboardNoUsefulText: 'Clipboard does not contain usable text',
    autoRunRunning: ({ runLabel }) => `Running${runLabel}`,
    autoRunPaused: ({ runLabel }) => `Paused${runLabel}`,
    autoRunInterrupted: ({ runLabel }) => `Interrupted${runLabel}`,
    versionChecking: 'Checking version...',
    versionTooltipLatest: ({ version }) => `You are on the latest version ${version}`,
    versionTooltipUpdateAvailable: ({ current, latest }) => `New version ${latest} available (current ${current}), click to view`,
    versionTooltipCheckFailed: 'Version check failed, click to view releases',
    newVersionFound: ({ latest }) => `New version ${latest} found. Click the header version badge to view`,
  },
};

function t(key, vars = {}) {
  const pack = I18N[currentLanguage] || I18N['zh-CN'];
  const fallbackPack = I18N['zh-CN'];
  const value = pack[key] ?? fallbackPack[key] ?? key;
  if (typeof value === 'function') return value(vars);
  return String(value).replace(/\{(\w+)\}/g, (_, name) => String(vars[name] ?? ''));
}

function setAutoRunButton(label) {
  btnAutoRun.innerHTML = `${AUTO_BUTTON_ICON} ${label}`;
}

function normalizeVersionValue(rawValue) {
  return String(rawValue || '')
    .trim()
    .replace(/^refs\/tags\//i, '')
    .replace(/^v/i, '');
}

function formatVersionLabel(versionValue) {
  const normalized = normalizeVersionValue(versionValue);
  return normalized ? `v${normalized}` : 'v0.0.0';
}

function parseVersionParts(versionValue) {
  const normalized = normalizeVersionValue(versionValue);
  if (!normalized) return [];

  return normalized
    .split('.')
    .map(part => {
      const matched = String(part || '').match(/\d+/);
      return matched ? Number(matched[0]) : 0;
    });
}

function compareVersionValues(leftVersion, rightVersion) {
  const left = parseVersionParts(leftVersion);
  const right = parseVersionParts(rightVersion);
  const length = Math.max(left.length, right.length);

  for (let index = 0; index < length; index++) {
    const leftPart = Number(left[index] || 0);
    const rightPart = Number(right[index] || 0);
    if (leftPart > rightPart) return 1;
    if (leftPart < rightPart) return -1;
  }

  return 0;
}

function getVersionBadgeTitle() {
  if (!isVersionCheckFinished) {
    return t('versionChecking');
  }

  if (hasNewRelease) {
    return t('versionTooltipUpdateAvailable', {
      current: currentManifestVersionLabel,
      latest: formatVersionLabel(latestReleaseVersion),
    });
  }

  if (latestReleaseVersion) {
    return t('versionTooltipLatest', {
      version: formatVersionLabel(latestReleaseVersion),
    });
  }

  return t('versionTooltipCheckFailed');
}

function renderVersionBadge() {
  if (appTitle) {
    appTitle.textContent = String(manifestInfo.name || 'MSRegFlow');
  }

  if (displayVersion) {
    displayVersion.textContent = currentManifestVersionLabel;
  }

  if (!btnVersion) return;

  btnVersion.classList.toggle('has-update', hasNewRelease);
  const title = getVersionBadgeTitle();
  btnVersion.title = title;
  btnVersion.setAttribute('aria-label', title);
}

async function checkLatestReleaseVersion() {
  if (versionCheckInFlight) return;

  versionCheckInFlight = true;
  isVersionCheckFinished = false;
  renderVersionBadge();

  try {
    const response = await fetch(`https://api.github.com/repos/${releaseRepo}/releases/latest`, {
      method: 'GET',
      headers: {
        Accept: 'application/vnd.github+json',
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const payload = await response.json();
    const latestTag = normalizeVersionValue(payload?.tag_name || '');
    const latestUrl = String(payload?.html_url || '').trim();

    latestReleaseVersion = latestTag;
    if (latestUrl) {
      latestReleaseUrl = latestUrl;
    }

    hasNewRelease = latestTag
      ? compareVersionValues(latestTag, currentManifestVersion) > 0
      : false;

    isVersionCheckFinished = true;
    renderVersionBadge();

    if (hasNewRelease && !hasShownNewReleaseToast) {
      hasShownNewReleaseToast = true;
      showToast(t('newVersionFound', { latest: formatVersionLabel(latestTag) }), 'warn', 4200);
    }
  } catch (err) {
    latestReleaseVersion = '';
    hasNewRelease = false;
    isVersionCheckFinished = true;
    renderVersionBadge();
    console.warn('Version check failed:', err);
  } finally {
    versionCheckInFlight = false;
  }
}

function getVersionOpenUrl() {
  const currentReleaseUrl = `https://github.com/${releaseRepo}/releases/tag/${currentManifestVersion}`;
  if (hasNewRelease && latestReleaseUrl) {
    return latestReleaseUrl;
  }
  if (latestReleaseVersion) {
    return currentReleaseUrl;
  }
  return latestReleaseUrl || currentReleaseUrl;
}

async function openVersionPage() {
  const url = getVersionOpenUrl();
  try {
    await chrome.tabs.create({ url, active: true });
  } catch {
    window.open(url, '_blank', 'noopener');
  }
}

function formatDuration(ms) {
  const totalSeconds = Math.max(0, Math.floor(Number(ms || 0) / 1000));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  if (hours > 0) {
    return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
  }
  return `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
}

function calculateAverageRunDurationMs() {
  if (!completedRunDurationsMs.length) return 0;
  const total = completedRunDurationsMs.reduce((sum, duration) => sum + duration, 0);
  return Math.round(total / completedRunDurationsMs.length);
}

function renderRunMetrics() {
  if (displayElapsed) {
    displayElapsed.textContent = activeRunStartMs
      ? formatDuration(Date.now() - activeRunStartMs)
      : '--:--';
  }

  if (displayAverageDuration) {
    const averageMs = calculateAverageRunDurationMs();
    displayAverageDuration.textContent = averageMs > 0 ? formatDuration(averageMs) : '--:--';
  }

  if (displaySuccessRate) {
    const attempts = finishedRunKeys.size;
    const successes = successfulRunKeys.size;
    displaySuccessRate.textContent = attempts > 0
      ? `${((successes / attempts) * 100).toFixed(1)}%`
      : '--';
  }
}

function startRunMetricsTicker() {
  if (runMetricsTicker !== null) return;
  runMetricsTicker = setInterval(() => {
    renderRunMetrics();
  }, 1000);
}

function stopRunMetricsTickerIfIdle() {
  if (activeRunStartMs) return;
  if (runMetricsTicker === null) return;
  clearInterval(runMetricsTicker);
  runMetricsTicker = null;
}

function beginRunMetrics(flowStartTimeMs) {
  const normalizedStartMs = Number(flowStartTimeMs || 0);
  if (!Number.isFinite(normalizedStartMs) || normalizedStartMs <= 0) {
    return;
  }

  const newRunKey = String(Math.trunc(normalizedStartMs));

  if (activeRunStartMs && activeRunKey && activeRunKey !== newRunKey) {
    finishActiveRunMetrics(false, Number(activeRunKey));
  }

  activeRunStartMs = normalizedStartMs;
  activeRunKey = newRunKey;
  startRunMetricsTicker();
  renderRunMetrics();
}

function finishActiveRunMetrics(success, flowStartTimeMs = 0) {
  if (!activeRunStartMs) return;

  const durationMs = Math.max(0, Date.now() - activeRunStartMs);
  const resolvedKey = String(Math.trunc(Number(flowStartTimeMs || activeRunKey || activeRunStartMs) || 0));

  if (resolvedKey && !finishedRunKeys.has(resolvedKey)) {
    finishedRunKeys.add(resolvedKey);
    if (success) {
      successfulRunKeys.add(resolvedKey);
      completedRunDurationsMs.push(durationMs);
    }
  }

  activeRunStartMs = 0;
  activeRunKey = '';
  stopRunMetricsTickerIfIdle();
  renderRunMetrics();
}

function syncRunMetricsFromState(state) {
  if (!state || !state.stepStatuses) return;

  const flowStartTimeMs = Number(state.flowStartTime || 0);
  const stepStatuses = Object.values(state.stepStatuses || {});
  const hasRunning = stepStatuses.includes('running');
  const hasFailedOrStopped = stepStatuses.some(status => status === 'failed' || status === 'stopped');
  const allProgressed = stepStatuses.length > 0
    && stepStatuses.every(status => status === 'completed' || status === 'skipped');

  if (flowStartTimeMs > 0) {
    beginRunMetrics(flowStartTimeMs);
  } else if (!hasRunning && activeRunStartMs) {
    finishActiveRunMetrics(false, 0);
  }

  if (!activeRunStartMs) {
    renderRunMetrics();
    return;
  }

  if (allProgressed) {
    finishActiveRunMetrics(true, flowStartTimeMs);
    return;
  }

  if (hasFailedOrStopped && !hasRunning) {
    finishActiveRunMetrics(false, flowStartTimeMs);
    return;
  }

  renderRunMetrics();
}

function resetRunMetrics() {
  activeRunStartMs = 0;
  activeRunKey = '';
  completedRunDurationsMs.length = 0;
  finishedRunKeys.clear();
  successfulRunKeys.clear();
  stopRunMetricsTickerIfIdle();
  renderRunMetrics();
}

function getCopyLabel(kind) {
  if (currentLanguage === 'zh-CN') {
    if (kind === 'email') return '邮箱';
    if (kind === 'password') return '密码';
    return '内容';
  }
  if (kind === 'email') return 'email';
  if (kind === 'password') return 'password';
  return 'value';
}

function isMicrosoftManagerProviderSelected() {
  return normalizeMailProviderValue(selectMailProvider.value) === 'microsoft-manager';
}

function shouldUseIcloudUi() {
  return !isMicrosoftManagerProviderSelected();
}

function getEmailSourceLabel() {
  return isMicrosoftManagerProviderSelected() ? t('microsoftManagerEmailName') : t('icloudAliasName');
}

function getFetchEmailTitle() {
  return isMicrosoftManagerProviderSelected() ? t('titleFetchEmailMicrosoftManager') : t('titleFetchEmail');
}

function getEmailPlaceholderText() {
  return isMicrosoftManagerProviderSelected() ? t('placeholderEmailMicrosoftManager') : t('placeholderEmail');
}

function getAutoHintText() {
  return isMicrosoftManagerProviderSelected() ? t('autoHintEmailMicrosoftManager') : t('autoHintEmail');
}

function isSub2apiOauthProviderSelected() {
  return selectOauthProvider.value === 'sub2api';
}

function updateOauthProviderUI() {
  const useSub2api = isSub2apiOauthProviderSelected();
  rowCpaAuthUrl.style.display = useSub2api ? 'none' : '';
  rowSub2apiBaseUrl.style.display = useSub2api ? '' : 'none';
  rowSub2apiApiKey.style.display = useSub2api ? '' : 'none';
}

function applyLanguage(language) {
  currentLanguage = I18N[language] ? language : 'zh-CN';
  localStorage.setItem('multipage-language', currentLanguage);
  document.documentElement.lang = currentLanguage;
  if (selectLanguage) {
    selectLanguage.value = currentLanguage;
  }

  document.querySelectorAll('[data-i18n]').forEach((node) => {
    const key = node.dataset.i18n;
    node.textContent = t(key);
  });
  document.querySelectorAll('[data-i18n-placeholder]').forEach((node) => {
    const key = node.dataset.i18nPlaceholder;
    node.placeholder = t(key);
  });
  document.querySelectorAll('[data-i18n-title]').forEach((node) => {
    const key = node.dataset.i18nTitle;
    node.title = t(key);
  });

  inputPassword.placeholder = t('placeholderPassword');
  if (!displayOauthUrl.classList.contains('has-value')) {
    displayOauthUrl.textContent = t('waiting');
  }
  if (!displayLocalhostUrl.classList.contains('has-value')) {
    displayLocalhostUrl.textContent = t('waiting');
  }
  updateOauthProviderUI();
  updateMailProviderUI();
  updateEmailSourceUI();
  syncPasswordToggleLabel();
  updateProgressCounter();
  renderVersionBadge();
  if (lastKnownState) {
    updateStatusDisplay(lastKnownState);
  } else {
    displayStatus.textContent = t('statusReady');
  }
  renderRunMetrics();
  renderIcloudAliases(lastRenderedIcloudAliases);
  if (!icloudSummary.textContent || icloudSummary.textContent === 'Load your Hide My Email aliases to manage them here.') {
    icloudSummary.textContent = t('icloudSummaryInitial');
  }
}

async function saveVpsUrlValue(value) {
  const vpsUrl = String(value || '').trim();
  inputVpsUrl.value = vpsUrl;
  if (!vpsUrl) return;
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { vpsUrl },
  });
}

async function copyTextValue(value, kind) {
  const trimmed = String(value || '').trim();
  const label = getCopyLabel(kind);
  if (!trimmed) {
    showToast(t('nothingToCopy', { label }), 'warn');
    return;
  }

  try {
    await navigator.clipboard.writeText(trimmed);
    showToast(t('copiedValue', { label }), 'success', 2000);
  } catch (err) {
    showToast(t('copyFailed', { label, message: err.message || err }), 'error');
  }
}

async function pasteCpaAuthFromClipboard(options = {}) {
  const { silentIfFilled = false } = options;
  if (silentIfFilled && inputVpsUrl.value.trim()) return;

  try {
    const text = String(await navigator.clipboard.readText() || '').trim();
    if (!text) {
      showToast(t('clipboardEmpty'), 'warn');
      return;
    }
    await saveVpsUrlValue(text);
    showToast(t('pastedCpaAuth'), 'success', 2000);
  } catch (err) {
    showToast(t('pasteFailed', { message: err.message || err }), 'warn');
  }
}

function showToast(message, type = 'error', duration = 4000) {
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `${TOAST_ICONS[type] || ''}<span class="toast-msg">${escapeHtml(message)}</span><button class="toast-close">&times;</button>`;

  toast.querySelector('.toast-close').addEventListener('click', () => dismissToast(toast));
  toastContainer.appendChild(toast);

  if (duration > 0) {
    setTimeout(() => dismissToast(toast), duration);
  }
}

function dismissToast(toast) {
  if (!toast.parentNode) return;
  toast.classList.add('toast-exit');
  toast.addEventListener('animationend', () => toast.remove());
}

// ============================================================
// State Restore on load
// ============================================================

async function restoreState() {
  try {
    const state = await chrome.runtime.sendMessage({ type: 'GET_STATE', source: 'sidepanel' });
    lastKnownState = state;
    applyLanguage(state.language || currentLanguage);

    if (state.oauthUrl) {
      displayOauthUrl.textContent = state.oauthUrl;
      displayOauthUrl.classList.add('has-value');
    }
    if (state.localhostUrl) {
      displayLocalhostUrl.textContent = state.localhostUrl;
      displayLocalhostUrl.classList.add('has-value');
    }
    if (state.email) {
      inputEmail.value = state.email;
    }
    syncPasswordField(state);
    if (state.vpsUrl) {
      inputVpsUrl.value = state.vpsUrl;
    }
    if (state.oauthProvider) {
      selectOauthProvider.value = state.oauthProvider;
    }
    if (state.sub2apiBaseUrl) {
      inputSub2apiBaseUrl.value = state.sub2apiBaseUrl;
    }
    if (state.sub2apiAdminApiKey) {
      inputSub2apiApiKey.value = state.sub2apiAdminApiKey;
    }
    checkboxAutoDeleteIcloud.checked = Boolean(state.autoDeleteUsedIcloudAlias);
    if (state.language) {
      selectLanguage.value = state.language;
    }
    if (state.mailProvider) {
      selectMailProvider.value = normalizeMailProviderValue(state.mailProvider);
    } else {
      selectMailProvider.value = 'microsoft-manager';
    }
    if (state.inbucketHost) {
      inputInbucketHost.value = state.inbucketHost;
    }
    if (state.inbucketMailbox) {
      inputInbucketMailbox.value = state.inbucketMailbox;
    }
    if (state.microsoftManagerUrl) {
      inputMicrosoftManagerUrl.value = state.microsoftManagerUrl;
    }
    if (state.microsoftManagerToken) {
      inputMicrosoftManagerToken.value = state.microsoftManagerToken;
    }
    if (state.microsoftManagerMode) {
      selectMicrosoftManagerMode.value = state.microsoftManagerMode;
    }
    if (state.microsoftManagerKeyword) {
      inputMicrosoftManagerKeyword.value = state.microsoftManagerKeyword;
    }

    if (state.stepStatuses) {
      for (const [step, status] of Object.entries(state.stepStatuses)) {
        updateStepUI(Number(step), status);
      }
    }

    if (state.logs) {
      for (const entry of state.logs) {
        appendLog(entry);
      }
    }

    updateStatusDisplay(state);
    updateProgressCounter();
    updateOauthProviderUI();
    updateMailProviderUI();
    updateEmailSourceUI();

    if (state.autoRunPausedPhase === 'waiting_email') {
      autoContinueBar.dataset.reason = 'waiting_email';
      autoHint.textContent = getAutoHintText();
      autoContinueBar.style.display = 'flex';
      btnAutoRun.disabled = false;
      inputRunCount.disabled = false;
    } else if (state.autoRunPausedPhase === 'error') {
      autoContinueBar.dataset.reason = 'error';
      autoHint.textContent = t('autoHintError');
      autoContinueBar.style.display = 'flex';
      btnAutoRun.disabled = false;
      inputRunCount.disabled = false;
    }
  } catch (err) {
    console.error('Failed to restore state:', err);
  }
}

function syncPasswordField(state) {
  inputPassword.value = state.customPassword || state.password || '';
}

function updateMailProviderUI() {
  selectMailProvider.value = normalizeMailProviderValue(selectMailProvider.value);
  const useInbucket = false;
  const useMicrosoftManager = true;
  rowMailProvider.style.display = '';
  selectMailProvider.disabled = selectMailProvider.options.length <= 1;
  rowInbucketHost.style.display = useInbucket ? '' : 'none';
  rowInbucketMailbox.style.display = useInbucket ? '' : 'none';
  rowMicrosoftManagerUrl.style.display = useMicrosoftManager ? '' : 'none';
  rowMicrosoftManagerToken.style.display = useMicrosoftManager ? '' : 'none';
  rowMicrosoftManagerMode.style.display = useMicrosoftManager ? '' : 'none';
  rowMicrosoftManagerKeyword.style.display = useMicrosoftManager ? '' : 'none';
}

function updateEmailSourceUI() {
  if (aliasSourceValue) {
    aliasSourceValue.textContent = getEmailSourceLabel();
  }
  inputEmail.placeholder = getEmailPlaceholderText();
  autoHint.textContent = getAutoHintText();
  btnFetchEmail.disabled = false;
  btnFetchEmail.title = getFetchEmailTitle();
  icloudSection.style.display = shouldUseIcloudUi() ? '' : 'none';
}

async function syncRuntimeSettingsBeforeExecution() {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: {
      oauthProvider: selectOauthProvider.value,
      vpsUrl: inputVpsUrl.value.trim(),
      sub2apiBaseUrl: inputSub2apiBaseUrl.value.trim(),
      sub2apiAdminApiKey: inputSub2apiApiKey.value.trim(),
      customPassword: inputPassword.value,
      mailProvider: normalizeMailProviderValue(selectMailProvider.value),
      inbucketHost: inputInbucketHost.value.trim(),
      inbucketMailbox: inputInbucketMailbox.value.trim(),
      microsoftManagerUrl: inputMicrosoftManagerUrl.value.trim(),
      microsoftManagerToken: inputMicrosoftManagerToken.value.trim(),
      microsoftManagerMode: selectMicrosoftManagerMode.value,
      microsoftManagerKeyword: inputMicrosoftManagerKeyword.value.trim(),
    },
  });
}

// ============================================================
// UI Updates
// ============================================================

function updateStepUI(step, status) {
  const statusEl = document.querySelector(`.step-status[data-step="${step}"]`);
  const row = document.querySelector(`.step-row[data-step="${step}"]`);

  if (statusEl) statusEl.textContent = STATUS_ICONS[status] || '';
  if (row) {
    row.className = `step-row ${status}`;
  }

  updateButtonStates();
  updateProgressCounter();
}

function updateProgressCounter() {
  let completed = 0;
  document.querySelectorAll('.step-row').forEach(row => {
    if (row.classList.contains('completed') || row.classList.contains('skipped')) completed++;
  });
  stepsProgress.textContent = `${completed} / ${TOTAL_STEPS}`;
}

function updateButtonStates() {
  const statuses = {};
  document.querySelectorAll('.step-row').forEach(row => {
    const step = Number(row.dataset.step);
    if (row.classList.contains('completed')) statuses[step] = 'completed';
    else if (row.classList.contains('skipped')) statuses[step] = 'skipped';
    else if (row.classList.contains('running')) statuses[step] = 'running';
    else if (row.classList.contains('failed')) statuses[step] = 'failed';
    else if (row.classList.contains('stopped')) statuses[step] = 'stopped';
    else statuses[step] = 'pending';
  });

  const anyRunning = Object.values(statuses).some(s => s === 'running');

  for (let step = 1; step <= TOTAL_STEPS; step++) {
    const btn = document.querySelector(`.step-btn[data-step="${step}"]`);
    const skipBtn = document.querySelector(`.step-skip-btn[data-step="${step}"]`);
    if (!btn) continue;

    const currentStatus = statuses[step];

    if (anyRunning) {
      btn.disabled = true;
      if (skipBtn) skipBtn.disabled = true;
    } else if (step === 1) {
      btn.disabled = false;
    } else {
      const prevStatus = statuses[step - 1];
      btn.disabled = !(
        prevStatus === 'completed'
        || prevStatus === 'skipped'
        || currentStatus === 'failed'
        || currentStatus === 'completed'
        || currentStatus === 'stopped'
        || currentStatus === 'skipped'
      );
    }

    if (skipBtn) {
      skipBtn.disabled = !(currentStatus === 'failed' || currentStatus === 'stopped');
    }
  }

  updateStopButtonState(anyRunning || autoContinueBar.style.display !== 'none');
}

function updateStopButtonState(active) {
  btnStop.disabled = !active;
}

function updateStatusDisplay(state) {
  if (!state || !state.stepStatuses) return;
  lastKnownState = state;
  syncRunMetricsFromState(state);

  statusBar.className = 'status-bar';

  const running = Object.entries(state.stepStatuses).find(([, s]) => s === 'running');
  if (running) {
    displayStatus.textContent = t('statusRunning', { step: running[0] });
    statusBar.classList.add('running');
    return;
  }

  const failed = Object.entries(state.stepStatuses).find(([, s]) => s === 'failed');
  if (failed) {
    displayStatus.textContent = t('statusFailed', { step: failed[0] });
    statusBar.classList.add('failed');
    return;
  }

  const stopped = Object.entries(state.stepStatuses).find(([, s]) => s === 'stopped');
  if (stopped) {
    displayStatus.textContent = t('statusStopped', { step: stopped[0] });
    statusBar.classList.add('stopped');
    return;
  }

  const entries = Object.entries(state.stepStatuses);
  const allProgressed = entries.every(([, s]) => s === 'completed' || s === 'skipped');
  if (allProgressed) {
    displayStatus.textContent = t('statusAllFinished');
    statusBar.classList.add('completed');
    return;
  }

  const lastProgressed = entries
    .filter(([, s]) => s === 'completed' || s === 'skipped')
    .map(([k]) => Number(k))
    .sort((a, b) => b - a)[0];

  if (lastProgressed) {
    displayStatus.textContent = state.stepStatuses[lastProgressed] === 'skipped'
      ? t('statusSkipped', { step: lastProgressed })
      : t('statusDone', { step: lastProgressed });
  } else {
    displayStatus.textContent = t('statusReady');
  }
}

function appendLog(entry) {
  const time = new Date(entry.timestamp).toLocaleTimeString('en-US', { hour12: false });
  const levelLabel = entry.level.toUpperCase();
  const line = document.createElement('div');
  line.className = `log-line log-${entry.level}`;

  const stepMatch = entry.message.match(/Step (\d+)/);
  const stepNum = stepMatch ? stepMatch[1] : null;

  let html = `<span class="log-time">${time}</span> `;
  html += `<span class="log-level log-level-${entry.level}">${levelLabel}</span> `;
  if (stepNum) {
    html += `<span class="log-step-tag step-${stepNum}">S${stepNum}</span>`;
  }
  html += `<span class="log-msg">${escapeHtml(entry.message)}</span>`;

  line.innerHTML = html;
  logArea.appendChild(line);
  logArea.scrollTop = logArea.scrollHeight;
}

function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

async function fetchConfiguredEmail() {
  const defaultLabel = t('btnAuto');
  btnFetchEmail.disabled = true;
  btnFetchEmail.textContent = '...';
  const sourceLabel = getEmailSourceLabel();

  try {
    const response = await chrome.runtime.sendMessage({
      type: 'FETCH_AUTO_EMAIL',
      source: 'sidepanel',
      payload: { generateNew: true },
    });

    if (response?.error) {
      throw new Error(response.error);
    }
    if (!response?.email) {
      throw new Error(`${sourceLabel} email was not returned.`);
    }

    inputEmail.value = response.email;
    showToast(t('fetchedEmail', { email: response.email }), 'success', 2500);
    return response.email;
  } catch (err) {
    showToast(t('autoFetchFailed', { message: err.message }), 'error');
    throw err;
  } finally {
    btnFetchEmail.disabled = false;
    btnFetchEmail.textContent = defaultLabel;
  }
}

function setIcloudLoadingState(loading, summary = '') {
  btnIcloudRefresh.disabled = loading;
  btnIcloudDeleteUsed.disabled = loading;
  btnIcloudLoginDone.disabled = loading;
  if (summary) icloudSummary.textContent = summary;
}

function showIcloudLoginHelp(payload = {}) {
  const loginUrl = String(payload.loginUrl || '').trim();
  const host = loginUrl ? new URL(loginUrl).host : 'icloud.com.cn / icloud.com';
  icloudLoginHelpTitle.textContent = t('icloudLoginHelpTitle');
  icloudLoginHelpText.textContent = t('icloudLoginHelpText', { host });
  icloudLoginHelp.style.display = 'flex';
}

function hideIcloudLoginHelp() {
  icloudLoginHelp.style.display = 'none';
}

function renderIcloudAliases(aliases = []) {
  lastRenderedIcloudAliases = Array.isArray(aliases) ? aliases : [];
  icloudList.innerHTML = '';

  if (!aliases.length) {
    icloudList.innerHTML = `<div class="icloud-empty">${escapeHtml(t('icloudEmpty'))}</div>`;
    icloudSummary.textContent = t('icloudSummaryInitial');
    btnIcloudDeleteUsed.disabled = true;
    return;
  }

  const usedCount = aliases.filter(alias => alias.used).length;
  icloudSummary.textContent = t('icloudAliasesLoaded', { count: aliases.length, usedCount });
  btnIcloudDeleteUsed.disabled = usedCount === 0;

  for (const alias of aliases) {
    const item = document.createElement('div');
    item.className = 'icloud-item';
    item.innerHTML = `
      <div class="icloud-item-main">
        <div class="icloud-item-email">${escapeHtml(alias.email)}</div>
        <div class="icloud-item-meta">
          ${alias.used ? `<span class="icloud-tag used">${escapeHtml(currentLanguage === 'zh-CN' ? '已用' : 'Used')}</span>` : ''}
          ${!alias.used && alias.active ? `<span class="icloud-tag active">${escapeHtml(currentLanguage === 'zh-CN' ? '可用' : 'Active')}</span>` : ''}
          ${alias.label ? `<span class="icloud-tag">${escapeHtml(alias.label)}</span>` : ''}
          ${alias.note ? `<span class="icloud-tag">${escapeHtml(alias.note)}</span>` : ''}
        </div>
      </div>
      <div class="icloud-item-actions">
        <button class="btn btn-outline btn-xs" type="button" data-action="toggle-used">${escapeHtml(alias.used ? t('btnMarkUnused') : t('btnMarkUsed'))}</button>
        <button class="btn btn-outline btn-xs" type="button" data-action="delete">${escapeHtml(t('btnDelete'))}</button>
      </div>
    `;

    item.querySelector('[data-action="toggle-used"]').addEventListener('click', async () => {
      await setSingleIcloudAliasUsedState(alias, !alias.used);
    });
    item.querySelector('[data-action="delete"]').addEventListener('click', async () => {
      await deleteSingleIcloudAlias(alias);
    });
    icloudList.appendChild(item);
  }
}

async function refreshIcloudAliases(options = {}) {
  const { silent = false } = options;

  if (!shouldUseIcloudUi()) {
    hideIcloudLoginHelp();
    return;
  }

  if (!silent) setIcloudLoadingState(true, t('icloudLoading'));
  try {
    const response = await chrome.runtime.sendMessage({
      type: 'LIST_ICLOUD_ALIASES',
      source: 'sidepanel',
      payload: {},
    });

    if (response?.error) throw new Error(response.error);
    hideIcloudLoginHelp();
    renderIcloudAliases(response?.aliases || []);
  } catch (err) {
    icloudList.innerHTML = `<div class="icloud-empty">${escapeHtml(currentLanguage === 'zh-CN' ? '无法加载 iCloud 别名。' : 'Could not load iCloud aliases.')}</div>`;
    icloudSummary.textContent = err.message;
    if (!silent) showToast(t('icloudLoadFailed', { message: err.message }), 'error');
  } finally {
    btnIcloudRefresh.disabled = false;
  }
}

function queueIcloudAliasRefresh() {
  if (!shouldUseIcloudUi()) return;
  if (icloudRefreshQueued) return;
  icloudRefreshQueued = true;
  setTimeout(async () => {
    icloudRefreshQueued = false;
    await refreshIcloudAliases({ silent: true });
  }, 150);
}

async function deleteSingleIcloudAlias(alias) {
  setIcloudLoadingState(true, t('deletingAlias', { email: alias.email }));
  try {
    const response = await chrome.runtime.sendMessage({
      type: 'DELETE_ICLOUD_ALIAS',
      source: 'sidepanel',
      payload: { email: alias.email, anonymousId: alias.anonymousId },
    });
    if (response?.error) throw new Error(response.error);
    showToast(t('deletedAlias', { email: alias.email }), 'success', 2500);
    await refreshIcloudAliases({ silent: true });
  } catch (err) {
    showToast(t('deleteFailed', { message: err.message }), 'error');
    icloudSummary.textContent = err.message;
  } finally {
    btnIcloudRefresh.disabled = false;
  }
}

async function setSingleIcloudAliasUsedState(alias, used) {
  setIcloudLoadingState(true, t('updatingAliasUsed', { email: alias.email, used }));
  try {
    const response = await chrome.runtime.sendMessage({
      type: 'SET_ICLOUD_ALIAS_USED_STATE',
      source: 'sidepanel',
      payload: { email: alias.email, used },
    });
    if (response?.error) throw new Error(response.error);
    showToast(t('updatedAliasUsed', { email: alias.email, used }), 'success', 2500);
    await refreshIcloudAliases({ silent: true });
  } catch (err) {
    showToast(t('updateAliasUsedFailed', { message: err.message }), 'error');
    icloudSummary.textContent = err.message;
  } finally {
    btnIcloudRefresh.disabled = false;
  }
}

async function deleteUsedIcloudAliases() {
  setIcloudLoadingState(true, t('deletingUsedAliases'));
  try {
    const response = await chrome.runtime.sendMessage({
      type: 'DELETE_USED_ICLOUD_ALIASES',
      source: 'sidepanel',
      payload: {},
    });
    if (response?.error) throw new Error(response.error);

    const deleted = response?.deleted || [];
    const skipped = response?.skipped || [];
    const summary = t('deletedUsedAliases', { deleted: deleted.length, skipped: skipped.length });
    showToast(summary, skipped.length ? 'warn' : 'success', 3000);
    await refreshIcloudAliases({ silent: true });
  } catch (err) {
    showToast(t('bulkDeleteFailed', { message: err.message }), 'error');
    icloudSummary.textContent = err.message;
  } finally {
    btnIcloudRefresh.disabled = false;
  }
}

function syncPasswordToggleLabel() {
  btnTogglePassword.textContent = inputPassword.type === 'password' ? t('btnShow') : t('btnHide');
}

// ============================================================
// Button Handlers
// ============================================================

document.querySelectorAll('.step-btn').forEach(btn => {
  btn.addEventListener('click', async () => {
    const step = Number(btn.dataset.step);
    await syncRuntimeSettingsBeforeExecution();
    if (step === 3) {
      const email = inputEmail.value.trim();
      if (!email) {
        showToast(t('pleaseEnterEmailFirst'), 'warn');
        return;
      }
      await chrome.runtime.sendMessage({ type: 'EXECUTE_STEP', source: 'sidepanel', payload: { step, email } });
    } else {
      await chrome.runtime.sendMessage({ type: 'EXECUTE_STEP', source: 'sidepanel', payload: { step } });
    }
  });
});

document.querySelectorAll('.step-skip-btn').forEach(btn => {
  btn.addEventListener('click', async () => {
    const step = Number(btn.dataset.step);
    const response = await chrome.runtime.sendMessage({
      type: 'SKIP_STEP',
      source: 'sidepanel',
      payload: { step },
    });
    if (response?.error) {
      showToast(t('skipFailed', { message: response.error }), 'error');
      return;
    }
    showToast(t('stepSkippedToast', { step }), 'warn', 2000);
  });
});

btnFetchEmail.addEventListener('click', async () => {
  await fetchConfiguredEmail().catch(() => {});
  await refreshIcloudAliases({ silent: true });
});

btnCopyEmail.addEventListener('click', async () => {
  await copyTextValue(inputEmail.value, 'email');
});

btnCopyPassword.addEventListener('click', async () => {
  await copyTextValue(inputPassword.value, 'password');
});

btnPasteVpsUrl.addEventListener('click', async () => {
  await pasteCpaAuthFromClipboard();
});

btnIcloudRefresh.addEventListener('click', async () => {
  await refreshIcloudAliases();
});

btnIcloudDeleteUsed.addEventListener('click', async () => {
  await deleteUsedIcloudAliases();
});

btnIcloudLoginDone.addEventListener('click', async () => {
  btnIcloudLoginDone.disabled = true;
  try {
    const response = await chrome.runtime.sendMessage({
      type: 'CHECK_ICLOUD_SESSION',
      source: 'sidepanel',
      payload: {},
    });
    if (response?.error) {
      throw new Error(response.error);
    }
    hideIcloudLoginHelp();
    showToast(t('icloudSessionReady'), 'success', 3000);
    await refreshIcloudAliases({ silent: true });
  } catch (err) {
    showToast(t('icloudStillNotSignedIn', { message: err.message }), 'warn', 4500);
  } finally {
    btnIcloudLoginDone.disabled = false;
  }
});

btnTogglePassword.addEventListener('click', () => {
  inputPassword.type = inputPassword.type === 'password' ? 'text' : 'password';
  syncPasswordToggleLabel();
});

btnVersion.addEventListener('click', async () => {
  if (!isVersionCheckFinished && !versionCheckInFlight) {
    await checkLatestReleaseVersion();
  }
  await openVersionPage();
});

btnStop.addEventListener('click', async () => {
  btnStop.disabled = true;
  await chrome.runtime.sendMessage({ type: 'STOP_FLOW', source: 'sidepanel', payload: {} });
  showToast(t('stoppingFlow'), 'warn', 2000);
});

// Auto Run
btnAutoRun.addEventListener('click', async () => {
  const totalRuns = parseInt(inputRunCount.value) || 1;
  resetRunMetrics();
  btnAutoRun.disabled = true;
  inputRunCount.disabled = true;
  setAutoRunButton(t('autoRunRunning', { runLabel: '' }));
  await syncRuntimeSettingsBeforeExecution();
  await chrome.runtime.sendMessage({ type: 'AUTO_RUN', source: 'sidepanel', payload: { totalRuns } });
});

btnAutoContinue.addEventListener('click', async () => {
  const reason = autoContinueBar.dataset.reason || 'waiting_email';
  const email = inputEmail.value.trim();
  if (reason === 'waiting_email' && !email) {
    showToast(t('continueNeedEmail'), 'warn');
    return;
  }
  const response = await chrome.runtime.sendMessage({
    type: 'CONTINUE_AUTO_RUN',
    source: 'sidepanel',
    payload: { email },
  });
  if (response?.error) {
    showToast(t('continueFailed', { message: response.error }), 'error');
    return;
  }
  autoContinueBar.style.display = 'none';
  autoContinueBar.dataset.reason = '';
});

// Reset
btnReset.addEventListener('click', async () => {
  if (confirm(t('confirmReset'))) {
    await chrome.runtime.sendMessage({ type: 'RESET', source: 'sidepanel' });
    displayOauthUrl.textContent = t('waiting');
    displayOauthUrl.classList.remove('has-value');
    displayLocalhostUrl.textContent = t('waiting');
    displayLocalhostUrl.classList.remove('has-value');
    inputEmail.value = '';
    displayStatus.textContent = t('statusReady');
    statusBar.className = 'status-bar';
    logArea.innerHTML = '';
    document.querySelectorAll('.step-row').forEach(row => row.className = 'step-row');
    document.querySelectorAll('.step-status').forEach(el => el.textContent = '');
    btnAutoRun.disabled = false;
    inputRunCount.disabled = false;
    setAutoRunButton(t('btnAuto'));
    autoContinueBar.style.display = 'none';
    updateStopButtonState(false);
    updateButtonStates();
    updateProgressCounter();
    resetRunMetrics();
  }
});

// Clear log
btnClearLog.addEventListener('click', () => {
  logArea.innerHTML = '';
});

// Save settings on change
inputEmail.addEventListener('change', async () => {
  const email = inputEmail.value.trim();
  if (email) {
    await chrome.runtime.sendMessage({ type: 'SAVE_EMAIL', source: 'sidepanel', payload: { email } });
  }
});

inputVpsUrl.addEventListener('change', async () => {
  const vpsUrl = inputVpsUrl.value.trim();
  if (vpsUrl) {
    await chrome.runtime.sendMessage({ type: 'SAVE_SETTING', source: 'sidepanel', payload: { vpsUrl } });
  }
});

inputVpsUrl.addEventListener('click', async () => {
  await pasteCpaAuthFromClipboard({ silentIfFilled: true });
});

selectOauthProvider.addEventListener('change', async () => {
  updateOauthProviderUI();
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { oauthProvider: selectOauthProvider.value },
  });
});

inputSub2apiBaseUrl.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { sub2apiBaseUrl: inputSub2apiBaseUrl.value.trim() },
  });
});

inputSub2apiApiKey.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { sub2apiAdminApiKey: inputSub2apiApiKey.value.trim() },
  });
});

inputPassword.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { customPassword: inputPassword.value },
  });
});

checkboxAutoDeleteIcloud.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { autoDeleteUsedIcloudAlias: checkboxAutoDeleteIcloud.checked },
  });
});

selectMailProvider.addEventListener('change', async () => {
  selectMailProvider.value = normalizeMailProviderValue(selectMailProvider.value);
  updateMailProviderUI();
  updateEmailSourceUI();
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING', source: 'sidepanel',
    payload: { mailProvider: normalizeMailProviderValue(selectMailProvider.value) },
  });

  if (shouldUseIcloudUi()) {
    await refreshIcloudAliases({ silent: true });
  } else {
    hideIcloudLoginHelp();
  }
});

selectLanguage.addEventListener('change', async () => {
  applyLanguage(selectLanguage.value || 'zh-CN');
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { language: currentLanguage },
  });
});

inputInbucketMailbox.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { inbucketMailbox: inputInbucketMailbox.value.trim() },
  });
});

inputInbucketHost.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { inbucketHost: inputInbucketHost.value.trim() },
  });
});

inputMicrosoftManagerUrl.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { microsoftManagerUrl: inputMicrosoftManagerUrl.value.trim() },
  });
});

inputMicrosoftManagerToken.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { microsoftManagerToken: inputMicrosoftManagerToken.value.trim() },
  });
});

selectMicrosoftManagerMode.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { microsoftManagerMode: selectMicrosoftManagerMode.value },
  });
});

inputMicrosoftManagerKeyword.addEventListener('change', async () => {
  await chrome.runtime.sendMessage({
    type: 'SAVE_SETTING',
    source: 'sidepanel',
    payload: { microsoftManagerKeyword: inputMicrosoftManagerKeyword.value.trim() },
  });
});

// ============================================================
// Listen for Background broadcasts
// ============================================================

chrome.runtime.onMessage.addListener((message) => {
  switch (message.type) {
    case 'LOG_ENTRY':
      appendLog(message.payload);
      if (message.payload.level === 'error') {
        showToast(message.payload.message, 'error');
      }
      break;

    case 'STEP_STATUS_CHANGED': {
      const { step, status } = message.payload;
      updateStepUI(step, status);
      chrome.runtime.sendMessage({ type: 'GET_STATE', source: 'sidepanel' }).then(updateStatusDisplay);
      if (status === 'completed') {
        chrome.runtime.sendMessage({ type: 'GET_STATE', source: 'sidepanel' }).then(state => {
          syncPasswordField(state);
          if (state.oauthUrl) {
            displayOauthUrl.textContent = state.oauthUrl;
            displayOauthUrl.classList.add('has-value');
          }
          if (state.localhostUrl) {
            displayLocalhostUrl.textContent = state.localhostUrl;
            displayLocalhostUrl.classList.add('has-value');
          }
        });
      }
      break;
    }

    case 'AUTO_RUN_RESET': {
      // Full UI reset for next run
      displayOauthUrl.textContent = t('waiting');
      displayOauthUrl.classList.remove('has-value');
      displayLocalhostUrl.textContent = t('waiting');
      displayLocalhostUrl.classList.remove('has-value');
      inputEmail.value = '';
      displayStatus.textContent = t('statusReady');
      statusBar.className = 'status-bar';
      logArea.innerHTML = '';
      document.querySelectorAll('.step-row').forEach(row => row.className = 'step-row');
      document.querySelectorAll('.step-status').forEach(el => el.textContent = '');
      icloudList.innerHTML = '';
      icloudSummary.textContent = t('icloudSummaryInitial');
      updateStopButtonState(false);
      updateProgressCounter();
      renderRunMetrics();
      break;
    }

    case 'DATA_UPDATED': {
      if (message.payload.email) {
        inputEmail.value = message.payload.email;
      }
      if (message.payload.password !== undefined) {
        inputPassword.value = message.payload.password || '';
      }
      if (message.payload.oauthUrl) {
        displayOauthUrl.textContent = message.payload.oauthUrl;
        displayOauthUrl.classList.add('has-value');
      }
      if (message.payload.localhostUrl) {
        displayLocalhostUrl.textContent = message.payload.localhostUrl;
        displayLocalhostUrl.classList.add('has-value');
      }
      if (message.payload.flowStartTime) {
        chrome.runtime.sendMessage({ type: 'GET_STATE', source: 'sidepanel' })
          .then(updateStatusDisplay)
          .catch(() => {});
      }
      break;
    }

    case 'ICLOUD_LOGIN_REQUIRED': {
      const loginMessage = t('icloudLoginRequiredToast');
      showToast(loginMessage, 'warn', 5000);
      icloudSummary.textContent = loginMessage;
      showIcloudLoginHelp(message.payload || {});
      break;
    }

    case 'ICLOUD_ALIASES_CHANGED':
      queueIcloudAliasRefresh();
      break;

    case 'AUTO_RUN_STATUS': {
      const { phase, currentRun, totalRuns } = message.payload;
      const runLabel = totalRuns > 1 ? ` (${currentRun}/${totalRuns})` : '';
      switch (phase) {
        case 'waiting_email':
          autoContinueBar.dataset.reason = 'waiting_email';
          autoHint.textContent = getAutoHintText();
          autoContinueBar.style.display = 'flex';
          setAutoRunButton(t('autoRunPaused', { runLabel }));
          btnAutoRun.disabled = false;
          inputRunCount.disabled = false;
          updateStopButtonState(true);
          break;
        case 'error':
          autoContinueBar.dataset.reason = 'error';
          autoHint.textContent = t('autoHintError');
          autoContinueBar.style.display = 'flex';
          setAutoRunButton(t('autoRunInterrupted', { runLabel }));
          btnAutoRun.disabled = false;
          inputRunCount.disabled = false;
          updateStopButtonState(false);
          finishActiveRunMetrics(false, Number(lastKnownState?.flowStartTime || 0));
          break;
        case 'running':
          autoContinueBar.dataset.reason = '';
          autoContinueBar.style.display = 'none';
          setAutoRunButton(t('autoRunRunning', { runLabel }));
          btnAutoRun.disabled = true;
          inputRunCount.disabled = true;
          updateStopButtonState(true);
          break;
        case 'complete':
          btnAutoRun.disabled = false;
          inputRunCount.disabled = false;
          setAutoRunButton(t('btnAuto'));
          autoContinueBar.style.display = 'none';
          autoContinueBar.dataset.reason = '';
          updateStopButtonState(false);
          break;
        case 'stopped':
          btnAutoRun.disabled = false;
          inputRunCount.disabled = false;
          setAutoRunButton(t('btnAuto'));
          autoContinueBar.style.display = 'none';
          autoContinueBar.dataset.reason = '';
          updateStopButtonState(false);
          finishActiveRunMetrics(false, Number(lastKnownState?.flowStartTime || 0));
          break;
      }
      chrome.runtime.sendMessage({ type: 'GET_STATE', source: 'sidepanel' })
        .then(updateStatusDisplay)
        .catch(() => {});
      break;
    }
  }
});

// ============================================================
// Theme Toggle
// ============================================================

const btnTheme = document.getElementById('btn-theme');

function setTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('multipage-theme', theme);
}

function initTheme() {
  const saved = localStorage.getItem('multipage-theme');
  if (saved) {
    setTheme(saved);
  } else if (window.matchMedia('(prefers-color-scheme: dark)').matches) {
    setTheme('dark');
  }
}

btnTheme.addEventListener('click', () => {
  const current = document.documentElement.getAttribute('data-theme');
  setTheme(current === 'dark' ? 'light' : 'dark');
});

// ============================================================
// Init
// ============================================================

initTheme();
applyLanguage(currentLanguage);
renderVersionBadge();
restoreState().then(() => {
  syncPasswordToggleLabel();
  updateButtonStates();
  if (shouldUseIcloudUi()) {
    refreshIcloudAliases({ silent: true });
  }
  checkLatestReleaseVersion().catch(() => {});
});
