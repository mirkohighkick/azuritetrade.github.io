window.CourseTracker = (function () {
  const msalConfig = {
    auth: {
      clientId: "650f8341-c424-49b8-9802-cffd488d487e",
      authority: "https://login.microsoftonline.com/4ef2b07f-ee6e-498a-89fd-9612b6ada89d",
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
  };

  let msalInstance = null;
  let currentAccount = null;

  async function initMSAL() {
    if (!msalInstance) {
      msalInstance = new msal.PublicClientApplication(msalConfig);
      await msalInstance.initialize();
    }
  }

  async function requireAuth(onReady) {
    await initMSAL();
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      currentAccount = accounts[0];
      onReady(currentAccount);
      return;
    }

    const overlay = document.createElement('div');
    overlay.id = 'csx-gate';
    overlay.innerHTML = `
      <div class="csx-gate-card">
        <h2>Добро пожаловать на курс</h2>
        <p>Для прохождения необходимо войти через корпоративный аккаунт Microsoft</p>
        <button id="csx-login-btn">Войти через Microsoft</button>
      </div>`;
    const style = document.createElement('style');
    style.textContent = `#csx-gate {position:fixed;inset:0;background:rgba(22,28,77,0.95);display:flex;align-items:center;justify-content:center;z-index:9999;}
      .csx-gate-card {background:#fff;border-radius:16px;padding:40px 32px;text-align:center;box-shadow:0 20px 50px rgba(0,0,0,0.3);max-width:420px;}
      #csx-login-btn {background:#0078D4;color:white;padding:14px 32px;border:none;border-radius:8px;font-size:16px;cursor:pointer;}`;
    document.head.appendChild(style);
    document.body.appendChild(overlay);

    document.getElementById('csx-login-btn').onclick = async () => {
      try {
        const resp = await msalInstance.loginPopup({scopes: ["User.Read"]});
        currentAccount = resp.account;
        overlay.remove();
        onReady(currentAccount);
      } catch (e) {
        alert("Не удалось войти. Попробуйте ещё раз.");
      }
    };
  }

  function getFullName() { return currentAccount ? (currentAccount.name || currentAccount.username || '') : ''; }
  function getEmail() { return currentAccount ? currentAccount.username : ''; }

  const PROGRESS_KEY = 'csx_progress';
  const MODULE_ORDER = ['module-1', 'module-2', 'module-3', 'module-4', 'module-5', 'module-6'];

  function getProgress() {
    try { return JSON.parse(localStorage.getItem(PROGRESS_KEY)) || {}; } catch (e) { return {}; }
  }

  function saveModuleResult(moduleId, moduleLabel, correct, total) {
    const progress = getProgress();
    progress[moduleId] = {label: moduleLabel, correct, total, date: new Date().toISOString()};
    localStorage.setItem(PROGRESS_KEY, JSON.stringify(progress));
  }

  function getTotals() {
    const progress = getProgress();
    let tc = 0, tq = 0, mc = 0;
    Object.keys(progress).forEach(k => {
      tc += progress[k].correct;
      tq += progress[k].total;
      mc++;
    });
    return {totalCorrect: tc, totalQuestions: tq, modulesCompleted: mc, pct: tq ? Math.round(tc / tq * 100) : 0};
  }

  function isModuleUnlocked(moduleNum) {
    if (moduleNum <= 1) return true;
    return !!getProgress()[MODULE_ORDER[moduleNum - 2]];
  }

  async function submitFinalResults() {
    const totals = getTotals();
    const progress = getProgress();
    const breakdownText = MODULE_ORDER.map(id => progress[id] ? `${progress[id].label}: ${progress[id].correct}/${progress[id].total}` : '').filter(Boolean).join('  |  ');

    const payload = {
      'form-name': 'course-results',
      name: getFullName(),
      email: getEmail(),
      modules_completed: totals.modulesCompleted + ' из 6',
      overall_score: totals.totalCorrect + ' из ' + totals.totalQuestions,
      overall_percent: totals.pct + '%',
      breakdown: breakdownText,
      date: new Date().toLocaleString('ru-RU')
    };

    return fetch('/', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: Object.keys(payload).map(k => encodeURIComponent(k) + '=' + encodeURIComponent(payload[k])).join('&')
    });
  }

  return {
    requireAuth,
    getFullName,
    getEmail,
    saveModuleResult,
    getTotals,
    submitFinalResults,
    isModuleUnlocked
  };
})();