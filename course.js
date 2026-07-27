/* ==========================================================================
   course.js — общая логика для всех страниц курса.
   Подключается на КАЖДОЙ странице (index.html и module-N.html) через:
   <script src="course.js"></script>

   Авторизация: вход через корпоративный аккаунт Microsoft (MSAL / Entra ID).
   Хранит: результат по каждому пройденному модулю (localStorage).
   Считает суммарный балл по всем модулям и отправляет его через Power
   Automate (HTTP-триггер) при нажатии "Завершить обучение" на модуле 6.
========================================================================== */
window.CourseTracker = (function () {
  var PROGRESS_KEY = 'csx_progress'; // { moduleId: {label, correct, total, date} }
  var MODULE_ORDER = ['module-1', 'module-2', 'module-3', 'module-4', 'module-5', 'module-6'];
  var PASS_THRESHOLD = 0.8; // нужно набрать минимум 80%, чтобы разблокировать следующий модуль

  // ---- Microsoft Entra ID / MSAL -----------------------------------------
  // Client ID и Tenant ID из вашего App registration в Azure Portal.
  // ВАЖНО: при переезде на GitHub Pages добавьте новый Redirect URI
  // (https://<username>.github.io/<repo>/) в Azure Portal → App registrations
  // → ваше приложение → Authentication → Single-page application.
  var msalConfig = {
    auth: {
      clientId: "650f8341-c424-49b8-9802-cffd488d487e",
      authority: "https://login.microsoftonline.com/4ef2b07f-ee6e-498a-89fd-9612b6ada89d",
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: { cacheLocation: "localStorage", storeAuthStateInCookie: true }
  };

  var msalInstance = null;
  var currentAccount = null;

  function initMSAL() {
    if (!msalInstance) {
      msalInstance = new msal.PublicClientApplication(msalConfig);
    }
    return msalInstance.initialize().then(function () {
      return msalInstance.handleRedirectPromise();
    });
  }

  function getName() {
    return currentAccount ? (currentAccount.name || currentAccount.username || '') : '';
  }
  function getEmail() {
    return currentAccount ? (currentAccount.username || '') : '';
  }

  // Показывает окно входа через Microsoft, если пользователь ещё не
  // авторизован в этой сессии браузера. Вызывать на КАЖДОЙ странице курса.
  function requireAuth(onReady) {
    initMSAL().then(function () {
      var accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        currentAccount = accounts[0];
        onReady(getName());
        return;
      }

      var overlay = document.createElement('div');
      overlay.id = 'csx-gate';
      overlay.innerHTML =
        '<div class="csx-gate-card">' +
        '  <h2>Добро пожаловать на курс</h2>' +
        '  <p>Для прохождения необходимо войти через корпоративный аккаунт Microsoft</p>' +
        '  <button id="csx-login-btn">Войти через Microsoft</button>' +
        '</div>';

      var style = document.createElement('style');
      style.textContent =
        '#csx-gate { position: fixed; inset: 0; background: rgba(22,28,77,0.92); ' +
        '  display: flex; align-items: center; justify-content: center; z-index: 9999; ' +
        '  font-family: Calibri, "Segoe UI", Arial, sans-serif; }' +
        '.csx-gate-card { background: #fff; border-radius: 16px; padding: 36px 32px; ' +
        '  max-width: 380px; width: 90%; text-align: center; box-shadow: 0 20px 50px rgba(0,0,0,0.3); }' +
        '.csx-gate-card h2 { font-family: Georgia, serif; color: #1E2761; margin: 0 0 10px 0; font-size: 22px; }' +
        '.csx-gate-card p { color: #5B6172; font-size: 13.5px; margin: 0 0 22px 0; line-height: 1.5; }' +
        '#csx-login-btn { width: 100%; background: #0078D4; color: #fff; border: none; ' +
        '  font-weight: 700; font-size: 14.5px; padding: 13px; border-radius: 8px; cursor: pointer; }' +
        '#csx-login-btn:hover { background: #006ABE; }';

      document.head.appendChild(style);
      document.body.appendChild(overlay);

      document.getElementById('csx-login-btn').addEventListener('click', function () {
        msalInstance.loginPopup({ scopes: ["User.Read"] }).then(function (resp) {
          currentAccount = resp.account;
          overlay.remove();
          onReady(getName());
        }).catch(function (e) {
          console.error(e);
          alert("Не удалось войти. Попробуйте ещё раз.");
        });
      });
    });
  }

  function getProgress() {
    try {
      return JSON.parse(localStorage.getItem(PROGRESS_KEY)) || {};
    } catch (e) {
      return {};
    }
  }

  function saveModuleResult(moduleId, moduleLabel, correct, total) {
    var progress = getProgress();
    progress[moduleId] = {
      label: moduleLabel,
      correct: correct,
      total: total,
      date: new Date().toISOString()
    };
    localStorage.setItem(PROGRESS_KEY, JSON.stringify(progress));
    return progress;
  }

  function getTotals() {
    var progress = getProgress();
    var totalCorrect = 0, totalQuestions = 0, modulesCompleted = 0;
    Object.keys(progress).forEach(function (key) {
      totalCorrect += progress[key].correct;
      totalQuestions += progress[key].total;
      modulesCompleted++;
    });
    return {
      totalCorrect: totalCorrect,
      totalQuestions: totalQuestions,
      modulesCompleted: modulesCompleted,
      pct: totalQuestions ? Math.round((totalCorrect / totalQuestions) * 100) : 0
    };
  }

  // Модуль под номером moduleNum (1-based) разблокирован, если это модуль 1,
  // либо предыдущий модуль уже есть в сохранённом прогрессе.
  function isModuleUnlocked(moduleNum) {
    if (moduleNum <= 1) return true;
    var progress = getProgress();
    return !!progress[MODULE_ORDER[moduleNum - 2]];
  }

  // ---- Отправка итогов через Power Automate -------------------------------
  // 1. В Power Automate создайте Flow: триггер "When an HTTP request is
  //    received" (Instant cloud flow).
  // 2. В качестве действия добавьте, например, "Add a row into a table"
  //    (Excel в SharePoint/OneDrive) или "Send an email (V2)" с телом из
  //    полей payload ниже.
  // 3. Сохраните Flow — Power Automate сгенерирует HTTP POST URL.
  //    Вставьте этот URL вместо строки ниже.
  var POWER_AUTOMATE_URL = "ВСТАВЬТЕ_СЮДА_URL_ИЗ_POWER_AUTOMATE";

  // Единственная точка отправки — вызывается один раз, когда пользователь
  // нажимает "Завершить обучение" на последнем модуле.
  function submitFinalResults() {
    var totals = getTotals();
    var progress = getProgress();
    var breakdown = MODULE_ORDER
      .map(function (id) { return progress[id] ? { id: id, label: progress[id].label, correct: progress[id].correct, total: progress[id].total } : null; })
      .filter(Boolean);

    var breakdownText = breakdown
      .map(function (row) { return row.label + ': ' + row.correct + '/' + row.total; })
      .join('  |  ');

    var payload = {
      name: getName(),
      email: getEmail(),
      modules_completed: totals.modulesCompleted + ' из 6',
      overall_score: totals.totalCorrect + ' из ' + totals.totalQuestions,
      overall_percent: totals.pct + '%',
      breakdown: breakdownText,
      date: new Date().toLocaleString('ru-RU')
    };

    // mode: 'no-cors' — HTTP-триггеры Power Automate обычно не отдают
    // CORS-заголовки, поэтому ответ прочитать нельзя (opaque response),
    // но сам запрос долетает и запускает Flow. Это стандартный обходной путь.
    return fetch(POWER_AUTOMATE_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    }).catch(function (err) {
      console.warn('Не удалось отправить итоговый результат:', err);
    }).then(function () {
      return {
        totalCorrect: totals.totalCorrect,
        totalQuestions: totals.totalQuestions,
        pct: totals.pct,
        breakdown: breakdown
      };
    });
  }

  return {
    getName: getName,
    getEmail: getEmail,
    getProgress: getProgress,
    getTotals: getTotals,
    saveModuleResult: saveModuleResult,
    submitFinalResults: submitFinalResults,
    isModuleUnlocked: isModuleUnlocked,
    requireAuth: requireAuth,
    PASS_THRESHOLD: PASS_THRESHOLD
  };
})();
