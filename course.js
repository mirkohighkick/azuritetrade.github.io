/* ==========================================================================
   course.js — общая логика для всех страниц курса.
   Подключается на КАЖДОЙ странице (index.html и module-N.html) через:
   <script src="course.js"></script>

   Хранит:
   - имя сотрудника (localStorage, на этом устройстве/браузере)
   - результат по каждому пройденному модулю
   Считает суммарный балл по всем модулям и отправляет его в Netlify Forms
   вместе с результатом за конкретный модуль при каждой отправке.
========================================================================== */
window.CourseTracker = (function () {
  var NAME_KEY = 'csx_employee_name';
  var PROGRESS_KEY = 'csx_progress'; // { moduleId: {label, correct, total, date} }
  var MODULE_ORDER = ['module-1', 'module-2', 'module-3', 'module-4', 'module-5', 'module-6'];
  var PASS_THRESHOLD = 0.8; // нужно набрать минимум 80%, чтобы разблокировать следующий модуль

  function getName() {
    return localStorage.getItem(NAME_KEY) || '';
  }
  function setName(name) {
    localStorage.setItem(NAME_KEY, name.trim());
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

  function encodeForm(data) {
    return Object.keys(data)
      .map(function (k) { return encodeURIComponent(k) + '=' + encodeURIComponent(data[k]); })
      .join('&');
  }

  // Единственная точка отправки на почту/в Netlify — вызывается один раз,
  // когда пользователь нажимает "Завершить обучение" на последнем модуле.
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
      'form-name': 'course-results',
      name: getName(),
      modules_completed: totals.modulesCompleted + ' из 6',
      overall_score: totals.totalCorrect + ' из ' + totals.totalQuestions,
      overall_percent: totals.pct + '%',
      breakdown: breakdownText,
      date: new Date().toLocaleString('ru-RU')
    };

    return fetch('/', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: encodeForm(payload)
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

  // Показывает окно-заглушку "Введите имя", если имя ещё не сохранено.
  // Вызывать на КАЖДОЙ странице курса (и на index, и на модулях) —
  // если имя уже есть, окно просто не появится.
  function requireName(onReady) {
    var existing = getName();
    if (existing) {
      onReady(existing);
      return;
    }

    var overlay = document.createElement('div');
    overlay.id = 'csx-gate';
    overlay.innerHTML =
      '<div class="csx-gate-card">' +
      '  <h2>Добро пожаловать на курс</h2>' +
      '  <p>Укажите имя и фамилию — это нужно, чтобы засчитать вам прохождение модулей</p>' +
      '  <input type="text" id="csx-gate-input" placeholder="Например: Иван Иванов">' +
      '  <div class="csx-gate-error" id="csx-gate-error">Пожалуйста, введите имя</div>' +
      '  <button id="csx-gate-btn">Продолжить</button>' +
      '</div>';

    var style = document.createElement('style');
    style.textContent =
      '#csx-gate { position: fixed; inset: 0; background: rgba(22,28,77,0.92); ' +
      '  display: flex; align-items: center; justify-content: center; z-index: 9999; ' +
      '  font-family: Calibri, "Segoe UI", Arial, sans-serif; }' +
      '.csx-gate-card { background: #fff; border-radius: 16px; padding: 36px 32px; ' +
      '  max-width: 380px; width: 90%; text-align: center; box-shadow: 0 20px 50px rgba(0,0,0,0.3); }' +
      '.csx-gate-card h2 { font-family: Georgia, serif; color: #1E2761; margin: 0 0 10px 0; font-size: 22px; }' +
      '.csx-gate-card p { color: #5B6172; font-size: 13.5px; margin: 0 0 20px 0; line-height: 1.5; }' +
      '.csx-gate-card input { width: 100%; padding: 12px 14px; border-radius: 8px; ' +
      '  border: 1px solid #D8DEEE; font-size: 14px; font-family: inherit; margin-bottom: 6px; }' +
      '.csx-gate-card input:focus { outline: none; border-color: #3454D1; }' +
      '.csx-gate-error { display: none; color: #C23B3B; font-size: 12.5px; text-align: left; margin-bottom: 10px; }' +
      '.csx-gate-error.is-shown { display: block; }' +
      '.csx-gate-card button { width: 100%; background: #1E2761; color: #fff; border: none; ' +
      '  font-weight: 700; font-size: 14.5px; padding: 13px; border-radius: 999px; cursor: pointer; margin-top: 6px; }';

    document.head.appendChild(style);
    document.body.appendChild(overlay);

    var input = document.getElementById('csx-gate-input');
    var error = document.getElementById('csx-gate-error');
    var btn = document.getElementById('csx-gate-btn');

    function submit() {
      var val = input.value.trim();
      if (!val) {
        error.classList.add('is-shown');
        input.focus();
        return;
      }
      setName(val);
      document.body.removeChild(overlay);
      onReady(val);
    }

    btn.addEventListener('click', submit);
    input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') submit();
    });
    input.focus();
  }

  return {
    getName: getName,
    setName: setName,
    getProgress: getProgress,
    getTotals: getTotals,
    saveModuleResult: saveModuleResult,
    submitFinalResults: submitFinalResults,
    isModuleUnlocked: isModuleUnlocked,
    requireName: requireName,
    PASS_THRESHOLD: PASS_THRESHOLD
  };
})();
