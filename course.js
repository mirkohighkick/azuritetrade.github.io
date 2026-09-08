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

  // ---- Авторизация: временно ПРОСТОЕ ИМЯ вместо входа через Microsoft ----
  // Вход через Microsoft (MSAL) временно отключён — тенант блокирует
  // согласие пользователя на разрешение User.Read, и нужен админ для
  // "Grant admin consent". Пока используется простой ввод ФИО, который
  // сохраняется в localStorage. Когда согласие администратора будет
  // получено, эту функцию можно заменить обратно на MSAL-логин.
  var NAME_KEY = 'csx_employee_name';

  function getName() {
    try { return localStorage.getItem(NAME_KEY) || ''; } catch (e) { return ''; }
  }
  function setName(name) {
    try { localStorage.setItem(NAME_KEY, name); } catch (e) {}
  }
  function getEmail() {
    return ''; // почта недоступна без входа через Microsoft
  }

  // Показывает окно с полем ФИО, если имя ещё не сохранено в этом браузере.
  // Вызывать на КАЖДОЙ странице курса.
  function requireAuth(onReady) {
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
      '  <p>Введите ваше имя и фамилию, чтобы начать обучение</p>' +
      '  <input id="csx-name-input" type="text" placeholder="Имя Фамилия" />' +
      '  <button id="csx-login-btn">Начать обучение</button>' +
      '</div>';

    document.body.appendChild(overlay);

    var input = document.getElementById('csx-name-input');
    var submit = function () {
      var name = input.value.trim();
      if (!name) {
        input.focus();
        return;
      }
      setName(name);
      overlay.remove();
      onReady(name);
    };

    document.getElementById('csx-login-btn').addEventListener('click', submit);
    input.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') submit();
    });
    input.focus();
  }

  // ---- Случайная база вопросов (общий "движок" для любого модуля) --------
  // Используется так: у каждого модуля своя база вопросов (массив объектов)
  // в отдельном файле, например question-bank-module6.js. На странице модуля
  // вызывается:
  //   CourseTracker.renderQuiz(MODULE6_QUESTIONS, container, 10, 'm6q')
  // — это отрисует 10 случайных вопросов из базы (с перемешанными
  // вариантами ответов) внутрь container и вернёт NodeList отрисованных
  // .question-элементов, с которыми дальше работает логика проверки
  // (checkBtn/retryBtn).
  //
  // Формат вопроса в базе:
  //   { section, text, options: [...], explain,
  //     type: 'single' (по умолчанию) | 'multi',
  //     correctIndex: N          — для type: 'single'
  //     correctIndexes: [N, M]   — для type: 'multi' (несколько верных)
  //   }

  function shuffle(arr) {
    var a = arr.slice();
    for (var i = a.length - 1; i > 0; i--) {
      var j = Math.floor(Math.random() * (i + 1));
      var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
    }
    return a;
  }

  function renderQuiz(bank, containerEl, count, namePrefix) {
    var howMany = count && count < bank.length ? count : bank.length;
    var picked = shuffle(bank).slice(0, howMany);
    containerEl.innerHTML = '';

    picked.forEach(function (q, idx) {
      var isMulti = q.type === 'multi';
      var correctSet = isMulti ? q.correctIndexes : [q.correctIndex];

      // перемешиваем порядок вариантов ответа, но не теряем, какие из них верные
      var optOrder = shuffle(q.options.map(function (_, i) { return i; }));
      var newCorrectSet = correctSet.map(function (origIndex) { return optOrder.indexOf(origIndex); });

      var qDiv = document.createElement('div');
      qDiv.className = 'question';
      qDiv.setAttribute('data-correct', newCorrectSet.join(','));
      qDiv.setAttribute('data-type', isMulti ? 'multi' : 'single');

      var qname = (namePrefix || 'q') + '_' + idx;
      var inputType = isMulti ? 'checkbox' : 'radio';
      var html = '';
      if (q.section) {
        html += '<span class="question-tag">' + q.section + '</span>';
      }
      html += '<p class="question-title">' + (idx + 1) + '. ' + q.text +
        (isMulti ? ' <span class="multi-hint">(возможно несколько вариантов)</span>' : '') + '</p>';
      optOrder.forEach(function (origIndex, i) {
        html += '<label class="option"><input type="' + inputType + '" name="' + qname + '" value="' + i + '"><span>' +
          q.options[origIndex] + '</span></label>';
      });
      html += '<div class="explain">' + q.explain + '</div>';

      qDiv.innerHTML = html;
      containerEl.appendChild(qDiv);
    });

    return containerEl.querySelectorAll('.question');
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
    shuffle: shuffle,
    renderQuiz: renderQuiz,
    PASS_THRESHOLD: PASS_THRESHOLD
  };
})();
