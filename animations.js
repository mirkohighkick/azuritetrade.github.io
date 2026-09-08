/* ==========================================================================
   animations.js — общие анимации для всех страниц курса.
   Подключается после styles.css, до course.js:
   <script src="animations.js"></script>

   Даёт два инструмента:
   - CourseAnim.drawPath(svgPathEl, durationMs) — анимированная прорисовка
     SVG-линии (используется для маршрута и рукописных завитков).
   - CourseAnim.revealOnScroll(selector) — плавное появление элементов при
     прокрутке до них (для секций, иконок, карточек).

   Уважает prefers-reduced-motion: если у пользователя эта настройка
   включена, все анимации отключаются, элементы показываются сразу.
========================================================================== */
(function () {
  function prefersReducedMotion() {
    return window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  }

  function drawPath(path, duration) {
    if (!path) return;
    var length;
    try {
      length = path.getTotalLength();
    } catch (e) {
      return; // элемент ещё не в DOM или не SVG-путь
    }

    if (prefersReducedMotion()) {
      path.style.strokeDasharray = 'none';
      path.style.strokeDashoffset = '0';
      return;
    }

    path.style.strokeDasharray = length;
    path.style.strokeDashoffset = length;
    // форсируем reflow, чтобы браузер применил стартовое состояние до transition
    path.getBoundingClientRect();
    path.style.transition = 'stroke-dashoffset ' + (duration || 1600) + 'ms cubic-bezier(0.65, 0, 0.35, 1)';
    requestAnimationFrame(function () {
      path.style.strokeDashoffset = '0';
    });
  }

  function revealOnScroll(selector, options) {
    var els = document.querySelectorAll(selector);
    if (!els.length) return;

    if (prefersReducedMotion() || !('IntersectionObserver' in window)) {
      els.forEach(function (el) { el.classList.add('is-revealed'); });
      return;
    }

    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (entry.isIntersecting) {
          entry.target.classList.add('is-revealed');
          io.unobserve(entry.target);
        }
      });
    }, Object.assign({ threshold: 0.15, rootMargin: '0px 0px -40px 0px' }, options || {}));

    els.forEach(function (el) { io.observe(el); });
  }

  window.CourseAnim = {
    drawPath: drawPath,
    revealOnScroll: revealOnScroll,
    prefersReducedMotion: prefersReducedMotion
  };
})();
