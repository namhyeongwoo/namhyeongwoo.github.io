(function () {
  "use strict";

  // Top bar border once the page scrolls.
  var bar = document.querySelector(".topbar");
  function onScroll() { bar.classList.toggle("scrolled", window.scrollY > 8); }
  onScroll();
  window.addEventListener("scroll", onScroll, { passive: true });

  if (!("IntersectionObserver" in window)) return;

  // Highlight the section currently in view.
  var links = Array.prototype.slice.call(document.querySelectorAll(".nav a"));
  var targets = links.map(function (a) { return document.querySelector(a.getAttribute("href")); });
  var spy = new IntersectionObserver(function (entries) {
    entries.forEach(function (e) {
      if (!e.isIntersecting) return;
      links.forEach(function (a, i) { a.classList.toggle("active", targets[i] === e.target); });
    });
  }, { rootMargin: "-40% 0px -55% 0px" });
  targets.forEach(function (t) { if (t) spy.observe(t); });
})();
