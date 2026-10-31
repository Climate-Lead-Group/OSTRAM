// OSTRAM site behaviour: theme toggle, phone menu, Ctrl K search, stepper,
// hover diagram and the choices sheet. No external libraries.
(function () {
  "use strict";
  var root = document.documentElement;

  function systemDark() {
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  }
  function currentTheme() {
    return root.dataset.theme || (systemDark() ? "dark" : "light");
  }
  var toggle = document.querySelector(".theme-toggle");
  if (toggle) {
    toggle.addEventListener("click", function () {
      var next = currentTheme() === "dark" ? "light" : "dark";
      root.dataset.theme = next;
      try { localStorage.setItem("ostram-theme", next); } catch (error) { /* ignore */ }
    });
  }

  var menu = document.querySelector(".menu-toggle");
  if (menu) {
    menu.addEventListener("click", function () {
      var open = document.body.classList.toggle("menu-open");
      menu.setAttribute("aria-expanded", open ? "true" : "false");
    });
  }

  // Search ---------------------------------------------------------------
  var overlay = document.querySelector(".search");
  var input = overlay && overlay.querySelector("input");
  var list = overlay && overlay.querySelector(".search-results");
  var index = window.OSTRAM_SEARCH || [];
  var base = (function () {
    var script = document.querySelector('script[src$="search-index.js"]');
    return script ? script.getAttribute("src").replace(/assets\/search-index\.js$/, "") : "";
  })();
  var selected = 0;

  function openSearch() {
    if (!overlay) return;
    overlay.hidden = false;
    input.value = "";
    render("");
    input.focus();
  }
  function closeSearch() { if (overlay) overlay.hidden = true; }
  function escapeHtml(text) {
    return text.replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; });
  }
  function score(entry, terms) {
    var h = entry.h.toLowerCase(), p = entry.p.toLowerCase(), t = entry.t.toLowerCase(), total = 0;
    for (var i = 0; i < terms.length; i++) {
      var term = terms[i], found = false;
      if (h.indexOf(term) >= 0) { total += 6; found = true; }
      if (p.indexOf(term) >= 0) { total += 3; found = true; }
      if (t.indexOf(term) >= 0) { total += 1; found = true; }
      if (!found) return 0;
    }
    return total;
  }
  function snippet(text, term) {
    var at = text.toLowerCase().indexOf(term);
    if (at < 0) return text.slice(0, 140);
    var start = Math.max(0, at - 50);
    return (start > 0 ? "..." : "") + text.slice(start, start + 160);
  }
  function render(query) {
    var terms = query.toLowerCase().split(/\s+/).filter(Boolean);
    var hits = [];
    if (terms.length) {
      for (var i = 0; i < index.length; i++) {
        var s = score(index[i], terms);
        if (s > 0) hits.push([s, i]);
      }
      hits.sort(function (a, b) { return b[0] - a[0] || a[1] - b[1]; });
    }
    selected = 0;
    list.innerHTML = hits.slice(0, 20).map(function (hit, n) {
      var entry = index[hit[1]];
      return '<li><a class="' + (n === 0 ? "sel" : "") + '" href="' + base + entry.u + '">' +
        '<span class="where">' + escapeHtml(entry.p) + "</span>" + escapeHtml(entry.h) +
        '<span class="snip">' + escapeHtml(snippet(entry.t, terms[0])) + "</span></a></li>";
    }).join("") || (terms.length ? '<li class="small">No match.</li>' : "");
  }
  function move(step) {
    var links = list.querySelectorAll("a");
    if (!links.length) return;
    links[selected].classList.remove("sel");
    selected = (selected + step + links.length) % links.length;
    links[selected].classList.add("sel");
    links[selected].scrollIntoView({ block: "nearest" });
  }
  document.addEventListener("keydown", function (event) {
    if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
      event.preventDefault();
      openSearch();
    } else if (overlay && !overlay.hidden) {
      if (event.key === "Escape") closeSearch();
      if (event.key === "ArrowDown") { event.preventDefault(); move(1); }
      if (event.key === "ArrowUp") { event.preventDefault(); move(-1); }
      if (event.key === "Enter") {
        var link = list.querySelectorAll("a")[selected];
        if (link) { event.preventDefault(); window.location.href = link.getAttribute("href"); }
      }
    }
  });
  if (overlay) {
    input.addEventListener("input", function () { render(input.value); });
    overlay.addEventListener("click", function (event) { if (event.target === overlay) closeSearch(); });
  }
  var opener = document.querySelector(".search-open");
  if (opener) opener.addEventListener("click", openSearch);

  // Stepper ----------------------------------------------------------------
  Array.prototype.forEach.call(document.querySelectorAll("[data-stepper]"), function (stepper) {
    var tabs = stepper.querySelectorAll("[data-tab]");
    var panels = stepper.querySelectorAll("[data-panel]");
    function show(panel, n) {
      var steps = panel.querySelectorAll(".step");
      n = Math.max(0, Math.min(steps.length - 1, n));
      panel.dataset.at = String(n);
      Array.prototype.forEach.call(steps, function (step, i) { step.classList.toggle("on", i === n); });
      panel.querySelector(".stepper-count").textContent = "Step " + (n + 1) + " of " + steps.length;
    }
    Array.prototype.forEach.call(panels, function (panel) {
      show(panel, 0);
      Array.prototype.forEach.call(panel.querySelectorAll("[data-step]"), function (button) {
        button.addEventListener("click", function () {
          show(panel, Number(panel.dataset.at) + Number(button.dataset.step));
        });
      });
    });
    Array.prototype.forEach.call(tabs, function (tab) {
      tab.addEventListener("click", function () {
        Array.prototype.forEach.call(tabs, function (other) { other.setAttribute("aria-selected", other === tab ? "true" : "false"); });
        Array.prototype.forEach.call(panels, function (panel) { panel.hidden = panel.dataset.panel !== tab.dataset.tab; });
      });
    });
  });

  // Hover diagram -------------------------------------------------------------
  Array.prototype.forEach.call(document.querySelectorAll("[data-hover]"), function (figure) {
    var items = figure.querySelectorAll("[data-note-for]");
    function show(key) {
      Array.prototype.forEach.call(items, function (item) { item.hidden = item.dataset.noteFor !== key; });
      Array.prototype.forEach.call(figure.querySelectorAll(".box"), function (box) { box.classList.toggle("on", box.dataset.note === key); });
    }
    Array.prototype.forEach.call(figure.querySelectorAll(".box"), function (box) {
      box.addEventListener("mouseenter", function () { show(box.dataset.note); });
      box.addEventListener("focus", function () { show(box.dataset.note); });
      box.addEventListener("click", function () { show(box.dataset.note); });
    });
  });

  // Choices sheet --------------------------------------------------------------
  Array.prototype.forEach.call(document.querySelectorAll("[data-sheet]"), function (sheet) {
    var buttons = sheet.querySelectorAll("[data-rule]");
    var cells = sheet.querySelectorAll("td.c");
    Array.prototype.forEach.call(buttons, function (button) {
      button.setAttribute("aria-pressed", "false");
      button.addEventListener("click", function () {
        var on = button.getAttribute("aria-pressed") !== "true";
        Array.prototype.forEach.call(buttons, function (other) { other.setAttribute("aria-pressed", "false"); });
        Array.prototype.forEach.call(cells, function (cell) { cell.classList.remove("hit"); });
        if (!on) return;
        button.setAttribute("aria-pressed", "true");
        var parts = button.dataset.hit.split(" "), mode = "";
        var wanted = { build: {}, run: {} };
        parts.forEach(function (part) {
          var m = part.match(/^(build|run):(.*)$/);
          if (m) { mode = m[1]; part = m[2]; }
          if (part) wanted[mode][part] = true;
        });
        Array.prototype.forEach.call(cells, function (cell) {
          var kind = cell.classList.contains("build") ? "build" : "run";
          if (wanted[kind][cell.dataset.g] || wanted[kind][cell.dataset.t]) cell.classList.add("hit");
        });
      });
    });
  });
})();
