// Apply a remembered light or dark choice before the page paints.
(function () {
  try {
    var theme = localStorage.getItem("ostram-theme");
    if (theme === "light" || theme === "dark") document.documentElement.dataset.theme = theme;
  } catch (error) { /* storage unavailable: follow the system setting */ }
})();
