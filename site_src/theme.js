(function () {
  window.getTheme = function () {
    var explicit = document.documentElement.getAttribute('data-theme');
    if (explicit) return explicit;
    return matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  };
  var stored = localStorage.getItem('theme');
  if (stored) document.documentElement.setAttribute('data-theme', stored);
})();
