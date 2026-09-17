(function () {
  function sincronizar() {
    var menu = document.getElementById('main-navigation');
    if (!menu) return;
    menu.querySelectorAll('a.menu-item').forEach(function (a) {
      var atual = a.dataset.atual === 'true';
      a.classList.toggle('active', atual);
      if (atual) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
    menu.querySelectorAll('.menu-folder > a.menu-item').forEach(function (a) {
      a.setAttribute('aria-expanded', String(a.parentElement.classList.contains('active')));
    });
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', sincronizar);
  } else {
    sincronizar();
  }
})();
