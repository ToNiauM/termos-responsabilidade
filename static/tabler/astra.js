/* astra.js — comportamentos do tema Tabler (o resto vem do Bootstrap embutido em tabler.min.js).
 * Substitui o que o dsgov.js fazia: busca em tabela, seleção de linhas, mensagens de sucesso que somem,
 * atalho da pesquisa e rolagem preservada em formulários GET. */
(function () {
  "use strict";

  /* ---------- Busca local em tabela: <input data-filtro-tabela="id-da-tabela"> ---------- */
  function normalizar(s) { return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }
  document.querySelectorAll('[data-filtro-tabela]').forEach(function (campo) {
    var tabela = document.getElementById(campo.dataset.filtroTabela);
    if (!tabela || !tabela.tBodies[0]) return;
    var contagem = document.querySelector('[data-contagem-tabela="' + tabela.id + '"]');
    var linhas = Array.from(tabela.tBodies[0].rows).filter(function (tr) { return tr.cells.length > 1; });
    campo.addEventListener('input', function () {
      var termo = normalizar(campo.value.trim());
      var visiveis = 0;
      linhas.forEach(function (tr) {
        var mostra = !termo || normalizar(tr.textContent).indexOf(termo) !== -1;
        tr.hidden = !mostra;
        if (mostra) visiveis++;
      });
      if (contagem) contagem.textContent = visiveis + ' de ' + linhas.length + ' registro(s)';
    });
    campo.addEventListener('keydown', function (e) { if (e.key === 'Enter') e.preventDefault(); });
  });

  /* ---------- Seleção de linhas: data-parent="grupo" marca/desmarca os data-child="grupo" ---------- */
  function atualizarSelecao(grupo) {
    var filhos = Array.from(document.querySelectorAll('[data-child="' + grupo + '"]'));
    var marcados = filhos.filter(function (c) { return c.checked; }).length;
    document.querySelectorAll('[data-parent="' + grupo + '"]').forEach(function (pai) {
      pai.checked = filhos.length > 0 && marcados === filhos.length;
      pai.indeterminate = marcados > 0 && marcados < filhos.length;
    });
    document.querySelectorAll('[data-selecao-info="' + grupo + '"]').forEach(function (info) {
      var n = info.querySelector('.count'), t = info.querySelector('.text');
      if (n) n.textContent = marcados;
      if (t) t.textContent = marcados === 1 ? 'item selecionado' : 'itens selecionados';
    });
    filhos.forEach(function (c) { var tr = c.closest('tr'); if (tr) tr.classList.toggle('table-active', c.checked); });
  }
  document.addEventListener('change', function (ev) {
    var el = ev.target;
    if (el.dataset && el.dataset.parent) {
      document.querySelectorAll('[data-child="' + el.dataset.parent + '"]').forEach(function (c) {
        var tr = c.closest('tr');
        if (!c.disabled && !(tr && tr.hidden)) c.checked = el.checked;
      });
      atualizarSelecao(el.dataset.parent);
    } else if (el.dataset && el.dataset.child) {
      atualizarSelecao(el.dataset.child);
    }
  });
  new Set(Array.from(document.querySelectorAll('[data-parent]'), function (p) { return p.dataset.parent; })).forEach(atualizarSelecao);

  /* ---------- Mensagens de sucesso somem sozinhas depois de 8 s ---------- */
  document.querySelectorAll('#mensagens .alert-success').forEach(function (msg) {
    setTimeout(function () {
      if (window.bootstrap && window.bootstrap.Alert) window.bootstrap.Alert.getOrCreateInstance(msg).close();
      else msg.remove();
    }, 8000);
  });

  /* ---------- Atalho Alt+Shift+P: foco na pesquisa do cabeçalho ---------- */
  document.addEventListener('keydown', function (ev) {
    if (ev.repeat || ev.isComposing) return;
    if (!(ev.altKey && ev.shiftKey && ev.code === 'KeyP')) return;
    if (document.activeElement && document.activeElement.closest('.modal.show')) return;
    var campo = document.getElementById('busca-header');
    if (!campo) return;
    ev.preventDefault();
    campo.focus();
    campo.select();
  });

  /* ---------- Rolagem preservada ao reenviar um <form method=get> da mesma tela ---------- */
  function chaveRolagem(pathname) { return 'astra-scroll:' + pathname; }
  document.addEventListener('submit', function (ev) {
    var form = ev.target;
    if (!form || form.tagName !== 'FORM' || (form.method || 'get').toLowerCase() !== 'get') return;
    try { window.sessionStorage.setItem(chaveRolagem(window.location.pathname), String(window.scrollY)); } catch (e) { /* sem sessionStorage */ }
  });
  (function restaurar() {
    var valor;
    try {
      var chave = chaveRolagem(window.location.pathname);
      valor = window.sessionStorage.getItem(chave);
      window.sessionStorage.removeItem(chave);
    } catch (e) { return; }
    if (valor === null) return;
    try { if (!document.referrer || new URL(document.referrer).pathname !== window.location.pathname) return; } catch (e) { return; }
    window.scrollTo(0, parseInt(valor, 10) || 0);
  })();

  /* ---------- Aparência: claro, escuro ou automático (segue o sistema); a escolha fica no navegador ---------- */
  var sistemaEscuro = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;

  function temaAtual() { return document.documentElement.getAttribute('data-bs-theme') === 'dark' ? 'dark' : 'light'; }

  /* Gráficos (graficos.js, tema "pca"): quando claro/escuro muda, avisa para redesenharem com os tons novos. */
  function temaDosGraficos() {
    document.dispatchEvent(new CustomEvent('pca:tema', { detail: { tema: temaAtual() } }));
  }

  /* Aparência por usuário (painel "Personalizar aparência"): claro/escuro, cor principal, tom dos cinzas e
     cantos. O _tema.html já aplicou a escolha antes de pintar; aqui cada troca muda a tela na hora e, com
     login, é gravada na conta (POST usuarios.aparencia). Cópia no navegador para a tela de login. */
  var raiz = document.documentElement;
  var PADRAO = { tema: 'auto', cor: 'blue', base: 'slate', cantos: '1' };
  function lerAparencia() {
    try { return Object.assign({}, PADRAO, JSON.parse(window.localStorage.getItem('termos-aparencia') || '{}')); }
    catch (e) { return Object.assign({}, PADRAO); }
  }
  var aparencia = lerAparencia();
  function aplicarAparencia(a) {
    var escuro = a.tema === 'escuro' || (a.tema !== 'claro' && !!(sistemaEscuro && sistemaEscuro.matches));
    var mudouTema = raiz.getAttribute('data-bs-theme') !== (escuro ? 'dark' : 'light');
    raiz.setAttribute('data-bs-theme', escuro ? 'dark' : 'light');
    raiz.setAttribute('data-bs-theme-primary', a.cor);
    raiz.setAttribute('data-bs-theme-base', a.base);
    raiz.setAttribute('data-bs-theme-radius', a.cantos);
    try { window.localStorage.setItem('termos-aparencia', JSON.stringify(a)); } catch (e) { /* sem localStorage */ }
    var form = document.querySelector('[data-aparencia]');
    if (form) ['tema', 'cor', 'base', 'cantos'].forEach(function (campo) {
      var marcado = form.querySelector('input[name="' + campo + '"][value="' + a[campo] + '"]');
      if (marcado) marcado.checked = true;
    });
    if (mudouTema) temaDosGraficos();
  }
  var gravacao = null;
  function gravarAparencia(a) {
    var form = document.querySelector('[data-aparencia]');
    if (!form || !form.hasAttribute('data-aparencia-grava')) return;
    var status = form.querySelector('[data-aparencia-status]');
    var csrf = document.querySelector('meta[name="csrf"]');
    clearTimeout(gravacao);
    gravacao = setTimeout(function () {
      var dados = new FormData();
      Object.keys(a).forEach(function (k) { dados.append(k, a[k]); });
      if (status) status.textContent = 'Salvando…';
      fetch(form.action, { method: 'POST', body: dados, credentials: 'same-origin',
        headers: { 'Accept': 'application/json', 'X-CSRF': csrf ? csrf.content : '' } })
        .then(function (r) { if (status) status.textContent = r.ok ? 'Salvo na sua conta.' : 'Não foi possível salvar.'; })
        .catch(function () { if (status) status.textContent = 'Não foi possível salvar.'; });
    }, 400);
  }
  document.addEventListener('change', function (ev) {
    var form = ev.target.closest && ev.target.closest('[data-aparencia]');
    if (!form || !ev.target.name || !(ev.target.name in PADRAO)) return;
    aparencia[ev.target.name] = ev.target.value;
    aplicarAparencia(aparencia);
    gravarAparencia(aparencia);
  });
  document.addEventListener('submit', function (ev) {
    if (!ev.target.matches || !ev.target.matches('[data-aparencia]')) return;
    ev.preventDefault();
    gravarAparencia(aparencia);
  });
  document.addEventListener('click', function (ev) {
    if (!ev.target.closest || !ev.target.closest('[data-aparencia-padrao]')) return;
    aparencia = Object.assign({}, PADRAO);
    aplicarAparencia(aparencia);
    gravarAparencia(aparencia);
  });
  if (sistemaEscuro && sistemaEscuro.addEventListener) sistemaEscuro.addEventListener('change', function () { if (aparencia.tema === 'auto') aplicarAparencia(aparencia); });
  /* Papel é branco: imprime sempre no claro e volta ao que estava depois */
  var temaAntesDeImprimir = null;
  window.addEventListener('beforeprint', function () {
    temaAntesDeImprimir = temaAtual();
    if (temaAntesDeImprimir === 'dark') { document.documentElement.setAttribute('data-bs-theme', 'light'); temaDosGraficos(); }
  });
  window.addEventListener('afterprint', function () {
    if (temaAntesDeImprimir === 'dark') { document.documentElement.setAttribute('data-bs-theme', 'dark'); temaDosGraficos(); }
    temaAntesDeImprimir = null;
  });
})();
