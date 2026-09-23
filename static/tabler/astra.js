/* astra.js — comportamentos do protótipo Tabler (o resto vem do Bootstrap embutido em tabler.min.js). */
(function () {
  // Busca dentro de uma tabela: <input data-filtro-tabela="id-da-tabela"> esconde as linhas que não contêm o texto
  function normalizar(s) { return s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }
  document.querySelectorAll('[data-filtro-tabela]').forEach(function (campo) {
    var tabela = document.getElementById(campo.dataset.filtroTabela);
    if (!tabela) return;
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
    // Enter no campo de busca não envia o formulário de filtro
    campo.addEventListener('keydown', function (e) { if (e.key === 'Enter') e.preventDefault(); });
  });
})();
