/* tabelas.js — colunas ordenáveis e com largura ajustável, nas duas aparências (DSGov e Tabler).
 *
 * Opt-in: <table data-tabela>. As macros cabecalho_tabela/grafico já marcam; tabela avulsa marca à mão.
 * Nada aqui é por tela: o mesmo código serve todas as tabelas marcadas.
 *
 * Ordenar — clique (ou Enter/Espaço) no botão do cabeçalho alterna crescente/decrescente; aria-sort no <th>.
 *   Cliente (padrão): reordena as linhas do <tbody>; compara em pt-BR com números ("R$ 1.234,56",
 *   patrimônio), datas DD/MM/AAAA [HH:MM] e vazios ("", "—") sempre por último.
 *   Servidor: <table data-tabela-ordem="servidor"> + <th data-ordem="campo">: a lista foi cortada no servidor
 *   (só os N primeiros), então o clique recarrega com ?ordem=campo&dir=asc|desc (a ordem é feita no SQL).
 *   <th> que já tem link/controle (ordenação antiga por link, seleção) ou data-ordenar="nao" não ganha botão.
 * Largura — alça na borda direita de cada <th> (role=separator, focável): arrastar, setas (Shift = passo maior),
 *   duplo clique ou Enter devolve a largura original da coluna. Ao ajustar, a tabela passa a table-layout:fixed.
 *   Larguras ficam no localStorage por página + tabela + nº de colunas. Alças somem na impressão (CSS).
 * Aparência: o <script> diz data-aparencia="tabler" (botão .table-sort do Tabler) ou "dsgov" (ícone fa-sort).
 * Conteúdo trocado depois (fetch/HTMX): um MutationObserver inicializa as tabelas novas. */
(function () {
  "use strict";

  var script = document.currentScript;
  var TABLER = !!(script && script.dataset.aparencia === "tabler");
  var MINIMO = 48, PASSO = 16, PASSO_GRANDE = 64;
  var colator = new Intl.Collator("pt-BR", { numeric: true, sensitivity: "base" });

  /* ---------- armazenamento (pode falhar em janela privada, prévia, site bloqueado) ---------- */
  function ler(chave) { try { return JSON.parse(window.localStorage.getItem(chave) || "null"); } catch (e) { return null; } }
  function gravar(chave, valor) {
    try {
      if (valor == null) window.localStorage.removeItem(chave);
      else window.localStorage.setItem(chave, JSON.stringify(valor));
    } catch (e) { /* sem armazenamento: a largura vale só nesta visita */ }
  }

  /* ---------- valores das células ---------- */
  function texto(celula) {
    if (!celula) return "";
    if (celula.dataset.valor != null) return celula.dataset.valor.trim();
    return (celula.textContent || "").replace(/\s+/g, " ").trim();
  }
  function vazio(t) { return t === "" || t === "—" || t === "-" || t === "–"; }
  function numero(t) {
    var s = t.replace(/^R\$\s*/, "").replace(/[\s ]/g, "").replace(/%$/, "");
    if (/^[-−]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, "").replace(",", ".");
    else if (/^[-−]?\d+,\d+$/.test(s)) s = s.replace(",", ".");
    else if (!/^[-−]?\d+(\.\d+)?$/.test(s)) return null;
    return parseFloat(s.replace("−", "-"));
  }
  function data(t) {
    var m = /^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2}))?/.exec(t);
    if (m) return Date.UTC(+m[3], +m[2] - 1, +m[1], +(m[4] || 0), +(m[5] || 0));
    m = /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/.exec(t);
    if (m) return Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0));
    return null;
  }
  /* Tipo da coluna pela maioria dos valores não vazios: data, número ou texto. */
  function chaves(valores) {
    var cheios = valores.filter(function (v) { return !vazio(v); });
    var datas = cheios.map(data), nums = cheios.map(numero);
    var nd = datas.filter(function (x) { return x !== null; }).length;
    var nn = nums.filter(function (x) { return x !== null; }).length;
    var tipo = cheios.length && nd * 2 > cheios.length ? data : (cheios.length && nn * 2 > cheios.length ? numero : null);
    return valores.map(function (v) {
      if (vazio(v)) return { grupo: 2, v: "" };
      if (tipo) { var x = tipo(v); if (x !== null) return { grupo: 0, v: x }; return { grupo: 1, v: v }; }
      return { grupo: 0, v: v };
    });
  }
  function comparar(a, b) {
    if (typeof a === "number" && typeof b === "number") return a - b;
    return colator.compare(String(a), String(b));
  }

  /* ---------- cabeçalho ---------- */
  function cabecalhos(tabela) {
    var linhas = tabela.tHead ? tabela.tHead.rows : null;
    return linhas && linhas.length ? Array.prototype.slice.call(linhas[linhas.length - 1].cells) : [];
  }
  function rotulo(th) { return (th.textContent || "").replace(/\s+/g, " ").trim(); }
  function ordenavel(th) {
    if (th.dataset.ordenar === "nao" || th.colSpan > 1) return false;
    if (th.querySelector("a, button, input, select, label, .br-checkbox")) return false;
    if (/(^|\s)(w-1|column-checkbox|astra-acoes|dsgov-acoes)(\s|$)/.test(th.className)) return false;
    return rotulo(th) !== "" && !th.querySelector(".visually-hidden, .sr-only");
  }
  function marcar(th, sentido) {
    if (sentido) th.setAttribute("aria-sort", sentido === "asc" ? "ascending" : "descending");
    else th.removeAttribute("aria-sort");
    var botao = th.querySelector(".tabela-ordenar");
    if (!botao) return;
    if (TABLER) {
      botao.classList.toggle("asc", sentido === "asc");
      botao.classList.toggle("desc", sentido === "desc");
    } else {
      var icone = botao.querySelector(".tabela-ordenar-icone");
      icone.className = "fas " + (sentido === "asc" ? "fa-sort-up" : sentido === "desc" ? "fa-sort-down" : "fa-sort") + " tabela-ordenar-icone";
    }
  }
  function sentidoAtual(th) {
    var s = th.getAttribute("aria-sort");
    return s === "ascending" ? "asc" : s === "descending" ? "desc" : null;
  }

  function ordenarCliente(tabela, th, sentido) {
    var ths = cabecalhos(tabela), indice = ths.indexOf(th), colunas = ths.length;
    Array.prototype.forEach.call(tabela.tBodies, function (corpo) {
      var linhas = Array.prototype.slice.call(corpo.rows);
      var dados = linhas.filter(function (tr) { return tr.cells.length === colunas; });
      var resto = linhas.filter(function (tr) { return tr.cells.length !== colunas; }); /* "Nenhum registro" etc. */
      var ks = chaves(dados.map(function (tr) { return texto(tr.cells[indice]); }));
      var fator = sentido === "desc" ? -1 : 1;
      var ordem = dados.map(function (tr, i) { return { tr: tr, k: ks[i], i: i }; });
      ordem.sort(function (a, b) {
        if (a.k.grupo !== b.k.grupo) return a.k.grupo - b.k.grupo;   /* vazios sempre no fim */
        return (comparar(a.k.v, b.k.v) * fator) || (a.i - b.i);      /* estável */
      });
      ordem.forEach(function (o) { corpo.appendChild(o.tr); });
      resto.forEach(function (tr) { corpo.appendChild(tr); });
    });
    ths.forEach(function (outro) { marcar(outro, outro === th ? sentido : null); });
  }

  function ordenarServidor(th, sentido) {
    var url = new URL(window.location.href);
    url.searchParams.set("ordem", th.dataset.ordem);
    url.searchParams.set("dir", sentido);
    url.searchParams.delete("pagina");
    window.location.assign(url.toString());
  }

  function prepararOrdenacao(tabela) {
    var servidor = tabela.dataset.tabelaOrdem === "servidor";
    cabecalhos(tabela).forEach(function (th) {
      if (!ordenavel(th) || (servidor && !th.dataset.ordem)) return;
      var botao = document.createElement("button");
      botao.type = "button";
      botao.className = TABLER ? "table-sort tabela-ordenar" : "tabela-ordenar";
      botao.title = "Ordenar por " + rotulo(th);
      while (th.firstChild) botao.appendChild(th.firstChild);
      if (!TABLER) {
        var icone = document.createElement("i");
        icone.setAttribute("aria-hidden", "true");
        botao.appendChild(icone);
        icone.className = "fas fa-sort tabela-ordenar-icone";
      }
      th.appendChild(botao);
      marcar(th, sentidoAtual(th));
      botao.addEventListener("click", function () {
        var sentido = sentidoAtual(th) === "asc" ? "desc" : "asc";
        if (servidor) ordenarServidor(th, sentido);
        else ordenarCliente(tabela, th, sentido);
      });
    });
  }

  /* ---------- largura das colunas ---------- */
  function chaveLarguras(tabela, ths) {
    var todas = Array.prototype.slice.call(document.querySelectorAll("table[data-tabela]"));
    var nome = tabela.id || ("n" + todas.indexOf(tabela));
    return "tabela-larguras:" + window.location.pathname + ":" + nome + ":" + ths.length;
  }

  function prepararLarguras(tabela) {
    var ths = cabecalhos(tabela);
    if (!ths.length) return;
    var chave = chaveLarguras(tabela, ths);
    var naturais = null, larguras = null;

    function medirNaturais() {  /* larguras do layout automático, medidas antes de fixar (tabela visível) */
      if (!naturais && !larguras && tabela.offsetWidth > 0) naturais = ths.map(function (th) { return Math.round(th.getBoundingClientRect().width); });
    }
    function aplicar() {
      if (!larguras) {
        tabela.classList.remove("tabela-fixa");
        tabela.style.width = "";
        ths.forEach(function (th) { th.style.width = ""; });
        return;
      }
      var total = 0;
      ths.forEach(function (th, i) {
        th.style.width = larguras[i] + "px";
        if (getComputedStyle(th).display !== "none") total += larguras[i];
      });
      tabela.classList.add("tabela-fixa");
      tabela.style.width = total + "px";
      ths.forEach(function (th, i) {
        var alca = th.querySelector(".tabela-alca");
        if (alca) alca.setAttribute("aria-valuenow", String(larguras[i]));
      });
    }
    function fixar() {
      if (larguras) return;
      medirNaturais();
      larguras = (naturais || ths.map(function (th) { return Math.round(th.getBoundingClientRect().width); })).slice();
    }
    function mudar(i, largura) {
      fixar();
      larguras[i] = Math.max(MINIMO, Math.round(largura));
      aplicar();
    }
    function salvar() {
      var iguais = larguras && naturais && larguras.every(function (l, i) { return l === naturais[i]; });
      if (iguais) { larguras = null; aplicar(); gravar(chave, null); }
      else gravar(chave, larguras);
    }
    function restaurar(i) {
      if (!larguras) return;
      if (!naturais || !(naturais[i] > 0)) { larguras = null; aplicar(); gravar(chave, null); return; }  /* sem medida: tudo automático */
      larguras[i] = naturais[i];
      aplicar();
      salvar();
    }

    ths.forEach(function (th, i) {
      if (th.querySelector(".tabela-alca")) return;
      var alca = document.createElement("span");
      alca.className = "tabela-alca";
      alca.tabIndex = 0;
      alca.setAttribute("role", "separator");
      alca.setAttribute("aria-orientation", "vertical");
      alca.setAttribute("aria-valuemin", String(MINIMO));
      alca.setAttribute("aria-label", "Largura da coluna " + (rotulo(th) || (i + 1)) + " (setas ajustam, Enter restaura)");
      alca.title = "Arraste para ajustar a largura; duplo clique restaura";
      th.appendChild(alca);

      alca.addEventListener("click", function (ev) { ev.stopPropagation(); ev.preventDefault(); });
      alca.addEventListener("pointerdown", function (ev) {
        if (ev.button !== 0) return;
        ev.preventDefault();
        ev.stopPropagation();
        medirNaturais();
        var inicioX = ev.clientX, inicio = th.getBoundingClientRect().width;
        alca.setPointerCapture(ev.pointerId);
        tabela.classList.add("tabela-redimensionando");
        function mover(e) { mudar(i, inicio + (e.clientX - inicioX)); }
        function soltar() {
          alca.removeEventListener("pointermove", mover);
          alca.removeEventListener("pointerup", soltar);
          alca.removeEventListener("pointercancel", soltar);
          tabela.classList.remove("tabela-redimensionando");
          if (larguras) salvar();
        }
        alca.addEventListener("pointermove", mover);
        alca.addEventListener("pointerup", soltar);
        alca.addEventListener("pointercancel", soltar);
      });
      alca.addEventListener("dblclick", function (ev) { ev.preventDefault(); ev.stopPropagation(); restaurar(i); });
      alca.addEventListener("keydown", function (ev) {
        var passo = ev.shiftKey ? PASSO_GRANDE : PASSO;
        if (ev.key === "ArrowLeft" || ev.key === "ArrowRight") {
          ev.preventDefault();
          medirNaturais();
          var atual = larguras ? larguras[i] : th.getBoundingClientRect().width;
          mudar(i, atual + (ev.key === "ArrowRight" ? passo : -passo));
          salvar();
        } else if (ev.key === "Enter" || ev.key === "Home") {
          ev.preventDefault();
          restaurar(i);
        }
      });
    });

    var salvas = ler(chave);
    if (Array.isArray(salvas) && salvas.length === ths.length && salvas.every(function (n) { return typeof n === "number" && n > 0; })) {
      medirNaturais();
      larguras = salvas.map(function (n) { return Math.max(MINIMO, n); });
      aplicar();
    }
    /* Colunas que aparecem/somem por breakpoint (d-none d-md-table-cell) mudam a soma. */
    window.addEventListener("resize", function () { if (larguras) aplicar(); });
  }

  /* ---------- inicialização ---------- */
  function iniciar(tabela) {
    if (tabela.dataset.tabelaPronta) return;
    tabela.dataset.tabelaPronta = "1";
    prepararOrdenacao(tabela);
    prepararLarguras(tabela);
  }
  function varrer(raiz) {
    if (raiz.matches && raiz.matches("table[data-tabela]")) iniciar(raiz);
    if (raiz.querySelectorAll) Array.prototype.forEach.call(raiz.querySelectorAll("table[data-tabela]"), iniciar);
  }
  function comecar() {
    varrer(document);
    if (window.MutationObserver) {
      new MutationObserver(function (mudancas) {
        mudancas.forEach(function (m) { Array.prototype.forEach.call(m.addedNodes, function (n) { if (n.nodeType === 1) varrer(n); }); });
      }).observe(document.body, { childList: true, subtree: true });
    }
    document.addEventListener("htmx:afterSettle", function (ev) { varrer(ev.target || document); });
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", comecar);
  else comecar();
})();
