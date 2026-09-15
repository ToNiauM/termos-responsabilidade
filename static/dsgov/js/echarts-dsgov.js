/* Tema ECharts "dsgov" — única forma permitida de colorir gráficos nos sistemas da skill.
 * Todas as cores vêm dos tokens do DSGov 3.7.0 (references/tokens.md). Registre uma vez e use
 * echarts.init(el, "dsgov"). Séries nunca definem cor própria, exceto via dsgov.graficos.status.
 */
(function () {
  "use strict";
  if (!window.echarts) return;

  var C = {
    interativo: "#1351b4",   /* blue-warm-vivid-70 */
    texto: "#333333",        /* gray-80 */
    textoSecundario: "#555555", /* gray-70 */
    eixo: "#888888",         /* gray-40 */
    grade: "#e6e6e6",        /* gray-10 */
    fundo: "#ffffff",        /* pure-0 */
    borda: "#cccccc",        /* gray-20 */
  };

  /* Categórica: 8 cores de famílias distintas do DS, todas com contraste >= 3:1 sobre branco. */
  var CATEGORICA = [
    "#1351b4", /* blue-warm-vivid-70 */
    "#168821", /* green-cool-vivid-50 */
    "#cf4900", /* orange-warm-vivid-50 */
    "#93348c", /* violet-warm-vivid-60 */
    "#00687d", /* cyan-vivid-60 */
    "#c2850c", /* gold-vivid-40 */
    "#d72d79", /* magenta-vivid-50 */
    "#5942d2", /* indigo-warm-vivid-60 */
  ];

  /* Sequencial monocromática (Blue Warm Vivid), do claro ao escuro. */
  var SEQUENCIAL = ["#d4e5ff", "#81aefc", "#2670e8", "#1351b4", "#0c326f"];

  /* Status: mesmas cores das tags br-tag status do DS. */
  var STATUS = {
    sucesso: "#168821", concluido: "#168821", ativo: "#168821",
    alerta: "#ffcd07", pendente: "#ffcd07",
    erro: "#e52207", atrasado: "#e52207", cancelado: "#757575", inativo: "#757575",
    info: "#155bcb", andamento: "#155bcb",
    neutro: "#757575",
  };

  var tema = {
    color: CATEGORICA,
    backgroundColor: "transparent",
    textStyle: { fontFamily: "Rawline, Raleway, sans-serif", color: C.texto, fontSize: 14 },
    title: { show: false }, /* o título fica no card-header, fora do canvas */
    legend: { bottom: 0, type: "scroll", icon: "roundRect", itemWidth: 12, itemHeight: 12, textStyle: { color: C.texto, fontSize: 12 }, pageIconColor: C.interativo, pageIconInactiveColor: C.borda, pageTextStyle: { color: C.textoSecundario, fontSize: 12 } },
    grid: { left: 8, right: 56, top: 24, bottom: 40, containLabel: true },
    tooltip: {
      backgroundColor: C.fundo, borderColor: C.borda, borderWidth: 1,
      textStyle: { color: C.texto, fontSize: 12 },
      extraCssText: "box-shadow: 0 1px 6px rgba(0,0,0,.16); border-radius: 4px;",
    },
    categoryAxis: {
      axisLine: { lineStyle: { color: C.eixo } }, axisTick: { show: false },
      axisLabel: { color: C.textoSecundario, fontSize: 12 }, splitLine: { show: false },
    },
    valueAxis: {
      axisLine: { show: false }, axisTick: { show: false }, splitNumber: 4, minInterval: 1,
      axisLabel: { color: C.textoSecundario, fontSize: 12, hideOverlap: true },
      splitLine: { lineStyle: { color: C.grade } },
    },
    bar: { itemStyle: { borderRadius: [2, 2, 0, 0] }, barMaxWidth: 40 },
    line: { symbol: "circle", symbolSize: 6, lineStyle: { width: 2 }, smooth: false },
    pie: { itemStyle: { borderColor: C.fundo, borderWidth: 2 }, label: { color: C.texto } },
  };

  echarts.registerTheme("dsgov", tema);

  /* Formatadores pt-BR compartilhados. */
  var fmtInteiro = new Intl.NumberFormat("pt-BR");
  var fmtMoeda = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  var fmtPct = new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 1 });

  /* Inicializa todo [data-grafico] cujo JSON de opções está num <script type="application/json"> irmão
     (padrão Django json_script). Redimensiona junto com a janela. */
  function montar(raiz) {
    (raiz || document).querySelectorAll("[data-grafico]").forEach(function (el) {
      if (el.dataset.dsgovGrafico) return;
      var script = document.getElementById(el.dataset.grafico);
      if (!script) return;
      var opcoes = aplicarPadroes(JSON.parse(script.textContent));
      var inst = echarts.init(el, "dsgov", { renderer: "canvas" });
      inst.setOption(opcoes);
      /* Segunda passada, com a opção JÁ RESOLVIDA (`inst.getOption()`), não com `opcoes` de novo:
         a primeira `setOption()` de uma instância NOVA desenha a legenda paginada
         (`legend.type: "scroll"`) errada quando há muitos itens (ex.: "UO por mês", 14 séries) —
         cada item na posição de OUTRO item, texto sobreposto sobre o eixo X e as barras — mesmo a
         opção já estando correta desde o início (`getOption()` já mostrava `type: "scroll"`; só o
         DESENHO da primeira passada saía errado). Chamar `setOption(opcoes)` de novo (o MESMO
         objeto raso que veio do `json_script`, sem os campos que o tema resolveu) não corrige em
         NENHUM atraso testado — o ECharts trata como "nada mudou" e pula o relayout da legenda.
         `setOption(inst.getOption())` força o relayout porque o objeto é visivelmente diferente
         (traz todos os campos já resolvidos pelo tema); corrige de imediato, na mesma tarefa, sem
         precisar de `setTimeout`/`requestAnimationFrame`/esperar fontes (medido empiricamente
         nesta investigação — não é fonte, não é tempo decorrido, é ESTE objeto específico). */
      inst.setOption(inst.getOption());
      el.dataset.dsgovGrafico = "1";
      window.addEventListener("resize", function () { inst.resize(); });
      /* Drill-down declarativo (references/graficos.md, regra 9): um item de dado que traga `url`
         ({"name": ..., "value": ..., "url": "/processos?situacao=atrasado"}) navega ao clique. É a
         única forma de clique permitida — nenhum handler em template, nenhum JS do projeto. */
      inst.on("click", function (p) {
        var url = p && p.data && typeof p.data === "object" ? p.data.url : null;
        if (typeof url === "string" && url) window.location.assign(url);
      });
    });
  }
  /* Formatação pt-BR nos eixos numéricos, rótulos e tooltips, salvo quando a view já definiu. */
  function fmtValor(v) { return typeof v === "number" ? fmtInteiro.format(Math.round(v * 100) / 100) : v; }
  function aplicarPadroes(op) {
    [].concat(op.xAxis || [], op.yAxis || []).forEach(function (eixo) {
      if (eixo && eixo.type === "value") { eixo.axisLabel = eixo.axisLabel || {}; if (!eixo.axisLabel.formatter) eixo.axisLabel.formatter = fmtValor; }
    });
    (op.series || []).forEach(function (s) {
      if (s.label && s.label.show && !s.label.formatter) s.label.formatter = function (p) { return fmtValor(p.value); };
    });
    op.tooltip = op.tooltip || {};
    if (!op.tooltip.valueFormatter) op.tooltip.valueFormatter = fmtValor;
    return op;
  }
  document.addEventListener("DOMContentLoaded", function () { montar(document); });
  document.addEventListener("htmx:afterSettle", function (ev) { montar(ev.detail && ev.detail.elt); });

  window.dsgov = window.dsgov || {};
  window.dsgov.graficos = {
    categorica: CATEGORICA, sequencial: SEQUENCIAL, status: STATUS, cores: C,
    inteiro: fmtInteiro.format.bind(fmtInteiro), moeda: fmtMoeda.format.bind(fmtMoeda), percentual: fmtPct.format.bind(fmtPct),
    montar: montar,
  };
})();
