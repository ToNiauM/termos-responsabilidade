/* graficos.js — tema ECharts "pca" da casca Tabler, o MESMO arquivo do PCA (core/static/tabler/js/echarts-pca.js):
 * gráficos iguais nos dois sistemas (paleta do Tabler, rosca com contorno fino, legendas e rótulos protegidos
 * contra sobreposição, números abreviados nos eixos, modo escuro). Ao mudar este arquivo, mude o do PCA também.
 * Uso: todo [data-grafico="id"] com um <script type="application/json" id="id"> é montado sozinho. A troca de
 * aparência (astra.js) dispara o evento "pca:tema" e os gráficos se redesenham com os tons novos.
 */
(function () {
  "use strict";
  if (!window.echarts) return;

  /* Tons de interface por aparência (claro/escuro); T aponta para a atual e o tema é registrado de novo a cada troca */
  var TONS = {
    light: { primaria: "#066fd1", texto: "#1d273b", textoSecundario: "#667382", eixo: "#dce1e7",
             grade: "#eef1f4", fundo: "#ffffff", borda: "#e6e7e9", dica: "rgba(29,39,59,.94)", dicaBorda: "transparent", sombra: "rgba(6,111,209,.06)" },
    dark:  { primaria: "#4299e1", texto: "#e5e7eb", textoSecundario: "#9ca3af", eixo: "#374151",
             grade: "#263041", fundo: "#1f2937", borda: "#374151", dica: "rgba(31,41,55,.97)", dicaBorda: "#374151", sombra: "rgba(66,153,225,.10)" },
  };
  function temaAtual() { return document.documentElement.getAttribute("data-bs-theme") === "dark" ? "dark" : "light"; }
  var T = TONS[temaAtual()];
  var CATEGORICA = ["#066fd1", "#2fb344", "#f76707", "#ae3ec9", "#17a2b8", "#f59f00", "#d6336c", "#4263eb", "#74b816", "#0ca678"];
  var SEQUENCIAL = ["#dbeafe", "#93c5fd", "#4299e1", "#066fd1", "#0a3d7a"];
  var STATUS = {
    sucesso: "#2fb344", concluido: "#2fb344", ativo: "#2fb344",
    alerta: "#f59f00", pendente: "#f59f00",
    erro: "#d63939", atrasado: "#d63939", cancelado: "#9aa0ac", inativo: "#9aa0ac",
    info: "#066fd1", andamento: "#066fd1", neutro: "#9aa0ac",
  };
  /* Cores que as views ainda mandam fixas → equivalente no Tabler */
  var TROCA = {
    "#1351b4": "#066fd1", "#155bcb": "#066fd1", "#168821": "#2fb344", "#ffcd07": "#f59f00", "#e52207": "#d63939",
    "#757575": "#9aa0ac", "#cf4900": "#f76707", "#93348c": "#ae3ec9", "#00687d": "#17a2b8", "#c2850c": "#f59f00",
    "#d72d79": "#d6336c", "#5942d2": "#4263eb", "#d4e5ff": "#dbeafe", "#81aefc": "#93c5fd", "#2670e8": "#4299e1",
    "#0c326f": "#0a3d7a", "#333333": "#1d273b", "#555555": "#667382", "#888888": "#9aa0ac", "#e6e6e6": "#eef1f4",
    "#cccccc": "#e6e7e9",
  };
  /* No escuro, os cinzas de texto/grade que as views mandam precisam clarear (senão somem no fundo) */
  var TROCA_ESCURO = {
    "#333333": "#e5e7eb", "#555555": "#9ca3af", "#888888": "#9ca3af", "#757575": "#6b7280",
    "#e6e6e6": "#263041", "#cccccc": "#374151", "#f8f8f8": "#1f2937", "#d4e5ff": "#1e3a5f",
  };
  var FONTE = '"Inter Var", -apple-system, "Segoe UI", Roboto, sans-serif';

  /* Cor de fundo real do card (o contorno das fatias da rosca "some" nela). A variável CSS vem como texto
     ("var(--tblr-…)"), que o ECharts não entende e trocava por preto — por isso se mede num elemento. */
  function corDoCard() {
    var el = document.createElement("div");
    el.className = "card";
    el.style.cssText = "position:absolute;visibility:hidden;width:0;height:0";
    document.body.appendChild(el);
    var cor = getComputedStyle(el).backgroundColor;
    el.remove();
    return cor && cor !== "rgba(0, 0, 0, 0)" ? cor : null;
  }

  function registrarTema() {
  T = Object.assign({}, TONS[temaAtual()]);
  /* A cor principal escolhida pelo usuário NÃO entra nos gráficos: as cores das séries têm significado
     (contratado, renovação, situação…) e precisam ser as mesmas para todos e iguais às legendas da tela. */
  /* a borda das fatias da pizza acompanha o fundo real do card */
  var superficie = corDoCard();
  echarts.registerTheme("pca", {
    color: CATEGORICA,
    backgroundColor: "transparent",
    textStyle: { fontFamily: FONTE, color: T.texto, fontSize: 12 },
    title: { show: false },
    legend: {
      bottom: 0, type: "scroll", icon: "circle", itemWidth: 8, itemHeight: 8, itemGap: 16,
      textStyle: { color: T.textoSecundario, fontSize: 12 },
      pageIconColor: T.primaria, pageIconInactiveColor: T.borda, pageTextStyle: { color: T.textoSecundario, fontSize: 11 },
    },
    grid: { left: 4, right: 16, top: 20, bottom: 36, containLabel: true },
    tooltip: {
      backgroundColor: T.dica, borderColor: T.dicaBorda, borderWidth: T.dicaBorda === "transparent" ? 0 : 1, padding: [8, 12],
      textStyle: { color: "#fff", fontSize: 12, fontFamily: FONTE },
      extraCssText: "box-shadow:0 8px 24px rgba(15,23,42,.18);border-radius:8px;",
      axisPointer: { type: "shadow", shadowStyle: { color: T.sombra }, lineStyle: { color: T.eixo } },
    },
    categoryAxis: {
      axisLine: { lineStyle: { color: T.eixo } }, axisTick: { show: false },
      axisLabel: { color: T.textoSecundario, fontSize: 11, margin: 10 }, splitLine: { show: false },
    },
    valueAxis: {
      axisLine: { show: false }, axisTick: { show: false }, splitNumber: 4,
      axisLabel: { color: T.textoSecundario, fontSize: 11, hideOverlap: true },
      splitLine: { lineStyle: { color: T.grade, type: [4, 4] } },
    },
    bar: { itemStyle: { borderRadius: [4, 4, 0, 0] }, barMaxWidth: 28, emphasis: { itemStyle: { shadowBlur: 0, opacity: .85 } } },
    line: { symbol: "circle", symbolSize: 6, showSymbol: false, lineStyle: { width: 2.5 }, smooth: .25, emphasis: { focus: "series" } },
    pie: {
      itemStyle: { borderColor: superficie || T.fundo, borderWidth: 1, borderRadius: 3 },
      label: { color: T.textoSecundario, fontSize: 11 }, labelLine: { lineStyle: { color: T.eixo } },
    },
    gauge: { axisLine: { lineStyle: { color: [[1, T.grade]] } } },
  });
  }
  registrarTema();

  var fmtInteiro = new Intl.NumberFormat("pt-BR");
  var fmtMoeda = new Intl.NumberFormat("pt-BR", { style: "currency", currency: "BRL" });
  var fmtPct = new Intl.NumberFormat("pt-BR", { style: "percent", maximumFractionDigits: 1 });
  function fmtValor(v) { return typeof v === "number" ? fmtInteiro.format(Math.round(v * 100) / 100) : v; }
  /* Eixo: números grandes abreviados (1,2 bi · 150 mi · 12 mil), como nos painéis corporativos */
  var fmtCurto = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 1 });
  function fmtEixo(v) {
    if (typeof v !== "number") return v;
    var a = Math.abs(v);
    if (a >= 1e9) return fmtCurto.format(v / 1e9) + " bi";
    if (a >= 1e6) return fmtCurto.format(v / 1e6) + " mi";
    if (a >= 1e4) return fmtCurto.format(v / 1e3) + " mil";
    return fmtInteiro.format(v);
  }

  var FORMATOS_DICA = {
    moeda: function (v) { return typeof v === "number" ? fmtMoeda.format(v) : v; },
    percentual: function (v) { return typeof v === "number" ? fmtCurto.format(v) + "%" : v; },
  };

  /* Troca, em qualquer ponto da opção, as cores fixas vindas das views pelas do Tabler */
  function trocarCores(no) {
    if (Array.isArray(no)) { for (var i = 0; i < no.length; i++) no[i] = trocarCores(no[i]); return no; }
    if (no && typeof no === "object") { for (var k in no) if (Object.prototype.hasOwnProperty.call(no, k)) no[k] = trocarCores(no[k]); return no; }
    if (typeof no === "string") { var k2 = no.toLowerCase(); var c = (temaAtual() === "dark" && TROCA_ESCURO[k2]) || TROCA[k2]; return c || no; }
    return no;
  }

  function degrade(cor) {
    return new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: cor + "33" }, { offset: 1, color: cor + "00" }]);
  }

  /* Modelos de texto do ECharts ("{b}: {c} ({d}%)") mostram o número cru (138183900.84). Viram função com o
     número formatado: curto nos rótulos do gráfico ("138,2 mi"), exato no tooltip ("138.183.900,84"). */
  function modelo(f, curto) {
    if (typeof f !== "string" || !/\{[bcd]\}/.test(f)) return f;
    return function (p) {
      var v = p.value && typeof p.value === "object" ? p.value.value : p.value;
      return f.replace(/\{b\}/g, p.name).replace(/R\$ \{c\}/g, curto ? "R$ " + fmtEixo(v) : fmtMoeda.format(v))
              .replace(/\{c\}/g, (curto ? fmtEixo : fmtValor)(v))
              .replace(/\{d\}/g, typeof p.percent === "number" ? fmtCurto.format(p.percent) : "");
    };
  }

  function aplicarPadroes(op) {
    trocarCores(op);
    [].concat(op.xAxis || [], op.yAxis || []).forEach(function (eixo) {
      if (eixo && eixo.type === "value") { eixo.axisLabel = eixo.axisLabel || {}; if (!eixo.axisLabel.formatter) eixo.axisLabel.formatter = fmtEixo; }
    });
    var paleta = op.color || CATEGORICA;
    (op.series || []).forEach(function (s, i) {
      /* Rótulo no gráfico: número curto (187 mi, 12,4 mil); o valor exato fica no tooltip e em "Ver dados" */
      if (s.label && s.label.show && !s.label.formatter) s.label.formatter = function (p) { return fmtEixo(p.value); };
      if (s.label) s.label.formatter = modelo(s.label.formatter, true);
      /* Rótulo fora da barra/ponto fica sobre o fundo do card: cor de texto da aparência e sem contorno */
      if (s.label && s.label.show && s.type !== "pie" && /^(top|bottom|left|right|outside)$/.test(s.label.position || "")) {
        if (!s.label.color || s.label.color === "inherit") s.label.color = T.texto;
        s.label.textBorderWidth = 0;
      }
      /* Linhas: área suave em degradê quando há uma ou duas séries (fica limpo; com muitas séries, sem área) */
      if (s.type === "line" && !s.areaStyle && (op.series || []).length <= 2 && !s.stack) {
        var cor = (s.itemStyle && s.itemStyle.color) || (s.lineStyle && s.lineStyle.color) || paleta[i % paleta.length];
        if (typeof cor === "string" && /^#[0-9a-f]{6}$/i.test(cor)) s.areaStyle = { color: degrade(cor) };
      }
      /* Rosca: anel fino, total no centro fica legível */
      if (s.type === "pie" && Array.isArray(s.radius) && s.radius.length === 2 && !s.padAngle) s.padAngle = 1.5;
      /* Barras empilhadas: só a do topo arredonda */
      if (s.type === "bar" && s.stack) { s.itemStyle = s.itemStyle || {}; if (s.itemStyle.borderRadius === undefined) s.itemStyle.borderRadius = 0; }
    });
    op.tooltip = op.tooltip || {};
    /* A view não manda função em JSON: "moeda" e "percentual" nomeiam o formato exato do tooltip */
    if (typeof op.tooltip.valueFormatter === "string") op.tooltip.valueFormatter = FORMATOS_DICA[op.tooltip.valueFormatter];
    if (!op.tooltip.valueFormatter) op.tooltip.valueFormatter = fmtValor;
    if (op.tooltip.trigger === "item") op.tooltip.formatter = modelo(op.tooltip.formatter, false);
    op.tooltip.confine = true;   /* tooltip nunca sai do card */
    return op;
  }

  /* Proteções de layout, calculadas pela largura real do card: legenda nunca sobrepõe o desenho,
     rótulos longos são cortados com reticências, rosca só mostra rótulos quando cabem. */
  function encaixar(op, largura) {
    var estreito = largura < 480;
    var series = op.series || [];
    var temLegenda = op.legend && op.legend.show !== false && (Array.isArray(op.legend) ? op.legend.length : true);
    if (op.legend && !Array.isArray(op.legend)) {
      op.legend.type = "scroll";
      if (op.legend.top === undefined && op.legend.bottom === undefined) op.legend.bottom = 0;
      op.legend.itemGap = estreito ? 10 : 16;
      /* lineHeight: a legenda paginada recorta o topo da linha, e o acento das maiúsculas (MÓVEIS) sumia */
      op.legend.textStyle = Object.assign({ overflow: "truncate", width: estreito ? 90 : 160, lineHeight: 16 }, op.legend.textStyle || {});
    }
    var grades = [].concat(op.grid || [{}]);
    grades.forEach(function (g) {
      g.containLabel = true;
      if (temLegenda && op.legend && op.legend.bottom !== undefined) g.bottom = Math.max(parseInt(g.bottom, 10) || 0, estreito ? 44 : 40);
      if (g.right === undefined || parseInt(g.right, 10) > 40) g.right = estreito ? 8 : 16;
      if (g.left === undefined) g.left = 4;
    });
    op.grid = Array.isArray(op.grid) ? grades : grades[0];
    function eixoCategoria(eixo, horizontal) {
      if (!eixo || eixo.type !== "category") return;
      eixo.axisLabel = eixo.axisLabel || {};
      eixo.axisLabel.hideOverlap = true;
      if (horizontal) {   /* categorias no eixo Y (barras horizontais): corta nomes longos */
        eixo.axisLabel.overflow = "truncate";
        eixo.axisLabel.width = eixo.axisLabel.width || (estreito ? 90 : 170);
      } else if (estreito && eixo.axisLabel.rotate === undefined) {
        eixo.axisLabel.overflow = "truncate";
        eixo.axisLabel.width = 70;
      }
      if (estreito) eixo.axisLabel.fontSize = 10;
    }
    [].concat(op.xAxis || []).forEach(function (e) { eixoCategoria(e, false); });
    [].concat(op.yAxis || []).forEach(function (e) { eixoCategoria(e, true); });
    /* Legenda no celular: em vez de paginar (cortava nomes na seta), quebra em linhas e reserva a altura delas */
    var linhasLegenda = 1;
    if (estreito && op.legend && !Array.isArray(op.legend) && temLegenda) {
      var nomes = op.legend.data || [];
      if (!nomes.length) series.forEach(function (s) {
        if (s.type === "pie") (s.data || []).forEach(function (d) { if (d && d.name != null) nomes.push(d.name); });
        else if (s.name != null) nomes.push(s.name);
      });
      var larguraUtil = Math.max(largura - 16, 200), linha = 0;
      nomes.forEach(function (n) {
        var w = Math.min(String(typeof n === "object" && n ? n.name : n).length * 6.2, 120) + 22;
        if (linha + w > larguraUtil && linha > 0) { linhasLegenda++; linha = 0; }
        linha += w;
      });
      if (linhasLegenda <= 4) {
        op.legend.type = "plain";
        op.legend.left = "center"; op.legend.width = larguraUtil;
        op.legend.textStyle = Object.assign({}, op.legend.textStyle, { width: 120 });
        grades.forEach(function (g) { g.bottom = Math.max(parseInt(g.bottom, 10) || 0, 22 * linhasLegenda + 18); });
      } else linhasLegenda = 1;
    }
    /* Meses "mmm/aaaa" no eixo: quebra em duas linhas para não encostarem */
    [].concat(op.xAxis || []).forEach(function (e) {
      if (!e || e.type !== "category" || !Array.isArray(e.data) || !e.data.length) return;
      if (e.data.every(function (v) { return typeof v === "string" && /^[a-zç]{3}\/\d{2,4}$/i.test(v); })) {
        e.axisLabel = e.axisLabel || {};
        if (!e.axisLabel.formatter) e.axisLabel.formatter = function (v) { return String(v).replace("/", "\n"); };
        e.axisLabel.lineHeight = 13;
      }
    });
    /* Celular com muitas barras rotuladas: os números se atropelam — ficam só no tooltip */
    if (estreito) {
      var categorias = 0;
      [].concat(op.xAxis || [], op.yAxis || []).forEach(function (e) { if (e && e.type === "category" && Array.isArray(e.data)) categorias = Math.max(categorias, e.data.length); });
      var barras = series.filter(function (s) { return s.type === "bar"; }).length;
      if (categorias * Math.max(barras, 1) > 8) series.forEach(function (s) {
        if (s.type === "bar" && s.label && s.label.show && !/inside/.test(s.label.position || "inside")) s.label.show = false;
      });
    }
    series.forEach(function (s) {
      if (s.type !== "pie") return;
      var n = (s.data || []).length;
      if (estreito || n > 6 || temLegenda) {   /* com legenda (ou pouco espaço/muitas fatias): nome e valor ficam na legenda e no tooltip */
        s.label = Object.assign({}, s.label || {}, { show: !!(s.label && s.label.position === "center") });
        s.labelLine = { show: false };
        if (s.center === undefined) s.center = ["50%", estreito && linhasLegenda > 1 ? (46 - 4 * linhasLegenda) + "%" : "45%"];
        /* anel cabe acima da legenda: raio externo no máximo 70% (interno encolhe na mesma proporção) */
        var r = s.radius, teto = estreito && linhasLegenda > 1 ? 62 : 70;
        var pct = function (v) { return typeof v === "string" && /%$/.test(v) ? parseFloat(v) : null; };
        if (Array.isArray(r) && pct(r[1]) && pct(r[1]) > teto) {
          var k = teto / pct(r[1]);
          s.radius = [pct(r[0]) != null ? (pct(r[0]) * k).toFixed(1) + "%" : r[0], teto + "%"];
        } else if (r === undefined || (pct(r) && pct(r) > teto)) {
          s.radius = r === undefined ? ["0%", teto + "%"] : teto + "%";
        }
      } else {
        s.label = Object.assign({ overflow: "truncate", width: 110 }, s.label || {});
        s.avoidLabelOverlap = true;
      }
      /* texto do furo da rosca: encolhe se for comprido (ex.: valores em R$) */
      if (s.label && s.label.show && s.label.position === "center" && !s.label.fontSize && !s.label.rich) s.label.fontSize = estreito ? 14 : 16;
    });
    return op;
  }

  /* Rosca com total no furo (graphic de texto "número\nlegenda"): se o número é a soma das fatias, ele passa a
     acompanhar a legenda — desligar uma fatia tira o valor dela do total. Descobre o formato em que a view
     escreveu o número (cru, inteiro pt-BR ou moeda) para reescrever igual; se não bate com a soma, não mexe. */
  function totalDaRosca(op) {
    var pizza = (op.series || []).filter(function (s) { return s.type === "pie"; })[0];
    var texto = [].concat(op.graphic || [])[0];
    if (!pizza || !texto || texto.type !== "text" || !texto.style || typeof texto.style.text !== "string") return null;
    var valor = function (d) { var v = d && typeof d === "object" ? d.value : d; return typeof v === "number" ? v : 0; };
    var soma = function (selecionados) {
      return (pizza.data || []).reduce(function (t, d) {
        return selecionados && d && selecionados[d.name] === false ? t : t + valor(d);
      }, 0);
    };
    var linhas = texto.style.text.split("\n");
    var limpo = function (s) { return String(s).replace(/\s/g, " "); };
    var fmtDecimal = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    var curta = function (v) {   /* moeda_curta do PCA: "165,4 mi", "12,4 mil", abaixo de mil "999,00" */
      var a = Math.abs(v), s = v < 0 ? "-" : "";
      if (a >= 1e6) return s + (a / 1e6).toFixed(1).replace(".", ",") + " mi";
      if (a >= 1e3) return s + (a / 1e3).toFixed(1).replace(".", ",") + " mil";
      return fmtDecimal.format(v);
    };
    var formatos = [String, fmtInteiro.format.bind(fmtInteiro), fmtMoeda.format.bind(fmtMoeda),
                    fmtDecimal.format.bind(fmtDecimal), curta];
    var formato = formatos.filter(function (f) { return limpo(f(soma())) === limpo(linhas[0]); })[0];
    if (!formato) return null;
    texto.id = texto.id || "pca-total-rosca";
    return function (selecionados) {
      return { id: texto.id, style: Object.assign({}, texto.style, { text: [formato(soma(selecionados))].concat(linhas.slice(1)).join("\n") }) };
    };
  }

  /* Texto do furo da rosca: a view o põe no meio do CARD, mas o anel sobe para dar lugar à legenda (center 45%
     ou menos). Posiciona o texto em pixels no centro real do anel; refeito a cada mudança de tamanho. */
  function centralizarTotal(op, largura, altura) {
    var pizza = (op.series || []).filter(function (s) { return s.type === "pie"; })[0];
    var texto = [].concat(op.graphic || [])[0];
    if (!pizza || !texto || texto.type !== "text" || !texto.style) return null;
    var c = Array.isArray(pizza.center) ? pizza.center : ["50%", "50%"];
    var px = function (v, total) {
      if (typeof v === "number") return v;
      return typeof v === "string" && /%$/.test(v) ? parseFloat(v) / 100 * total : total / 2;
    };
    texto.id = texto.id || "pca-total-rosca";
    ["left", "top", "right", "bottom"].forEach(function (k) { delete texto[k]; });
    texto.x = px(c[0], largura); texto.y = px(c[1], altura);
    texto.style.textAlign = "center"; texto.style.textVerticalAlign = "middle";
    /* furo pequeno (card estreito): o total encolhe para caber dentro do anel */
    if (largura < 480) { texto.style.fontSize = 16; texto.style.lineHeight = 20; }
    return { id: texto.id, x: texto.x, y: texto.y };
  }

  function montar(raiz) {
    (raiz || document).querySelectorAll("[data-grafico]").forEach(function (el) {
      if (el.dataset.graficoMontado) return;
      var script = document.getElementById(el.dataset.grafico);
      if (!script) return;
      var bruto = script.textContent;
      var opcoes = encaixar(aplicarPadroes(JSON.parse(bruto)), el.clientWidth || 600);
      centralizarTotal(opcoes, el.clientWidth || 600, el.clientHeight || 300);
      el.__pcaOpcoes = opcoes;
      var totalRosca = totalDaRosca(opcoes);
      el.__pcaBruto = bruto;
      var inst = echarts.init(el, "pca", { renderer: "canvas" });
      inst.setOption(opcoes);
      inst.setOption(inst.getOption());   /* segunda passada: corrige o desenho da legenda paginada na primeira montagem */
      el.dataset.graficoMontado = "1";
      /* Ao mudar de largura (girar o celular, abrir o menu), refaz as proteções de layout para a nova largura */
      if (!el.__pcaObservado) {
        el.__pcaObservado = true;
        var ultimaFaixa = (el.clientWidth || 600) < 480;
        var ajustar = function () {
          var atual = echarts.getInstanceByDom(el);
          if (!atual) return;
          var faixa = (el.clientWidth || 600) < 480;
          if (faixa !== ultimaFaixa) {
            ultimaFaixa = faixa;
            var novas = encaixar(aplicarPadroes(JSON.parse(el.__pcaBruto)), el.clientWidth);
            centralizarTotal(novas, el.clientWidth, el.clientHeight);
            el.__pcaOpcoes = novas;
            totalDaRosca(novas);   /* mesmo id no texto do furo; a legenda volta toda ligada, e o total também */
            atual.setOption(novas, true);
          }
          atual.resize();
          var pos = el.__pcaOpcoes && centralizarTotal(el.__pcaOpcoes, el.clientWidth, el.clientHeight);
          if (pos) atual.setOption({ graphic: { elements: [pos] } });
        };
        if (window.ResizeObserver) new ResizeObserver(ajustar).observe(el);
        else window.addEventListener("resize", ajustar);
      }
      if (totalRosca) inst.on("legendselectchanged", function (ev) {
        inst.setOption({ graphic: { elements: [totalRosca(ev.selected)] } });
      });
      /* Drill-down declarativo: item de dado com `url` navega ao clique */
      inst.on("click", function (p) {
        var url = p && p.data && typeof p.data === "object" ? p.data.url : null;
        if (typeof url === "string" && url) window.location.assign(url);
      });
      if ((opcoes.series || []).some(function (s) { return s.data && s.data.some && s.data.some(function (d) { return d && d.url; }); })) el.classList.add("cursor-pointer");
    });
  }
  /* Seletor dentro do card: <select data-grafico-seletor="base"> (um ou mais) escolhe qual opção o gráfico
     [data-grafico-base="base"] mostra. O id do <script type="application/json"> é a base mais os valores dos
     seletores, na ordem da página, unidos por "--" (ex.: g-explorar--ccusto--valor). A tabela "Ver dados" de
     cada combinação vem marcada com [data-grafico-tabela="id"] e só a da combinação atual fica visível. */
  function trocarVariante(base) {
    var valores = [];
    document.querySelectorAll('[data-grafico-seletor="' + base + '"]').forEach(function (s) { valores.push(s.value); });
    var id = [base].concat(valores).join("--");
    var el = document.querySelector('[data-grafico-base="' + base + '"]');
    if (!el || !document.getElementById(id) || el.dataset.grafico === id) return;
    var atual = echarts.getInstanceByDom(el);
    if (atual) echarts.dispose(el);
    delete el.dataset.graficoMontado;
    el.classList.remove("cursor-pointer");
    el.dataset.grafico = id;
    var rotulo = document.getElementById(id).getAttribute("data-resumo");
    if (rotulo) el.setAttribute("aria-label", rotulo);
    document.querySelectorAll("[data-grafico-tabela]").forEach(function (t) {
      var alvo = t.getAttribute("data-grafico-tabela");
      if (alvo.indexOf(base + "--") === 0) t.hidden = alvo !== id;
    });
    montar(el.parentNode);
  }
  document.addEventListener("change", function (ev) {
    var base = ev.target && ev.target.getAttribute && ev.target.getAttribute("data-grafico-seletor");
    if (base) trocarVariante(base);
  });

  document.addEventListener("DOMContentLoaded", function () { montar(document); });
  document.addEventListener("htmx:afterSettle", function (ev) { montar(ev.detail && ev.detail.elt ? ev.detail.elt.parentNode || document : document); });
  /* Gráficos dentro de abas/colapsos recolhidos nascem com largura zero: redimensiona ao abrir */
  ["shown.bs.tab", "shown.bs.collapse", "shown.bs.modal"].forEach(function (nome) {
    document.addEventListener(nome, function (ev) {
      var alvo = ev.target && ev.target.getAttribute && (ev.target.getAttribute("data-bs-target") || ev.target.getAttribute("href"));
      var raiz = (alvo && alvo.charAt(0) === "#" && document.querySelector(alvo)) || ev.target;
      if (raiz && raiz.querySelectorAll) raiz.querySelectorAll("[data-grafico]").forEach(function (el) { var i = echarts.getInstanceByDom(el); if (i) i.resize(); });
    });
  });

  /* Troca de aparência: redesenha todos os gráficos com os tons novos (as opções originais continuam no json_script) */
  document.addEventListener("pca:tema", function () {
    registrarTema();
    document.querySelectorAll("[data-grafico][data-grafico-montado]").forEach(function (el) {
      echarts.dispose(el);
      delete el.dataset.graficoMontado;
    });
    montar(document);
  });

  window.pca = window.pca || {};
  window.pca.graficos = {
    categorica: CATEGORICA, sequencial: SEQUENCIAL, status: STATUS, get cores() { return T; },
    inteiro: fmtInteiro.format.bind(fmtInteiro), moeda: fmtMoeda.format.bind(fmtMoeda), percentual: fmtPct.format.bind(fmtPct),
    montar: montar, eixo: fmtEixo, trocar: trocarVariante,
  };
})();
