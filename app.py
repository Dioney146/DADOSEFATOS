"""
Modo painel — os cinco gráficos numa grade que se arrasta e redimensiona
com o mouse, no estilo do canvas do Power BI.

Como funciona: as figuras Plotly são geradas normalmente em Python, convertidas
para JSON e desenhadas dentro de um único componente HTML. A grade é a
biblioteca Gridstack; o Plotly.js redesenha cada gráfico quando o painel muda
de tamanho. O layout que o usuário montar fica salvo no próprio navegador
(localStorage), separado por estado.

Requer acesso a dois CDNs (cdn.jsdelivr.net e cdn.plot.ly). Se a rede da
empresa bloquear, o componente avisa na tela e o modo apresentação continua
funcionando normalmente.
"""

from __future__ import annotations

import json
from string import Template

import streamlit as st

# Layout inicial: grade de 12 colunas, mesma disposição do slide
LAYOUT_PADRAO = [
    {"id": "veiculos", "x": 0, "y": 0, "w": 4, "h": 4},
    {"id": "ocupacao", "x": 4, "y": 0, "w": 4, "h": 4},
    {"id": "drop", "x": 8, "y": 0, "w": 4, "h": 4},
    {"id": "paradas", "x": 0, "y": 4, "w": 6, "h": 4},
    {"id": "peso", "x": 6, "y": 4, "w": 6, "h": 4},
]

_MODELO = Template("""
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/gridstack.min.css"/>
<script src="https://cdn.jsdelivr.net/npm/gridstack@10.3.1/dist/gridstack-all.js"></script>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>

<style>
  #raiz-painel { font-family: 'Archivo', Arial, sans-serif; }
  #raiz-painel .grid-stack { background: transparent; }
  #raiz-painel .grid-stack-item-content {
      background: #FFFFFF; border: 1px solid #14161A; overflow: hidden;
      display: flex; flex-direction: column; padding: 8px 10px;
  }
  #raiz-painel .cabeca {
      display: flex; align-items: baseline; justify-content: space-between;
      padding-bottom: 6px; cursor: move; user-select: none;
  }
  #raiz-painel .titulo {
      font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif;
      font-weight: 700; font-size: 21px; text-transform: uppercase; color: #14161A;
  }
  #raiz-painel .nota {
      font-size: 10px; letter-spacing: .14em; text-transform: uppercase; color: #7C858D;
  }
  #raiz-painel .tela { flex: 1; min-height: 0; }
  #raiz-painel .barra { display: flex; gap: 8px; align-items: center; margin-bottom: 8px; }
  #raiz-painel .botao {
      font-size: 11px; letter-spacing: .08em; text-transform: uppercase;
      background: #FFFFFF; border: 1px solid #14161A; padding: 5px 12px; cursor: pointer;
  }
  #raiz-painel .botao:hover { background: #14161A; color: #FFFFFF; }
  #raiz-painel .aviso {
      border: 1px solid #C2492E; background: #FDF1EE; color: #8A2F1B;
      padding: 12px 14px; font-size: 13px; line-height: 1.5;
  }
  #raiz-painel .dica { font-size: 11px; color: #7C858D; margin-left: auto; }
</style>

<div id="raiz-painel" style="min-height: ${altura_total}px;">
  <div class="barra">
    <button class="botao" onclick="restaurarLayout()">Restaurar layout</button>
    <span class="dica">Arraste pelo título · redimensione pelo canto inferior direito · o layout fica salvo neste navegador</span>
  </div>
  <div class="grid-stack"></div>
</div>

<script>
  const PAINEIS = $paineis;
  const LAYOUT_PADRAO = $layout_padrao;
  const CHAVE = "layout_painel_$chave";

  function bibliotecasOk() {
    return typeof GridStack !== "undefined" && typeof Plotly !== "undefined";
  }

  function layoutSalvo() {
    try {
      const bruto = window.localStorage.getItem(CHAVE);
      return bruto ? JSON.parse(bruto) : null;
    } catch (e) { return null; }
  }

  function salvarLayout(grade) {
    try {
      const atual = grade.save(false).map(n => ({ id: n.id, x: n.x, y: n.y, w: n.w, h: n.h }));
      window.localStorage.setItem(CHAVE, JSON.stringify(atual));
    } catch (e) { /* navegador sem localStorage: segue sem salvar */ }
  }

  function restaurarLayout() {
    try { window.localStorage.removeItem(CHAVE); } catch (e) {}
    window.location.reload();
  }

  function desenhar(painel) {
    const alvo = document.getElementById("tela-" + painel.id);
    const layout = Object.assign({}, painel.figura.layout, {
      autosize: true, height: null, width: null,
      margin: { l: 8, r: 8, t: 6, b: 26 }
    });
    Plotly.newPlot(alvo, painel.figura.data, layout,
                   { displayModeBar: false, responsive: true });
  }

  function iniciar() {
    if (!bibliotecasOk()) {
      document.getElementById("raiz-painel").innerHTML =
        '<div class="aviso"><b>Não foi possível carregar a grade arrastável.</b><br>' +
        'A rede bloqueou o acesso a cdn.jsdelivr.net ou cdn.plot.ly. ' +
        'Desmarque "Modo painel" na barra lateral para voltar ao modo apresentação, ' +
        'que não depende de nenhum recurso externo.</div>';
      return;
    }

    const grade = document.querySelector("#raiz-painel .grid-stack");
    const posicoes = layoutSalvo() || LAYOUT_PADRAO;
    const porId = {};
    posicoes.forEach(p => porId[p.id] = p);

    PAINEIS.forEach(painel => {
      const pos = porId[painel.id] || { x: 0, y: 0, w: 4, h: 4 };
      const item = document.createElement("div");
      item.className = "grid-stack-item";
      item.setAttribute("gs-id", painel.id);
      item.setAttribute("gs-x", pos.x); item.setAttribute("gs-y", pos.y);
      item.setAttribute("gs-w", pos.w); item.setAttribute("gs-h", pos.h);
      item.innerHTML =
        '<div class="grid-stack-item-content">' +
          '<div class="cabeca"><span class="titulo">' + painel.titulo + '</span>' +
          '<span class="nota">' + painel.nota + '</span></div>' +
          '<div class="tela" id="tela-' + painel.id + '"></div>' +
        '</div>';
      grade.appendChild(item);
    });

    const gs = GridStack.init({
      column: 12, cellHeight: $altura_celula, margin: 6,
      handle: ".cabeca", float: true, resizable: { handles: "se, sw, e, s" }
    }, grade);

    PAINEIS.forEach(desenhar);

    function redimensionar() {
      PAINEIS.forEach(p => Plotly.Plots.resize(document.getElementById("tela-" + p.id)));
    }
    gs.on("resizestop", () => { redimensionar(); salvarLayout(gs); });
    gs.on("dragstop", () => salvarLayout(gs));
    window.addEventListener("resize", redimensionar);
  }

  // st.html injeta o bloco depois do "load" da página; espera as bibliotecas do CDN
  (function aguardar(tentativas) {
    if (bibliotecasOk() || tentativas <= 0) { iniciar(); }
    else { setTimeout(() => aguardar(tentativas - 1), 200); }
  })(25);
</script>
""")


def renderizar_painel(figuras: list[dict], chave: str, altura_celula: int = 70,
                      altura_total: int = 700) -> None:
    """
    Desenha a grade arrastável.

    figuras: lista de dicionários com id, titulo, nota e figura (objeto Plotly).
    chave: identificador do layout salvo (usar o estado, ex.: "AM").
    """
    paineis = [
        {
            "id": item["id"],
            "titulo": item["titulo"],
            "nota": item["nota"],
            "figura": json.loads(item["figura"].to_json()),
        }
        for item in figuras
    ]

    html = _MODELO.substitute(
        paineis=json.dumps(paineis),
        layout_padrao=json.dumps(LAYOUT_PADRAO),
        chave="".join(c for c in chave if c.isalnum()) or "geral",
        altura_celula=altura_celula,
        altura_total=altura_total,
    )
    st.html(html, unsafe_allow_javascript=True)
