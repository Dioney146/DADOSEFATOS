"""
Modo painel — os cinco gráficos numa grade que se arrasta e redimensiona
com o mouse, no estilo do canvas do Power BI.

Sem nenhum recurso externo: o Plotly.js vai embutido na própria página (vem
junto do pacote plotly instalado pelo requirements.txt) e o arrastar/
redimensionar é código próprio, em JavaScript puro. Assim funciona mesmo em
rede que bloqueia CDN e Google Fonts.

O layout montado pelo usuário fica salvo no navegador (localStorage),
separado por estado.
"""

from __future__ import annotations

import json
from string import Template

import streamlit as st
from plotly.offline import get_plotlyjs

# Grade de 12 colunas: largura (colunas) e altura (linhas) iniciais de cada painel
LAYOUT_PADRAO = [
    {"id": "veiculos", "w": 4, "h": 4},
    {"id": "ocupacao", "w": 4, "h": 4},
    {"id": "drop", "w": 4, "h": 4},
    {"id": "paradas", "w": 6, "h": 4},
    {"id": "peso", "w": 6, "h": 4},
]

_MODELO = Template("""
<style>
  body { margin: 0; background: #F4F4F2;
         font-family: 'Archivo', 'Segoe UI', Arial, sans-serif; color: #14161A; }
  .barra { display: flex; align-items: center; gap: 10px; margin: 0 0 8px 0; }
  .botao { font-size: 11px; letter-spacing: .08em; text-transform: uppercase;
           background: #FFFFFF; border: 1px solid #14161A; padding: 5px 12px; cursor: pointer; }
  .botao:hover { background: #14161A; color: #FFFFFF; }
  .dica { font-size: 11px; color: #7C858D; margin-left: auto; }

  .grade { display: grid; grid-template-columns: repeat(12, 1fr);
           grid-auto-rows: ${altura_celula}px; gap: 8px; }
  .painel { background: #FFFFFF; border: 1px solid #14161A; position: relative;
            display: flex; flex-direction: column; padding: 8px 10px; overflow: hidden; }
  .painel.arrastando { opacity: .35; }
  .painel.alvo { outline: 2px dashed #5D87B0; outline-offset: -3px; }
  .cabeca { display: flex; align-items: baseline; justify-content: space-between;
            padding-bottom: 6px; cursor: grab; user-select: none; }
  .titulo { font-family: 'Barlow Condensed', 'Arial Narrow', Arial, sans-serif;
            font-weight: 700; font-size: 21px; text-transform: uppercase; }
  .nota { font-size: 10px; letter-spacing: .14em; text-transform: uppercase; color: #7C858D; }
  .tela { flex: 1; min-height: 0; }
  .puxador { position: absolute; right: 0; bottom: 0; width: 16px; height: 16px;
             cursor: nwse-resize; background:
             linear-gradient(135deg, transparent 50%, #B9BEC4 50%, #B9BEC4 100%); }
</style>

<div class="barra">
  <button class="botao" onclick="restaurarLayout()">Restaurar layout</button>
  <span class="dica">Arraste um painel pelo título para trocar de lugar · puxe o canto inferior direito para redimensionar</span>
</div>
<div class="grade" id="grade"></div>

<script>${plotlyjs}</script>
<script>
  const PAINEIS = ${paineis};
  const LAYOUT_PADRAO = ${layout_padrao};
  const CHAVE = "painel_${chave}";

  function lerLayout() {
    try {
      const bruto = window.localStorage.getItem(CHAVE);
      if (!bruto) return null;
      const salvo = JSON.parse(bruto);
      const ids = PAINEIS.map(p => p.id).sort().join(",");
      return salvo.map(s => s.id).sort().join(",") === ids ? salvo : null;
    } catch (e) { return null; }
  }

  function gravarLayout() {
    try {
      const atual = [...document.querySelectorAll(".painel")].map(el => ({
        id: el.dataset.id, w: Number(el.dataset.w), h: Number(el.dataset.h)
      }));
      window.localStorage.setItem(CHAVE, JSON.stringify(atual));
    } catch (e) { /* navegador sem localStorage: apenas não salva */ }
  }

  function restaurarLayout() {
    try { window.localStorage.removeItem(CHAVE); } catch (e) {}
    document.getElementById("grade").innerHTML = "";
    montar(LAYOUT_PADRAO);
  }

  function aplicarTamanho(el) {
    el.style.gridColumn = "span " + el.dataset.w;
    el.style.gridRow = "span " + el.dataset.h;
  }

  function redesenhar(el) {
    const tela = el.querySelector(".tela");
    if (tela && tela.data) { Plotly.Plots.resize(tela); }
  }

  function montar(layout) {
    const grade = document.getElementById("grade");
    const porId = {};
    PAINEIS.forEach(p => porId[p.id] = p);

    layout.forEach(pos => {
      const painel = porId[pos.id];
      if (!painel) return;

      const el = document.createElement("div");
      el.className = "painel";
      el.dataset.id = painel.id;
      el.dataset.w = pos.w;
      el.dataset.h = pos.h;
      el.innerHTML =
        '<div class="cabeca" draggable="true">' +
          '<span class="titulo">' + painel.titulo + '</span>' +
          '<span class="nota">' + painel.nota + '</span>' +
        '</div>' +
        '<div class="tela"></div>' +
        '<div class="puxador"></div>';
      aplicarTamanho(el);
      grade.appendChild(el);

      const tela = el.querySelector(".tela");
      const layoutFig = Object.assign({}, painel.figura.layout,
        { autosize: true, height: null, width: null });
      Plotly.newPlot(tela, painel.figura.data, layoutFig,
                     { displayModeBar: false, responsive: true });

      ligarArrastar(el);
      ligarRedimensionar(el);
    });
  }

  /* ── trocar painéis de lugar ─────────────────────────────────────────── */
  let arrastado = null;

  function ligarArrastar(el) {
    const cabeca = el.querySelector(".cabeca");
    cabeca.addEventListener("dragstart", ev => {
      arrastado = el;
      el.classList.add("arrastando");
      ev.dataTransfer.effectAllowed = "move";
      ev.dataTransfer.setData("text/plain", el.dataset.id);
    });
    cabeca.addEventListener("dragend", () => {
      el.classList.remove("arrastando");
      document.querySelectorAll(".painel").forEach(p => p.classList.remove("alvo"));
      arrastado = null;
    });
    el.addEventListener("dragover", ev => {
      if (arrastado && arrastado !== el) { ev.preventDefault(); el.classList.add("alvo"); }
    });
    el.addEventListener("dragleave", () => el.classList.remove("alvo"));
    el.addEventListener("drop", ev => {
      ev.preventDefault();
      el.classList.remove("alvo");
      if (!arrastado || arrastado === el) return;
      const grade = document.getElementById("grade");
      const itens = [...grade.children];
      const posOrigem = itens.indexOf(arrastado);
      const posDestino = itens.indexOf(el);
      if (posOrigem < posDestino) { grade.insertBefore(arrastado, el.nextSibling); }
      else { grade.insertBefore(arrastado, el); }
      [arrastado, el].forEach(redesenhar);
      gravarLayout();
    });
  }

  /* ── redimensionar pelo canto ────────────────────────────────────────── */
  function ligarRedimensionar(el) {
    const puxador = el.querySelector(".puxador");
    puxador.addEventListener("mousedown", ev => {
      ev.preventDefault();
      const grade = document.getElementById("grade");
      const larguraColuna = (grade.clientWidth + 8) / 12;
      const alturaLinha = ${altura_celula} + 8;
      const x0 = ev.clientX, y0 = ev.clientY;
      const w0 = Number(el.dataset.w), h0 = Number(el.dataset.h);

      function mover(e) {
        const w = Math.min(12, Math.max(2, w0 + Math.round((e.clientX - x0) / larguraColuna)));
        const h = Math.max(2, h0 + Math.round((e.clientY - y0) / alturaLinha));
        if (w !== Number(el.dataset.w) || h !== Number(el.dataset.h)) {
          el.dataset.w = w; el.dataset.h = h;
          aplicarTamanho(el);
          redesenhar(el);
        }
      }
      function soltar() {
        document.removeEventListener("mousemove", mover);
        document.removeEventListener("mouseup", soltar);
        redesenhar(el);
        gravarLayout();
      }
      document.addEventListener("mousemove", mover);
      document.addEventListener("mouseup", soltar);
    });
  }

  montar(lerLayout() || LAYOUT_PADRAO);
  window.addEventListener("resize", () => document.querySelectorAll(".painel").forEach(redesenhar));
</script>
""")


def _desenhar_html(html: str, altura: int) -> None:
    """Usa a API de componente disponível na versão instalada do Streamlit."""
    try:
        import streamlit.components.v1 as components
        components.html(html, height=altura, scrolling=True)
    except Exception:  # noqa: BLE001
        st.html(html, unsafe_allow_javascript=True)


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
        plotlyjs=get_plotlyjs(),
        paineis=json.dumps(paineis),
        layout_padrao=json.dumps(LAYOUT_PADRAO),
        chave="".join(c for c in chave if c.isalnum()) or "geral",
        altura_celula=altura_celula,
    )
    _desenhar_html(html, altura_total)
