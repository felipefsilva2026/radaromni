"""Etapa 2: lê a planilha de curadoria já revisada e gera o HTML executivo da diretoria.

Só entram as linhas com STATUS = SIM. Impacto, texto de impacto, área responsável
e comentário vêm exatamente do que estiver preenchido na planilha.
"""

import argparse
import html
import sys
import unicodedata
from collections import Counter
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

from gerar_planilha import SEARCHES
from imagens import embed_many, enrich_many

ASSETS_DIR = Path(__file__).parent / "assets"
PLANILHAS_DIR = Path("planilhas")

CATEGORY_ORDER = list(SEARCHES)
CATEGORY_ICON = {
    "Omni": "🏢",
    "Concorrentes": "🎯",
    "Selic e macroeconomia": "📈",
    "Mercado automotivo": "🚗",
    "Seguros": "🛡️",
    "Financiamento e crédito": "💳",
    "Bancos e parceiros": "🤝",
    "Riscos e crises": "⚠️",
}
IMPACT_ORDER = {"Alto": 0, "Médio": 1, "Baixo": 2}
IMPACT_SLUG = {"Alto": "alto", "Médio": "medio", "Baixo": "baixo"}
IMPACT_ICON = {"Alto": "▲", "Médio": "◆", "Baixo": "▼"}
NOT_CLASSIFIED = "Não classificado"
MONTHS = ("jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez")


def strip_accents(value: str) -> str:
    return unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")


def norm(value) -> str:
    return strip_accents(str(value or "").strip()).upper()


def text(value) -> str:
    return str(value).strip() if value is not None else ""


def esc(value) -> str:
    return html.escape(text(value), quote=True)


def normalize_impact(value) -> str:
    key = norm(value)
    if key.startswith("ALT"):
        return "Alto"
    if key.startswith("MED"):
        return "Médio"
    if key.startswith("BAIX"):
        return "Baixo"
    return NOT_CLASSIFIED


def is_selected(value) -> bool:
    return norm(value) in ("SIM", "S", "YES", "X", "OK")


def top_position(value) -> int | None:
    """Coluna TOP 6: número = posição escolhida; SIM/X = entra sem posição; vazio = não entra."""
    if value is None or text(value) == "":
        return None
    try:
        return int(float(str(value).replace(",", ".")))
    except ValueError:
        return 99 if is_selected(value) else None


def latest_spreadsheet() -> Path:
    files = [p for p in PLANILHAS_DIR.glob("*.xlsx") if not p.name.startswith("~$")]
    if not files:
        sys.exit(f"Nenhuma planilha encontrada em '{PLANILHAS_DIR}'. Rode primeiro '1 - Gerar planilha.bat'.")
    return max(files, key=lambda p: p.stat().st_mtime)


def read_selected(path: Path) -> tuple[list[dict], int]:
    wb = load_workbook(path, data_only=True, read_only=True)
    ws = wb["Pesquisa"] if "Pesquisa" in wb.sheetnames else wb.worksheets[0]
    rows = ws.iter_rows(values_only=True)
    headers = [norm(h) for h in next(rows)]
    for required in ("TITULO", "STATUS"):
        if required not in headers:
            sys.exit(f"A planilha '{path.name}' não tem a coluna '{required.title()}'. Não renomeie os cabeçalhos da aba 'Pesquisa'.")

    def get(row, name):
        key = norm(name)
        return row[headers.index(key)] if key in headers and headers.index(key) < len(row) else None

    total = 0
    selected = []
    for row in rows:
        if not any(row) or not text(get(row, "Título")):
            continue
        total += 1
        top = top_position(get(row, "TOP 6"))
        # Marcar TOP 6 já leva a notícia para a diretoria, mesmo sem SIM no STATUS.
        if not is_selected(get(row, "STATUS")) and top is None:
            continue
        score = get(row, "Score")
        selected.append({
            "date": get(row, "Data"),
            "category": text(get(row, "Categoria")) or "Outros",
            "title": text(get(row, "Título")),
            "summary": text(get(row, "Resumo")),
            "source": text(get(row, "Fonte")) or "Fonte não informada",
            "link": text(get(row, "Link")),
            "image": text(get(row, "Imagem")),
            "score": score if isinstance(score, (int, float)) else 0,
            "impact": normalize_impact(get(row, "Impacto")),
            "impact_text": text(get(row, "Impacto para a Omni")),
            "area": text(get(row, "Área responsável")),
            "comment": text(get(row, "Comentário da equipe")),
            "top": top,
        })
    wb.close()
    selected.sort(key=lambda i: (IMPACT_ORDER.get(i["impact"], 3), -i["score"]))
    return selected, total


def date_label(value) -> str:
    if isinstance(value, datetime):
        return f"{value.day:02d} {MONTHS[value.month - 1]} {value.year}"
    return text(value)


def category_sort_key(category: str) -> int:
    return CATEGORY_ORDER.index(category) if category in CATEGORY_ORDER else len(CATEGORY_ORDER)


def impact_badge(impact: str) -> str:
    slug = IMPACT_SLUG.get(impact, "nc")
    icon = IMPACT_ICON.get(impact, "●")
    return f'<span class="badge badge--{slug}"><span aria-hidden="true">{icon}</span> Impacto {esc(impact)}</span>'


def read_logo(name: str) -> str:
    path = ASSETS_DIR / name
    return path.read_text(encoding="utf-8") if path.exists() else ""


def render_kpis(items: list[dict], total: int) -> str:
    alto = sum(1 for i in items if i["impact"] == "Alto")
    omni = sum(1 for i in items if i["category"] == "Omni")
    areas = len({i["area"] for i in items if i["area"]})
    tiles = [
        (len(items), "notícias relevantes", f"de {total} analisadas", ""),
        (alto, "de alto impacto", "exigem atenção da diretoria", " kpi--alto"),
        (omni, "sobre a Omni", "marcas do grupo na mídia", " kpi--omni"),
        (areas, "áreas envolvidas", "com temas para acompanhar", ""),
    ]
    return "".join(
        f'<div class="kpi{extra}"><span class="kpi-value">{value}</span><span class="kpi-label">{label}</span><span class="kpi-sub">{sub}</span></div>'
        for value, label, sub, extra in tiles
    )


def render_impact_meter(items: list[dict]) -> str:
    counts = Counter(i["impact"] for i in items)
    levels = [lvl for lvl in ("Alto", "Médio", "Baixo", NOT_CLASSIFIED) if counts[lvl]]
    total = len(items) or 1
    segments = "".join(
        f'<div class="meter-seg meter-seg--{IMPACT_SLUG.get(lvl, "nc")}" style="flex:{counts[lvl]}" title="{esc(lvl)}: {counts[lvl]} notícia(s) ({round(counts[lvl] * 100 / total)}%)"></div>'
        for lvl in levels
    )
    legend = "".join(
        f'<div class="meter-legend-item"><span class="dot dot--{IMPACT_SLUG.get(lvl, "nc")}"></span>'
        f'<span class="meter-legend-label">{IMPACT_ICON.get(lvl, "●")} {esc(lvl)}</span>'
        f'<strong>{counts[lvl]}</strong><span class="muted">{round(counts[lvl] * 100 / total)}%</span></div>'
        for lvl in levels
    )
    return f'<div class="meter">{segments}</div><div class="meter-legend">{legend}</div>'


def render_category_bars(items: list[dict]) -> str:
    counts = Counter(i["category"] for i in items)
    ordered = sorted(counts.items(), key=lambda kv: (-kv[1], category_sort_key(kv[0])))
    peak = max(counts.values()) if counts else 1
    rows = []
    for category, count in ordered:
        altos = sum(1 for i in items if i["category"] == category and i["impact"] == "Alto")
        width = max(count * 100 / peak, 4)
        hint = f" · {altos} alto impacto" if altos else ""
        rows.append(
            f'<div class="bar-row" title="{esc(category)}: {count} notícia(s){hint}">'
            f'<span class="bar-label">{CATEGORY_ICON.get(category, "•")} {esc(category)}</span>'
            f'<span class="bar-track"><span class="bar-fill" style="width:{width:.1f}%"></span></span>'
            f'<span class="bar-value">{count}</span></div>'
        )
    return "".join(rows)


def render_areas(items: list[dict]) -> str:
    counts = Counter(i["area"] for i in items if i["area"])
    if not counts:
        return '<p class="muted">Nenhuma área preenchida na planilha.</p>'
    chips = []
    for area, count in counts.most_common():
        altos = sum(1 for i in items if i["area"] == area and i["impact"] == "Alto")
        alert = f'<span class="chip-alert">{altos} alto</span>' if altos else ""
        chips.append(f'<span class="chip"><strong>{count}</strong> {esc(area)}{alert}</span>')
    return "".join(chips)


def render_visual(item: dict, big: bool) -> str:
    icon = CATEGORY_ICON.get(item["category"], "📰")
    fallback = f'<div class="visual-fallback"><span class="visual-icon">{icon}</span><span class="visual-cat">{esc(item["category"])}</span></div>'
    if not item["image"]:
        return f'<div class="visual{" visual--big" if big else ""}">{fallback}</div>'
    return (
        f'<div class="visual{" visual--big" if big else ""}">{fallback}'
        f'<img src="{esc(item["image"])}" alt="" loading="lazy" onerror="this.nextElementSibling.remove(); this.remove()">'
        f'<span class="visual-tag">{icon} {esc(item["category"])}</span></div>'
    )


def render_comment(item: dict) -> str:
    if not item["comment"]:
        return ""
    return f'<blockquote class="comment"><span class="comment-label">Comentário da equipe</span>{esc(item["comment"])}</blockquote>'


def render_highlight(item: dict) -> str:
    link_open = f'<a href="{esc(item["link"])}" target="_blank" rel="noopener">' if item["link"] else "<span>"
    link_close = "</a>" if item["link"] else "</span>"
    area = f'<span class="area">👤 {esc(item["area"])}</span>' if item["area"] else ""
    return f"""
      <article class="highlight highlight--{IMPACT_SLUG.get(item['impact'], 'nc')}">
        {render_visual(item, True)}
        <div class="highlight-body">
          <div class="meta-row">{impact_badge(item['impact'])}<span class="cat">{esc(item['category'])}</span></div>
          <h3>{link_open}{esc(item['title'])}{link_close}</h3>
          <p class="source">{esc(item['source'])} · {esc(date_label(item['date']))}</p>
          <div class="impact-box"><span class="impact-box-label">O que significa para a Omni</span>{esc(item['impact_text']) or '<span class="muted">Sem texto de impacto preenchido.</span>'}</div>
          {render_comment(item)}
          <div class="card-foot">{area}{f'<a class="more" href="{esc(item["link"])}" target="_blank" rel="noopener">Ler notícia →</a>' if item['link'] else ''}</div>
        </div>
      </article>"""


def render_card(item: dict) -> str:
    title = f'<a href="{esc(item["link"])}" target="_blank" rel="noopener">{esc(item["title"])}</a>' if item["link"] else esc(item["title"])
    area = f'<span class="area">👤 {esc(item["area"])}</span>' if item["area"] else ""
    return f"""
        <article class="card card--{IMPACT_SLUG.get(item['impact'], 'nc')}">
          {render_visual(item, False)}
          <div class="card-body">
            <div class="meta-row">{impact_badge(item['impact'])}</div>
            <h4>{title}</h4>
            <p class="source">{esc(item['source'])} · {esc(date_label(item['date']))}</p>
            <p class="impact-text">{esc(item['impact_text'])}</p>
            {render_comment(item)}
            <div class="card-foot">{area}</div>
          </div>
        </article>"""


def clip(value: str, limit: int) -> str:
    value = text(value)
    if len(value) <= limit:
        return value
    return value[:limit].rsplit(" ", 1)[0].rstrip(",.;:—-") + "…"


RESUMO_TOP = 6       # notícias na página executiva
RESUMO_FEATURED = 2  # as primeiras vão em destaque, com foto grande


def render_resumo_card(item: dict, rank: int, featured: bool) -> str:
    slug = IMPACT_SLUG.get(item["impact"], "nc")
    title = esc(clip(item["title"], 150 if featured else 115))
    if item["link"]:
        title = f'<a href="{esc(item["link"])}" target="_blank" rel="noopener">{title}</a>'
    comment = f'<p class="rcard-comment">💬 {esc(clip(item["comment"], 120 if featured else 80))}</p>' if item["comment"] else ""
    kind = "rcard--featured" if featured else "rcard--compact"
    return f"""
          <article class="rcard {kind} rcard--{slug}">
            <div class="rcard-media">{render_visual(item, featured)}<span class="rank">{rank}</span></div>
            <div class="rcard-body">
              <div class="rcard-tags">{impact_badge(item['impact'])}<span class="rcard-cat">{CATEGORY_ICON.get(item['category'], '•')} {esc(item['category'])}</span></div>
              <h3 class="rcard-title">{title}</h3>
              {comment}
              <p class="rcard-meta">{esc(item['source'])} · {esc(date_label(item['date']))}</p>
            </div>
          </article>"""


def render_resumo(items: list[dict], total: int, period: str, now: datetime) -> str:
    """Página executiva de uma folha (A4): números do período e as Top 6 notícias com foto.

    Quem entra é escolhido na coluna TOP 6 da planilha (na ordem das posições digitadas).
    Se ninguém foi marcado, entram as 6 primeiras pela ordem automática (impacto, depois Score).
    """
    chosen = sorted((i for i in items if i["top"] is not None), key=lambda i: (i["top"], IMPACT_ORDER.get(i["impact"], 3), -i["score"]))
    top = chosen[:RESUMO_TOP] if chosen else items[:RESUMO_TOP]
    cards = "".join(render_resumo_card(item, rank, rank <= RESUMO_FEATURED) for rank, item in enumerate(top, start=1))
    rest = len(items) - len(top)
    rest_note = f"+ {rest} notícia(s) na versão completa" if rest > 0 else ""
    return f"""
  <div class="resumo">
    <div class="sheet">
      <div class="sheet-head">
        <div class="sheet-top"><div class="sheet-logo">{read_logo("logo-omni-white.svg")}</div><span class="eyebrow">RESUMO EXECUTIVO · DIRETORIA</span></div>
        <h1>Radar de Mercado Omni</h1>
        <p>{now.day:02d} {MONTHS[now.month - 1]} {now.year}{f" · notícias de {period}" if period else ""}</p>
      </div>
      <div class="sheet-body">
        <div class="sheet-section-title"><h2>🔥 Top {len(top)} do período</h2><span>{rest_note}</span></div>
        <div class="top6 top6--{len(top)}">{cards}</div>
      </div>
      <div class="sheet-foot">Curadoria da equipe sobre notícias públicas (Google Notícias) · {len(items)} de {total} notícias aprovadas · gerado em {now:%d/%m/%Y %H:%M}</div>
    </div>
  </div>"""


def render_by_category(items: list[dict]) -> tuple[str, str]:
    groups: dict[str, list[dict]] = {}
    for item in items:
        groups.setdefault(item["category"], []).append(item)
    sections, nav = [], []
    for index, category in enumerate(sorted(groups, key=category_sort_key)):
        group = groups[category]
        anchor = f"tema-{index}"
        nav.append(f'<a href="#{anchor}">{CATEGORY_ICON.get(category, "•")} {esc(category)} <span>{len(group)}</span></a>')
        cards = "".join(render_card(i) for i in group)
        omni = " theme--omni" if category == "Omni" else ""
        title = "A Omni na mídia" if category == "Omni" else category
        sections.append(f"""
    <section class="theme{omni}" id="{anchor}">
      <div class="theme-head"><span class="theme-icon">{CATEGORY_ICON.get(category, "•")}</span><h3>{esc(title)}</h3><span class="theme-count">{len(group)} notícia(s)</span></div>
      <div class="cards">{cards}</div>
    </section>""")
    return "".join(sections), "".join(nav)


CSS = """
:root {
  --blue: #013188; --blue-soft: #2453A9; --blue-pale: #E6ECF7;
  --orange: #E95814; --orange-pale: #FDEBE3;
  --ink: #1F2330; --ink-2: #4A4F5C; --muted: #6B7180;
  --line: #E3E5EA; --bg: #F4F5F8; --card: #FFFFFF;
  --alto: #C43B2C; --alto-pale: #FBE4E1;
  --medio: #B86E00; --medio-pale: #FCEFD9;
  --baixo: #5B6B82; --baixo-pale: #E9EDF3;
  --nc: #8A8F99; --nc-pale: #EFEFF1;
}
* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { margin: 0; font-family: "Segoe UI", -apple-system, Helvetica, Arial, sans-serif; color: var(--ink); background: var(--bg); line-height: 1.45; }
a { color: inherit; }
.wrap { max-width: 1180px; margin: 0 auto; padding: 0 24px; }
.muted { color: var(--muted); }

.hero { background: radial-gradient(circle at 85% 20%, #1c4bb0 0%, var(--blue) 45%, #01195a 100%); color: #fff; padding: 44px 0 120px; position: relative; overflow: hidden; }
.hero::after { content: ""; position: absolute; right: -120px; bottom: -160px; width: 420px; height: 420px; border-radius: 50%; border: 60px solid rgba(233, 88, 20, 0.9); opacity: .9; }
.hero .wrap { position: relative; z-index: 1; }
.hero-logo svg { height: 38px; width: auto; display: block; margin-bottom: 36px; }
.eyebrow { display: inline-block; background: var(--orange); color: #fff; font-weight: 700; font-size: 12px; letter-spacing: 1.2px; padding: 5px 12px; border-radius: 4px; margin: 0 0 16px; }
.hero h1 { font-size: 46px; line-height: 1.05; margin: 0 0 12px; letter-spacing: -0.5px; }
.hero p { font-size: 18px; color: #C9D6EF; margin: 0; max-width: 620px; }
.hero .when { margin-top: 22px; font-size: 14px; color: #fff; opacity: .85; }

.kpis { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-top: -78px; position: relative; z-index: 2; }
.kpi { background: var(--card); border-radius: 14px; padding: 22px 22px 20px; box-shadow: 0 10px 30px rgba(1, 25, 90, .12); border-top: 5px solid var(--blue); }
.kpi--alto { border-top-color: var(--alto); }
.kpi--omni { border-top-color: var(--orange); }
.kpi-value { display: block; font-size: 48px; font-weight: 800; line-height: 1; color: var(--ink); }
.kpi--alto .kpi-value { color: var(--alto); }
.kpi--omni .kpi-value { color: var(--orange); }
.kpi-label { display: block; font-size: 15px; font-weight: 700; margin-top: 8px; }
.kpi-sub { display: block; font-size: 12.5px; color: var(--muted); margin-top: 2px; }

.nav { position: sticky; top: 0; z-index: 10; background: rgba(244, 245, 248, .95); backdrop-filter: blur(6px); border-bottom: 1px solid var(--line); margin-top: 28px; }
.nav .wrap { display: flex; gap: 8px; overflow-x: auto; padding-top: 10px; padding-bottom: 10px; }
.nav a { white-space: nowrap; text-decoration: none; font-size: 13px; font-weight: 600; color: var(--ink-2); background: var(--card); border: 1px solid var(--line); padding: 6px 12px; border-radius: 20px; }
.nav a:hover { border-color: var(--blue-soft); color: var(--blue); }
.nav a span { color: var(--muted); font-weight: 700; margin-left: 4px; }

section.block { margin-top: 40px; }
.block-title { display: flex; align-items: baseline; gap: 12px; margin: 0 0 16px; }
.block-title h2 { font-size: 24px; margin: 0; color: var(--blue); letter-spacing: -0.2px; }
.block-title p { margin: 0; font-size: 13.5px; color: var(--muted); }

.overview { display: grid; grid-template-columns: 1fr 1.2fr; gap: 16px; }
.panel { background: var(--card); border-radius: 14px; padding: 22px 24px; border: 1px solid var(--line); }
.panel h3 { font-size: 13px; text-transform: uppercase; letter-spacing: .6px; color: var(--muted); margin: 0 0 16px; }
.meter { display: flex; gap: 2px; height: 26px; border-radius: 6px; overflow: hidden; }
.meter-seg { min-width: 8px; }
.meter-seg--alto, .dot--alto { background: var(--alto); }
.meter-seg--medio, .dot--medio { background: var(--medio); }
.meter-seg--baixo, .dot--baixo { background: var(--baixo); }
.meter-seg--nc, .dot--nc { background: var(--nc); }
.meter-legend { display: grid; gap: 10px; margin-top: 18px; }
.meter-legend-item { display: grid; grid-template-columns: 14px 1fr auto 48px; align-items: center; gap: 10px; font-size: 14px; }
.meter-legend-item strong { font-size: 20px; text-align: right; }
.meter-legend-item .muted { text-align: right; font-size: 13px; }
.dot { width: 12px; height: 12px; border-radius: 3px; }
.bar-row { display: grid; grid-template-columns: 210px 1fr 32px; align-items: center; gap: 12px; padding: 5px 0; }
.bar-label { font-size: 13.5px; color: var(--ink-2); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.bar-track { height: 14px; background: var(--bg); border-radius: 4px; overflow: hidden; }
.bar-fill { display: block; height: 100%; background: var(--blue-soft); border-radius: 0 4px 4px 0; }
.bar-row:hover .bar-fill { background: var(--blue); }
.bar-value { font-weight: 800; font-size: 15px; text-align: right; }
.areas { margin-top: 16px; }
.chips { display: flex; flex-wrap: wrap; gap: 8px; }
.chip { background: var(--blue-pale); color: var(--blue); font-size: 13.5px; padding: 7px 14px; border-radius: 20px; display: inline-flex; align-items: center; gap: 6px; }
.chip-alert { background: var(--alto); color: #fff; font-size: 11px; font-weight: 700; padding: 1px 8px; border-radius: 10px; }

.badge { display: inline-flex; align-items: center; gap: 5px; font-size: 11.5px; font-weight: 700; padding: 3px 10px; border-radius: 20px; }
.badge--alto { background: var(--alto-pale); color: var(--alto); }
.badge--medio { background: var(--medio-pale); color: var(--medio); }
.badge--baixo { background: var(--baixo-pale); color: var(--baixo); }
.badge--nc { background: var(--nc-pale); color: var(--nc); }
.meta-row { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 8px; }
.cat { font-size: 12px; font-weight: 600; color: var(--muted); text-transform: uppercase; letter-spacing: .4px; }
.source { font-size: 12.5px; color: var(--muted); margin: 0 0 12px; }
.area { font-size: 12.5px; font-weight: 600; color: var(--blue); background: var(--blue-pale); padding: 4px 10px; border-radius: 6px; }
.card-foot { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-top: auto; padding-top: 12px; }
.more { font-size: 13px; font-weight: 700; color: var(--orange); text-decoration: none; }
.more:hover { text-decoration: underline; }

.visual { position: relative; background: linear-gradient(135deg, var(--blue) 0%, var(--blue-soft) 100%); height: 120px; overflow: hidden; }
.visual--big { height: auto; min-height: 100%; }
.visual img { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; }
.visual-fallback { position: absolute; inset: 0; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 6px; color: #fff; }
.visual-icon { font-size: 38px; }
.visual--big .visual-icon { font-size: 56px; }
.visual-tag { position: absolute; left: 10px; bottom: 10px; background: rgba(1, 25, 90, .85); color: #fff; font-size: 10.5px; font-weight: 700; text-transform: uppercase; letter-spacing: .5px; padding: 4px 9px; border-radius: 5px; }
.visual-cat { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: .8px; opacity: .85; }

.highlights { display: grid; gap: 18px; }
.highlight { display: grid; grid-template-columns: 240px 1fr; background: var(--card); border-radius: 16px; overflow: hidden; border: 1px solid var(--line); box-shadow: 0 6px 20px rgba(1, 25, 90, .06); border-left: 6px solid var(--alto); }
.highlight--medio { border-left-color: var(--medio); }
.highlight--baixo { border-left-color: var(--baixo); }
.highlight-body { padding: 22px 26px; display: flex; flex-direction: column; }
.highlight h3 { font-size: 21px; line-height: 1.3; margin: 0 0 6px; }
.highlight h3 a { text-decoration: none; }
.highlight h3 a:hover { color: var(--blue); }
.impact-box { background: var(--orange-pale); border-radius: 10px; padding: 14px 16px; font-size: 15px; line-height: 1.5; }
.impact-box-label, .comment-label { display: block; font-size: 11px; font-weight: 800; text-transform: uppercase; letter-spacing: .6px; color: var(--orange); margin-bottom: 4px; }
.comment { margin: 12px 0 0; padding: 12px 16px; border-left: 4px solid var(--blue); background: var(--blue-pale); border-radius: 0 10px 10px 0; font-size: 14px; font-style: italic; color: var(--ink-2); }
.comment-label { color: var(--blue); font-style: normal; }

.theme { margin-top: 30px; scroll-margin-top: 64px; }
.theme-head { display: flex; align-items: center; gap: 10px; padding-bottom: 10px; margin-bottom: 14px; border-bottom: 3px solid var(--blue); }
.theme-head h3 { margin: 0; font-size: 19px; color: var(--blue); }
.theme-icon { font-size: 22px; }
.theme-count { margin-left: auto; font-size: 12.5px; font-weight: 700; color: var(--muted); }
.theme--omni .theme-head { border-bottom-color: var(--orange); }
.theme--omni .theme-head h3 { color: var(--orange); }
.cards { display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr)); gap: 14px; }
.card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; overflow: hidden; display: flex; flex-direction: column; border-top: 4px solid var(--nc); transition: transform .15s, box-shadow .15s; }
.card:hover { transform: translateY(-2px); box-shadow: 0 10px 24px rgba(1, 25, 90, .10); }
.card--alto { border-top-color: var(--alto); }
.card--medio { border-top-color: var(--medio); }
.card--baixo { border-top-color: var(--baixo); }
.card-body { padding: 16px 18px 18px; display: flex; flex-direction: column; flex: 1; }
.card h4 { font-size: 15.5px; line-height: 1.35; margin: 0 0 4px; }
.card h4 a { text-decoration: none; }
.card h4 a:hover { color: var(--blue); text-decoration: underline; }
.impact-text { font-size: 13.5px; color: var(--ink-2); margin: 0; }

footer { margin-top: 56px; padding: 28px 0 40px; border-top: 1px solid var(--line); font-size: 12.5px; color: var(--muted); }

@media (max-width: 900px) {
  .kpis { grid-template-columns: repeat(2, 1fr); }
  .overview { grid-template-columns: 1fr; }
  .highlight { grid-template-columns: 1fr; }
  .visual--big { height: 150px; min-height: 0; }
}
@media (max-width: 560px) {
  .wrap { padding: 0 16px; }
  .hero h1 { font-size: 34px; }
  .kpi-value { font-size: 38px; }
  .bar-row { grid-template-columns: 1fr 28px; }
  .bar-label { grid-column: 1 / -1; }
  .cards { grid-template-columns: 1fr; }
}

/* Botões flutuantes: alternam entre versão completa e resumida */
.toolbar { position: fixed; top: 14px; right: 14px; z-index: 50; display: flex; gap: 8px; }
.tbtn { font: inherit; font-size: 13.5px; font-weight: 700; border: none; cursor: pointer; padding: 10px 18px; border-radius: 24px; box-shadow: 0 6px 18px rgba(1, 25, 90, .25); background: var(--orange); color: #fff; }
.tbtn:hover { filter: brightness(1.08); }
.tbtn--ghost { background: #fff; color: var(--blue); }
.resumo, .only-resumo { display: none; }
body.modo-resumo .full, body.modo-resumo .only-full { display: none; }
body.modo-resumo .resumo { display: block; }
body.modo-resumo .only-resumo { display: inline-block; }
body.modo-resumo { background: #DCDFE6; }

/* Versão resumida: uma folha A4 retrato */
.resumo { padding: 28px 0 40px; }
.sheet { page: resumo; width: 210mm; height: 297mm; margin: 0 auto; background: #fff; box-shadow: 0 16px 50px rgba(1, 25, 90, .22); overflow: hidden; display: flex; flex-direction: column; }
.sheet-head { background: radial-gradient(circle at 90% 10%, #1c4bb0 0%, var(--blue) 50%, #01195a 100%); color: #fff; padding: 5.5mm 12mm 5mm; position: relative; overflow: hidden; }
.sheet-head::after { content: ""; position: absolute; right: -70px; top: -90px; width: 230px; height: 230px; border-radius: 50%; border: 34px solid var(--orange); }
.sheet-top { display: flex; align-items: center; gap: 12px; margin-bottom: 3mm; }
.sheet-logo svg { height: 20px; width: auto; display: block; }
.sheet-head .eyebrow { font-size: 9.5px; padding: 3px 9px; margin: 0; }
.sheet-head h1 { font-size: 23px; margin: 0 0 3px; letter-spacing: -0.3px; }
.sheet-head p { margin: 0; font-size: 12.5px; color: #C9D6EF; }
.sheet-body { padding: 5mm 12mm 3mm; display: flex; flex-direction: column; gap: 2.5mm; flex: 1; min-height: 0; }
.sheet-section-title { display: flex; align-items: baseline; justify-content: space-between; border-bottom: 2px solid var(--blue); padding-bottom: 1.5mm; }
.sheet-section-title h2 { margin: 0; font-size: 15px; color: var(--blue); }
.sheet-section-title span { font-size: 10.5px; color: var(--muted); font-weight: 600; }
/* Top 6: as 2 primeiras em destaque (foto grande), as outras 4 com miniatura */
.top6 { flex: 1; min-height: 0; display: grid; grid-template-columns: 1fr 1fr; grid-template-rows: 1.5fr 1fr 1fr; gap: 3.5mm; }
.top6--1 { grid-template-columns: 1fr; grid-template-rows: 1fr; }
.top6--2 { grid-template-rows: 1fr; }
.top6--3, .top6--4 { grid-template-rows: 1.5fr 1fr; }
.top6--3 .rcard:last-child, .top6--5 .rcard:last-child { grid-column: 1 / -1; }  /* ímpar: último ocupa a linha toda */
.rcard { background: #fff; border: 1px solid var(--line); border-bottom: 4px solid var(--nc); border-radius: 12px; overflow: hidden; display: flex; min-height: 0; box-shadow: 0 4px 14px rgba(1, 25, 90, .08); transition: transform .15s, box-shadow .15s; }
.rcard:hover { transform: translateY(-2px); box-shadow: 0 10px 24px rgba(1, 25, 90, .16); }
.rcard--alto { border-bottom-color: var(--alto); }
.rcard--medio { border-bottom-color: var(--medio); }
.rcard--baixo { border-bottom-color: var(--baixo); }
.rcard--featured { flex-direction: column; }
.rcard--featured .rcard-media { flex: 1; min-height: 30mm; }
.rcard--compact { flex-direction: row; }
.rcard--compact .rcard-media { width: 40%; flex-shrink: 0; }
.rcard-media { position: relative; overflow: hidden; background: var(--blue); }
.rcard .visual { position: absolute; inset: 0; height: auto; min-height: 0; }
.rcard .visual-icon { font-size: 34px; }
.rcard--featured .visual-icon { font-size: 52px; }
.rcard--compact .visual-cat { display: none; }
.rcard .visual-tag { display: none; }
.rank { position: absolute; top: 8px; left: 8px; z-index: 2; width: 30px; height: 30px; border-radius: 50%; background: var(--orange); color: #fff; font-weight: 800; font-size: 15px; display: flex; align-items: center; justify-content: center; box-shadow: 0 3px 10px rgba(0, 0, 0, .35); }
.rcard--featured .rank { width: 40px; height: 40px; font-size: 20px; top: 10px; left: 10px; }
.rcard-body { padding: 3mm 3.5mm 3mm; display: flex; flex-direction: column; gap: 1.6mm; min-height: 0; min-width: 0; }
.rcard-tags { display: flex; align-items: center; gap: 6px; flex-wrap: wrap; }
.rcard-tags .badge { font-size: 9.5px; padding: 2px 8px; }
.rcard-cat { font-size: 9px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: .4px; }
.rcard-title { margin: 0; font-size: 14px; line-height: 1.3; display: -webkit-box; -webkit-line-clamp: 5; -webkit-box-orient: vertical; overflow: hidden; }
.rcard--featured .rcard-title { font-size: 16px; -webkit-line-clamp: 3; }
.rcard-title a { color: var(--ink); text-decoration: none; }
.rcard-title a:hover { color: var(--blue); text-decoration: underline; }
.rcard-comment { margin: 0; font-size: 10px; font-style: italic; color: var(--blue); display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden; }
.rcard-meta { margin: auto 0 0; font-size: 9.5px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sheet-foot { padding: 3mm 12mm 5mm; font-size: 9px; color: var(--muted); border-top: 1px solid var(--line); }

@page resumo { size: A4 portrait; margin: 0; }
@media print {
  body { background: #fff; }
  .nav, .toolbar { display: none !important; }
  .card, .highlight, .panel, .kpi { break-inside: avoid; box-shadow: none; }
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  body.modo-resumo { background: #fff; }
  body.modo-resumo .resumo { padding: 0; }
  body.modo-resumo .sheet { box-shadow: none; margin: 0; }
}
"""

SCRIPT = """
<script>
(function () {
  function setMode(resumo) {
    document.body.classList.toggle('modo-resumo', resumo);
    history.replaceState(null, '', resumo ? '#resumo' : location.pathname + location.search);
    window.scrollTo(0, 0);
  }
  document.getElementById('btn-resumo').addEventListener('click', function () { setMode(true); });
  document.getElementById('btn-completo').addEventListener('click', function () { setMode(false); });
  document.getElementById('btn-print').addEventListener('click', function () { window.print(); });
  if (location.hash === '#resumo') setMode(true);
})();
</script>
"""


def build_html(items: list[dict], total: int, source_file: Path) -> str:
    now = datetime.now()
    dates = [i["date"] for i in items if isinstance(i["date"], datetime)]
    period = f"{date_label(min(dates))} a {date_label(max(dates))}" if dates else ""
    highlights = [i for i in items if i["impact"] == "Alto"]
    highlight_note = "Notícias marcadas pela equipe como de alto impacto para a Omni."
    if not highlights:
        highlights = items[:3]
        highlight_note = "Nenhuma notícia foi marcada como alto impacto; abaixo, as três primeiras selecionadas."
    themes_html, nav_html = render_by_category(items)

    return f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Radar Executivo Omni</title>
<style>{CSS}</style>
</head>
<body>
  <div class="toolbar">
    <button type="button" class="tbtn only-full" id="btn-resumo">📄 Versão resumida (1 página)</button>
    <button type="button" class="tbtn only-resumo" id="btn-print">🖨️ Imprimir / salvar PDF</button>
    <button type="button" class="tbtn tbtn--ghost only-resumo" id="btn-completo">← Versão completa</button>
  </div>
{render_resumo(items, total, period, now)}
  <div class="full">
  <header class="hero">
    <div class="wrap">
      <div class="hero-logo">{read_logo("logo-omni-white.svg")}</div>
      <span class="eyebrow">RADAR EXECUTIVO · DIRETORIA</span>
      <h1>Radar de Mercado Omni</h1>
      <p>As notícias que a equipe considerou relevantes e o que cada uma significa para o negócio.</p>
      <div class="when">{now.day:02d} {MONTHS[now.month - 1]} {now.year}{f" · notícias de {period}" if period else ""}</div>
    </div>
  </header>

  <div class="wrap"><div class="kpis">{render_kpis(items, total)}</div></div>

  <nav class="nav"><div class="wrap">
    <a href="#visao-geral">📊 Visão geral</a>
    <a href="#destaques">🔥 Destaques <span>{len(highlights)}</span></a>
    {nav_html}
  </div></nav>

  <main class="wrap">
    <section class="block" id="visao-geral">
      <div class="block-title"><h2>Visão geral</h2><p>Onde estão os sinais e quanto pesam</p></div>
      <div class="overview">
        <div class="panel"><h3>Nível de impacto</h3>{render_impact_meter(items)}</div>
        <div class="panel"><h3>Notícias por tema</h3>{render_category_bars(items)}</div>
      </div>
      <div class="panel areas"><h3>Áreas que devem acompanhar</h3><div class="chips">{render_areas(items)}</div></div>
    </section>

    <section class="block" id="destaques">
      <div class="block-title"><h2>🔥 Destaques</h2><p>{highlight_note}</p></div>
      <div class="highlights">{"".join(render_highlight(i) for i in highlights)}</div>
    </section>

    <section class="block">
      <div class="block-title"><h2>Todas as notícias selecionadas</h2><p>Organizadas por tema, da mais para a menos impactante</p></div>
      {themes_html}
    </section>
  </main>

  <footer><div class="wrap">
    Curadoria feita pela equipe a partir de notícias públicas (Google Notícias). Nível de impacto, textos, áreas e comentários
    foram revisados manualmente na planilha <strong>{esc(source_file.name)}</strong>. {len(items)} de {total} notícias analisadas foram aprovadas.
    Gerado em {now:%d/%m/%Y %H:%M}.
  </div></footer>
  </div>
{SCRIPT}
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser(description="Gera o HTML executivo da diretoria a partir da planilha revisada.")
    parser.add_argument("planilha", nargs="?", type=Path, help="Planilha revisada. Padrão: a mais recente da pasta 'planilhas'.")
    parser.add_argument("--output", type=Path, help="Arquivo HTML. Padrão: html_diretoria/radar_diretoria_AAAA-MM-DD_HHMM.html")
    args = parser.parse_args()

    source = args.planilha or latest_spreadsheet()
    if not source.exists():
        sys.exit(f"Planilha não encontrada: {source}")
    try:
        items, total = read_selected(source)
    except PermissionError:
        sys.exit(f"Não consegui abrir '{source.name}'. Feche a planilha no Excel e tente de novo.")
    if not items:
        sys.exit(f"Nenhuma notícia com STATUS = SIM em '{source.name}'. Marque SIM na última coluna, salve e rode de novo.")

    chosen = sum(1 for item in items if item["top"] is not None)
    if chosen > RESUMO_TOP:
        print(f"Aviso: {chosen} notícias marcadas na coluna TOP 6; a página executiva mostra só as {RESUMO_TOP} primeiras posições.")
    elif chosen == 0:
        print("Coluna TOP 6 vazia: o Top 6 foi montado automaticamente (Alto impacto primeiro, depois maior Score).")

    # Planilhas antigas (ou fotos que falharam no passo 1): tenta buscar agora só para as selecionadas.
    missing = [item for item in items if not item["image"] and item["link"]]
    if missing:
        print(f"Buscando foto de {len(missing)} notícia(s) selecionada(s) sem imagem...")
        for item, (real_url, image) in zip(missing, enrich_many([item["link"] for item in missing])):
            item["link"] = real_url
            item["image"] = image

    with_image = [item for item in items if item["image"]]
    if with_image:
        print(f"Baixando e embutindo {len(with_image)} foto(s) no HTML...")
        for item, embedded in zip(with_image, embed_many([item["image"] for item in with_image])):
            item["image"] = embedded

    output = args.output or Path("html_diretoria") / f"radar_diretoria_{datetime.now():%Y-%m-%d_%H%M}.html"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(build_html(items, total, source), encoding="utf-8")
    print(f"Planilha lida: {source} ({len(items)} de {total} notícias com SIM)")
    print(output.resolve())


if __name__ == "__main__":
    main()
