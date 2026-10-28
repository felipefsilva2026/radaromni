"""Pesquisa notícias públicas sobre a Omni usando o RSS do Google Notícias.

Cópia independente do motor do projeto "PESQUISA DE MERCADO" (pesquisa_mercado.py),
para que mudanças em um projeto não quebrem o outro.
"""

import argparse
import csv
import html
import json
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import quote_plus
from urllib.request import Request, urlopen
from xml.etree import ElementTree


DEFAULT_KEYWORDS = [
    '"Omni S.A."',
    '"Omni financeira"',
    '"Omni crédito"',
    '"Omni financiamento de veículos"',
    '"Banco Omni"',
]
RSS_URL = "https://news.google.com/rss/search?q={query}&hl=pt-BR&gl=BR&ceid=BR:pt-419"

CATEGORY_RULES = {
    "Omni": ("omni", "banco omni", "finsol", "crediário omni", "crediario omni", "omni1", "grupo omni", "omni&co"),
    "Concorrentes": ("banco bv", "banco pan", "santander financiamentos", "itaú", "bradesco financiamentos"),
    "Selic e macroeconomia": ("selic", "copom", "juros", "economia", "desemprego"),
    "Mercado automotivo": ("veículo", "veiculos", "carro", "automotivo", "seminovo", "feirão", "fenabrave"),
    "Seguros": ("seguro", "susep", "insurtech", "bancassurance"),
    "Financiamento e crédito": ("financiamento", "crédito", "credito", "inadimplência", "funding", "cdc"),
    "Bancos e parceiros": ("parceria", "parceiro", "concessionária", "concessionaria", "fintech", "marketplace"),
    "Riscos e crises": ("fraude", "regulação", "regulacao", "crise", "risco", "documentação", "documentacao"),
}

# Termos que definem o impacto potencial (Alto/Médio/Baixo) e o sinal de
# acompanhamento (Atenção/Monitorar/Contexto) de cada notícia. Avaliados nesta
# ordem: Alto vence Médio, que vence o padrão Baixo.
HIGH_IMPACT_TERMS = ("crise", "fraude", "inadimplência", "regulação", "susep", "selic", "licença")
MEDIUM_IMPACT_TERMS = ("concorr", "parceria", "expansão", "funding", "juros", "financiamento", "crédito", "seguro")

# Texto fixo de "por que importa" exibido em cada notícia, um por categoria.
CATEGORY_REASONS = {
    "Omni": "Pode afetar posicionamento, reputação, originação ou relacionamento com clientes e parceiros.",
    "Concorrentes": "Pode alterar preço, aprovação, canais de distribuição e disputa por parceiros.",
    "Selic e macroeconomia": "Pode alterar custo de funding, demanda, margem e inadimplência.",
    "Mercado automotivo": "Pode indicar mudança no volume de veículos, preços e oportunidades de originação.",
    "Seguros": "Pode abrir oportunidades de venda combinada e aumentar a disputa no pós-venda.",
    "Financiamento e crédito": "Pode afetar apetite de risco, carteira, funding e conversão de propostas.",
    "Bancos e parceiros": "Pode criar ou enfraquecer canais de distribuição e alianças comerciais.",
    "Riscos e crises": "Pode exigir resposta de compliance, crédito, prevenção a perdas ou continuidade.",
}
DEFAULT_REASON = "Ajuda a acompanhar o ambiente competitivo e econômico."

# Nota de confiabilidade editorial de cada fonte (1 a 10), definida pela Omni.
# Fontes que não aparecem aqui recebem DEFAULT_SOURCE_NOTE (tratadas como baixa
# confiabilidade até serem avaliadas e cadastradas manualmente).
SOURCE_CREDIBILITY = {
    # Nota 10 - Máxima confiabilidade
    "estadão": 10, "estadao": 10,
    "valor econômico": 10, "valor economico": 10, "valor.globo.com": 10,
    "folha de s.paulo": 10, "folha de s. paulo": 10, "folha de sao paulo": 10,
    # Nota 9 - Muito alta confiabilidade
    "o globo": 9,
    "bbc": 9, "bbc brasil": 9, "bbc news brasil": 9,
    "cnn brasil": 9,
    "poder360": 9,
    "consultor jurídico": 9, "consultor juridico": 9, "conjur": 9,
    "senado": 9, "senado notícias": 9, "senado noticias": 9,
    "portal da câmara dos deputados": 9, "portal da camara dos deputados": 9, "câmara dos deputados": 9, "camara dos deputados": 9,
    "gov.br": 9,
    # Nota 8 - Alta confiabilidade
    "infomoney": 8,
    "exame": 8,
    "autoesporte": 8,
    "investnews": 8,
    "serasa": 8,
    "safra": 8,
    "forbes brasil": 8, "forbes": 8,
    # Nota 7 - Boa confiabilidade
    "times brasil | cnbc": 7, "times brasil": 7, "cnbc": 7,
    "money times": 7,
    "portal contábeis": 7, "portal contabeis": 7, "contabeis.com.br": 7,
    "finsiders brasil": 7, "finsiders": 7,
    "cqcs": 7,
    "autoindústria": 7, "autoindustria": 7,
    "canal rural": 7,
    "seu dinheiro": 7,
    "istoé dinheiro": 7, "istoe dinheiro": 7,
    "migalhas": 7,
    # Nota 6 - Confiabilidade moderada
    "terra": 6,
    "uol economia": 6, "uol": 6,
    "correio braziliense": 6,
    "o tempo": 6,
    "diário de pernambuco": 6, "diario de pernambuco": 6,
    "nd mais": 6, "ndmais": 6,
    "metrópoles": 6, "metropoles": 6,
    "cartacapital": 6,
    "vrum": 6,
    "spacemoney": 6,
    "capital reset": 6,
    "movimento econômico": 6, "movimento economico": 6,
    "tribuna do norte": 6,
    # Nota 5 - Confiabilidade média
    "xp investimentos": 5, "xp": 5,
    "mundo rh": 5,
    "mundocoop": 5,
    "jornal dci": 5,
    "economic news brasil": 5,
    # Nota 4 - Baixa prioridade (ainda considerada, mas com peso reduzido)
    "bra1.com.br": 4,
    "seu crédito digital": 4, "seu credito digital": 4,
    "acionista.com.br": 4,
    "advfn": 4,
    "agenda do poder": 4,
    "rede brasil de televisão": 4, "rede brasil de televisao": 4,
    "bahia.ba": 4,
    "radar digital brasília": 4, "radar digital brasilia": 4,
    "click petróleo e gás": 4, "click petroleo e gas": 4,
    "dol": 4,
    "tnh1": 4,
    "o dia": 4,
    "fast company brasil": 4,
}
DEFAULT_SOURCE_NOTE = 3  # fonte não cadastrada: tratada como nota 3 (baixa) até ser avaliada
MIN_SOURCE_NOTE = 4  # notícias de fontes com nota abaixo disso são descartadas por padrão

# Score de impacto (0-10) por nível de impacto já calculado em classify_item.
IMPACT_SCORE = {"Alto": 10, "Médio": 5, "Baixo": 2}

PRIORITY_LEVELS = (
    (16, "🔴 Crítico"),
    (12, "🟠 Muito importante"),
    (8, "🟡 Importante"),
    (4, "🔵 Acompanhar"),
    (0, "⚪ Baixa prioridade"),
)


def normalize_source(value: str) -> str:
    return re.sub(r"\s+", " ", (value or "").strip().lower())


def source_credibility_note(source: str) -> int:
    """Nota de confiabilidade (1-10) da fonte, por nome cadastrado ou aproximação."""
    normalized = normalize_source(source)
    if normalized in SOURCE_CREDIBILITY:
        return SOURCE_CREDIBILITY[normalized]
    for name, note in SOURCE_CREDIBILITY.items():
        if name in normalized or normalized in name:
            return note
    return DEFAULT_SOURCE_NOTE


def source_credibility_weight(note: int) -> float:
    """Nota 10 -> peso 2.0, nota 1 -> peso 0.2 (peso = nota * 0.2)."""
    return round(note * 0.2, 2)


def priority_label(score: float) -> str:
    for threshold, label in PRIORITY_LEVELS:
        if score >= threshold:
            return label
    return PRIORITY_LEVELS[-1][1]


def clean_text(value: str) -> str:
    value = html.unescape(value or "")
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", value)).strip()


def parse_date(value: str) -> datetime | None:
    try:
        return datetime.strptime(value, "%a, %d %b %Y %H:%M:%S %Z").replace(
            tzinfo=timezone.utc
        )
    except ValueError:
        return None


def classify_item(item: dict[str, str]) -> dict[str, str]:
    text = f"{item['title']} {item['keyword']} {item['summary']}".lower()
    category = item.get("category", "")
    if not category:
        category = next((name for name, terms in CATEGORY_RULES.items() if any(term in text for term in terms)), "Mercado")
    if any(term in text for term in HIGH_IMPACT_TERMS):
        impact, signal = "Alto", "Atenção"
    elif any(term in text for term in MEDIUM_IMPACT_TERMS):
        impact, signal = "Médio", "Monitorar"
    else:
        impact, signal = "Baixo", "Contexto"
    source_note = source_credibility_note(item.get("source", ""))
    source_weight = source_credibility_weight(source_note)
    impact_score = IMPACT_SCORE.get(impact, IMPACT_SCORE["Baixo"])
    final_score = round(impact_score * source_weight, 2)

    item["category"] = category
    item["impact"] = impact
    item["signal"] = signal
    item["why_it_matters"] = CATEGORY_REASONS.get(category, DEFAULT_REASON)
    item["source_note"] = source_note
    item["source_weight"] = source_weight
    item["impact_score"] = impact_score
    item["final_score"] = final_score
    item["priority"] = priority_label(final_score)
    return item


def fetch_keyword(keyword: str, timeout: int = 20) -> list[dict[str, str]]:
    url = RSS_URL.format(query=quote_plus(keyword))
    request = Request(url, headers={"User-Agent": "PesquisaMercadoOmni/1.0"})
    with urlopen(request, timeout=timeout) as response:
        root = ElementTree.fromstring(response.read())

    results = []
    for item in root.findall("./channel/item"):
        title = clean_text(item.findtext("title"))
        link = clean_text(item.findtext("link"))
        published = clean_text(item.findtext("pubDate"))
        description = clean_text(item.findtext("description"))
        source = clean_text(item.findtext("source"))
        image_url = ""
        for element in item.iter():
            if element.tag.endswith("content") or element.tag.endswith("thumbnail"):
                image_url = element.attrib.get("url", "")
                if image_url:
                    break
        if not image_url:
            image_match = re.search(r'<img[^>]+src=["\']([^"\']+)', item.findtext("description") or "", re.IGNORECASE)
            image_url = image_match.group(1) if image_match else ""
        results.append(
            {
                "keyword": keyword,
                "title": title,
                "source": source or "Não informado",
                "published": published,
                "link": link,
                "summary": description,
                "image_url": image_url,
            }
        )
    return results


def collect(keywords: list[str], days: int, limit: int, min_source_note: int = MIN_SOURCE_NOTE) -> list[dict[str, str]]:
    cutoff = datetime.now(timezone.utc) - timedelta(days=days)
    unique = {}
    for keyword in keywords:
        items = None
        for attempt in range(3):  # o Google às vezes derruba a conexão; tenta de novo antes de desistir
            try:
                items = fetch_keyword(keyword)
                break
            except Exception as error:
                if attempt == 2:
                    print(f"Aviso: não foi possível pesquisar '{keyword}': {error}", file=sys.stderr)
                else:
                    time.sleep(3 * (attempt + 1))
        if items is None:
            continue
        for item in items:
            date = parse_date(item["published"])
            if date and date < cutoff:
                continue
            classified = classify_item(item)
            if classified["source_note"] < min_source_note:
                continue
            unique[item["link"]] = classified

    results = list(unique.values())
    results.sort(
        key=lambda item: (
            item["final_score"],
            parse_date(item["published"]) or datetime.min.replace(tzinfo=timezone.utc),
        ),
        reverse=True,
    )
    return results[:limit]


def write_outputs(results: list[dict[str, str]], output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    generated_at = datetime.now().astimezone().strftime("%d/%m/%Y %H:%M")

    with (output_dir / "resultados.json").open("w", encoding="utf-8") as file:
        json.dump(results, file, ensure_ascii=False, indent=2)

    with (output_dir / "resultados.csv").open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.DictWriter(file, fieldnames=results[0].keys() if results else ["keyword", "title", "source", "published", "link", "summary"])
        writer.writeheader()
        writer.writerows(results)

    with (output_dir / "relatorio.md").open("w", encoding="utf-8") as file:
        file.write(f"# Pesquisa de mercado: Omni\n\nGerado em {generated_at}. {len(results)} resultado(s).\n\n")
        if not results:
            file.write("Nenhum resultado encontrado no período informado.\n")
            return
        for index, item in enumerate(results, start=1):
            file.write(f"## {index}. {item['title']}\n\n")
            file.write(f"- **Fonte:** {item['source']}\n")
            file.write(f"- **Data:** {item['published']}\n")
            file.write(f"- **Palavra-chave:** {item['keyword']}\n")
            file.write(f"- **Link:** {item['link']}\n\n")


def main() -> None:
    parser = argparse.ArgumentParser(description="Pesquisa notícias públicas sobre a Omni.")
    parser.add_argument("--keywords", nargs="+", default=DEFAULT_KEYWORDS, help="Palavras ou frases para pesquisar.")
    parser.add_argument("--days", type=int, default=30, help="Quantidade de dias a considerar (padrão: 30).")
    parser.add_argument("--limit", type=int, default=100, help="Número máximo de resultados (padrão: 100).")
    parser.add_argument("--output", type=Path, default=Path("resultado"), help="Pasta de saída.")
    parser.add_argument("--min-credibility", type=int, default=MIN_SOURCE_NOTE, help=f"Nota mínima de confiabilidade da fonte, de 1 a 10 (padrão: {MIN_SOURCE_NOTE}). Notícias de fontes com nota menor são descartadas.")
    args = parser.parse_args()
    results = collect(args.keywords, max(args.days, 1), max(args.limit, 1), args.min_credibility)
    write_outputs(results, args.output)
    print(f"Pesquisa concluída: {len(results)} resultado(s) em '{args.output}'.")


if __name__ == "__main__":
    main()