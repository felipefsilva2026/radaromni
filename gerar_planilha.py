"""Etapa 1: pesquisa as notícias e gera a planilha de curadoria (Excel).

A equipe revisa a planilha, ajusta as colunas editáveis (impacto, texto de
impacto, área responsável, comentário) e marca SIM/NÃO na coluna STATUS.
Depois, gerar_html_diretoria.py lê a planilha salva e monta o HTML executivo
só com as notícias marcadas como SIM.
"""

import argparse
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from imagens import enrich_many
from motor_pesquisa import HIGH_IMPACT_TERMS, MEDIUM_IMPACT_TERMS, MIN_SOURCE_NOTE, collect, parse_date

# Mesmas 39 buscas do jornal diário (projeto PESQUISA DE MERCADO).
SEARCHES = {
    "Omni": [
        "Omni Financeira",
        "Banco Omni",
        "Omni crédito",
        "Omni seguros",
        "Omni veículos financiamento",
        "Omni1",
        "Finsol",
        "Grupo Omni",
        "Omni Empresas",
        "Omni&Co",
    ],
    "Concorrentes": [
        '"Banco BV" financiamento veículos',
        '"Banco PAN" financiamento veículos',
        '"Santander Financiamentos" veículos',
        '"Itaú" financiamento veículos',
        '"Bradesco Financiamentos" veículos',
    ],
    "Selic e macroeconomia": [
        "Selic Copom crédito",
        "juros inadimplência pessoa física",
        "economia brasileira crédito 2026",
        "taxa Selic financiamento veículos",
    ],
    "Mercado automotivo": [
        "mercado automotivo brasileiro veículos usados",
        "Fenabrave vendas veículos usados",
        "vendas carros seminovos Brasil",
        "preço carro usado Brasil",
    ],
    "Seguros": [
        "seguro auto Brasil mercado",
        "SUSEP seguro auto 2026",
        "insurtech seguro veículos Brasil",
        "bancassurance seguros Brasil",
    ],
    "Financiamento e crédito": [
        "financiamento de veículos crédito Brasil",
        "CDC veículos carteira de crédito",
        "funding financeiras crédito",
        "inadimplência financiamento veículos",
    ],
    "Bancos e parceiros": [
        "bancos concessionárias parceria financiamento",
        "fintech parceria concessionárias veículos",
        "marketplace automotivo financiamento",
        "parceria seguradora banco veículos",
    ],
    "Riscos e crises": [
        "crise crédito Brasil financeiras",
        "fraude financiamento veículos Brasil",
        "regulação financeira crédito veículos",
        "desemprego inadimplência crédito",
    ],
}

# Sugestão de área que deve olhar cada tipo de notícia (editável na planilha).
CATEGORY_AREA = {
    "Omni": "Marketing / Comunicação",
    "Concorrentes": "Comercial",
    "Selic e macroeconomia": "Tesouraria / Funding",
    "Mercado automotivo": "Comercial / Originação",
    "Seguros": "Seguros / Produtos",
    "Financiamento e crédito": "Crédito / Risco",
    "Bancos e parceiros": "Comercial / Parcerias",
    "Riscos e crises": "Compliance / Prevenção a Fraudes",
}
DEFAULT_AREA = "Diretoria"

IMPACT_LEVELS = ("Alto", "Médio", "Baixo")
STATUS_VALUES = ("SIM", "NÃO")

# (cabeçalho, largura). Todas as colunas são editáveis. A ordem aqui é a ordem das colunas
# no Excel; STATUS fica sempre por último. gerar_html_diretoria.py lê pelo nome do cabeçalho.
COLUMNS = [
    ("ID", 6),
    ("Data", 11),
    ("Categoria", 20),
    ("Título", 60),
    ("Resumo", 50),
    ("Fonte", 20),
    ("Nota da fonte", 9),
    ("Score", 8),
    ("Prioridade sugerida", 18),
    ("Link", 30),
    ("Imagem", 30),
    ("Impacto", 11),
    ("Impacto para a Omni", 60),
    ("Área responsável", 24),
    ("Comentário da equipe", 40),
    ("TOP 6", 8),
    ("STATUS", 10),
]
DECISION_COLUMNS = ("TOP 6", "STATUS")  # cabeçalho laranja: colunas de decisão da equipe
TOP_POSITIONS = ("1", "2", "3", "4", "5", "6")

BLUE = "013188"
ORANGE = "E95814"
HEADER_FILL = PatternFill("solid", fgColor=BLUE)
HEADER_STATUS_FILL = PatternFill("solid", fgColor=ORANGE)
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def matched_terms(item: dict, terms: tuple[str, ...]) -> list[str]:
    text = f"{item['title']} {item['keyword']} {item['summary']}".lower()
    return [term for term in terms if term in text]


def suggest_impact_text(item: dict) -> str:
    """Sugestão de texto de impacto, por regra fixa (sem IA). A equipe pode reescrever."""
    base = item.get("why_it_matters", "")
    high = matched_terms(item, HIGH_IMPACT_TERMS)
    medium = matched_terms(item, MEDIUM_IMPACT_TERMS)
    if high:
        return f"{base} Envolve {', '.join(high)} — sugere-se avaliação imediata pela área responsável."
    if medium:
        return f"{base} Envolve {', '.join(medium)} — acompanhar a evolução do tema."
    return f"{base} Contexto de mercado, sem ação imediata sugerida."


def run_search(days: int, limit: int, min_credibility: int) -> list[dict]:
    keywords = [keyword for terms in SEARCHES.values() for keyword in terms]
    results = collect(keywords, days, limit, min_credibility)
    categories = {keyword: category for category, terms in SEARCHES.items() for keyword in terms}
    for item in results:
        item["category"] = categories.get(item["keyword"], "Outros")

    print(f"Buscando endereço real e foto de {len(results)} notícia(s)...")
    for item, (real_url, image) in zip(results, enrich_many([item["link"] for item in results])):
        item["link"] = real_url
        item["image_url"] = item.get("image_url") or image
    print(f"{sum(1 for item in results if item['image_url'])} de {len(results)} notícia(s) com foto.")
    return results


def write_instructions(wb: Workbook, generated_at: str, total: int) -> None:
    ws = wb.create_sheet("Como usar")
    ws.column_dimensions["A"].width = 110
    lines = [
        ("Curadoria da pesquisa de mercado — Omni", Font(bold=True, size=14, color=BLUE)),
        (f"Gerada em {generated_at} · {total} notícia(s) para revisar", Font(italic=True, color="676D79")),
        ("", None),
        ("1. Na aba 'Pesquisa', leia cada notícia (a coluna Link abre a matéria original).", None),
        ("2. TODAS as colunas podem ser editadas. O que estiver escrito é o que vai para o HTML:", None),
        ("     • Data, Categoria, Título, Fonte e Link: aparecem em cada notícia. Categoria tem lista suspensa, mas aceita um tema novo.", None),
        ("     • Imagem: endereço da foto. Para trocar, cole o endereço de outra imagem; para tirar, apague.", None),
        ("     • Impacto: Alto, Médio ou Baixo (lista suspensa). Vem com sugestão.", None),
        ("     • Impacto para a Omni: texto que aparece para a diretoria. Vem com sugestão; pode reescrever à vontade.", None),
        ("     • Área responsável: quem deve olhar o tema. Vem com sugestão.", None),
        ("     • Comentário da equipe: opcional. Se preenchido, aparece em destaque no HTML.", None),
        ("     • Score: define a ordem dentro de cada nível de impacto (maior primeiro).", None),
        ("     • ID, Resumo, Nota da fonte e Prioridade sugerida: referência para a revisão; não aparecem no HTML.", None),
        ("3. Na coluna TOP 6, escolha as notícias da página executiva (versão resumida): digite a posição, de 1 a 6.", None),
        ("     • Pode marcar menos de 6. Marcar TOP 6 já leva a notícia para a diretoria, mesmo sem SIM no STATUS.", None),
        ("     • Se ninguém for marcado, o Top 6 é montado sozinho (Alto impacto primeiro, depois maior Score).", None),
        ("4. Na última coluna, STATUS, marque SIM para o que deve ir para a diretoria e NÃO para o que não é relevante.", None),
        ("     • Só as linhas com SIM entram no HTML. Linha em branco = não entra.", None),
        ("5. Salve a planilha (pode renomear e mandar para outras pessoas filtrarem).", None),
        ("6. Rode '2 - Gerar HTML diretoria.bat' (ou arraste a planilha em cima dele).", None),
        ("", None),
        ("O HTML sempre usa o que estiver preenchido na planilha — as sugestões só servem de ponto de partida.", Font(bold=True, color=ORANGE)),
        ("Não mude os nomes dos cabeçalhos da aba 'Pesquisa'; o resto (ordem, filtros, cores) pode mexer.", None),
    ]
    for row, (text, font) in enumerate(lines, start=1):
        cell = ws.cell(row=row, column=1, value=text)
        if font:
            cell.font = font


def build_workbook(results: list[dict], output_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Pesquisa"

    for col, (header, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_STATUS_FILL if header in DECISION_COLUMNS else HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
        ws.column_dimensions[cell.column_letter].width = width
    ws.row_dimensions[1].height = 32

    col_index = {header: i for i, (header, _) in enumerate(COLUMNS, start=1)}

    for row, item in enumerate(results, start=2):
        published = parse_date(item.get("published", ""))
        values = {
            "ID": row - 1,
            "Data": published.replace(tzinfo=None) if published else None,
            "Categoria": item.get("category", ""),
            "Título": item.get("title", ""),
            "Resumo": item.get("summary", "") if item.get("summary", "").lower() not in item.get("title", "").lower() else "",
            "Fonte": item.get("source", ""),
            "Nota da fonte": item.get("source_note", ""),
            "Score": item.get("final_score", ""),
            "Prioridade sugerida": item.get("priority", ""),
            "Link": item.get("link", ""),
            "Imagem": item.get("image_url", ""),
            "Impacto": item.get("impact", "Baixo"),
            "Impacto para a Omni": suggest_impact_text(item),
            "Área responsável": CATEGORY_AREA.get(item.get("category", ""), DEFAULT_AREA),
            "Comentário da equipe": "",
            "TOP 6": None,
            "STATUS": "",
        }
        for header, value in values.items():
            col = col_index[header]
            cell = ws.cell(row=row, column=col, value=value)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="top", wrap_text=header in ("Título", "Resumo", "Impacto para a Omni", "Comentário da equipe"))
        date_cell = ws.cell(row=row, column=col_index["Data"])
        date_cell.number_format = "DD/MM/YYYY"
        link_cell = ws.cell(row=row, column=col_index["Link"])
        if values["Link"]:
            link_cell.hyperlink = values["Link"]
            link_cell.font = Font(color="2453A9", underline="single")
        for header in DECISION_COLUMNS:
            decision_cell = ws.cell(row=row, column=col_index[header])
            decision_cell.alignment = Alignment(horizontal="center", vertical="top")
            decision_cell.font = Font(bold=True)

    last_row = max(len(results) + 1, 2)
    last_col_letter = ws.cell(row=1, column=len(COLUMNS)).column_letter
    impact_letter = ws.cell(row=1, column=col_index["Impacto"]).column_letter
    status_letter = ws.cell(row=1, column=col_index["STATUS"]).column_letter

    impact_dv = DataValidation(type="list", formula1=f'"{",".join(IMPACT_LEVELS)}"', allow_blank=True)
    impact_dv.add(f"{impact_letter}2:{impact_letter}{last_row}")
    status_dv = DataValidation(type="list", formula1=f'"{",".join(STATUS_VALUES)}"', allow_blank=True)
    status_dv.add(f"{status_letter}2:{status_letter}{last_row}")
    # Categoria: lista suspensa só como atalho — aceita digitar um tema novo sem erro.
    category_letter = ws.cell(row=1, column=col_index["Categoria"]).column_letter
    category_dv = DataValidation(type="list", formula1=f'"{",".join(SEARCHES)}"', allow_blank=True, showErrorMessage=False)
    category_dv.add(f"{category_letter}2:{category_letter}{last_row}")
    # TOP 6: posição 1 a 6 na página executiva (lista suspensa; aceita também "X" sem ordem).
    top_letter = ws.cell(row=1, column=col_index["TOP 6"]).column_letter
    top_dv = DataValidation(type="list", formula1=f'"{",".join(TOP_POSITIONS)}"', allow_blank=True, showErrorMessage=False)
    top_dv.add(f"{top_letter}2:{top_letter}{last_row}")
    ws.add_data_validation(impact_dv)
    ws.add_data_validation(status_dv)
    ws.add_data_validation(category_dv)
    ws.add_data_validation(top_dv)
    top_range = f"{top_letter}2:{top_letter}{last_row}"
    ws.conditional_formatting.add(top_range, FormulaRule(formula=[f'LEN({top_letter}2)>0'], fill=PatternFill("solid", fgColor="FDEBE3"), font=Font(bold=True, color=ORANGE)))

    status_range = f"{status_letter}2:{status_letter}{last_row}"
    ws.conditional_formatting.add(status_range, FormulaRule(formula=[f'UPPER({status_letter}2)="SIM"'], fill=PatternFill("solid", fgColor="C6EFCE"), font=Font(bold=True, color="006100")))
    ws.conditional_formatting.add(status_range, FormulaRule(formula=[f'OR(UPPER({status_letter}2)="NÃO",UPPER({status_letter}2)="NAO")'], fill=PatternFill("solid", fgColor="FFC7CE"), font=Font(bold=True, color="9C0006")))
    impact_range = f"{impact_letter}2:{impact_letter}{last_row}"
    ws.conditional_formatting.add(impact_range, FormulaRule(formula=[f'{impact_letter}2="Alto"'], font=Font(bold=True, color="C43B2C")))

    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    write_instructions(wb, generated_at, len(results))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def main():
    parser = argparse.ArgumentParser(description="Pesquisa as notícias e gera a planilha de curadoria.")
    parser.add_argument("--days", type=int, default=30, help="Quantidade de dias a considerar (padrão: 30).")
    parser.add_argument("--limit", type=int, default=150, help="Número máximo de notícias (padrão: 150).")
    parser.add_argument("--min-credibility", type=int, default=MIN_SOURCE_NOTE, help=f"Nota mínima da fonte, de 1 a 10 (padrão: {MIN_SOURCE_NOTE}).")
    parser.add_argument("--output", type=Path, help="Caminho da planilha. Padrão: planilhas/curadoria_AAAA-MM-DD_HHMM.xlsx")
    args = parser.parse_args()

    output = args.output or Path("planilhas") / f"curadoria_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    results = run_search(max(args.days, 1), max(args.limit, 1), args.min_credibility)
    build_workbook(results, output)
    print(f"Planilha criada com {len(results)} notícia(s): {output}")
    print(output.resolve())


if __name__ == "__main__":
    main()
