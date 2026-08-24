"""
Gera faturas (uma planilha .xlsx por hub) a partir da Pré-fatura, usando
o layout de "BETIM - BRMG02.xlsx" como template.

Regras de agregação e mapeamento entre abas foram validadas manualmente
contra a fatura BETIM - BRMG02.xlsx (ver Relatorio_Analise_Pre-Fatura.md).
Período e Line Haul não existem na Pré-fatura e precisam ser passados na
linha de comando.
"""
from __future__ import annotations

import argparse
import unicodedata
from pathlib import Path

import openpyxl
from num2words import num2words


def normalize_city(name: str) -> str:
    """Sem acento/maiúsculas: a aba '1. MELI' grafa cidades como 'SIMÕES FILHO'
    e a 'CT-es' como 'SIMOES FILHO', então o casamento direto falha."""
    n = unicodedata.normalize("NFKD", name.strip().upper())
    return "".join(c for c in n if not unicodedata.combining(c))


PRESTADOR = {
    "razao_social": "MULTICOM LOG TRANSPORTES LTDA",
    "inscr_est": 150140233110,
    "cnpj": "55.939.622/0001-71",
    "municipio": "Ribeirão Preto/SP",
    "endereco": "R CERQUEIRA CESAR, 1625 - CXPST 44",
    "cep": "14.025-120",
}

CLIENTE = {
    "razao_social": "EBAZAR.COM.BR. LTDA",
    "endereco": "AV DAS NACOES UNIDAS, 3000, BONFIM - PARTE A",
    "municipio": "OSASCO",
    "cnpj": "03.007.331/0001-41",
}

# Aba "4. Lançamento": bloco "FATURAMENTO MERCADO LIVRE" (K:O), não o bloco
# "FATURAMENTO LÖSUNG" (Q:U), que dá um valor de ICMS levemente diferente.
LANC_ORIGEM = 7
LANC_FRETE = 11
LANC_GRIS = 12
LANC_ICMS = 13
LANC_TOTAL = 15
LANC_DOC = 23
LANC_NUMERO = 24
LANC_FIRST_ROW = 7
LANC_LAST_ROW = 651

MELI_VEICULO = 6
MELI_ROTA = 7
MELI_ORIGEM = 8
MELI_CIDADE = 9
MELI_FIRST_ROW = 3
MELI_LAST_ROW = 647

CTES_NUMERO = 3
CTES_CIDADE_ORIGEM = 40
CTES_FIRST_ROW = 2


def format_brl(value: float) -> str:
    s = f"{value:,.2f}"
    s = s.replace(",", "§").replace(".", ",").replace("§", ".")
    return f"BRL {s}"


def valor_por_extenso(value: float) -> str:
    reais = int(value)
    centavos = round((value - reais) * 100)
    texto = f"{num2words(reais, lang='pt_BR')} reais e {num2words(centavos, lang='pt_BR')} centavos"
    return texto[0].upper() + texto[1:]


def load_meli_maps(ws):
    """Retorna (hub -> cidade, hub -> [(veiculo, nome_rota), ...] em ordem de 1a aparição)."""
    hub_to_city: dict[str, str] = {}
    hub_to_routes: dict[str, dict[tuple, None]] = {}
    for r in range(MELI_FIRST_ROW, MELI_LAST_ROW + 1):
        hub = ws.cell(row=r, column=MELI_ORIGEM).value
        cidade = ws.cell(row=r, column=MELI_CIDADE).value
        veiculo = ws.cell(row=r, column=MELI_VEICULO).value
        rota = ws.cell(row=r, column=MELI_ROTA).value
        if not hub:
            continue
        hub = str(hub).strip()
        if cidade and hub not in hub_to_city:
            hub_to_city[hub] = str(cidade).strip()
        if veiculo and rota:
            hub_to_routes.setdefault(hub, {})[(str(veiculo).strip(), str(rota).strip())] = None
    return hub_to_city, {h: list(d.keys()) for h, d in hub_to_routes.items()}


def load_lancamento_groups(ws):
    """Agrupa a aba 4. Lançamento por hub, só linhas Doc.=='Fatura'."""
    groups: dict[str, dict] = {}
    for r in range(LANC_FIRST_ROW, LANC_LAST_ROW + 1):
        hub = ws.cell(row=r, column=LANC_ORIGEM).value
        doc = ws.cell(row=r, column=LANC_DOC).value
        if not hub or doc != "Fatura":
            continue
        hub = str(hub).strip()
        g = groups.setdefault(hub, {"subtotal": 0.0, "icms": 0.0, "total": 0.0, "numeros": set(), "linhas": 0})
        frete = ws.cell(row=r, column=LANC_FRETE).value or 0
        gris = ws.cell(row=r, column=LANC_GRIS).value or 0
        icms = ws.cell(row=r, column=LANC_ICMS).value or 0
        total = ws.cell(row=r, column=LANC_TOTAL).value or 0
        numero = ws.cell(row=r, column=LANC_NUMERO).value
        g["subtotal"] += frete + gris
        g["icms"] += icms
        g["total"] += total
        g["linhas"] += 1
        if numero is not None:
            g["numeros"].add(numero)
    return groups


def load_ctes_by_city(ws):
    """Indexado pela cidade normalizada (sem acento) para não depender de
    grafia idêntica entre a aba CT-es e a aba 1. MELI."""
    by_city: dict[str, list[int]] = {}
    for r in range(CTES_FIRST_ROW, ws.max_row + 1):
        cidade = ws.cell(row=r, column=CTES_CIDADE_ORIGEM).value
        numero = ws.cell(row=r, column=CTES_NUMERO).value
        if cidade and numero is not None:
            by_city.setdefault(normalize_city(str(cidade)), []).append(numero)
    for city in by_city:
        by_city[city].sort()
    return by_city


def build_fatura(template_path: Path, out_path: Path, *, hub: str, cidade: str,
                  emissao, periodo: str, line_haul: str,
                  group: dict, ctes: list[int], rotas: list[tuple]):
    wb = openpyxl.load_workbook(template_path)
    ws = wb.worksheets[0]

    ws["C4"] = PRESTADOR["razao_social"]
    ws["F4"] = PRESTADOR["inscr_est"]
    ws["C5"] = PRESTADOR["cnpj"]
    ws["F5"] = PRESTADOR["municipio"]
    ws["C6"] = PRESTADOR["endereco"]
    ws["F6"] = PRESTADOR["cep"]

    ws["C8"] = CLIENTE["razao_social"]
    ws["C9"] = CLIENTE["endereco"]
    ws["C10"] = CLIENTE["municipio"]
    ws["C11"] = CLIENTE["cnpj"]

    ws["C14"] = emissao  # B14 mantém a fórmula '=90+C14' do template
    numeros_fatura = sorted(str(n) for n in group["numeros"])
    if len(numeros_fatura) != 1:
        print(f"  [!] hub {hub}: número de fatura não é único na coluna X: {numeros_fatura!r}, "
              f"usando o primeiro.")
    ws["D14"] = numeros_fatura[0] if numeros_fatura else ""
    total = round(group["total"], 2)
    ws["F14"] = format_brl(total)

    ws["B17"] = f"Transporte CTE - {periodo} - Line Haul N. {line_haul}"
    ws["B20"] = valor_por_extenso(total)

    ctes_str = " / ".join(str(n) for n in ctes)
    rotas_str = "\n".join(f"{veic.upper()} - {rota}" for veic, rota in rotas)
    ws["B22"] = (
        "Pela sua PRESTAÇÃO DE SERVIÇO conforme Dacte(s) abaixo. Na falta de pagamento "
        "no vencimento, serão cobrados juros de mora.\n\n"
        f"CTE: {ctes_str}\n\n"
        f"{rotas_str}"
    )

    ws["F24"] = round(group["subtotal"], 2)
    ws["F25"] = None
    ws["F26"] = round(group["icms"], 2)
    # F27 mantém a fórmula '=SUM(F24:F26)' do template

    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--prefatura", required=True, type=Path,
                     help="Caminho da Pré-fatura (.xlsx)")
    ap.add_argument("--template", required=True, type=Path,
                     help="Caminho do template de fatura (ex.: BETIM - BRMG02.xlsx)")
    ap.add_argument("--out-dir", required=True, type=Path,
                     help="Pasta onde salvar as faturas geradas")
    ap.add_argument("--emissao", required=True,
                     help="Data de emissão (YYYY-MM-DD), usada em todas as faturas geradas")
    ap.add_argument("--periodo", required=True,
                     help='Código de período p/ Observação (ex.: "202608Q1")')
    ap.add_argument("--linehaul", required=True,
                     help='Número "Line Haul" p/ Observação (ex.: "6745030")')
    ap.add_argument("--hub", default=None,
                     help="Gerar só um hub específico (ex.: BRMG02). Default: todos.")
    args = ap.parse_args()

    import datetime
    emissao = datetime.date.fromisoformat(args.emissao)

    print(f"Carregando {args.prefatura} ...")
    wb = openpyxl.load_workbook(args.prefatura, data_only=True)
    hub_to_city, hub_to_routes = load_meli_maps(wb["1. MELI"])
    groups = load_lancamento_groups(wb["4. Lançamento"])
    ctes_by_city = load_ctes_by_city(wb["CT-es"])

    args.out_dir.mkdir(parents=True, exist_ok=True)

    hubs = [args.hub] if args.hub else sorted(groups.keys())
    for hub in hubs:
        group = groups.get(hub)
        if not group:
            print(f"  [!] hub {hub}: nenhuma linha 'Fatura' encontrada na aba 4. Lançamento, pulando.")
            continue
        cidade = hub_to_city.get(hub)
        if not cidade:
            print(f"  [!] hub {hub}: cidade não encontrada na aba 1. MELI, pulando.")
            continue
        ctes = ctes_by_city.get(normalize_city(cidade), [])
        if not ctes:
            print(f"  [!] hub {hub} ({cidade}): nenhum CT-e encontrado na aba CT-es para essa cidade.")
        rotas = hub_to_routes.get(hub, [])

        out_path = args.out_dir / f"{cidade} - {hub}.xlsx"
        build_fatura(
            args.template, out_path,
            hub=hub, cidade=cidade, emissao=emissao,
            periodo=args.periodo, line_haul=args.linehaul,
            group=group, ctes=ctes, rotas=rotas,
        )
        print(f"  OK  {out_path.name}  "
              f"(subtotal={group['subtotal']:.2f} icms={group['icms']:.2f} total={group['total']:.2f} "
              f"linhas={group['linhas']} ctes={len(ctes)})")

    print("Concluído.")


if __name__ == "__main__":
    main()
