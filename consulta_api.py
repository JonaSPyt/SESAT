"""
Módulo de consulta à API do TRE-CE para buscar dados de patrimônio.
Faz scraping do HTML retornado e extrai informações do equipamento.
"""
import requests
from bs4 import BeautifulSoup

BASE_URL = (
    "https://intranet.tre-ce.jus.br/area-de-dados/outros/"
    "consultas-sequi/consulta-por-numero-do-patrimonio/"
    "html_view_2?patrimonio={tombamento}"
)


def consultar_patrimonio(tombamento: str) -> dict | None:
    """
    Consulta o patrimônio na intranet do TRE-CE.
    Retorna dict com os dados ou None se não encontrado / erro.

    Campos retornados:
        - patrimonio
        - sigla (seção/zona, ex: '066 ZE')
        - nome_setor (ex: '66ª ZONA ELEITORAL - AQUIRAZ')
        - local (endereço)
        - nome_responsavel
        - ds_bem (descrição curta, ex: 'APARELHO TELEFONICO IP COM LCD')
        - descricao_completa (descrição com marca/modelo)
        - vl_unitario
    """
    tombamento = str(tombamento).strip()
    if not tombamento:
        return None

    url = BASE_URL.format(tombamento=tombamento)

    try:
        resp = requests.get(url, timeout=15, verify=False)
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"[ERRO] Falha na consulta: {e}")
        return None

    soup = BeautifulSoup(resp.text, "html.parser")
    table = soup.find("table", id="tabela-consultaoracle")
    if not table:
        return None

    tbody = table.find("tbody")
    if not tbody:
        return None

    row = tbody.find("tr")
    if not row:
        return None

    cells = row.find_all("td")
    if len(cells) < 8:
        return None

    dados = {
        "patrimonio": cells[0].get_text(strip=True),
        "sigla": cells[1].get_text(strip=True),
        "nome_setor": cells[2].get_text(strip=True),
        "local": cells[3].get_text(strip=True),
        "nome_responsavel": cells[4].get_text(strip=True),
        "ds_bem": cells[5].get_text(strip=True),
        "descricao_completa": cells[6].get_text(strip=True),
        "vl_unitario": cells[7].get_text(strip=True),
    }

    return dados


# Suprimir warnings de SSL (intranet pode ter cert auto-assinado)
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
