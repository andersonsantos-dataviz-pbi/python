import json
from pathlib import Path
import pandas as pd


def json_to_xlsx(
    json_path: str,
    xlsx_path: str,
    record_path: str | None = None,
    meta: list[str] | None = None,
    sheet_name: str = "dados",
):
    """
    Lê um JSON (ex.: exportado do Postman) e gera um Excel (.xlsx).

    - Se o JSON for uma LISTA de objetos -> vira uma tabela direto.
    - Se for um OBJETO com lista interna (ex.: {"data":[...]}) -> use record_path="data".
    - Para JSONs aninhados, usa pandas.json_normalize para "achatar" campos.

    record_path:
      - caminho até a lista principal (string com '.'), ex.: "data.items"
    meta:
      - campos para repetir como colunas, ex.: ["status", "requestId"]
    """

    json_path = str(Path(json_path))
    xlsx_path = str(Path(xlsx_path))

    with open(json_path, "r", encoding="utf-8") as f:
        payload = json.load(f)

    # Se record_path for informado, usamos json_normalize para extrair a lista principal
    if record_path:
        rp = record_path.split(".")
        df = pd.json_normalize(payload, record_path=rp, meta=meta, errors="ignore")
    else:
        # Caso comum: payload é lista -> DataFrame direto
        if isinstance(payload, list):
            df = pd.json_normalize(payload)
        # Caso comum: payload é dict -> tenta "adivinhar" a primeira lista interna
        elif isinstance(payload, dict):
            # tenta achar uma chave que seja lista de dicts (data, items, results, etc.)
            possible_lists = [
                (k, v) for k, v in payload.items()
                if isinstance(v, list) and (len(v) == 0 or isinstance(v[0], dict))
            ]
            if possible_lists:
                k, v = possible_lists[0]
                df = pd.json_normalize(v)
            else:
                # se não tiver lista, exporta como 1 linha "achatada"
                df = pd.json_normalize(payload)
        else:
            raise ValueError("JSON precisa ser lista ou objeto (dict).")

    # Ajuste opcional: garantir colunas em ordem estável
    df = df.reindex(sorted(df.columns), axis=1)

    # Exporta para Excel
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

    print(f"OK: gerado {xlsx_path} com {len(df)} linhas e {len(df.columns)} colunas.")


if __name__ == "__main__":
    json_to_xlsx(
        json_path="response.json",
        xlsx_path="dados_api.xlsx",
        # Exemplos (descomente se precisar):
        # record_path="data.items",
        # meta=["status", "requestId"],
        sheet_name="Dados_API"
    )
