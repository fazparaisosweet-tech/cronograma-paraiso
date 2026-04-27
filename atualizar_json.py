import pandas as pd
import json
import math
import subprocess
import os
from datetime import datetime, date

# =========================
# CONFIGURAÇÕES
# =========================
ARQUIVO_EXCEL = r"\\SERV13-BKP\Serv13 Fazendas Arquivos & BKP\Fazenda Sweet Paraiso\13 - Planejamento Agrícola\Cronograma_agricola_Paraíso.xlsx"

PASTA_GITHUB = r"C:\Users\emanuel.rodrigues\Documents\GitHub\cronograma-paraiso"

ARQUIVO_JSON = os.path.join(PASTA_GITHUB, "dados.json")

ABA_ANO = "2026"

# =========================
# LER EXCEL
# =========================
df_raw = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=ABA_ANO,
    header=None
)

linha_cabecalho = None

for i, row in df_raw.iterrows():
    if row.astype(str).str.lower().str.contains("parcela").any():
        linha_cabecalho = i
        break

if linha_cabecalho is None:
    raise Exception("Cabeçalho com PARCELA não encontrado.")

df = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=ABA_ANO,
    header=linha_cabecalho
)

# =========================
# LIMPAR COLUNAS
# =========================
df.columns = df.columns.astype(str).str.strip()
df = df.loc[:, ~df.columns.str.contains("Unnamed", case=False)]

# =========================
# FORÇAR PARCELA INTEIRO
# =========================
coluna_parcela = next(
    (c for c in df.columns if "parcela" in c.lower()),
    None
)

if coluna_parcela:
    df[coluna_parcela] = pd.to_numeric(
        df[coluna_parcela],
        errors="coerce"
    ).fillna(0).astype(int)

# =========================
# CONVERTER DADOS
# =========================
def converter(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, int):
        return int(valor)

    if isinstance(valor, float):
        if math.isnan(valor):
            return None

        if valor.is_integer():
            return int(valor)

        return valor

    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")

    if isinstance(valor, str):
        v = valor.strip()

        if v == "" or v == "-":
            return None

        return v

    return valor

df = df.apply(lambda col: col.apply(converter))

# =========================
# GERAR JSON
# =========================
dados = df.to_dict(orient="records")

# =========================
# LIMPAR NAN FINAL
# =========================
def limpar_nan(obj):
    if isinstance(obj, float) and math.isnan(obj):
        return None

    if isinstance(obj, dict):
        return {k: limpar_nan(v) for k, v in obj.items()}

    if isinstance(obj, list):
        return [limpar_nan(v) for v in obj]

    return obj

dados = limpar_nan(dados)

# =========================
# SALVAR JSON
# =========================
with open(ARQUIVO_JSON, "w", encoding="utf-8") as f:
    json.dump(
        dados,
        f,
        ensure_ascii=False,
        indent=2,
        allow_nan=False
    )

print("✅ JSON atualizado | PARCELA sem decimal")

# =========================
# ENVIAR GITHUB
# =========================
os.chdir(PASTA_GITHUB)

subprocess.run(["git", "add", "."], check=True)

# VERIFICA SE EXISTE ALTERAÇÃO
status = subprocess.run(
    ["git", "status", "--porcelain"],
    capture_output=True,
    text=True
)

if status.stdout.strip():
    subprocess.run(
        ["git", "commit", "-m", "Atualização automática"],
        check=True
    )

    subprocess.run(
        ["git", "push"],
        check=True
    )

    print("✅ GitHub atualizado com sucesso.")

else:
    print("ℹ️ Nenhuma alteração detectada. Nada para enviar.")