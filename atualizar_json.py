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
df_raw = pd.read_excel(ARQUIVO_EXCEL, sheet_name=ABA_ANO, header=None)

linha_cabecalho = None

for i, row in df_raw.iterrows():
    if row.astype(str).str.lower().str.contains("parcela").any():
        linha_cabecalho = i
        break

df = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=ABA_ANO,
    header=linha_cabecalho
)

df.columns = df.columns.astype(str).str.strip()
df = df.loc[:, ~df.columns.str.contains("Unnamed", case=False)]

# =========================
# LIMPAR
# =========================
def converter(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, float):
        return int(valor) if valor.is_integer() else valor

    if isinstance(valor, (datetime, date)):
        return valor.strftime("%Y-%m-%d")

    return valor

df = df.apply(lambda col: col.apply(converter))

dados = df.to_dict(orient="records")

# =========================
# SALVAR JSON
# =========================
with open(ARQUIVO_JSON, "w", encoding="utf-8") as f:
    json.dump(dados, f, ensure_ascii=False, indent=2)

print("JSON atualizado.")

# =========================
# ENVIAR GITHUB
# =========================
os.chdir(PASTA_GITHUB)

subprocess.run(["git", "add", "."])
subprocess.run(["git", "commit", "-m", "Atualização automática"])
subprocess.run(["git", "push"])

print("GitHub atualizado com sucesso.")