import pandas as pd
import json
import math
import subprocess
import os
from datetime import datetime, date
from openpyxl import load_workbook

# =========================
# CONFIGURAÇÕES
# =========================
ARQUIVO_EXCEL = r"\\SERV13-BKP\Serv13 Fazendas Arquivos & BKP\Fazenda Sweet Paraiso\13 - Planejamento Agrícola\Cronograma_agricola_Paraíso.xlsx"

PASTA_GITHUB = r"C:\Users\emanuel.rodrigues\Documents\GitHub\cronograma-paraiso"

ARQUIVO_JSON = os.path.join(PASTA_GITHUB, "dados.json")

ABA_ANO = "2026"

# =========================
# LER COMENTÁRIOS DO EXCEL
# =========================
def ler_comentarios_excel(arquivo_excel, aba_nome):
    """
    Lê todos os comentários do Excel e retorna um dicionário
    onde a chave é (linha, coluna) e o valor é o comentário
    """
    comentarios = {}
    
    try:
        # Carrega o workbook com openpyxl para acessar comentários
        workbook = load_workbook(arquivo_excel, data_only=True)
        
        if aba_nome in workbook.sheetnames:
            planilha = workbook[aba_nome]
            
            for cell in planilha:
                if cell.comment:
                    # Comentário existe nesta célula
                    # Converte linha e coluna para índices (0-based)
                    linha = cell.row - 1  # openpyxl usa 1-based
                    coluna = cell.column - 1  # openpyxl usa 1-based
                    
                    # Pega o texto do comentário
                    texto_comentario = cell.comment.text
                    
                    # Armazena no dicionário
                    comentarios[(linha, coluna)] = texto_comentario
                    
                    print(f"📝 Comentário encontrado na linha {linha+1}, coluna {coluna+1}: {texto_comentario[:50]}...")
        else:
            print(f"⚠️ Aba '{aba_nome}' não encontrada para leitura de comentários")
        
        workbook.close()
        
    except Exception as e:
        print(f"⚠️ Erro ao ler comentários: {e}")
        print("   Continuando sem comentários...")
    
    return comentarios

# =========================
# LER EXCEL (DADOS)
# =========================
print("📖 Lendo dados do Excel...")

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

print(f"📌 Cabeçalho encontrado na linha {linha_cabecalho + 1}")

# =========================
# LER COMENTÁRIOS
# =========================
print("\n📝 Lendo comentários do Excel...")
comentarios = ler_comentarios_excel(ARQUIVO_EXCEL, ABA_ANO)

# =========================
# CARREGAR DADOS COM CABEÇALHO
# =========================
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
# ADICIONAR COMENTÁRIOS AOS DADOS
# =========================
print("\n🔗 Vinculando comentários às parcelas...")

# Mapeia o índice de cada coluna
colunas = list(df.columns)
coluna_parcela_idx = colunas.index(coluna_parcela) if coluna_parcela in colunas else 0

# Adiciona uma coluna de comentários
df['COMENTARIOS'] = None

# Converte o DataFrame para lista de dicionários
dados = df.to_dict(orient="records")

# Adiciona comentários a cada registro
for idx, linha in enumerate(dados):
    parcela_num = linha.get(coluna_parcela)
    
    # Calcula a linha real no Excel (cabeçalho + 1 + índice)
    linha_excel = linha_cabecalho + 2 + idx  # +2 porque cabeçalho é 0-based e tem a linha do cabeçalho
    
    # Busca comentários para esta linha
    comentarios_linha = {}
    
    for (linha_comment, coluna_comment), texto in comentarios.items():
        # Se a linha do comentário corresponde à linha atual
        if linha_comment == linha_excel:
            # Obtém o nome da coluna correspondente
            if coluna_comment < len(colunas):
                nome_coluna = colunas[coluna_comment]
                # Evita colunas especiais ou índice
                if nome_coluna not in ['COMENTARIOS', coluna_parcela]:
                    comentarios_linha[nome_coluna] = texto
    
    # Se houver comentários, adiciona ao registro
    if comentarios_linha:
        linha['COMENTARIOS'] = comentarios_linha
        print(f"📌 Parcela {parcela_num}: {len(comentarios_linha)} comentário(s) vinculado(s)")

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

print(f"\n✅ JSON atualizado com {len(dados)} registros e comentários incluídos")

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
        ["git", "commit", "-m", "Atualização automática com comentários"],
        check=True
    )

    subprocess.run(
        ["git", "push"],
        check=True
    )

    print("✅ GitHub atualizado com sucesso.")

else:
    print("ℹ️ Nenhuma alteração detectada. Nada para enviar.")