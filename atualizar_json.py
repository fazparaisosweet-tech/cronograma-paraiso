import pandas as pd
import json
import math
import subprocess
import os
import re
from datetime import datetime, date
from openpyxl import load_workbook
from openpyxl.comments import Comment

# =========================
# CONFIGURAÇÕES
# =========================
ARQUIVO_EXCEL = r"\\SERV13-BKP\Serv13 Fazendas Arquivos & BKP\Fazenda Sweet Paraiso\13 - Planejamento Agrícola\Cronograma_agricola_Paraíso.xlsx"

PASTA_GITHUB = r"C:\Users\emanuel.rodrigues\Documents\GitHub\cronograma-paraiso"

ARQUIVO_JSON = os.path.join(PASTA_GITHUB, "dados.json")

ABA_ANO = "2026"

# =========================
# FUNÇÃO PARA LIMPAR COMENTÁRIO
# =========================
def limpar_comentario(texto):
    """
    Remove texto padrão do Excel como '[Comentário encadeado]' e links
    """
    if not texto:
        return texto
    
    # Remove o marcador de comentário encadeado
    texto = re.sub(r'\[Comentário encadeado\]\s*', '', texto)
    
    # Remove o aviso do Excel
    texto = re.sub(r'Sua versão do Excel permite que você leia este comentário encadeado.*?https?://[^\s]+', '', texto, flags=re.DOTALL)
    
    # Remove links restantes
    texto = re.sub(r'https?://[^\s]+', '', texto)
    
    # Remove linhas em branco extras no início
    texto = texto.strip()
    
    # Se o texto ficou vazio após limpeza, retorna None
    if not texto or texto == "":
        return None
    
    return texto

# =========================
# LER COMENTÁRIOS DO EXCEL
# =========================
def ler_comentarios_excel(arquivo_excel, aba_nome):
    """
    Lê todos os comentários do Excel e retorna um dicionário
    onde a chave é (linha, coluna) e o valor é o comentário limpo
    """
    comentarios = {}
    
    try:
        # Carrega o workbook com openpyxl para acessar comentários
        workbook = load_workbook(arquivo_excel, data_only=True)
        
        if aba_nome in workbook.sheetnames:
            planilha = workbook[aba_nome]
            
            # Itera sobre todas as células que têm comentários
            for row in range(1, planilha.max_row + 1):
                for col in range(1, planilha.max_column + 1):
                    cell = planilha.cell(row=row, column=col)
                    
                    if cell.comment and isinstance(cell.comment, Comment):
                        # Comentário existe nesta célula
                        texto_original = cell.comment.text
                        texto_limpo = limpar_comentario(texto_original)
                        
                        if texto_limpo:  # Só armazena se tiver conteúdo útil
                            comentarios[(row, col)] = texto_limpo
                            print(f"📝 Comentário na linha {row}, coluna {col}: {texto_limpo[:50]}...")
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

print(f"📊 Total de comentários encontrados: {len(comentarios)}")

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

# Lista todas as colunas
colunas = list(df.columns)

# Cria um mapa de colunas para índices (1-based para Excel)
coluna_indices = {}
for i, col in enumerate(colunas):
    coluna_indices[col] = i + 1  # +1 porque Excel é 1-based

# Adiciona uma coluna de comentários
df['COMENTARIOS'] = None

# Converte o DataFrame para lista de dicionários
dados = df.to_dict(orient="records")

# Mapeia comentários por parcela e coluna
comentarios_por_parcela = {}

for (linha_excel, col_excel), texto in comentarios.items():
    # A linha do cabeçalho está na posição linha_cabecalho + 1 (1-based)
    # As linhas de dados começam em linha_cabecalho + 2
    linha_dados = linha_excel - (linha_cabecalho + 2)  # Índice do DataFrame (0-based)
    
    # Verifica se é uma linha de dados (não é cabeçalho)
    if linha_dados >= 0 and linha_dados < len(dados):
        # Obtém o nome da coluna pelo índice
        for nome_col, idx_col in coluna_indices.items():
            if idx_col == col_excel:
                # Adiciona o comentário ao mapa
                parcela_num = dados[linha_dados].get(coluna_parcela)
                
                if parcela_num not in comentarios_por_parcela:
                    comentarios_por_parcela[parcela_num] = {}
                
                comentarios_por_parcela[parcela_num][nome_col] = texto
                print(f"📌 Parcela {parcela_num}: Comentário em '{nome_col}'")
                break

# Adiciona os comentários aos dados
for idx, linha in enumerate(dados):
    parcela_num = linha.get(coluna_parcela)
    
    if parcela_num in comentarios_por_parcela and comentarios_por_parcela[parcela_num]:
        linha['COMENTARIOS'] = comentarios_por_parcela[parcela_num]
        print(f"✅ Parcela {parcela_num}: {len(comentarios_por_parcela[parcela_num])} comentário(s) vinculado(s)")

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