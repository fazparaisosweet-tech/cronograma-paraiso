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
    
    # Remove linhas em branco extras
    texto = texto.strip()
    
    if not texto or texto == "":
        return None
    
    return texto

# =========================
# LER CABEÇALHO E MAPEAR COLUNAS
# =========================
print("📖 Lendo dados do Excel...")

# Primeiro, lê o cabeçalho para saber os nomes das colunas
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

# Lê o cabeçalho para mapear colunas
df_header = pd.read_excel(
    ARQUIVO_EXCEL,
    sheet_name=ABA_ANO,
    header=linha_cabecalho,
    nrows=0  # Lê apenas o cabeçalho
)

# Cria um mapeamento de número da coluna (1-based) para nome da coluna
colunas = list(df_header.columns)
coluna_mapeamento = {}
for idx, col in enumerate(colunas):
    # Excel é 1-based, então coluna = idx + 1
    coluna_mapeamento[idx + 1] = str(col).strip()

print(f"📊 {len(colunas)} colunas encontradas")

# =========================
# LER COMENTÁRIOS DO EXCEL
# =========================
print("\n📝 Lendo comentários do Excel...")

def ler_comentarios_excel_com_precisao(arquivo_excel, aba_nome, coluna_mapeamento, linha_cabecalho):
    """
    Lê comentários e mapeia para (linha do Excel, nome da coluna).
    A mesma parcela pode ter duas linhas (ciclo colhido e ciclo atual), por isso
    o comentário fica preso à LINHA, e não ao número da parcela.
    """
    comentarios = {}
    
    try:
        workbook = load_workbook(arquivo_excel, data_only=True)
        
        if aba_nome in workbook.sheetnames:
            planilha = workbook[aba_nome]
            
            for row in range(linha_cabecalho + 2, planilha.max_row + 1):
                for col in range(1, planilha.max_column + 1):
                    cell = planilha.cell(row=row, column=col)
                    
                    if cell.comment and isinstance(cell.comment, Comment):
                        texto_limpo = limpar_comentario(cell.comment.text)
                        
                        if texto_limpo:
                            nome_coluna = coluna_mapeamento.get(col, f"Coluna_{col}")
                            comentarios[(row, nome_coluna)] = texto_limpo
                            print(f"📝 Linha {row}, Coluna '{nome_coluna}': {texto_limpo[:50]}...")
        else:
            print(f"⚠️ Aba '{aba_nome}' não encontrada")
        
        workbook.close()
        
    except Exception as e:
        print(f"⚠️ Erro ao ler comentários: {e}")
    
    return comentarios

comentarios = ler_comentarios_excel_com_precisao(
    ARQUIVO_EXCEL, 
    ABA_ANO, 
    coluna_mapeamento, 
    linha_cabecalho
)

print(f"📊 Total de comentários encontrados: {len(comentarios)}")

# =========================
# DATAS DIGITADAS x FÓRMULA
# =========================
# Regra da planilha:
#   data DIGITADA na célula  -> operação EXECUTADA
#   data vinda de FÓRMULA    -> PREVISÃO (programada)
# O pandas só enxerga o valor final, então abrimos o Excel com
# data_only=False para ver o que é fórmula.
print("\n🔎 Verificando datas digitadas x fórmulas...")

def eh_formula(cell):
    v = cell.value
    if cell.data_type == "f":
        return True
    if isinstance(v, str) and v.startswith("="):
        return True
    # Fórmulas de matriz/tabela vêm como objetos especiais do openpyxl
    return v is not None and not isinstance(v, (str, int, float, bool, datetime, date))

def ler_datas_digitadas(arquivo_excel, aba_nome, coluna_mapeamento, linha_cabecalho):
    """
    Retorna:
      digitadas: {linha_excel: [nomes das colunas com data digitada]}
      parcelas:  {linha_excel: número da parcela digitado na linha} (para conferência)
    """
    digitadas = {}
    parcelas = {}
    total_formulas = 0
    total_digitadas = 0

    col_parcela = next(
        (c for c, nome in coluna_mapeamento.items() if "parcela" in nome.lower()),
        None
    )

    workbook = load_workbook(arquivo_excel, data_only=False)
    try:
        planilha = workbook[aba_nome]
        for linha in planilha.iter_rows(min_row=linha_cabecalho + 2, max_row=planilha.max_row):
            for cell in linha:
                if cell.value is None:
                    continue
                nome = coluna_mapeamento.get(cell.column)
                if not nome:
                    continue

                if cell.column == col_parcela and isinstance(cell.value, (int, float)):
                    parcelas[cell.row] = int(cell.value)

                if eh_formula(cell):
                    total_formulas += 1
                elif isinstance(cell.value, (datetime, date)):
                    digitadas.setdefault(cell.row, []).append(nome)
                    total_digitadas += 1
    finally:
        workbook.close()

    print(f"   ✍️  Datas digitadas (executadas): {total_digitadas}")
    print(f"   🧮 Células com fórmula: {total_formulas}")
    return digitadas, parcelas

datas_digitadas, parcelas_por_linha = ler_datas_digitadas(
    ARQUIVO_EXCEL,
    ABA_ANO,
    coluna_mapeamento,
    linha_cabecalho
)

# =========================
# CARREGAR DADOS COMPLETOS
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

# Converte o DataFrame para lista de dicionários
dados = df.to_dict(orient="records")

# Cada item de "dados" corresponde a uma linha do Excel, na mesma ordem:
# item 0 = linha logo abaixo do cabeçalho.
padrao_data = re.compile(r"^\d{4}-\d{2}-\d{2}$")
divergencias = 0

for idx, linha in enumerate(dados):
    linha_excel = linha_cabecalho + 2 + idx
    parcela_num = linha.get(coluna_parcela)

    # Conferência: a parcela da linha do Excel tem que bater com a do pandas
    parcela_excel = parcelas_por_linha.get(linha_excel)
    if parcela_excel is not None and parcela_num and parcela_excel != parcela_num:
        divergencias += 1

    # Comentários desta linha
    comentarios_parcela = {
        coluna: texto
        for (lin, coluna), texto in comentarios.items()
        if lin == linha_excel
    }
    if comentarios_parcela:
        linha['COMENTARIOS'] = comentarios_parcela
        print(f"✅ Parcela {parcela_num} (linha {linha_excel}): {len(comentarios_parcela)} comentário(s)")

    # Operações executadas = colunas com data digitada nesta linha
    if parcela_num:
        linha['EXECUTADAS'] = [
            coluna for coluna in datas_digitadas.get(linha_excel, [])
            if isinstance(linha.get(coluna), str) and padrao_data.match(linha[coluna])
        ]

if divergencias:
    raise Exception(
        f"{divergencias} linha(s) com parcela diferente entre Excel e pandas. "
        "As linhas não estão alinhadas; JSON NÃO foi gerado para não marcar execuções erradas."
    )

print(f"✍️  Campo EXECUTADAS adicionado em {sum(1 for l in dados if 'EXECUTADAS' in l)} registros")

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

print(f"\n✅ JSON atualizado com {len(dados)} registros")

# =========================
# ENVIAR GITHUB
# =========================
os.chdir(PASTA_GITHUB)

subprocess.run(["git", "add", "."], check=True)

status = subprocess.run(
    ["git", "status", "--porcelain"],
    capture_output=True,
    text=True
)

if status.stdout.strip():
    subprocess.run(
        ["git", "commit", "-m", "Atualização automática (comentários + executadas)"],
        check=True
    )

    subprocess.run(
        ["git", "push"],
        check=True
    )

    print("✅ GitHub atualizado com sucesso.")

else:
    print("ℹ️ Nenhuma alteração detectada. Nada para enviar.")