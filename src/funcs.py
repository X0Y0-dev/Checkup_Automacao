import streamlit as st
import pandas as pd
import re
import unicodedata

from io import BytesIO
from datetime import date, timedelta
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule


#region CONFIRGUAÇÃO DE ELEBILIDADE

# Empresas que necessitam de elegibilidade
EMPRESAS = ["Bradesco", "Santander", "Itau", "Mediservice"]

# Convênios que necessitam de elegibilidade
CONVENIOS = ["Central Nacional Unimed", "Unimed Seguros Saude", "Care Plus", "Unafisco", "Camara", "Gama", "Senado", "Intermedici"]

# Opções de status
STATUS_OPC = ["-", "ELEGÍVEL", "NÃO ELEGÍVEL", "SEM GUIA"]

# Opções de situação
SITUACAO_OPC = ["-", "AGENDADO", "REAGENDADO", "CANCELADO", "PENDENTE", "VENCIDA"]

#endregion

#region CONFIGURAÇÕES DE ESTILIZAÇÃO

# Função de definir a cor da célula
def fill(cor: str) -> PatternFill:
    return PatternFill(
        start_color = cor,
        end_color = cor,
        fill_type = "solid"
    )

# Função de definir cor da fonte
def fonte(cor: str, bold: bool = True) -> Font:
    return Font(
        color = cor,
        bold = True
    )


# Cores das células
VERDE = fill("92D050")
VERDE_ESCURO = fill("006400")
AMARELO = fill("FFFF00")
VERMELHO = fill("FF0000")
CINZA_ESCURO = fill("5A5A5A")
VINHO = fill("8B0000")
LARANJA = fill("FF5F15")
ROXO = fill("7030A0")
BRANCO_FILL = fill("FFFFFF")
CINZA = fill("D9D9D9")

# Cores das fontes
PRETO = fonte("000000")
BRANCO = fonte("FFFFFF")

# Dicionário de estilização de status
STATUS_ESTILOS = {
    "-": (BRANCO_FILL, PRETO),
    "ELEGÍVEL": (VERDE, PRETO),
    "NÃO ELEGÍVEL": (VERMELHO, BRANCO),
    "SEM GUIA": (AMARELO, PRETO),
}

# Dicionário de estilização de situação
SITUACAO_ESTILOS = {
    "-": (BRANCO_FILL, PRETO),
    "AGENDADO": (VERDE_ESCURO, BRANCO),
    "REAGENDADO": (ROXO, BRANCO),
    "CANCELADO": (CINZA_ESCURO, BRANCO),
    "PENDENTE": (LARANJA, BRANCO),
    "VENCIDA": (VINHO, BRANCO),
}

#endregion

estado = {"df": pd.DataFrame(), "historico": []}

hora_regex = re.compile(r'^\d{2}:\d{2}$')
secoes_regex = re.compile(r'^\d+\s*/\s*\d+$')
IDX_SECOES = 4

#region FUNÇÕES E TRATAMENTO DE DADOS

def parse_checkup(texto: str) -> pd.DataFrame:

    #===========================
    #|| Normalização do texto ||
    #===========================
        
    # Remove caracteres invisíveis que podem aparecer ao copiar
    texto = re.sub(r'[\u200b\u200c\u200d\u202a\u202b\u202c\u202d\u202e\ufeff]', '', texto)
    texto = texto.replace('\r\n', '\n').replace('\r', '\n')
    
    todas_linhas = [l.strip() for l in texto.split('\n')] # Divide o texto em linhas e remove espaços extras

    #==============================
    #|| Identificação dos blocos ||
    #==============================
       
    # Cada paciente inivia em uma linha contendo um horário válido.
    # A segunda condição evita que números pertencentes a outras
    # seções sejam interpretados como um novo paciente
    
    inicios = [
        i for i, linha in enumerate(todas_linhas)
        if hora_regex.match(linha)
        and i + 1 < len(todas_linhas)
        and not todas_linhas[i + 1][:1].isdigit()
    ]

    pacientes = []

    # Percorre cada bloco identificado
    for idx, inicio in enumerate(inicios):

        fim = inicios[idx + 1] if idx + 1 < len(inicios) else len(todas_linhas) # Fim do bloco indica um novo paciente ou final do texto
        linhas = [l for l in todas_linhas[inicio:fim] if l] # Remova linhas vazias do bloco

        # Ignora blocos incompletos
        if len(linhas) <= IDX_SECOES:
            continue
        # Valida se a seção esperada realmente existe e descarta blocos mal formados
        if not secoes_regex.match(linhas[IDX_SECOES]):
            continue

        # Os quatro ptimeiros campos possuem posição fixa dentro de um bloco válido
        hora, paciente, convenio, categoria = linhas[:4]

        # Armazena os dados do paciente
        pacientes.append({
            "Hora": hora,
            "Paciente": paciente,
            "Convênio": convenio,
            "Categoria": categoria,
            "Status": "",
            "Situação": "",
        })
    return pd.DataFrame(pacientes) # Retorna todos os pacientes encontrador em formato DataFrame

def normalizar(texto: str) -> str:
    texto = unicodedata.normalize("NFD", texto)        # Decompõe acentos
    texto = texto.encode("ascii", "ignore").decode()   # Remove acentos
    texto = texto.lower()                              # Deixa em minúsculo
    texto = re.sub(r'\s+', ' ', texto).strip()         # Remove espaços extras
    return texto

def identificar_empresa(convenio: str) -> str | None:

    # Comparação com o que é passado com o dicionário
    for empresa in EMPRESAS:
        if normalizar(empresa) in normalizar(convenio):
            return empresa
    return None

def identificar_convenio(convenio: str) -> str | None:

    # Comparação com o que é passado com o dicionário
    for convs in CONVENIOS:
        if normalizar(convs) in normalizar(convenio):
            return convs
    return None

def gerar_txt_convenios(df_convenios: pd.DataFrame) -> bytes:
    linhas = []

    # Percorre cada convênio em ordem alfabética
    for convenio in sorted(df_convenios["Convênio"].unique()):
        linhas.append(f"\n\n{convenio}:") # Adiciona o nome do convênio como título da seção
        df_conv = df_convenios[df_convenios["Convênio"] == convenio] # Seleciona apenas os pacientes do convênio atual

        # Agrupa os pacientes por data
        for data in sorted(df_conv["Data"].unique()):
            linhas.append(f"\n  {data}:\n") # Adiciona a data como subtítulo da seção
            df_data = df_conv[(df_conv["Data"] == data) & (~df_conv["Hora"].str.startswith("10:"))] # Seleciona apenas os pacientes da data atual
            
            # Lista os pacientes e suas respectivas categorias
            for _, row in df_data.iterrows():
                linhas.append(f"    {row['Paciente']} - {row['Categoria']}")
        linhas.append("") # Linha em branco

    return "\n\n".join(linhas).encode("utf-8")  # Retorna bytes ao invés de salvar

def gerar_txt_brasilia(df_brasilia: pd.DataFrame) -> bytes:
    
    linhas = []

    # Agrupa os pacientes por data
    for data in sorted(df_brasilia["Data"].unique()):
        linhas.append(f"{data}:")
        df_data = df_brasilia[(df_brasilia["Data"] == data) & (~df_brasilia["Hora"].str.startswith("10:"))] # Seleciona apenas os pacientes da data atual
        
        # Lista os pacientes e suas respectivas categorias
        for _, row in df_data.iterrows():
            linhas.append(f"    {row['Paciente']} - {row['Convênio']} - {row['Categoria']}")
        linhas.append("")

    return "\n\n".join(linhas).encode("utf-8")  # Retorna bytes ao invés de salvar

def gerar_txt_empresa(df_empresa: pd.DataFrame, nome_empresa: str) -> bytes:
    linhas = []

    # Agrupa os pacientes por data
    for data in sorted(df_empresa["Data"].unique()):
        linhas.append(f"{data}:")
        df_data = df_empresa[(df_empresa["Data"] == data) & (~df_empresa["Hora"].str.startswith("10:"))] # Seleciona apenas os pacientes da data atual

        # Lista os pacientes e suas respectivas categorias
        for _, row in df_data.iterrows():
            linhas.append(f"    {row['Paciente']} - {row['Convênio']} - {row['Categoria']}")
        linhas.append("")
    
    return "\n\n".join(linhas).encode("utf-8") # Retorna bytes ao invés de salvar

def estilizar_header(worksheet):

    worksheet.freeze_panes = "A2"

    for cell in worksheet[1]:
        cell.fill = CINZA
        cell.font = PRETO
    
    for column in worksheet.columns:

        max_length = 0
        column_letter = column[0].column_letter
        header = column[0].value

        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        
        # Largura automárica
        largura = max_length + 2

        # Ajuste de largura para status
        if header == "Status" or header == "Situação":
            largura = 20
        
        worksheet.column_dimensions[column_letter].width = largura


def estilizar_validacoes(worksheet):

    # Validações
    dv_status = DataValidation(
        type = "list",
        formula1 = f'"{",".join(STATUS_OPC)}"',
        allow_blank = True
    )

    dv_situacao = DataValidation(
        type = "list",
        formula1 = f'"{",".join(SITUACAO_OPC)}"',
        allow_blank = True
    )

    worksheet.add_data_validation(dv_status)
    worksheet.add_data_validation(dv_situacao)

    # Localiza as colunas
    colunas = {
        cell.value: cell.column_letter
        for cell in worksheet[1]
    }

    # ===========================
    # ||     Coluna Status     ||
    # ===========================

    if "Status" in colunas:

        coluna = colunas["Status"]
        intervalo = f"{coluna}2:{coluna}{worksheet.max_row}"

        dv_status.add(intervalo)

        for status, (fill, font) in STATUS_ESTILOS.items():
            worksheet.conditional_formatting.add(
                intervalo,
                FormulaRule(
                    formula=[f'{coluna}2="{status}"'],
                    fill=fill,
                    font=font
                )
            )

    # ===========================
    # ||    Coluna Situação    ||
    # ===========================

    if "Situação" in colunas:

        coluna = colunas["Situação"]
        intervalo = f"{coluna}2:{coluna}{worksheet.max_row}"

        dv_situacao.add(intervalo)

        for situacao, (fill, font) in SITUACAO_ESTILOS.items():
            worksheet.conditional_formatting.add(
                intervalo,
                FormulaRule(
                    formula=[f'{coluna}2="{situacao}"'],
                    fill=fill,
                    font=font
                )
            )
            
#endregion
