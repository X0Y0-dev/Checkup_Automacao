import streamlit as st
import pandas as pd
import re
import pathlib

from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from datetime import date, timedelta
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

from src.menu import render_header, render_sidebar
from src.funcs import parse_checkup, identificar_empresa, identificar_convenio, gerar_txt_convenios, gerar_txt_brasilia, estilizar_header, estilizar_status, EMPRESAS


BASE_DIR = Path(__file__).resolve().parent.parent

css = pathlib.Path(BASE_DIR / "pages" / "assets" / "style.css").read_text()
st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)

render_header("Elegibilidade")
render_sidebar()

with st.expander("Configuração da **Agenda de Serviços**", expanded = False):
    st.markdown(f"""
    Antes de utilizar o sistema, verifique se a **Agenda de Serviços** do TASY HTML5 está configurada na seguinte ordem:

    1. **Hora**
    2. **Paciente**
    3. **Convênio**
    4. **Categoria**
    5. (Demais campos podem estar em qualquer posição)
    6. **Sessões**

    > **Observação:** Os campos entre **Categoria** e **Sessões** podem variar conforme a configuração da agenda. Entretanto, o campo **Sessões** deve permanecer sempre como o último item exibido, pois ele é utilizado pelo sistema para identificar o término das informações de cada paciente.

    ---

    ### ⚠️ Importante

    Os campos em **negrito** são utilizados pelo sistema para identificar e tratar os pacientes.

    Caso a ordem dos campos esteja diferente da listada acima, o sistema poderá:
    - importar informações incorretas;
    - misturar dados dos pacientes;
    - ou não funcionar corretamente.

    ---

    ## ❓ Dados não identificados

    O sistema foi desenvolvido para automatizar o tratamento dos dados dos pacientes de forma confiável, auxiliando no processo de elegibilidade e otimizando o tempo de execução dessa tarefa.  

    No entanto, como qualquer sistema automatizado, podem ocorrer situações em que determinadas informações não sejam identificadas corretamente. Nesses casos, o campo será preenchido com:

    ```text
    ?
    ```

    Isso normalmente acontece quando:
    - a ordem dos campos no TASY está diferente da configuração esperada;
    - algum paciente possui informações incompletas;
    - ou o TASY retorna os dados de forma inconsistente.

    ---

    ## ✔️ Conferência Final

    Mesmo com a automação, é altamente recomendado conferir a planilha final antes do envio.

    Verifique principalmente:
    - pacientes sem convênio;
    - categorias incorretas;
    - campos preenchidos com `"?"`;
    - e possíveis inconsistências de data ou unidade.
    """)

with st.expander("Utilizando o **Sistema de Elegibilidade**", expanded = False):
    st.markdown(f"""
    # Passo a Passo

    1. Na interface, selecione:
    - a **data**
    - e a **unidade** que será analisada.

    2. No **TASY HTML5**, selecione os pacientes:
    - começando pela primeira célula da coluna **Hora**
    - até a última célula da coluna **Sessões**.

    3. Copie os pacientes selecionados e cole-os no **Campo de Texto** do sistema.

    4. Clique no botão **Adicionar Dados**.

    5. Repita o processo até finalizar todos os períodos, datas e unidades necessárias.

    6. Ao finalizar, clique em:
    - **Baixar Planilha** → para gerar a planilha de controle.
    - **Baixar Scripts** → para gerar os arquivos de solicitação de elegibilidade.

    ---

    # Explicando a Interface

    ## 📅 Data
    Permite selecionar a data dos pacientes.

    A data pode ser:
    - digitada manualmente;
    - ou selecionada pelo calendário.

    Após finalizar um dia, utilize os botões:
    - **"⬅️"** → voltar um dia;
    - **"➡️"** → avançar um dia.

    Isso evita precisar abrir o calendário repetidamente.

    ---

    ## 🏥 Unidade
    Permite selecionar a unidade dos pacientes:
    - **Itaim**
    - **Brasília III**

    Antes de adicionar os dados, confirme se a unidade selecionada está correta.

    ---

    ## 📝 Campo de Texto
    Área utilizada para colar os pacientes copiados do **TASY HTML5**.

    ⚠️ **Importante:**  
    Não cole todos os pacientes de uma única vez, pois isso pode travar o sistema.

    O ideal é separar por períodos, por exemplo:
    - Manhã Masculino
    - Manhã Feminino
    - Intermediário Masculino
    - Intermediário Feminino
    - Brasília Masculino
    - Brasília Feminino

    ---

    ## ➕ Adicionar Dados
    Botão azul responsável por:
    - ler os dados colados;
    - tratar as informações;
    - salvar os pacientes na tabela.

    Após adicionar os dados, uma visualização será exibida abaixo da interface.

    ---

    ## ↩️ Desfazer
    Botão vermelho utilizado para desfazer a última ação realizada.

    Útil em casos como:
    - data incorreta;
    - unidade incorreta;
    - pacientes adicionados errados.

    Funciona de forma semelhante ao **Ctrl + Z**.

    ---

    ## 📊 Baixar Planilha
    Botão verde responsável por gerar a planilha de controle de elegibilidade.

    A planilha é separada por:
    - Convênios
    - Brasília III
    - Itaú
    - Bradesco
    - Mediservice
    - Santander

    ---

    ## 🌐 Como abrir a planilha

    Como o setor não possui acesso ao Microsoft Excel, a planilha deverá ser aberta pelo Google Sheets.

    Acesse o:
    [Google Sheets](https://docs.google.com/spreadsheets/u/0/)

    Depois:

    1. Crie uma **Planilha em Branco**
    2. Clique em:
    - **Arquivo**
    - **Abrir**
    - **Upload**
    - **Procurar**
    3. Selecione a planilha baixada na pasta **Downloads**

    O Google Sheets abrirá automaticamente a planilha com todas as abas separadas.

    ---

    ## 📄 Baixar Scripts
    Botão ciano utilizado para gerar os arquivos de solicitação de elegibilidade.

    Os arquivos serão enviados para:
    - Andréia
    - Central de Guias
    - Equipe de Brasília

    ---

    ## 🗑️ Recomeçar
    Botão amarelo utilizado para apagar todos os dados inseridos até o momento.

    ⚠️ Esta ação não pode ser desfeita.
    """)

#region ESTADO GLOBAL

# Dias da semana em que o Check-up é realizado + Domingo para tratativa de erro
DIAS_PTBR = {
    "Monday": "Segunda-feira", 
    "Tuesday": "Terça-feira", 
    "Wednesday": "Quarta-feira", 
    "Thursday": "Quinta-feira", 
    "Friday": "Sexta-feira", 
    "Saturday": "Sábado", 
    "Sunday": "Domingo"
}


if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()
if "historico" not in st.session_state:
    st.session_state.historico = []
if "area" not in st.session_state:
    st.session_state.area = ""
if "msg" not in st.session_state:
    st.session_state.msg = ""
if "btn_rollback" not in st.session_state:
    st.session_state.btn_rollback = True
if "btn_finalizar" not in st.session_state:
    st.session_state.btn_finalizar = True
if "btn_exportar" not in st.session_state:
    st.session_state.btn_exportar = True
if "campo_data" not in st.session_state:
    st.session_state.campo_data = date.today()
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = b""
if "zip_scripts" not in st.session_state:
    st.session_state.zip_scripts = b""

def carregar_css(css_path):
    with open(css_path) as f:
        st.html(f"<style>{f.read()}</style>")

css_path = pathlib.Path(BASE_DIR / "pages" / "assets" / "style.css")
carregar_css(css_path)

#endregion

#region FUNCIONALIDADES

def dia_anterior():
    st.session_state.campo_data -= timedelta(days=1)

def dia_seguinte():
    st.session_state.campo_data += timedelta(days=1)

def adicionar():    
    texto = st.session_state.area.strip()

    if not texto:
        st.session_state.msg = "⚠️ Cole algum dado antes de adicionar!"
        return
    
    if not st.session_state.campo_data:
        st.session_state.msg = "⚠️ Selecione uma data antes de adicionar!"
        return

    data_selecionada = st.session_state.campo_data
    dia_mes = data_selecionada.strftime("%d/%m")
    dia_semana = DIAS_PTBR[data_selecionada.strftime("%A")]

    if dia_semana == "Domingo":
        st.session_state.msg = "🚨 ATENÇÃO 🚨 Check-up não é realizado aos domingos!"
        return
    
    novo_df = parse_checkup(texto)
    novo_df["Data"] = f"{dia_mes} ({dia_semana})"
    novo_df["_unidade"] = st.session_state.spinner_unidade

    chave = ["Paciente", "Convênio"]
    novo_df = novo_df.drop_duplicates(subset = chave, keep = "first")

    if not st.session_state.df.empty:
        ja_existentes = st.session_state.df[chave].apply(tuple, axis=1)
        chaves_novas = novo_df[chave].apply(tuple, axis=1)
        duplicatas = novo_df[chaves_novas.isin(ja_existentes)]

        if not duplicatas.empty:
            nomes = ", ".join(duplicatas["Paciente"].tolist())
            novo_df = novo_df[~chaves_novas.isin(ja_existentes)]

            if novo_df.empty:
                st.session_state.msg = (
                    f"⚠️ {len(duplicatas)} duplicata(s) ignorada(s): {nomes}"
                )
                return

    st.session_state.historico.append(st.session_state.df.copy())
    st.session_state.df = pd.concat([st.session_state.df, novo_df], ignore_index=True)
    st.session_state.area = ""
    st.session_state.btn_finalizar = False
    st.session_state.btn_rollback = False

    total = len(st.session_state.df)
    novos_count = len(novo_df)

    st.session_state.msg = (
        f"✅ {novos_count} paciente(s) adicionado(s) ── Total acumulado: {total}\n\n"
        f"📍 Pacientes por unidade:\n\n"
        f"Itaim: {(st.session_state.df['_unidade'] == 'Itaim').sum()}\n\n"
        f"Brasília III: {(st.session_state.df['_unidade'] == 'Brasília III').sum()}"
    )

def rollback():
    if not st.session_state.historico:
        st.session_state.msg = "⚠️ Nada para desfazer!"
        return

    st.session_state.df = st.session_state.historico.pop()

    if st.session_state.df.empty:
        st.session_state.btn_finalizar = True
        st.session_state.btn_exportar = True

    if not st.session_state.historico:
        st.session_state.btn_rollback = True

    total = len(st.session_state.df)
    st.session_state.msg = f"↩️ Último lote desfeito ── total acumulado: {total}"

def finalizar():
    st.session_state.msg = f"📋 Total final: {len(st.session_state.df)} paciente(s)"

    colunas = ["Data", "Hora", "Paciente", "Convênio", "Categoria", "Status"]
    st.session_state.df = st.session_state.df[colunas + ["_unidade"]]
    st.session_state.df = st.session_state.df.sort_values(by = ["Data", "Hora"]).reset_index(drop=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        df_bsb = st.session_state.df[st.session_state.df["_unidade"] == "Brasília III"].drop(columns="_unidade")
        df_bsb_conv = df_bsb.copy()
        df_bsb_conv["_convenio"] = df_bsb_conv["Convênio"].apply(identificar_convenio)
        df_bsb_conv = df_bsb_conv[df_bsb_conv["_convenio"].notna()]
        df_bsb_conv = df_bsb_conv.drop(columns="_convenio")
        if not df_bsb_conv.empty:
            df_bsb_conv.to_excel(writer, sheet_name = "Brasília III", index = False)
            worksheet = writer.sheets["Brasília III"]
            estilizar_header(worksheet)
            estilizar_status(worksheet)

        df_itaim = st.session_state.df[st.session_state.df["_unidade"] == "Itaim"].drop(columns = "_unidade").copy()
        if not df_itaim.empty:
            df_itaim["_sheet"] = df_itaim["Convênio"].apply(identificar_empresa)

            df_convenios = df_itaim[df_itaim["_sheet"].isna()].drop(columns = "_sheet").copy()
            df_convenios["_convenio"] = df_convenios["Convênio"].apply(identificar_convenio)
            df_convenios = df_convenios[df_convenios["_convenio"].notna()]
            df_convenios.drop(columns = "_convenio").to_excel(writer, sheet_name = "Convênios", index = False)
            worksheet = writer.sheets["Convênios"]
            estilizar_header(worksheet)
            estilizar_status(worksheet)

            for empresa in df_itaim["_sheet"].dropna().unique():
                df_itaim[df_itaim["_sheet"] == empresa].drop(columns = "_sheet").to_excel(writer, sheet_name = empresa, index = False)
                worksheet = writer.sheets[empresa]
                estilizar_header(worksheet)
                estilizar_status(worksheet)

    st.session_state.excel_bytes = output.getvalue()
    st.session_state.btn_exportar = False

def exportar():
    st.session_state.msg = "📄 Gerando arquivos de elegibilidade..."
    st.session_state.arquivos_txt = {}

    df = st.session_state.df.copy()
    df["_sheet"] = df["Convênio"].apply(identificar_empresa)

    df_bsb = df[df["_unidade"] == "Brasília III"].drop(columns="_sheet").copy()
    df_bsb["_convenio_nome"] = df_bsb["Convênio"].apply(identificar_convenio)
    df_bsb = df_bsb[df_bsb["_convenio_nome"].notna()].drop(columns="_convenio_nome")  # <- filtro aplicado
    if not df_bsb.empty:
        st.session_state.arquivos_txt["Brasilia_III.txt"] = gerar_txt_brasilia(df_bsb)

    df_itaim = df[df["_unidade"] == "Itaim"].copy()

    for empresa in df_itaim["_sheet"].dropna().unique():
        df_empresa = df_itaim[df_itaim["_sheet"] == empresa].drop(columns="_sheet").copy()
        st.session_state.arquivos_txt[f"{empresa}.txt"] = gerar_txt_empresa(df_empresa, empresa)

    df_conv_itaim = df_itaim[df_itaim["_sheet"].isna()].drop(columns="_sheet").copy()
    df_conv_itaim["_convenio_nome"] = df_conv_itaim["Convênio"].apply(identificar_convenio)
    df_conv_itaim = df_conv_itaim[df_conv_itaim["_convenio_nome"].notna()]  # <- filtro aplicado
    if not df_conv_itaim.empty:
        st.session_state.arquivos_txt["Convenios_Itaim.txt"] = gerar_txt_convenios(df_conv_itaim)

    zip_buffer = BytesIO()
    with ZipFile(zip_buffer, "w") as zip_file:
        for nome, conteudo in st.session_state.arquivos_txt.items():
            zip_file.writestr(nome, conteudo)

    zip_buffer.seek(0)
    st.session_state.zip_scripts = zip_buffer.getvalue()
    st.session_state.msg = "✅ Arquivos exportados!"

def limpar():
    st.session_state.df = pd.DataFrame()
    st.session_state.historico = []
    st.session_state.area = ""
    st.session_state.btn_finalizar = True
    st.session_state.btn_rollback = True
    st.session_state.btn_exportar = True
    st.session_state.txt_brasilia = None
    st.session_state.txt_convenios = None
    st.session_state.excel_bytes = b""
    st.session_state.zip_scripts = b""
    st.session_state.msg = "🗑️ Dados apagados. Pode começar do zero!"

#endregion

#region DATA E UNIDADE

col_ant, col_data, col_prox, col_unidade = st.columns([0.5, 1.5, 0.5, 1.5])

with col_ant:
    st.button(
        "⬅️", 
        on_click = dia_anterior,
        type = "primary"
    )

with col_data:
    st.date_input(
        "Data",
        value = date.today(),
        format = "DD/MM/YYYY", 
        key = "campo_data"
    )

with col_prox:
    st.button(
        "➡️", 
        on_click = dia_seguinte,
        type = "primary"
    )

with col_unidade:
    st.selectbox(
        "Unidade", 
        ["Itaim", "Brasília III"], 
        key = "spinner_unidade"
    )

#endregion

#region CAMPO DE TEXTO

st.text_area(
    "Cole os dados aqui (Ctrl+V ou Win+V)...",
    height = 300,
    key = "area"
)

#endregion

#region BOTÕES DE EDIÇÃO

col_add, col_roll, col_end, col_export, col_clear = st.columns(5)

with col_add:
    st.button(
        "➕ Adicionar Dados",
        use_container_width = True,
        on_click = adicionar
    )

with col_roll:
    st.button(
        "↩️ Desfazer", 
        key = "btn_roll", 
        use_container_width = True,
        disabled = st.session_state.get("btn_rollback", True), 
        on_click = rollback
    )

with col_end:
    if st.button(
        "📊 Gerar Planilha",
        use_container_width=True,
        disabled=st.session_state.get("btn_finalizar", True),
    ):
        finalizar()
        st.rerun()

if st.session_state.excel_bytes:
    st.download_button(
        "⬇️ Baixar Planilha",
        data=st.session_state.excel_bytes,
        file_name="Elegibilidade.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

with col_export:
    if st.button(
        "📄 Gerar Scripts",
        use_container_width=True,
        disabled=st.session_state.get("btn_exportar", True),
    ):
        exportar()
        st.rerun()

if st.session_state.zip_scripts:
    st.download_button(
        "⬇️ Baixar Scripts",
        data = st.session_state.zip_scripts,
        file_name = "Scripts.zip",
        mime = "application/zip",
        use_container_width = True,
    )

with col_clear:
    st.button(
        "🗑️ Recomeçar", 
        key = "btn_clear", 
        use_container_width = True, 
        on_click = limpar
    )

#endregion

if st.session_state.msg:
    st.write(st.session_state.msg)

if not st.session_state.df.empty:
    st.dataframe(st.session_state.df)