import pandas as pd
import streamlit as st
from io import BytesIO
import warnings
import plotly.express as px
import urllib.parse
import base64  # Usado para embutir a imagem do logo

warnings.filterwarnings('ignore')

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Dashboard de Ativos Contábeis",
    page_icon=" ",
    layout="wide"
)

# --- FUNÇÕES DE LÓGICA (sem alteração) ---


def padronizar_nome_filial(nome_filial):
    if not isinstance(nome_filial, str):
        return "Não Identificado"
    nome_upper = nome_filial.upper().strip()
    mapa_nomes = {
        "GENERAL WATER": "General Water S/A", "GW S/A": "General Water S/A",
        "G W AGUAS": "GW Águas", "GW ÁGUAS": "GW Águas",
        "GW SANEAMENTO": "GW Saneamento", "GW SANEA": "GW Saneamento",
        "GW SISTEMAS": "GW Sistemas", "GW SISTEM": "GW Sistemas",
        "MATRIZ": "GW Sistemas Matriz"
    }
    return mapa_nomes.get(nome_upper, nome_filial)


def converter_valor(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        valor_str = str(valor).replace('R$', '').strip()
        if ',' in valor_str and '.' in valor_str:
            valor_str = valor_str.replace('.', '')
        valor_str = valor_str.replace(',', '.')
        return float(valor_str)
    except (ValueError, TypeError):
        return 0.0


def formatar_valor(valor):
    try:
        return f"R$ {float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except (ValueError, TypeError):
        return "R$ 0,00"


def corrigir_filiais_nao_identificadas(df_arquivo):
    if df_arquivo.empty:
        return df_arquivo
    contagem_filiais = df_arquivo[df_arquivo['Filial']
                                  != 'Não Identificado']['Filial'].mode()
    if not contagem_filiais.empty:
        filial_predominante = contagem_filiais[0]
        df_arquivo['Filial'] = df_arquivo['Filial'].replace(
            'Não Identificado', filial_predominante)
    return df_arquivo


def processar_planilha(file):
    try:
        xl = pd.ExcelFile(file)
        dados_processados = []
        for sheet_name in xl.sheet_names:
            sheet_df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            categoria_atual, filial_atual = "Não Identificado", "Não Identificado"
            for _, row in sheet_df.iterrows():
                if pd.notna(row.iloc[0]) and str(row.iloc[0]).startswith('1.2.3.'):
                    categoria_atual = str(row.iloc[1]).strip() if pd.notna(
                        row.iloc[1]) else str(row.iloc[0]).strip()
                elif pd.notna(row.iloc[0]) and 'Filial :' in str(row.iloc[0]):
                    nome_extraido = str(row.iloc[0]).split(
                        'Filial :')[-1].split(' - ')[-1].strip()
                    filial_atual = padronizar_nome_filial(nome_extraido)
                elif pd.notna(row.iloc[0]) and str(row.iloc[0]).strip() == 'R$':
                    valores = [converter_valor(v) for v in row.iloc[1:8]]
                    valor_atualizado = valores[2] if len(valores) > 2 else 0
                    deprec_acumulada = valores[5] if len(valores) > 5 else 0
                    if valor_atualizado > 0:
                        dados_processados.append({
                            'Arquivo': file.name, 'Filial': filial_atual, 'Categoria': categoria_atual,
                            'Valor Original': valores[1] if len(valores) > 1 else 0,
                            'Valor Atualizado': valor_atualizado,
                            'Deprec. no mês': valores[3] if len(valores) > 3 else 0,
                            'Deprec. no Exercício': valores[4] if len(valores) > 4 else 0,
                            'Deprec. Acumulada': deprec_acumulada,
                            'Valor Residual': valor_atualizado - deprec_acumulada
                        })
        if dados_processados:
            df_final = pd.DataFrame(dados_processados)
            return corrigir_filiais_nao_identificadas(df_final), None
        return None, f"Nenhum dado relevante encontrado em {file.name}."
    except Exception as e:
        return None, f"Erro crítico ao processar {file.name}: {e}"


# --- ESTRUTURA DA APLICAÇÃO ---
st.title("Dashboard de Ativos Contábeis")

with st.sidebar:
    try:
        st.image("logo_GW.png", width=200)
    except Exception:
        st.title("General Water")

    st.header("Instruções")
    st.info("1. **Carregue** os arquivos.\n2. **Aguarde** o processamento.\n3. **Filtre** e analise os dados.\n4. **Explore** os gráficos interativos.\n5. **Baixe** o relatório.")

    st.header("Ajuda & Suporte")
    email1 = "bruce@generalwater.com.br"
    email2 = "nathalia.vidal@generalwater.com.br"
    mensagem_inicial = "Olá, preciso de ajuda com o Dashboard de Ativos Contábeis."
    link_teams = f"https://teams.microsoft.com/l/chat/0/0?users={email1},{email2}&message={urllib.parse.quote(mensagem_inicial)}"
    st.markdown(f'<a href="{link_teams}" target="_blank" style="display: inline-block; padding: 10px 20px; background-color: #4B53BC; color: white; text-align: center; text-decoration: none; border-radius: 5px; font-weight: bold;">Abrir Chat no Teams</a>', unsafe_allow_html=True)

uploaded_files = st.file_uploader("Escolha os arquivos Excel de ativos", type=[
                                  'xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    all_data, errors = [], []
    progress_bar = st.progress(0, text="Iniciando...")
    for i, file in enumerate(uploaded_files):
        progress_bar.progress((i + 1) / len(uploaded_files),
                              text=f"Processando: {file.name}")
        dados, erro = processar_planilha(file)
        if dados is not None and not dados.empty:
            all_data.append(dados)
        if erro:
            errors.append(erro)

    if all_data:
        dados_combinados = pd.concat(all_data, ignore_index=True)
        st.success(
            f"Processamento concluído! {len(all_data)} arquivo(s) válidos.")

        # --- CONTAINER PARA O CONTEÚDO DO PDF ---
        # Envolvemos o conteúdo que queremos imprimir em um container com um ID específico
        with st.container():
            st.markdown('<div id="conteudo-para-pdf">', unsafe_allow_html=True)

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Registros Filtrados", f"{len(dados_combinados):,}")
            col2.metric("Valor Total Atualizado", formatar_valor(
                dados_combinados["Valor Atualizado"].sum()))
            col3.metric("Depreciação Acumulada", formatar_valor(
                dados_combinados["Deprec. Acumulada"].sum()))
            col4.metric("Valor Residual Total", formatar_valor(
                dados_combinados["Valor Residual"].sum()))

            # Filtros (movidos para dentro do container para aparecerem no PDF, se desejado)
            col1, col2, col3 = st.columns(3)
            arquivos_options = sorted(dados_combinados['Arquivo'].unique())
            filiais_options = sorted(dados_combinados['Filial'].unique())
            categorias_options = sorted(dados_combinados['Categoria'].unique())
            with col1:
                selecao_arquivo = st.multiselect(
                    "Arquivo:", ["Selecionar Todos"] + arquivos_options, default="Selecionar Todos")
            with col2:
                selecao_filial = st.multiselect(
                    "Filial:", ["Selecionar Todos"] + filiais_options, default="Selecionar Todos")
            with col3:
                selecao_categoria = st.multiselect(
                    "Categoria:", ["Selecionar Todos"] + categorias_options, default="Selecionar Todos")

            filtro_arquivo = arquivos_options if "Selecionar Todos" in selecao_arquivo else selecao_arquivo
            filtro_filial = filiais_options if "Selecionar Todos" in selecao_filial else selecao_filial
            filtro_categoria = categorias_options if "Selecionar Todos" in selecao_categoria else selecao_categoria
            dados_filtrados = dados_combinados[
                (dados_combinados['Arquivo'].isin(filtro_arquivo)) &
                (dados_combinados['Filial'].isin(filtro_filial)) &
                (dados_combinados['Categoria'].isin(filtro_categoria))
            ]

            # Gráfico
            opcoes_eixo_y = ["Valor Atualizado",
                             "Deprec. Acumulada", "Valor Residual"]
            col_graf1, col_graf2, col_graf3 = st.columns(3)
            with col_graf1:
                tipo_grafico = st.selectbox("Escolha o Tipo de Gráfico:", [
                                            "Barras", "Pizza", "Linhas"])
            with col_graf2:
                eixo_x = st.selectbox("Agrupar por (Eixo X):", [
                                      "Filial", "Categoria", "Arquivo"], key="eixo_x_selectbox")
            with col_graf3:
                if tipo_grafico == "Pizza":
                    eixos_y = st.selectbox(
                        "Analisar Valor (Eixo Y):", opcoes_eixo_y, index=0)
                    eixos_y = [eixos_y]
                else:
                    eixos_y = st.multiselect("Analisar Valores (Eixo Y):", opcoes_eixo_y, default=[
                                             "Valor Atualizado", "Valor Residual"])

            if not dados_filtrados.empty and eixo_x and eixos_y:
                dados_para_grafico = dados_filtrados.copy()
                dados_agrupados = dados_para_grafico.groupby(
                    eixo_x)[eixos_y].sum().reset_index()
                titulo = f"Comparativo de Métricas por {eixo_x}"
                fig = None
                if tipo_grafico == "Barras":
                    dados_grafico_melted = pd.melt(dados_agrupados, id_vars=[
                                                   eixo_x], value_vars=eixos_y, var_name='Métrica', value_name='Valor')
                    fig = px.bar(dados_grafico_melted, x=eixo_x, y='Valor', color='Métrica', title=titulo, labels={
                                 eixo_x: eixo_x, 'Valor': "Soma dos Valores", 'Métrica': "Métrica Financeira"}, text_auto='.2s', barmode='group')
                    fig.update_traces(textposition='outside')
                elif tipo_grafico == "Linhas":
                    dados_grafico_melted = pd.melt(dados_agrupados, id_vars=[
                                                   eixo_x], value_vars=eixos_y, var_name='Métrica', value_name='Valor')
                    fig = px.line(dados_grafico_melted, x=eixo_x, y='Valor', color='Métrica', title=titulo, labels={
                                  eixo_x: eixo_x, 'Valor': "Soma dos Valores", 'Métrica': "Métrica Financeira"}, markers=True)
                elif tipo_grafico == "Pizza":
                    metrica_unica = eixos_y[0]
                    titulo_pizza = f"Distribuição de '{metrica_unica}' por {eixo_x}"
                    fig = px.pie(dados_agrupados, names=eixo_x,
                                 values=metrica_unica, title=titulo_pizza, hole=0.3)
                    fig.update_traces(textposition='outside',
                                      textinfo='percent+label')

                if fig:
                    fig.update_layout(uniformtext_minsize=8, uniformtext_mode='hide', margin=dict(
                        t=80, b=50), plot_bgcolor='rgba(0,0,0,0)', legend_title_text='')
                    st.plotly_chart(fig, use_container_width=True)

            # Tabela de dados agregados
            st.markdown("### Dados Agregados")
            colunas_para_somar = ['Valor Atualizado',
                                  'Deprec. Acumulada', 'Valor Residual']
            df_agregado = dados_filtrados.groupby(['Filial', 'Categoria'])[
                colunas_para_somar].sum().reset_index()
            for col in colunas_para_somar:
                df_agregado[col] = df_agregado[col].apply(formatar_valor)
            st.dataframe(df_agregado, use_container_width=True)

            # Fecha o container do PDF
            st.markdown('</div>', unsafe_allow_html=True)

        # --- BOTÕES DE DOWNLOAD ---
        st.markdown("---")
        st.header("Exportar Relatório")

        col_download1, col_download2 = st.columns(2)

        with col_download1:
            output_excel = BytesIO()
            df_display_excel = dados_filtrados.copy()
            for col in ['Valor Original', 'Valor Atualizado', 'Deprec. no mês', 'Deprec. no Exercício', 'Deprec. Acumulada', 'Valor Residual']:
                df_display_excel[col] = df_display_excel[col].apply(
                    formatar_valor)
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                df_display_excel.to_excel(
                    writer, sheet_name='Dados_Filtrados', index=False)
            st.download_button(
                label="Baixar Relatório em Excel",
                data=output_excel.getvalue(),
                file_name="relatorio_ativos_filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with col_download2:
            # Função para carregar a imagem do logo e converter para base64
            def get_image_as_base64(path):
                try:
                    with open(path, "rb") as img_file:
                        return base64.b64encode(img_file.read()).decode()
                except FileNotFoundError:
                    return None

            logo_base64 = get_image_as_base64("logo_GW.png")
            logo_html = f'<img src="data:image/png;base64,{logo_base64}" style="width: 150px; margin-bottom: 20px;">' if logo_base64 else '<h1>General Water</h1>'

            # O botão agora aciona o JavaScript
            st.markdown(f"""
            <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
            <button id="btn-pdf" onclick="gerarPdf( )">Baixar Relatório em PDF</button>
            <script>
                function gerarPdf() {{
                    const elemento = document.getElementById('conteudo-para-pdf');
                    const logoHtml = `{logo_html}`;
                    
                    // Clona o elemento para não modificar o original
                    const elementoClonado = elemento.cloneNode(true);
                    
                    // Cria um container para o cabeçalho
                    const cabecalho = document.createElement('div');
                    cabecalho.innerHTML = logoHtml + "<h1>Relatório de Ativos Contábeis</h1><hr>";
                    
                    // Insere o cabeçalho no topo do elemento clonado
                    elementoClonado.insertBefore(cabecalho, elementoClonado.firstChild);

                    const opt = {{
                        margin:       1,
                        filename:     'relatorio_ativos.pdf',
                        image:        {{ type: 'jpeg', quality: 0.98 }},
                        html2canvas:  {{ scale: 2, useCORS: true }},
                        jsPDF:        {{ unit: 'in', format: 'letter', orientation: 'landscape' }}
                    }};

                    html2pdf().set(opt).from(elementoClonado).save();
                }}
            </script>
            <style>
                #btn-pdf {{
                    width: 100%;
                    padding: 0.5rem 1rem;
                    font-weight: 600;
                    border-radius: 0.5rem;
                    border: 1px solid rgba(49, 51, 63, 0.2);
                    background-color: transparent;
                    color: inherit;
                    cursor: pointer;
                }}
                #btn-pdf:hover {{
                    border-color: #4B53BC;
                    color: #4B53BC;
                }}
            </style>
            """, unsafe_allow_html=True)

    if errors:
        st.warning("Alguns arquivos apresentaram problemas:", icon="❗")
        for error in errors:
            st.error(error)
else:
    st.info("Aguardando o upload dos arquivos para iniciar o processamento.")

st.markdown("---")
st.caption("Desenvolvido para General Water | v26.0 - Suporte via Teams")
