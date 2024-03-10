import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.dates as mdates
import numpy as np

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

import plotly.graph_objects as go

APP_TITLE = 'Agência Turmalina'
APP_SUB_TITLE = 'Aqui a porra toda acontece'

# Função para carregar os nomes das abas da planilha
@st.cache_data
def load_sheet_names(filepath):
    xls = pd.ExcelFile(filepath)
    return xls.sheet_names

# Função para carregar planilha e deixá-la em cache
@st.cache_data
def load_data(filepath, sheet_name):
    df = pd.read_excel(filepath, index_col='Day', sheet_name=sheet_name)
    return df

# Função para retornar a soma das campanhas
def soma_das_campanhas(df, selected_columns):
    if selected_columns:
        # Filtra o DataFrame para incluir apenas as colunas selecionadas mais a 'Campaign Name'
        df_filtered = df[['Campaign Name'] + selected_columns]
    else:
        # Se nenhuma coluna for selecionada, usa o DataFrame como está
        df_filtered = df
    return df_filtered.groupby('Campaign Name')[selected_columns].sum()

def criar_grafico_barras(df_operacao, coluna_agrupamento, selected_columns):
    # Verifica se a lista de colunas selecionadas não está vazia
    if not selected_columns:
        st.error('Nenhuma coluna foi selecionada.')
        return

    # Usa apenas a primeira coluna selecionada
    primeira_coluna = selected_columns[0]

    # Verifica se a primeira coluna está presente no DataFrame
    if primeira_coluna not in df_operacao.columns:
        st.error(f"A coluna '{primeira_coluna}' não foi encontrada no DataFrame.")
        return

    # Obtém os valores da primeira coluna selecionada
    values = df_operacao[primeira_coluna]

    plt.figure(figsize=(10, 8))

    # Determina as cores das barras
    cores = ['blue' if x == values.max() else 'lightgray' for x in values]

    # Cria o gráfico de barras horizontais
    plt.barh(values.index, values, color=cores)

    plt.xlabel('Valores')
    plt.ylabel(coluna_agrupamento)
    plt.title(f'{primeira_coluna}: ')

    plt.tight_layout()
    st.pyplot(plt.gcf())

def criar_grafico_pizza(df_operacao, coluna):
    # Assume que df_operacao é um DataFrame resultante de uma operação de agregação
    # e coluna é a coluna selecionada para visualização
    
    # Ordena os valores da coluna selecionada em ordem decrescente
    valores_ordenados = df_operacao[coluna].sort_values(ascending=False)
    
    # Separa os 3 maiores valores
    tres_maiores = valores_ordenados[:3]
    
    # Soma todos os outros valores como 'Outras'
    outras = valores_ordenados[3:].sum()
    tres_maiores['Outras'] = outras  # Adiciona a soma ao DataFrame

    # Prepara os dados para o gráfico
    valores = tres_maiores.values
    labels = tres_maiores.index

    # Cria o gráfico de pizza
    plt.figure(figsize=(8, 8))
    plt.pie(valores, labels=labels, autopct='%1.1f%%', startangle=140)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.title(f'Distribuição de {coluna}')

    st.pyplot(plt.gcf())
    
def fornecer_e_receber_atributos_tabela(df):
    # Seleção das colunas para filtragem na sidebar
    selected_columns = st.sidebar.multiselect('Selecione as colunas para visualizar:', 
                                  options=df.columns[3:].tolist(),
                                  default=df.columns[3:6].tolist()) 
    
    # Opções de agrupamento incluem o índice (ex: "Data") e mais três colunas
    opcoes_agrupamento = ['Data'] + df.columns[0:3].tolist()  # Substitua 'Coluna 1', 'Coluna 2', 'Coluna 3' pelos nomes reais das suas colunas
    
    coluna_agrupamento = st.sidebar.selectbox('Escolha a coluna para agrupamento:', opcoes_agrupamento)
    
    # Converte o índice do DataFrame para uma coluna se "Data" for a opção selecionada para o agrupamento
    if coluna_agrupamento == 'Data':
        df_reset = df.reset_index().rename(columns={df.index.name: 'Data'})
        coluna_agrupamento = 'Data'  # Define 'Data' como a coluna de agrupamento
    else:
        df_reset = df
    
    # Continua a partir daqui com a seleção da operação estatística e a aplicação da operação
    operacao = st.sidebar.selectbox('Escolha a operação:', 
                                    ['Soma', 'Média', 'Desvio Padrão', 'Mediana', 'Máximo', 'Mínimo'])
    
    # Aplica a operação escolhida
    operacoes = {
        'Soma': 'sum',
        'Média': 'mean',
        'Desvio Padrão': 'std',
        'Mediana': 'median',
        'Máximo': 'max',
        'Mínimo': 'min',
    }
    df_operacao = getattr(df_reset.groupby(coluna_agrupamento)[selected_columns], operacoes[operacao])()
    
    # Dividindo a tela em duas colunas
    col1, col2, col3 = st.columns([1, 1, 1])
    
    # Apresenta uma tabela com os resultados da operação escolhida, formatados com 1 casa decimal
    with col1:
        if selected_columns:
            df_display = df_operacao.style.format("{:.1f}")
            st.dataframe(df_display)
        else:
            st.write("Selecione uma ou mais colunas para visualizar os dados.")
    with col2:
        criar_grafico_barras(df_operacao, coluna_agrupamento, selected_columns)
    with col3:
        criar_grafico_pizza(df_operacao, selected_columns[0])   
        
def criar_graficos_simples(df):
    # Seção de gráficos
    st.header('Gráfico Simples')
    column_to_plot = st.sidebar.selectbox('Escolha a coluna para plotar:', df.columns[3:])
    
    fig, ax = plt.subplots()
    
    # Assegura que o índice é do tipo datetime
    df_nonan = df.dropna(subset=[column_to_plot])
    # Plota o gráfico
    sns.lineplot(data=df_nonan, x=df_nonan.index, y=column_to_plot, ax=ax)
    
    # Define o formato do label do eixo x para mostrar uma data por semana no formato "Mes-Dia"
    ax.xaxis.set_major_locator(mdates.WeekdayLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%b-%d'))
    
    # Melhora a apresentação girando os labels e ajustando o espaçamento
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('./grafico.png')
    st.pyplot(fig)

def criar_grafico_velocimetro(df):
    # Calcula a média da coluna "Frequência"
    media_frequencia = df['Frequency'].sum()

    # Define o limite para o aviso
    limite_aviso = 800  # Este é um exemplo, ajuste conforme necessário

    # Cria o gráfico de velocímetro
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = media_frequencia,
        domain = {'x': [0, 1], 'y': [0, 1]},
        title = {'text': "Média de Frequência"},
        gauge = {'axis': {'range': [None, limite_aviso+200]},  # Ajuste o limite superior conforme necessário
                 'steps' : [
                     {'range': [0, limite_aviso], 'color': "lightgray"},
                     {'range': [limite_aviso, limite_aviso+200], 'color': "red"}],  # Ajuste o limite superior conforme necessário
                 'threshold' : {'line': {'color': "red", 'width': 4}, 'thickness': 0.75, 'value': limite_aviso}}))

    # Atualiza o fundo para branco
    fig.update_layout(paper_bgcolor = "white", font = {'color': "darkblue", 'family': "Arial"})

    st.plotly_chart(fig)

# Envio de email       
def enviar_imagem_email(destinatario, imagem_path):
    # Configurações do servidor de e-mail
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_user = "thiago.luiz.lx@gmail.com"
    smtp_password = "qwhcnikrteeqwgwn"

    # Criando a mensagem
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = destinatario
    msg['Subject'] = "Gráfico da Agência Turmalina"

    # Anexando a imagem
    with open(imagem_path, 'rb') as file:
        img = MIMEImage(file.read())
        img.add_header('Content-Disposition', 'attachment; filename="grafico.png"')
        msg.attach(img)

    # Enviando o e-mail
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.sendmail(smtp_user, destinatario, msg.as_string())
    server.quit()

    print("E-mail enviado com sucesso para", destinatario)
    
# Assegura que o Seaborn esteja configurado para um estilo mais bonito
sns.set_theme(style="whitegrid")

# Interface do Streamlit
def main():
    st.set_page_config(page_title=APP_TITLE, layout='wide')

    filepath = './SEP.xlsx'
    
    # Sidebar
    st.sidebar.header(APP_SUB_TITLE)
    # Carregar nomes das abas
    sheet_names = load_sheet_names(filepath)
    
    # Selecionador de aba da planilha na sidebar
    sheet_name = st.sidebar.selectbox('Selecione a aba da planilha:', sheet_names)

    # Carregar dados
    df = load_data(filepath, sheet_name)
    
    # Criar tabelas de atributos
    fornecer_e_receber_atributos_tabela(df)
    
    # Dividindo a tela em duas colunas
    col1,col2 = st.columns([1, 1])

    # Criar um gráfico simples
    with col1:
        criar_graficos_simples(df)
    with col2:
        criar_grafico_velocimetro(df)   
        
    # Opção para o usuário inserir o endereço de e-mail na sidebar
    email_destino = st.sidebar.text_input("Insira o e-mail para enviar o gráfico:")
    
    # Botão para enviar o e-mail, agora localizado na sidebar
    if st.sidebar.button('Enviar Gráfico por E-mail'):
        enviar_imagem_email(email_destino, './grafico.png')
        st.sidebar.success('Gráfico enviado com sucesso!')


if __name__ == "__main__":
    main()
