# Definindo bibliotecas que utilizaremos
# pip install pandas
# pip install openpyxl
# pip install tabulate
# pip install pandas matplotlib openpyxl
import pandas as pd
from tabulate import tabulate
import matplotlib.pyplot as plt


# Definir onde está a planilha Excel 
arquivo_excel = r'D:\\Caroline\\Analise-de-Dados\\arquivos\\excelEmrpesa.xlsx'

# Definindo a planilha como um DataFrame a ser manipulado
#df = pd.read_excel(arquivo_excel, sheet_name="Outubro - 2024")
df = pd.read_excel(arquivo_excel, sheet_name="Novembro - 2024")


def grafico_pizza_idades():
    # Contar quantas vezes cada idade se repete
    coluna_idade = df['Idade']
    contagem_idades = coluna_idade.value_counts()

    # Exibir as contagens das idades
    print("Contagem das idades:")
    print(tabulate(contagem_idades.reset_index(), headers=['Idade', 'Vezes repetidas'], tablefmt='pretty'))
    
    # Adicionar a palavra "anos" aos rótulos das idades
    labels = [f"{idade} anos" for idade in contagem_idades.index]

    # Gráfico de pizza
    plt.figure(figsize=(10, 10))  
    plt.pie(contagem_idades, 
            labels=labels,  
            autopct='%1.1f%%', 
            startangle=90, 
            colors=plt.cm.Paired.colors,
            wedgeprops={'edgecolor': 'black', 'linewidth': 1.5},  # 
            pctdistance=0.85)  

    # Ajustar a disposição dos rótulos
    plt.title('Distribuição das Idades', loc='left')
    plt.axis('equal')  

    # Ajustar a distância dos rótulos
    plt.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.05)  

    plt.show()

def idade_repetida():
    # Sabendo qual idade mais se repete na coluna B (Idade) do Excel
    coluna_idade = df['Idade']
    # Contar quantas vezes cada idade se repete
    contagem_idades = coluna_idade.value_counts().reset_index()
    contagem_idades.columns = ['Idade', 'Vezes repetidas']
    contagem_idades = contagem_idades.sort_values(by='Vezes repetidas', ascending=False)  
    top_3_idades = contagem_idades.head(3)

    # Exibir o top 3 das idades mais repetidas
    print("Top 3 idades mais repetidas:")
    print(tabulate(top_3_idades, headers='keys', tablefmt='pretty', showindex=False))

    # Gráfico de barras para as idades mais repetidas
    plt.figure(figsize=(8, 6))
    plt.bar(top_3_idades['Idade'].astype(str), top_3_idades['Vezes repetidas'], color='skyblue')
    plt.title('Top 3 Idades Mais Repetidas')
    plt.xlabel('Idade')
    plt.ylabel('Vezes Repetidas')
    plt.xticks(rotation=0)
    plt.show()

def servico_repetido():
    coluna_servico = df['Serviço']
    # Contar quantas vezes cada serviço se repete
    contagem_servicos = coluna_servico.value_counts().reset_index()
    contagem_servicos.columns = ['Serviço', 'Vezes utilizados']
    contagem_servicos = contagem_servicos.sort_values(by='Vezes utilizados', ascending=False)
    top_3_servicos = contagem_servicos.head(3)

    # Exibir o top 3 dos serviços mais utilizados
    print("\nTop 3 serviços mais utilizados:")
    print(tabulate(top_3_servicos, headers='keys', tablefmt='pretty', showindex=False))

    # Gráfico de barras para os serviços mais utilizados
    plt.figure(figsize=(8, 6))
    plt.bar(top_3_servicos['Serviço'], top_3_servicos['Vezes utilizados'], color='lightcoral')
    plt.title('Top 3 Serviços Mais Utilizados')
    plt.xlabel('Serviço')
    plt.ylabel('Vezes Utilizados')
    plt.xticks(rotation=45)
    plt.show()

def data_repetida():
    coluna_data = df['Data']
    contagem_datas = coluna_data.value_counts().reset_index()
    contagem_datas.columns = ['Data', 'Vezes repetidas']
    contagem_datas = contagem_datas.sort_values(by='Vezes repetidas', ascending=False)
    top_3_datas = contagem_datas.head(3)

    # Exibir o top 3 das datas mais repetidas
    print("\nTop 3 das datas mais repetidas:")
    print(tabulate(top_3_datas, headers='keys', tablefmt='pretty', showindex=False))

    # Gráfico de barras para as datas mais repetidas
    plt.figure(figsize=(8, 6))
    plt.bar(top_3_datas['Data'].astype(str), top_3_datas['Vezes repetidas'], color='lightgreen')
    plt.title('Top 3 Datas Mais Repetidas')
    plt.xlabel('Data')
    plt.ylabel('Vezes Repetidas')
    plt.xticks(rotation=45)
    plt.show()

def data_repetida_semana(df):
    coluna_data = df['Data']
    contagem_datas = coluna_data.value_counts().reset_index()
    contagem_datas.columns = ['Data', 'Vezes repetidas']
    contagem_datas = contagem_datas.sort_values(by='Vezes repetidas', ascending=False)
    contagem_datas['Dia da Semana'] = pd.to_datetime(contagem_datas['Data'], errors='coerce').dt.day_name()

    # Traduzindo os dias para português
    traducao_dias = {
        'Monday': 'Segunda-feira',
        'Tuesday': 'Terça-feira',
        'Wednesday': 'Quarta-feira',
        'Thursday': 'Quinta-feira',
        'Friday': 'Sexta-feira',
        'Saturday': 'Sábado',
        'Sunday': 'Domingo'
    }
    contagem_datas['Dia da Semana'] = contagem_datas['Dia da Semana'].map(traducao_dias)
    top_3_datas = contagem_datas.head(3)

    # Exibir o top 3 das datas mais repetidas com o dia da semana correspondente
    print("\nTop 3 das datas mais repetidas com o dia da semana correspondente:")
    print(tabulate(top_3_datas, headers='keys', tablefmt='pretty', showindex=False))

    # Gráfico de barras para as datas mais repetidas com o dia da semana
    plt.figure(figsize=(8, 6))
    plt.bar(top_3_datas['Data'].astype(str), top_3_datas['Vezes repetidas'], color='lightblue')
    plt.title('Top 3 Datas Mais Repetidas com Dia da Semana')
    plt.xlabel('Data')
    plt.ylabel('Vezes Repetidas')
    plt.xticks(rotation=45)
    plt.show()
    
    
def grafico_pizza_dia_semana():
    # Garantir que a coluna 'Data' esteja no formato datetime
    df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
    
    # Obter o dia da semana em inglês
    df['Dia da Semana'] = df['Data'].dt.day_name()
    
    # Dicionário de tradução para português
    traducao_dias = {
        'Monday': 'Segunda-feira',
        'Tuesday': 'Terça-feira',
        'Wednesday': 'Quarta-feira',
        'Thursday': 'Quinta-feira',
        'Friday': 'Sexta-feira',
        'Saturday': 'Sábado',
        'Sunday': 'Domingo'
    }
    
    # Aplicar o mapeamento de tradução para o português
    df['Dia da Semana'] = df['Dia da Semana'].map(traducao_dias)
    
    # Contar quantas vezes cada dia da semana se repete
    contagem_dias_semana = df['Dia da Semana'].value_counts()

    # Exibir as contagens dos dias da semana
    print("Contagem dos dias da semana:")
    print(tabulate(contagem_dias_semana.reset_index(), headers=['Dia da Semana', 'Vezes Repetidas'], tablefmt='pretty'))

    # Adicionar "dias" aos rótulos dos dias da semana
    labels = [f"{dia}" for dia in contagem_dias_semana.index]

    # Gráfico de pizza
    plt.figure(figsize=(10, 10))  
    plt.pie(contagem_dias_semana, 
            labels=labels,  
            autopct='%1.1f%%', 
            startangle=90, 
            colors=plt.cm.Paired.colors,
            wedgeprops={'edgecolor': 'black', 'linewidth': 1.5}, 
            pctdistance=0.85)  

    # Ajustar a disposição dos rótulos
    plt.title('Distribuição dos Dias da Semana', loc='left')  
    plt.axis('equal')  

    # Ajustar a distância dos rótulos
    plt.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.05)  # Ajuste da área visível para o gráfico

    plt.show()
    

# Iniciar as funções
grafico_pizza_idades()
idade_repetida()
servico_repetido() 
data_repetida()
data_repetida_semana(df)
grafico_pizza_dia_semana()