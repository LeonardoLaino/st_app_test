import streamlit as st
import pandas as pd
import re
import numpy as np
import pickle
from io import BytesIO
from PIL import Image
from unidecode import unidecode


RELATIVE_IMG_PATH = './img_instrucoes/'

def substring(texto):
    return re.sub(r'^\w+\s*-\s*', '', texto)

footer = """
<footer style="position: fixed; bottom: 0; width: 100%; background-color: #f8f9fa; text-align: center; padding: 10px;">
    Desenvolvido por Leonardo
    <span style="padding-left: 20px;"></span>
    (Processos Escolares)
    <span style="padding-left: 20px;"></span>
    Versão: 0.2.1
</footer>
"""

st.set_page_config(
    page_title= "Cópias e Impressões",
    page_icon= ':chart_with_upwards_trend:',
    layout= 'wide',
    initial_sidebar_state= 'auto'
)


def instrucoes_copias_impressoes():

    st.title("1. Instruções")
    st.write("""Nesta página você irá encontrar instruções importantes para funcionamento desta ferramenta.""")
    
    st.subheader("""**`Planilha do Papercut`**""")
    st.write("Após baixar o arquivo `.csv` do site do Papercut no **intervalo de data adequado**, execute os passos a seguir:")

    st.markdown(">**Passo 1**: Abra um arquivo novo no Excel e carregue o arquivo `.csv`.")
    st.image(Image.open(RELATIVE_IMG_PATH+"ingestao_csv_papercut.png"))

    st.markdown(">**Passo 2**: Após selecionar o arquivo, clique em `Carregar` na nova janela.")
    st.image(Image.open(RELATIVE_IMG_PATH+"carregar_dados_papercut.png"))

    st.markdown(">**Passo 3**: Após carregado os dados, elimine a **planilha adicional** que estará vazia.")
    st.image(Image.open(RELATIVE_IMG_PATH+"remover_plan_1.png"))

    st.markdown(">**Passo 4**: Na planilha com os dados, **selecione e remova as duas primeiras linhas**.")
    st.image(Image.open(RELATIVE_IMG_PATH+"remover_linhas.png"))

    st.markdown(">**Passo 5**: Selecione o conteúdo da linha 2 e copie no cabeçalho da tabela.")
    st.image(Image.open(RELATIVE_IMG_PATH+"selecione_linha_2.png"))

    st.markdown(">**Passo 6**: Remova a linha 2.")
    st.image(Image.open(RELATIVE_IMG_PATH+"elimine_a_linha_2.png"))

    st.write("**Após executar todos os passos acima, salve o arquivo na extensão padrão do Excel (`.xlsx`)**")
    st.write("**Com o arquivo salvo, acesse a página `2.Relatório` para gerar o relatório de cópias e impressões.**")



def gerar_relatorio_impressoes(ped, pcut):
    # Carregando os dados das impressoras
    with open('./MAPPER_IMPRESSORAS.pickle', 'rb') as f:
        MAPPER_IMPRESSORAS = pickle.load(f)


    # Pegando info das unidades através da impressora
    try:
        pcut['unidade'] = pcut['identificador_de_impressora_fisica'].apply(lambda x: str(x).replace("net://", "")).map(MAPPER_IMPRESSORAS)
    except:
        st.error("Verifique se a coluna 'identificador de impressora fisica' existe na planilha do papercut.")

    # Normalizando as datas (removendo hora:min:seg)
    pcut['data'] = pd.to_datetime(pcut['data']).dt.date

    try:
        ped['data_da_utilizacao'] = pd.to_datetime(ped['data_da_utilizacao']).dt.date
    except:
        st.error("Verifique se existem erros na coluna 'data da utilização'. Isso inclui datas fora do padrão ou valores de dia/mes/ano inválidos.")

    try:
        ped['data_da_solicitacao'] = pd.to_datetime(ped['data_da_solicitacao']).dt.date
    except:
        st.error("Verifique se existem erros na coluna 'data da solicitação'. Isso inclui datas fora do padrão ou valores de dia/mes/ano inválidos.")
    
    # Instanciando a coluna
    ped['total_impressoes_pcut'] = 0

    nomes_ped = np.sort(ped['nome'].unique())

    for nome_iter in nomes_ped:

        ped_temp_aux = ped.loc[ped['nome'] == nome_iter]

        pcut_temp_aux = pcut.loc[pcut['nome_conta_normalizado'] == nome_iter]

        for idx in ped_temp_aux.index:
            
            if ped_temp_aux.empty:
                ped.loc[idx, 'status'] == 'NOME NÃO LOCALIZADO NO PAPERCUT'
                ped.loc[idx, 'documento_corresp_pcut'] = ''
                ped.loc[idx, 'linhas_corresp_pcut'] = ''
                continue

            data_inf, data_sup, unidade = ped_temp_aux['data_da_solicitacao'].loc[idx], ped_temp_aux['data_da_utilizacao'].loc[idx], ped_temp_aux['unidade'].loc[idx]

            # Catching Time not Informed.
            if ((str(data_inf) == 'NaT') or (str(data_sup) == 'NaT')):
                ped.loc[idx, 'status'] = 'DATA NÃO INFORMADA'
                ped.loc[idx, 'documento_corresp_pcut'] = ''
                ped.loc[idx, 'linhas_corresp_pcut'] = ''
                continue

            pcut_temp_iter = pcut_temp_aux.loc[
                (pcut_temp_aux['data']>= data_inf) & (pcut_temp_aux['data'] <= data_sup) & (pcut_temp_aux['unidade'] == unidade)
            ]

            # Caso o df filtrado seja vazio, pula pra próxima iteração
            if pcut_temp_iter.empty:
                ped.loc[idx, 'status'] = 'NÃO ENCONTRADO NO PAPERCUT'
                ped.loc[idx, 'documento_corresp_pcut'] = ''
                ped.loc[idx, 'linhas_corresp_pcut'] = ''
                continue
            
            ped.loc[idx, 'status'] = 'CORRESPONDENCIA ENCONTRADA'
            ped.loc[idx, 'documento_corresp_pcut'] = ', '.join(map(str, pcut_temp_iter['documento']))
            ped.loc[idx, 'total_impressoes_pcut'] = pcut_temp_iter['total_paginas_impressas'].sum()
            ped.loc[idx, 'linhas_corresp_pcut'] = ', '.join(map(str, pcut_temp_iter.index+2))
    
    ped['fl_diff_impressoes'] = (ped['impressoes_totais'] - ped['total_impressoes_pcut']).apply(lambda x: "SIM" if x != 0 else "NAO")
    
    return ped


def relatorio_copias_impressoes():
    st.title("2. Gerar Relatório")
    st.write("""Nesta página você irá encontrar um passo a passo para gerar o relatório de cópias e impressões.""")

    st.subheader("""**2.1. Upload: Planilha do Papercut**""")
    st.write("Arraste ou carregue a planilha do papercut, após seguir as instruções da página 1. (Apenas 1 arquivo.)")

    # Upload: Papercut
    papercut = st.file_uploader(
        label= ' ',
        type= 'xlsx',
        key= 'papercut_upload'
    )

    # Carregando os dados do Papercut
    if papercut is not None:
        st.subheader('Dados do arquivo:')
        
        try:
            pcut = pd.read_excel(papercut)
        except:
            st.error("Falha em carregar a Planilha do Papercut. Verifique se o arquivo foi salvo em uma planilha do excel válida.")


        pcut.columns = [unidecode(col.lower().replace(' ','_')) for col in pcut.columns]
        try:
            pcut['nome_conta_normalizado'] = pcut['nome_da_conta_compartilhada'].apply(lambda x: unidecode(x.lower()) if type(x) == str else '').apply(lambda x: substring(x))
        except:
            st.error("A coluna 'nome da conta compartilhada' não foi encontrada na planilha do papercut.")
        st.dataframe(pcut)

    st.subheader("""**2.2. Upload: Planilha do Pedagógico**""")
    st.write("Carregue a planilha do pedagógico (Apenas 1 arquivo, correspondente a sua unidade).")

    # Upload: Ped
    pedagogico = st.file_uploader(
        label= ' ',
        type= 'xlsx',
        key= 'pedagogico_upload'
    )

    # Carregando os dados do PED
    if pedagogico is not None:
        st.subheader('Dados do arquivo:')

        try:
            ped = pd.read_excel(pedagogico)
        except:
            st.error("Falha em carregar o arquivo do PED. Verifique se o arquivo selecionado está sem os padrões de formatação.")

        ped.columns = [unidecode(str(col).lower().strip().replace(' ','_')) for col in ped.columns]
        
        try:
            ped = ped[['unidade', 'nome_do_arquivo', 'impressoes_totais', 'data_da_solicitacao', 'data_da_utilizacao', 'nome']].copy()
        except:
            st.error("Verifique se as seguintes colunas estão presentes na planilha do PED: unidade, nome do arquivo, impressoes totais, data da solicitação, data da utilização e nome. Verifique também se você fez o upload da planilha contendo apenas a tabela, sem as formatações!")
        ped['nome'] = ped['nome'].apply(lambda x: unidecode(str(x).lower()))
        st.dataframe(ped)


    if (pedagogico is not None) and (papercut is not None):
        st.subheader("""**2.3. Download: Relatório**""")

        if st.button("Gerar Relatório"):
            relatorio = gerar_relatorio_impressoes(
                ped = ped,
                pcut = pcut
            )
            st.dataframe(relatorio)

            with BytesIO() as buffer:
                relatorio.to_excel(buffer, sheet_name= 'relatorio', index= False)
                buffer.seek(0)
                download = st.download_button(
                    label= "**`Baixar Relatório`**",
                    data= buffer,
                    file_name= 'relatorio_copias_impressoes.xlsx',
                    mime= 'application/vnd.ms-excel',
                    key= 'relatorio_download'
                )

def main():
  
    # Título do Menu de Navegação
    st.sidebar.title("Menu de Navegação")

    # Definindo a página inicial
    pagina_selecionada = st.sidebar.radio("Páginas", ["1.Instruções", "2.Relatório"])

    # Limpar a área de exibição
    st.sidebar.markdown("---")

    # Exibir o conteúdo da página selecionada
    if pagina_selecionada == "1.Instruções":
        instrucoes_copias_impressoes()
    
    if pagina_selecionada == "2.Relatório":
        relatorio_copias_impressoes()

    # Rodapé
    st.markdown(footer, unsafe_allow_html= True)


if __name__ == "__main__":
    main()