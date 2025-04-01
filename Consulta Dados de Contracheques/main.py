import os
import pandas as pd
from easygui import msgbox
from selenium.webdriver.common.by import By
from selenium import webdriver
from pathlib import Path
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
from reportlab.pdfgen import canvas
import base64
import glob
import re
import unicodedata
from pypdf import PdfWriter, PdfReader
from time import sleep
# INPUT UNIVERSAL
tabela = pd.read_excel("Dados.xlsx", dtype=str)


# DEFINIÇÃO NAVEGADOR PRA TODAS AS FUNÇÕES
options = Options()
options.add_experimental_option('prefs', {
'download.default_directory': str(Path.cwd()),
'download.prompt_for_download': False,
'download.directory_upgrade': True,
'plugins.always_open_pdf_externally': True,
})
service = Service()
navegador = webdriver.Chrome(service=service, options=options)
navegador.implicitly_wait(10)

def acesso(login, senha):
    link = 'editado'
    navegador.get(url=link)

    botao_petinicial = navegador.find_element(by=By.XPATH, value='//*[@id="pensionistaScroll"]/ul[1]/li[1]/a')
    botao_petinicial.click()

    campo_login = navegador.find_element(by=By.XPATH, value='//*[@id="conterConsulta"]/form/div[1]/div[2]/input')   # CAMPO LOGIN
    campo_login.send_keys(login)

    campo_senha = navegador.find_element(by=By.XPATH, value='//*[@id="conterConsulta"]/form/div[2]/div[2]/input')   # CAMPO SENHA
    campo_senha.send_keys(senha)

    btn_entrar = navegador.find_element(by=By.XPATH, value='//*[@id="conterConsulta"]/form/div[3]/div[2]/input[1]') # BOTÃO DE LOGIN
    btn_entrar.click()

    try:
        btn_contracheque = (navegador.find_element(by=By.XPATH, value='//*[@id="pensionista"]/ul[1]/li[4]/a'))
        btn_contracheque.click()    # BOTÃO CONTRACHEQUE
    except:
        btn_contracheque = (navegador.find_element(by=By.XPATH, value='//*[@id="pensionistaScroll"]/ul[1]/li[4]/a'))
        btn_contracheque.click()    # BOTÃO CONTRACHEQUE
    return

def converte_pdf(caminho_pdf, caminho_png):
    # pega as dimensões da pagina para o print
    with Image.open(caminho_png) as img:
        largura, altura = img.size

    # salvando o pdf com as dimensões totais da pagina
    c = canvas.Canvas(caminho_pdf, pagesize=(largura, altura))
    c.drawImage(caminho_png, 0, 0, width=largura, height=altura)
    c.showPage()
    c.save()
    return

def html_excel(mes_inicial, ano_inicial,caminho_nome,tipo_folha):
    dados_contracheque = {}

    #extraindo os dados
    informacoes_gerais = navegador.find_elements(By.CSS_SELECTOR, "div.ConteudoCampo + div.Titulo3")
    for info in informacoes_gerais:
        #fazendo a lista pro df
        campo = info.find_element(By.XPATH,"preceding-sibling::div").text.strip()
        valor = info.text.strip()
        dados_contracheque[campo] = valor

    # dataframe com os dados
    df_gerais = pd.DataFrame(list(dados_contracheque.items()), columns=['Descrição', 'Valor'])
    df_gerais['Valor'] = df_gerais['Valor'].astype(str)
    df_gerais['Valor'] = df_gerais['Valor'].str.replace('*', '')
    # salvando em excel
    df_gerais.to_excel(f'./{caminho_nome}/{ano_inicial}_{str(mes_inicial).zfill(2)}_xlsx_contracheque_{tipo_folha}.xlsx', sheet_name='Dados', index=False)

    return

def scraping(nome,cont,repeticao,mes_inicial,ano_inicial, tipo_folha):
    navegador.find_element(by=By.XPATH, value='//*[@id="conterConsulta"]/div[3]/div[2]/input').click()        #CLICA EM ENVIAR APÓS SELECIONAR COMPETÊNCIA 
    elemento = WebDriverWait(navegador, 3).until(
        EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ConteudoCampo + div.Titulo3"))                #ESPERA A TABELA APARECER PRA SEGUIR
    )
    
    # SALVAR CAPTURA DA TELA EM PNG
    nome_sem_acento = ''.join(c for c in unicodedata.normalize('NFD', nome) if unicodedata.category(c) != 'Mn') 
    nome_seguro = re.sub(r'[^a-zA-Z0-9]+', '_', nome_sem_acento).strip('_')  # coloca underline em vez de espaço (conflito na hora de salvar a imagem pdf -> biblioteca que eu usei nao achava o caminho por causa dos espaços)
    caminho_nome = f'./Contracheques_{nome_seguro}'
    os.makedirs(caminho_nome, exist_ok=True)  # Cria a pasta se não existir
    
    caminho_png = os.path.join(caminho_nome, 'captura.png')
    captura = navegador.execute_cdp_cmd("Page.captureScreenshot", {"format": "png", "captureBeyondViewport": True})
    with open(caminho_png, 'wb') as f:
        f.write(base64.b64decode(captura['data']))

    # CONVERTER PNG EM PDF
    caminho_pdf = os.path.join(caminho_nome, f'{ano_inicial}_{str(mes_inicial).zfill(2)}_pdf_contracheque_{tipo_folha}.pdf')  #especifica o caminho pra salvar o pdf
    converte_pdf(caminho_pdf, caminho_png)
    try:
        os.remove(f'./{caminho_nome}/captura.png')
    except OSError as e:
        print(f"Error:{ e.strerror}")


    # EXTRAÇÃO DE HTML EM EXCEL
    html_excel(mes_inicial,ano_inicial,caminho_nome,tipo_folha)

    #RETOMA PRA HOME DO SITE PRA NOVA CONSULTA
    navegador.switch_to.default_content()       # Retorna para a janela principal (fora do iframe)
    navegador.find_element(by=By.XPATH, value='//*[@id="pensionista"]/ul[1]/li[4]/a').click()


    #ATUALIZA O MES E CHECA SE VIROU O ANO
    mes_inicial = int(mes_inicial)
    ano_inicial = int(ano_inicial)
    if repeticao == 0:
        if int(ano_inicial)<=2014:
            if mes_inicial == 13:
                mes_inicial = 0

                ano_inicial = int(ano_inicial)
                ano_inicial += 1
                ano_inicial = str(ano_inicial)

                mes_inicial = int(mes_inicial)
                mes_inicial += 1
                mes_inicial = str(mes_inicial)
            else:
                mes_inicial += 1  # Incrementa o mês
                if mes_inicial == 13:
                    unir_arquivos(caminho_nome, ano_inicial)
                    mes_inicial = 1  # Se for Dezembro, muda para Janeiro do próximo ano e incrementa o ano
                    ano_inicial += 1
        else:
            if mes_inicial == 12 and repeticao == 0:
                unir_arquivos(caminho_nome, ano_inicial)
                ano_inicial += 1
                mes_inicial = 1  # Janeiro do próximo ano
            else:
                
                mes_inicial += 1  # Incrementa o mês
                if mes_inicial == 13:
                    unir_arquivos(caminho_nome, ano_inicial)
                    mes_inicial = 1  # Se for Dezembro, muda para Janeiro do próximo ano
                    ano_inicial += 1
        # Formatação do mês como string com dois dígitos
        mes_inicial = str(mes_inicial).zfill(2)
        ano_inicial = str(ano_inicial)

    cont+=1
    

    return mes_inicial,ano_inicial,cont,caminho_nome

def extrair_dados(mes_inicial,mes_final,ano_inicial,ano_final,nome):
    #CALCULAR A QUANTIDADE DE MESES NECESSÁRIA PRA CONTROLAR QUANTAS VEZES FAZER
    qntd_meses = (int(ano_final) - int(ano_inicial)) * 12 + (int(mes_final) - int(mes_inicial) + 1)


    if ano_inicial == ano_final and int(mes_final) >= 12 >= int(mes_inicial):
        qntd_meses += 1  # Adiciona o 13º salário

    #VARIÁVEIS DE CONTROLE DO WHILE
    cont = 0
    repeticao = 0
    while cont < qntd_meses:
        iframe = navegador.find_element(by=By.XPATH, value='/html/body/div/div[2]/div[2]/iframe')     # Pega o XPath do iframe e atribui a uma variável
        navegador.switch_to.frame(iframe)       # Muda o foco para o iframe
        data = str(mes_inicial).zfill(2) + str(ano_inicial) #monta a data
        navegador.find_element(by=By.XPATH, value='//*[@id="N6_P_MMAAAA"]').send_keys(data)       #PREENCHE CAMPO DATA
        botao_envia = navegador.find_element(by=By.XPATH, value='//*[@id="conterConsulta"]/form/div[3]/div[2]/input[1]') #CLICA NO BOTÃO ENVIAR
        botao_envia.click()

        folha=Select(navegador.find_element(by=By.XPATH, value='//*[@id="FOLHA"]'))

        if len(folha.options)==2:
            if repeticao == 0:
                folha.select_by_index('0')
                print('Folha mensal')
                repeticao+=1
                qntd_meses+=1
                tipo_folha = 'mensal'                                          
                mes_inicial,ano_inicial,cont,caminho_nome = scraping(nome,cont,repeticao,mes_inicial,ano_inicial,tipo_folha)
            else:
                folha.select_by_index('1')
                print('Folha 13')
                repeticao = 0
                tipo_folha = '13 mensal'                                          
                mes_inicial,ano_inicial,cont,caminho_nome = scraping(nome,cont,repeticao,mes_inicial,ano_inicial,tipo_folha)
        else:
            folha.select_by_index('0')
            print('Folha mensal única')
            tipo_folha = 'mensal unica'                                       
            mes_inicial,ano_inicial,cont,caminho_nome = scraping(nome,cont,repeticao,mes_inicial,ano_inicial,tipo_folha)
                    
        if cont == qntd_meses:
            unir_arquivos(caminho_nome, ano_inicial)
        

    return

def unir_arquivos(caminho_nome,ano_inicial):
    #UNE TODOS OS XLSX
    arquivos = glob.glob(fr'{caminho_nome}/{ano_inicial}*.xlsx')
    dfs = []
    with pd.ExcelWriter(f'{caminho_nome}/Contracheques {ano_inicial}.xlsx') as writer:
        for arquivo in arquivos:
            # Ler o arquivo Excel
            df = pd.read_excel(arquivo)
            partes = arquivo.split('_')
            partes_corrigidas = [parte.replace('.xlsx', '') for parte in partes]
            nome_sheet = partes_corrigidas[-2] + ' ' + partes_corrigidas[-1] + ' ' + partes_corrigidas[-4]
            nome_sheet = re.sub(r'[<>:"/\\|?*]', '', nome_sheet)   
            # Adicionar o DataFrame como uma nova aba
            df.to_excel(writer, sheet_name=nome_sheet, index=False)
            
    #UNE TODOS OS PDF
    writer = PdfWriter()
    pasta = f'{caminho_nome}'
    comum = f'{ano_inicial}'
    output_pdf = f'{caminho_nome}/Capturas {ano_inicial}.pdf'
    
    for arquivo in os.listdir(pasta):
    # Verifica se o arquivo é PDF e contém a parte do nome em comum
        if arquivo.endswith('.pdf') and comum in arquivo:
            caminho_completo = os.path.join(pasta, arquivo)
            reader = PdfReader(caminho_completo)
            for page in reader.pages:
                writer.add_page(page)
            
    with open(output_pdf, 'wb') as pdf_final:
        writer.write(pdf_final)    
    return

def execucao(): #função principal
    for indice, linha in tabela.iterrows():
        #extrai os dados da planilha (1 por vez = cada for executa 1 linha da planilha)
        nome = linha[1]
        mes_inicial = linha[2]
        ano_inicial = linha[3]
        mes_final = linha[4]
        ano_final = linha[5]
        login = linha[6]
        senha = linha[7]
        acesso(login,senha)
        extrair_dados(mes_inicial,mes_final,ano_inicial,ano_final,nome)
        
    msgbox('finalizado')
    return

execucao()

