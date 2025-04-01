from time import sleep
import xlrd
from selenium import webdriver
import pandas as pd
from easygui import *
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import json

with open('credentials.json', 'r') as read_file:
    credenciais = json.load(read_file)

message = "Qual grau de jurisdição?"
title = "Jurisdição"
if boolbox(message, title, ["1º Grau", "2º Grau"]):
    url = 'editado'
else:
    url = 'editado'

options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.prompt_for_download': False,
    'plugins.always_open_pdf_externally': True,
})
navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador.implicitly_wait(5)


# COLETA DADOS EXCEL
wb = xlrd.open_workbook('intimacoes.xls')
planilha = wb.sheet_by_name('Processos Pendentes - Urgente')
total_linhas = planilha.nrows
total_colunas = planilha.ncols
login, senha = credenciais['credentials']
navegador.get(url)  # ABRE NAVEGADOR
navegador.find_element(by=By.XPATH, value='//*[@id="txtUsuario"]').send_keys(login)  # CAMPO LOGIN
navegador.find_element(by=By.XPATH, value='//*[@id="pwdSenha"]').send_keys(senha)  # CAMPO SENHA
navegador.find_element(by=By.XPATH, value='//*[@id="sbmEntrar"]').click()  # BOTAO ENTRAR

msgbox("Resolva o Captcha caso apareça!")

navegador.find_element(by=By.XPATH, value='//*[@id="tr0"]').click()  # SELECIONA PERFIL ADVOGADO
sleep(1.5)

# LISTAS
dados = []
precatorios = []
precatorios2 = []
processos = []
procsadm = []
data_dist = []
comarcas = []

titulos = []
tipos = []
requerentes = []
requeridos = []
notas = []
textoVermelho = []
cj_descricao = []
cj_cpf = []

# PESQUISA PROCESSOS PARA CADA LINHA DO EXCEL
for i in range(2, total_linhas):
    processo = planilha.cell_value(rowx=i, colx=0)
    processo = str(processo)
    print('xxxxxxxxxxxxxxxxxxxxxxxxxxxxxx')
    print(processo)

    navegador.find_element(by=By.XPATH, value='//*[@id="navbar"]/div/div[3]/div[3]/form/input[1]').send_keys(processo)
    sleep(0.5)

    navegador.find_element(by=By.XPATH,
                           value='//*[@id="navbar"]/div/div[3]/div[3]/form/button[1]').click()  # BOTAO PARA PESQUISAR
    sleep(1)

    # COLETA ORIGINARIO
    try:
        processo_originario = navegador.find_element(by=By.XPATH,
                                                     value='//*[@id="tableRelacionado"]/tbody/tr/td[1]/font/a').text
        sleep(0.5)
    except NoSuchElementException:
        processo_originario = 'Sem dados de processo originário'

    # COLETA DATA DE DISTRIBUIÇÃO
    data = navegador.find_element(by=By.XPATH, value='//*[@id="txtAutuacao"]').text
    sleep(0.5)
    data_reduc = data[0:10]

    # NOME PARTES
    nomeParteTodos = navegador.find_elements(by=By.CLASS_NAME, value='infraNomeParte')  # COLETA NOME DA PARTE
    nomeParte = nomeParteTodos[0].text
    sleep(0.5)
    print(nomeParte)
    cpf = navegador.find_element(by=By.XPATH, value='//*[@id="spnCpfParteAutor0"'
                                                    ']').text  # COLETA CPF DO AUTOR
    print(cpf)

    try:
        precatorio = navegador.find_element(by=By.XPATH,
                                            value='/html/body/div[1]/div[2]/div[2]/div[1]/div[1]/form[2]/div[2]/div[1]/'
                                                  'div/fieldset[1]/div/table/tbody/tr[2]/td[1]/font/a').text
        sleep(0.5)
    except:
        precatorio = 'Sem número de precatório'

    try:
        precatorio2 = navegador.find_element(by=By.XPATH,
                                             value='/html/body/div[1]/div[2]/div[2]/div[1]/div[1]/form[2]/div[2]'
                                                   '/div[1]/div/fieldset[1]/div/table/tbody/tr[1]/td[1]/font/a').text
        sleep(0.5)
    except:
        precatorio2 = 'Sem número de originário 2'

    precatorios.append(precatorio)
    precatorios2.append(precatorio2)

    # COMEÇA COLETA DAS NOTAS
    # CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE CLARA
    controleClara = len(navegador.find_elements(by=By.CSS_SELECTOR,
                                                value='[class="infraTrClara infraEventoPrazoAguardando"]'))

    # CHECA QUANTOS ELEMENTOS VERMELHOS TEM NA PARTE ESCURA
    controleEscura = len(navegador.find_elements(by=By.CSS_SELECTOR,
                                                 value='[class="infraTrEscura infraEventoPrazoAguardando"]'))

    # CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE CLARA
    controleClaraAmarelo = len(navegador.find_elements(by=By.CSS_SELECTOR,
                                                       value='[class="infraTrClara infraEventoPrazoAberto"]'))

    # CHECA QUANTOS ELEMENTOS AMARELOS TEM NA PARTE ESCURA
    controleEscuraAmarelo = len(navegador.find_elements(by=By.CSS_SELECTOR,
                                                        value='[class="infraTrEscura infraEventoPrazoAberto"]'))

    if controleClara > 0 or controleEscura > 0:
        if controleClara > 0:
            print("Prazo Aguardando Abertura - linha clara")
            vermelho = navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrClara infraEventoPrazoAguardando"]')
        elif controleEscura > 0:
            print("Prazo Aguardando Abertura - linha escura")
            vermelho = navegador.find_elements(by=By.CSS_SELECTOR, value='[class="infraTrEscura infraEventoPrazoAguardando"]')
        for x in range(0, len(vermelho)):
            textoVermelho.append(vermelho[x].text)
            if nomeParte[0] in vermelho[x].text:
                controleTexto = x
                break

        print(textoVermelho[controleTexto])

        textoVermelho[controleTexto] = textoVermelho[controleTexto].replace('\n', '')  # TIRA AS QUEBRAS DE LINHA
        posterior = textoVermelho[controleTexto].split(':')  # ARMAZENAMENTO DO PRIMEIRO TERMO ATÉ O PRÓXIMO
        anterior = textoVermelho[controleTexto].split('(')  # menos um do posterior
        indiceParenteses = posterior[3].find('(')
        resultado1 = posterior[3][0:indiceParenteses]  # UNIÃO DOS RESULTADOS
        resultado2 = resultado1.split(' ')

        if len(resultado1) > 2:
            resultado = resultado2[1]
        else:
            resultado = resultado1.replace(' ', '')

        print('RESULTADO: \n', resultado)
        navegador.find_element(by=By.CSS_SELECTOR, value=f'[id="trEvento{resultado}"]')
        descricao = navegador.find_element(by=By.XPATH, value=f'//*[@id="trEvento{resultado}"]/td[3]/label').text
        cj_descricao.append(descricao)
        # um evento só
        descricaoNota = navegador.find_element(by=By.XPATH, value=f'//*[@id="trEvento{resultado}"]/td[3]/label').text
        print('Descrição da Nota: \n', descricaoNota)
        if descricaoNota == "Juntada de íntegra do processo":
            paragrafoUnificado = 'Digitalização de Processo'
            requerente = 'Sem dados'
            requerido = 'Sem dados'
            tipo_acao = 'Sem dados'
        elif descricaoNota == 'PETIÇÃO':
            paragrafoUnificado = 'Intimação sobre manifestação de outra parte'
            requerente = 'Sem dados'
            requerido = 'Sem dados'
            tipo_acao = 'Sem dados'
        else:
            try:
                # ABRE O ÚLTIMO DOCUMENTO(EVENTO)
                navegador.find_element(by=By.XPATH,
                                       value=f'//*[@id="trEvento{resultado}"]/td[5]/a').click()
                sleep(1)
                navegador.switch_to.window(navegador.window_handles[1])  # TROCA DE GUIA

                req2 = navegador.find_element(by=By.CLASS_NAME, value='nome_parte')  # COLETA TODOS OS REQUERIDOS
                prec = navegador.find_element(by=By.CLASS_NAME, value='identificacao_processo')  # COLETA O NUMERO DO PROCESSO
            except:
                req2 = 'Sem requeridos'
                prec = 'Sem dados processo'

            try:
                acao = navegador.find_element(by=By.CLASS_NAME, value='assunto_processo').text  # COLETA A ACAO TODA
            except:
                acao = 'Sem dados'

            if req2 != 'Sem requeridos':
                requerente = req2[0].text  # COLETA SÓ O REQUERENTE
            else:
                requerente = 'Sem requerente'

            paragrafo = navegador.find_element(by=By.CLASS_NAME, value='paragrafoPadrao')
            contLinha = len(paragrafo)
            x = 0
            paragrafoLinhas = []

            for j in range(0, contLinha):
                linha = paragrafo[x].text
                x += 1
                paragrafoLinhas.append(linha)
            paragrafoUnificado = "".join(paragrafoLinhas)

            requerido = req2[1].text  # COLETA SÓ O REQUERENTE
            if acao == 'Sem dados':
                tipo_acao = 'Sem dados'
            else:
                tipo_acao = acao[14:]  # DIVIDE O TIPO DA AÇÃO PARA O DATAFRAME
            sleep(1)
            navegador.close()
            navegador.switch_to.window(navegador.window_handles[0])

        print('Prazo em aberto')
        processo_originario = 'Prazo em aberto'  # JOGA O VALOR EXTRAIDO PARA A LISTA
        processo = 'Prazo em aberto'  # JOGA O VALOR DO PROCESSO PARA A LISTA
        data_reduc = 'Prazo em aberto'  # JOGA VALOR DATA PARA A LISTA
        tipo_acao = 'Prazo em aberto'
        requerente = 'Prazo em aberto'
        requerido = 'Prazo em aberto'
        paragrafoUnificado = 'Prazo em aberto'
    else:
        print('Prazo fechado ou erro')

    dados.append(processo_originario)  # JOGA O VALOR EXTRAIDO PARA A LISTA
    processos.append(processo)  # JOGA O VALOR DO PROCESSO PARA A LISTA
    data_dist.append(data_reduc)  # JOGA VALOR DATA PARA A LISTA
    tipos.append(tipo_acao)
    requerentes.append(requerente)
    requeridos.append(requerido)
    notas.append(paragrafoUnificado)
    cj_cpf.append(cpf)

    textoVermelho = []
    print(f'Contagem: {i}/{total_linhas}')


# EXPORT
df = {'Principal': processos, 'Originário': dados, 'Data distribuição': data_dist, 'Título': titulos, 'Tipo': tipos,
      'Requerente': requerentes, 'CPF': cj_cpf, 'Requerido': requeridos,
      'Certidão': notas}  # ARMAZENA OS VALORES EM UMA MATRIZ
df1 = pd.DataFrame(df)  # CRIA O DATAFRAME
df1.to_excel('./Lista de notas.xlsx')  # CRIA O ARQUIVO EXCEL

print('FINALIZADO')
navegador.quit()