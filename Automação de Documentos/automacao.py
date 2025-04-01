from docx import Document
import pandas as pd
import docxedit
from datetime import datetime
import locale 
import os
def faz_peticao():
    
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
    data_atual = datetime.now()
    data_formatada = data_atual.strftime("%d de %B de %Y")

    os.makedirs('./Peticoes')
    tabela = pd.read_excel("Resultado.xlsx")
    mask = tabela['Data do Pagamento'].str.contains('Sem dados', na = False)
    tabela = tabela[mask]


    for linha in tabela.index:
        documento = Document("modelo.docx")

        nome = tabela.loc[linha, "Nome"].upper()
        try:
            juizo = tabela.loc[linha, "Juizo"].upper()
        except:
            juizo = 'Sem dados'
        comarca = tabela.loc[linha, "Comarca"].upper()
        data = str(' ' + data_formatada)
        processo = str(tabela.loc[linha,"Processos"])

        referencias = {
            "XXXX": nome,
            "YYYY": comarca,
            "WWWW": data,
            "QQQQ": processo,
            "ZZZZ": juizo,    
        }

        for paragrafo in documento.paragraphs:
            for codigo in referencias:
                valor = referencias[codigo]
                docxedit.replace_string(documento, old_string=codigo, new_string=valor)

        
        documento.save(f"./Peticoes/Petição - {nome}.docx")

faz_peticao()