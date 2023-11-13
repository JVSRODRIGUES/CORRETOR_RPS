from bs4 import BeautifulSoup as bs
import pandas as pd
import os

cont = 0

# Mapeando local do sistema
loc1 = os.getcwd() 
loc = os.getcwd()
local = loc + '\LOTES'
os.chdir(local)
lote = os.listdir()
local = loc1
os.chdir(local)
num_lote = len(lote)

while cont <= (num_lote - 1):

    endereço_xml = 'LOTES/' + lote[cont]
    endereço_xml_s = 'LOTES_PRONTOS/' + lote[cont]
    endereço_excel = 'BASE.xlsx'

    # Pega os arquivos XML
    file = open(endereço_xml,"r")
    dados = file.read()
    soup = bs(dados, "xml")
    lista_xml = soup.find_all('Numero')

    # Pega os arquivos EXCEL
    lista_ex = pd.read_excel(endereço_excel)

    num = len(lista_xml)

    # Tranforma excel em lista
    excel = lista_ex.head(num // 2)
    excel = excel['Sequencia'].to_list()

    # Laço para fazer a alteração do RPS
    y = -2
    for x in range(num//2):
        i = 0
        i = i + 1
        y = y + 2
        #print('-'*20)
        #print(y, x)
        #print(int(lista_xml[y].text),'=' ,int(excel[x]))
        lista_xml[y].string.replace_with(str(excel[x]))

    # Salva arquivo XML editado
    with open(endereço_xml_s, 'w') as f:
        f.write(soup.prettify())


    # Apagar RPS já utilizados na tabela
    for x in range(num // 2):
        lista_ex = lista_ex.drop(x, axis = 'index')

    # Salvar alteração do EXCEL
    lista_ex.to_excel('BASE.xlsx', index = False)    
    
    print('-'*20)
    print('')
    print('Total de RPS usados:',num//2)
    print('Sendo o ultimo: ',int(lista_xml[y].text))
    print(lote[cont],' OK')
    print('')
    print('{} Linas apagadas'.format(num // 2))
    cont= cont + 1

print('')
print('Sobraram {} RPS'.format(len(lista_ex)))    
print('')
print('-'*8,'FIM','-'*7)
