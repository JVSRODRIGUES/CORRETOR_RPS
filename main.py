from bs4 import BeautifulSoup as bs
import pandas as pd

condição = True

while condição :
    condição = False
    fim = ''
    print('Nome dos arquivos tem que ser com a extensão!!')
    #endereço_xml = str(input('Nome do Arquivo XML: '))
    #endereço_excel = str(input('Nome do Arquivo EXCEL: '))

    endereço_xml = 'LOTE_08_20231103_000050867.xml'
    endereço_xml_s = 'LOTES_PRONTOS/LOTE_08_20231103_000050867.xml'
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

    print(len(lista_xml), len(lista_ex))

    # Laço para fazer a alteração do RPS
    y = -2
    for x in range(num//2):
        i = 0
        i = i + 1
        y = y + 2
        print('-'*20)
        print(y, x)
        print(int(lista_xml[y].text),'=' ,int(excel[x]))
        lista_xml[y].string.replace_with(str(excel[x]))

    # Salva arquivo XML editado
    with open(endereço_xml_s, 'w') as f:
        f.write(soup.prettify())

    # Apagar coluna excel
    lista_ex.drop(num // 2, axis = 'index')

    # Apagar RPS já utilizados na tabela
    for x in range(num // 2):
        lista_ex = lista_ex.drop(x, axis = 'index')

    # Salvar alteração do EXCEL
    lista_ex.to_excel('BASE.xlsx', index = False)    
    
    print('-'*20)
    print('')
    print('Total de RPS usados:',num//2)
    print('Sendo o ultimo: ',int(lista_xml[y].text))
    print('')
    print('-'*8,'FIM','-'*7)
    print('')
    print('{} Linas apagadas'.format(num // 2))

    #fim = input('Deseja alterar mais alguns [s/n]? ').upper().strip()

    #if fim == 'S':
    #    condição = True   
