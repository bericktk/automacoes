# Instalando bibliotecas
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os

webbrowser.open('https://web.whatsapp.com/')
sleep(30)

# Ler os dados de uma planilha e pegar dados como nome, telefone e empresa
workbook = openpyxl.load_workbook('planilha.xlsx')
pagina_socios = workbook['nome_da_planilha']

for linha in pagina_socios.iter_rows(min_row=2):
    # Pegando os dados
    nome = linha[0].value
    telefone = linha[1].value
    vencimento = linha[2].value
    link_pagamento = linha[3].value
    mensagem = f'Olá {nome} seu boleto venceu no dia {vencimento.strftime("%d/%m/%Y")}. Por favor acesse o link para realizar o pagamento da sua fatura: {link_pagamento}'

    # Criando o link personalizado para mensagem no whatsapp do sócio
    try:    
        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
        webbrowser.open(link_mensagem_whatsapp)
        sleep(10)
    
        seta = pyautogui.locateCenterOnScreen('seta.png')
        sleep(5)
        pyautogui.click(seta[0],seta[1])
        sleep(5)
        pyautogui.hotkey('ctrl', 'w')
        sleep(5)
    except:
        print(f'Não foi possivel enviar mensagem para {nome}')
        with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
            arquivo.write(f'{nome},{telefone}')
