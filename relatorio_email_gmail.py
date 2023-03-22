import pyautogui
import time
import pyperclip
import docx2txt
import pandas as pd


def abrir_email():
# Abrindo o navegador
    pyautogui.press('win')
    time.sleep(2)
    pyautogui.hotkey('win', 's')
    pyautogui.write('Chrome')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(5)

# Abrindo o e-mail    
    pyperclip.copy('gmail.com')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('enter')
    time.sleep(3)
    return


def abrir_compose():
#Abrindo um novo compose
    pyautogui.press('c')
    time.sleep(3)
    return


def copy_paste_email():
#Copiando e colando cada e-mail da lista_de_emails
    for x in lista_de_emails():
        pyperclip.copy(x)
        pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('Tab')
    time.sleep(3)
    return


def copy_paste_assunto():
#Copiando e colando cada e-mail da lista_de_emails
    pyperclip.copy(assunto_email())
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('Tab')
    time.sleep(3)
    return


def copy_paste_relatorio():
#Copiando e colando cada e-mail da lista_de_emails
    pyperclip.copy(relatorio_email())
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(3)
    return


def lista_de_emails():
#Abrindo arquivo excel e retornando os valores da coluna "E-mail"
    lista_de_emails = pd.read_excel('cadastro.xlsx')
    email_format = []
    for email in lista_de_emails['E-mail'].unique():
        email_format.append((f'{email}, '))
    return email_format


def assunto_email():
#Abrindo o arquivo assunto.txt e retornando o valor
    with open ('assunto.txt', 'r') as arquivo2:
        assunto = arquivo2.read()
    return assunto


def relatorio_email():
#Abrindo e retornando os valores do arquivo word
    relatorio = docx2txt.process('relatorio.docx')
    return relatorio


def enviar():
    pyautogui.hotkey('ctrl', 'enter')
    print('E-mails enviados.')
    return

while True:
    abrir_email()
    time.sleep(1)
    abrir_compose()
    time.sleep(1)
    copy_paste_email()
    time.sleep(1)
    copy_paste_assunto()
    time.sleep(1)
    copy_paste_relatorio()
    time.sleep(1)
    enviar()
    break