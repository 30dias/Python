import win32com.client as win32
import pandas as pd
from time import sleep
from random import randint

# CÓDIGO PARA ENVIAR EMAILS

cont_email = 0
whatsapp_number = input('Digite seu número whatsapp: ')
phone_number = input('Digite seu número de telefone')

link_whats = f'https://wa.me/{whatsapp_number}'

#tentativa de localizar planilha base  
while True:
    try:
        fonte = pd.read_excel("g://Meu Drive/Emails_vendedores/email_magali_1907_valeria-o.xlsx")
        break
    except:
        print('Planilha fonte não encontrada, pare o programa e adicione uma planilha base antes de continuar')
        continue

#dados da planilha fonte
while True:

    for n, emaill in enumerate (fonte["Email"]):
        # Empresa = fonte.loc[n, "Empresa"]

        try:
            #tempo entre cada email
            timer = randint(60,120)
            sleep(timer)

            #outlook
            outlook = win32.Dispatch('outlook.application')
            #email
            email = outlook.CreateItem(0)
            #informações do e-mail
            email.To = str(emaill)
            email.Subject = "VIVO EMPRESAS"
            email.HTMLBody = f""" Digite aqui a mensagem do email no formato HTML
                                  Dados de contato: {phone_number}
                                  Whatsapp: {whatsapp_number}     """

            email.Send()
            cont_email += 1
            print(f'{cont_email} emails enviados')
            print(f'Delay desde o último email - {timer} segundos')
        except:
            print('DADOS DA PLANILHA BASE INVÁLIDOS')
            break

    break
#contator de emails enviados
print(f'FINALIZADO!! Foram enviados {cont_email} emails.')