import win32com.client as win32
import pandas as pd
import PySimpleGUI as sg
import tkinter as tk

#Exception treatment
try:
    #Excel sheet read as data input
    df = ('C:/Users/GD5CX71/OneDrive - Deere & Co/Desktop/Automação portaria/Dados termo.xlsm')
    for i, contact in enumerate(df['EMAIL']):
        indice = df.loc[i, "ID_ARQUIVO"]

        path_attach= 'C:/Users/GD5CX71/OneDrive - Deere & Co/Desktop/Termos de responsabilidade/'
        file = path_attach + df.loc[i, "ANEXO"]
    
        date = df.loc[i, "DATA"]

        barcode = df.loc[i, "BARCODE"]
    
        # creation of outlook interaction
        outlook = win32.Dispatch('outlook.application')

        #create an e-mail
        email = outlook.CreateItem(0)

        #Configure e-mail parameters
        email.to = contact
        email.Subject = f"Termo de responsabilidade - {indice}"

        email.HTMLBody = f'''
            <p> Olá, 
            <br>
            <br>
            <p> Segue em anexo neste e-mail o termo de responsabilidade, em arquivo PDF, referente ao 
            laptop que o usuário retirou no dia <strong> {date} </strong>. Máquina de barcode <strong> {barcode}.</strong>
            <br>
            <br>
            Atenciosamente; </p><br>

            <p><strong> <font color = "#006600">Time de TI – Catalão <br>
            IT I&O Edge Operations<br>
            John Deere Catalão - NW </strong></font><br>
            John Deere Catalão – Catalão, Goiás, Brazil<br>
            Rua Quadra 11, s/n, Eixo 3, LOTE 000
            </p>
        '''

        #Item attach
        email.Attachments.Add(file)

    #Janela que informa ao user que o e-mail foi enviado
    sg.theme('DarkGreen1')
    layout = [[sg.Text('E-mail and database have been succesfully sent/updated!')]]
    window = sg.Window('E-mail/Database confirmation', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break
        print('You entered ', values[0])

    window.close()

#indica que ocorreu alguma falha ao gerar um e-mail
except Exception as e:
    #Janela para o usuário saber que determinado e-mail não será enviado devido a erros
    sg.theme('DarkRed1')
    layout = [[sg.Text(f'E-mail/database connection{i+1} has an error and will not be sent/')]]
    window = sg.Window('E-mail/Database', layout)

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break
        print('You entered ', values[0])

    window.close()

    print(e)
    pass
