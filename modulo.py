import pandas as pd
import numpy as np
import win32com.client as win32
from datetime import date


def manipulador():
    data_atual = date.today()
    data = data_atual.strftime('%d.%m')
    df_geral = pd.read_excel('INADIMPLENCIA.xlsx')

    sureg = [
        '37 - SUREG PORTO ALEGRE', 
        '6 - SUREG CENTRO',
        '5 - SUREG SERRA',
        '4 - SUREG SUL',
        '8 - SUREG NOROESTE',
        '7 - SUREG FRONTEIRA',
        '38 - SUREG OUTROS ESTADOS',
        '14 - SUREG LESTE',
        '9 - SUREG ALTO URUGUAI']

    emails = [
        'email 1',
        'email 2',
        'email 3',
        'email 4',
        'email 5',
        'email 6',
        'email 7',
        'email 8',
        'email 9'
        ]
    
    copias = 'filano@gmail.com; ciclano@gmail.com; beltrano@gmail.com; deltrano@gmail.com'
    
    i=1
    while i <= len(sureg):
        df_manipulado = df_geral[df_geral['SUREG'] == f'{sureg[i-1]}']
        df_manipulado = df_manipulado.sort_values(by="Agência", ascending = True).reset_index(drop = True)
        df_manipulado.to_excel(f'INADIMPLENCIA FGI {sureg[i-1]} {data}.xlsx', index=False)

    
        #Contando operações sem repetições
        df_parcelas = df_manipulado.drop_duplicates(subset="Nome Cliente")
        df_parcelas['Qtd.Operações'] = df_parcelas.groupby('Agência')['Agência'].transform('count')
        df_parcelas = df_parcelas[['Agência', 'Qtd.Operações']].sort_values(by ='Qtd.Operações', ascending = False).reset_index(drop = True).drop_duplicates(subset='Agência')
        df_parcelas = format(df_parcelas.head().to_html(index=False))


        #Somando valores das operações
        df_soma = df_manipulado[['Agência', 'Valor Vencido']]
        df_soma['Valor Vencido'] = df_manipulado.groupby('Agência')['Valor Vencido'].transform(np.sum)
        df_soma = df_soma.drop_duplicates(subset="Agência").sort_values(by ='Valor Vencido', ascending = False).reset_index(drop = True)
        df_soma = format(df_soma.head().to_html(index=False))


        #Utilizando o Openxl para escrever em VBA
        outlook = win32.Dispatch('outlook.application')
        email = outlook.CreateItem(0)
        anexo = fr"C:\Users\eusou\OneDrive\Documentos\Python\Projetos\E-mails inadimplencia\INADIMPLENCIA FGI {sureg[i-1]} {data}.xlsx"

        #email.SentOnBehalfOfName = ""
        email.to = f'{emails[i-1]}'
        email.cc = copias
        email.Display()
        assinatura = email.HTMLBody
        email.subject = f"Ranking Inadimplencia {sureg[i-1]}"
        email.Attachments.Add (anexo)
        email.htmlbody = f'''
        <font color = "007FFF" size = "3"> Bom dia!

        <p> Prezados(as), seguem as planilhas com todas as operações em andamento, com informações da SUREG, agência e etapa da operação. </p>

        <p> Implementamos o arquivo parcelas em atraso, no qual estão presentes o número de parcelas em atraso, valores contratados, vencidos, multas e demais
        informações pertinente, o relatório tem informação separada por parcela, sendo uma linha para cada, os valores apresentandos são os de liquidação na 
        data de hoje. </p>

        <p> Enviaremos estes arquivos diariamente, para seus acompanhamentos.</p>

        <p> Ranking operações em inadimplência: </p>
        {df_parcelas}
        
        <p> Ranking valores vencidos </p> </font>
        {df_soma}
        '''
        #email.send

        i += 1