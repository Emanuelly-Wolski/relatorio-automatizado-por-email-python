import pandas as pd

#Importar a base de dados:

tabela_vendas = pd.read_excel('Vendas.xlsx')
print(tabela_vendas)

#visualizar a base de dados:

pd.set_option('display.max_columns', None)

#Faturamento por loja:

print('-' * 50) #imprime 50x o tracinho

faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#Quantidade de produtos vendidos por loja:

print('-' * 50)

quantidade_produtos_vendidos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade_produtos_vendidos)


#Ticket médio por produto em cada loja (faturamento/qtde de produtos vendidos):

print('-' * 50)

ticket_medio = (faturamento['Valor Final'] / quantidade_produtos_vendidos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0 : 'Ticket Médio'})
print(ticket_medio)

#Enviar email com relatório:

import win32com.client as win32

outlook = win32.Dispatch('outlook.application') #conecta o python com o outlook
mail = outlook.CreateItem(0)     #cria um email
mail.To = 'endereço de email que receberá o relatório'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f''' 
<p>Prezados,</p>

<p>Segue Relatório de Vendas atualizado de acordo com cada uma das Lojas.</p>

<h4>Faturamento:</h4>
{faturamento.to_html(formatters={'Valor final': 'R${:,.2f}'.format})}

<h4>Quantidade Vendida:</h4>
{quantidade_produtos_vendidos.to_html()}

<h4>Ticket médio dos produtos em cada Loja:<h4>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou a disposição!</p>.
<p>Atenciosamente, nome</p>

'''

mail.Send()

print('Email enviado')
