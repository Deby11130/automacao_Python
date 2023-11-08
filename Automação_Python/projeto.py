import pandas as pd  #instale pip install pandas 

#importar base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')


#visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-'*50)

#Faturamento por loja

faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()

print(faturamento)

print('-'*50)

#Quantidade de produtos vendido por loja

qtd_produto = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()

print(qtd_produto)

print('-'*50)

#Ticket Médio por produto de cada loja

ticket_medio = (faturamento['Valor Final'] / qtd_produto['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print( ticket_medio)

#enviar email com relatório

import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'your_email'
mail.Subject = 'Relatório'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue p Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
<p>Quantidade Vendida:</p>
{qtd_produto.to_html()}
<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
<p>Qualquer dúvida estou á disposição.</p>
<p>Att..</p>'
'''

mail.Send()

print('Email enviado')