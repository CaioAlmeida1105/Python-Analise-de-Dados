import pandas as pd
import win32com.client as win32

# Importar Base de dados
planilha = pd.read_excel('Python/Pandas/Projeto Simples/Planilha/Vendas.xlsx')



# Visualizar Base de dados
pd.set_option('display.max_columns',None)
print(planilha)

# faturamento por loja
faturamento = planilha [['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print (faturamento)

#quantidade de produtos vendidos por loja
quantidade = planilha [['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print (quantidade)

#ticket médio por produto em cada loja
ticket_medio = (faturamento ['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print (ticket_medio)
#enviar um e-mail com relatorio

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'santos.almeida.caio@gmail.com'
mail.Subject = 'Relatorio de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relátorio de Vendas por cada Loja.</p>


<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produto em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposição.</p>

<p>
    Att.,
    Caio
</p>
'''
mail.Send()

print ('E-mail enviado!')