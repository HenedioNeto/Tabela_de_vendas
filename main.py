import pandas as pd
import win32com.client as win32

# Importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar base de dados
pd.set_option('display.max_columns', None)

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

# Enviar um email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'user@email.com'
mail.Subject = 'Relatório de Vendas'
mail.HTMLBody = f'''
<p>Segue o relatório de vendas</p>

<p>Faturamento</p>
{faturamento.to_html(formatters={'Valor Final': 'R$:{:,.2f}'.format})}

<p>Quantidade vendida</p>
{quantidade.to_html()}

<p>Ticket Médio dos produtos em cada loja</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R$:{:,.2f}'.format})}

<p>Att.,</p>
<p>Henédio</p>
'''

mail.Send()
