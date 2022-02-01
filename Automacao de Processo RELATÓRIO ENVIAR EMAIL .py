#!/usr/bin/env python
# coding: utf-8

# ### Passo 1 - Importar Arquivos e Bibliotecas

# In[13]:


#IMPORTANDO AS BIBLIOTECAS #

import pandas as pd
import pathlib
import win32com.client as win32


emails = pd.read_excel(r'Bases de Dados\Emails.xlsx')
lojas =  pd.read_csv(r'Bases de Dados\Lojas.csv', encoding='latin1', sep=';')
vendas = pd.read_excel(r'Bases de Dados\Vendas.xlsx')
display(emails)
display(lojas)
display(vendas)

vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja']==loja, :]
display(dicionario_lojas['Rio Mar Recife'])
display(dicionario_lojas['Salvador Shopping']) 


dia_indicador = vendas['Data'].max()
print(dia_indicador)
print(dia_indicador.day)
print(dia_indicador.month)               #FORMAS DE IMPRIMIR AS DATA, MÊS E ANO
print(dia_indicador.year)
print('{}/{}'.format(dia_indicador.day, dia_indicador.month))
print(f'{dia_indicador.day}/{dia_indicador.month}')


   # CRIANDO ARQUIVOS DE BACKUP #


#dia_indicador = vendas['Data'].min()  #Calcula o primeiro dia do mês e ano
caminho_backup = pathlib.Path(r'Backup Arquivos Lojas')

arquivos_pasta_backup = caminho_backup.iterdir()
lista_nomes_backup = [arquivo.name for arquivo in arquivos_pasta_backup]

#Salvar dentro da pasta
for loja in dicionario_lojas:
    if loja not in lista_nomes_backup:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir()
    
 #Cria arquivos de backup 
nome_arquivo = '{}_{}_{}.xlsx'.format(dia_indicador.month, dia_indicador.day, loja)
local_arquivo = caminho_backup / loja /  nome_arquivo                
dicionario_lojas[loja].to_excel(local_arquivo)
    

    
        #Definições metas
meta_faturamento_dia = 1000
meta_faturamento_ano = 1650000
meta_qtdaprodutos_dia = 4
meta_qtdaprodutos_ano = 120
meta_tecketmedio_dia = 500
meta_tecketmedio_ano = 500


for loja in dicionario_lojas:
    
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    # Calcular o Faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)


    # Calcular os Indicadores de Produtos "diversidade"
    qtda_produtos_ano = len(vendas_loja['Produto'].unique())
    #print(qtda_produtos_ano)
    qtda_produtos_dia = len(vendas_loja_dia['Produto'].unique()) # .unique() Tira os itens Duplicado
    #print(qtda_produtos_dia)


    #Calcular o Ticket Médio
    valor_venda = vendas_loja.groupby('Código Venda').sum() # .sum() Soma os valores 
    ticket_medio_ano = valor_venda['Valor Final'].mean() # .mean() Dar a média
    #print(ticket_medio_ano)
    #ticket_medio_dia
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum() # .sum() Soma os valores 
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean() # .mean() Dar a média
    #print(ticket_medio_dia)
    
    
    #Enviar o E-mail
    outlook = win32.Dispatch('outlook.application')

    nome = emails.loc[emails['Loja']==loja, 'Gerente'].values[0]
    mail = outlook.CreateItem(0)
    mail.to = emails.loc[emails['Loja']==loja, 'E-mail'].values[0]
    mail.Subject = f'OnePage Dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}'
    mail.Body = 'Texto do E-mail'


    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'

    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if qtda_produtos_dia >= meta_qtdaprodutos_dia:    
        cor_qtda_dia = 'green'
    else:
        cor_qtda_dia = 'red'

    if qtda_produtos_ano >= meta_qtdaprodutos_ano:    
        cor_qtda_ano = 'green'
    else:
        cor_qtda_ano = 'red'

    if ticket_medio_dia >= meta_tecketmedio_dia:    
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'

    if ticket_medio_ano >= meta_tecketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    mail.HTMLBody = f'''
    <p> Bom Dia!!! {nome}</p>

    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month}) Loja {loja}</strong> foi de:</p>

    <table>
      <tr>
        <th style="text-align: center">Indicador</th>
        <th style="text-align: center">Valor Dia</th>
        <th style="text-align: center">Meta Dia</th>
        <th style="text-align: center">Cenário Dia</th>
      </tr>
      <tr>
        <td style="text-align: center">Faturamento</td>
        <td style="text-align: center">R${faturamento_dia:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_dia}">◙</font></td>
      </tr>
      <tr>
         <td style="text-align: center">Diversidade de Produtos</td>
        <td style="text-align: center">{qtda_produtos_dia}</td>
        <td style="text-align: center">{meta_qtdaprodutos_dia}</td>
        <td style="text-align: center"><font color="{cor_qtda_dia}">◙</font></td>
      </tr>
      <tr>
         <td style="text-align: center">Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_dia:.2f}</td>
        <td style="text-align: center">R${meta_tecketmedio_dia:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_dia}">◙</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th style="text-align: center">Indicador</th>
        <th style="text-align: center">Valor Ano</th>
        <th style="text-align: center">Meta Ano</th>
        <th style="text-align: center">Cenário Ano</th>
      </tr>
      <tr>
        <td style="text-align: center">Faturamento</td>
        <td style="text-align: center">R${faturamento_ano:.2f}</td>
        <td style="text-align: center">R${meta_faturamento_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_fat_ano}">◙</font></td>
      </tr>
      <tr>
         <td style="text-align: center">Diversidade de Produtos</td>
        <td style="text-align: center">{qtda_produtos_ano}</td>
        <td style="text-align: center">{meta_qtdaprodutos_ano}</td>
        <td style="text-align: center"><font color="{cor_qtda_ano}">◙</font></td>
      </tr>
      <tr>
         <td style="text-align: center">Ticket Médio</td>
        <td style="text-align: center">R${ticket_medio_ano:.2f}</td>
        <td style="text-align: center">R${meta_tecketmedio_ano:.2f}</td>
        <td style="text-align: center"><font color="{cor_ticket_ano}">◙</font></td>
      </tr>
    </table>

    <p>Segue em anexo a Planilha com todos os dados para mais detalhes.</P>

    <p>Qualquer dúvida Estou à Disposição.</p>
    <p><strong>Att., Fabricio Freitas</strong></p>
    '''

    attachment = pathlib.Path.cwd() / caminho_backup / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    mail.Attachments.Add(str(attachment)) #nao pode tirar o STR() pois ele transforma em texto

    mail.Send()
    print('E-Mail da loja {} Enviado com Sucesso!!!'.format(loja))
    
    
    
                            #CRIA O FATURAMENTO ANUAL e DIA DAS LOJAS E SALVA UM ARQUIVO BACKUP          
    
faturamento_lojas = vendas.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False) #ascending=False ordena em ordem decrescente
display(faturamento_lojas_ano)

nome_arquivo = '{}_{}_Ranking Anual.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_ano.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))


vendas_dia = vendas.loc[vendas['Data']==dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)
display(faturamento_lojas_dia)

nome_arquivo = '{}_{}_Ranking Dia.xlsx'.format(dia_indicador.month, dia_indicador.day)
faturamento_lojas_dia.to_excel(r'Backup Arquivos Lojas\{}'.format(nome_arquivo))



                                   #ENVIA E-MAIL PARA DIRETORIA 

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.to = emails.loc[emails['Loja']=='Diretoria', 'E-mail'].values[0]
mail.Subject = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
mail.Body = f'''
Prezados, Bom Dia!

Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R$:{faturamento_lojas_dia.iloc[0, 0]:.2f}
Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R$:{faturamento_lojas_dia.iloc[-1, 0]:.2f}

Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R$:{faturamento_lojas_ano.iloc[0, 0]:.2f}
Pior loja do Amo em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R$:{faturamento_lojas_ano.iloc[-1, 0]:.2f}


Segue em anexo os ranking do ano e do dia de todas as lojas.

Qualquer Dúvida estou a disposição.

Att.,
Frabricio Freitas
'''

attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
mail.Attachments.Add(str(attachment)) #nao pode tirar o STR() pois ele transforma em texto
attachment = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Dia.xlsx'
mail.Attachments.Add(str(attachment))

mail.Send()
print('E-Mail da Diretoria Enviado com Sucesso!!!')


# ### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# ### Passo 3 - Salvar a planilha na pasta de backup

# In[ ]:





# ### Passo 4 - Calcular o indicador para 1 loja

# In[ ]:





# In[ ]:





# ### Passo 5 - Enviar por e-mail para o gerente

# In[ ]:





# ### Passo 6 - Automatizar todas as lojas

# In[ ]:





# ### Passo 7 - Criar ranking para diretoria

# In[ ]:





# ### Passo 8 - Enviar e-mail para diretoria

# In[ ]:




