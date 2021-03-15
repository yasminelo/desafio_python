#Extrair banco de dados
#Primeiro desafio - por Yasmine Lopes, 09/02/21

import pandas as pd
tabela_vendas = pd.read_excel("/content/drive/MyDrive/Vendas.xlsx")
display (tabela_vendas)

#Calcular Faturamento
#ID Loja, Valor Final #São colunas, então -> [[]]
#Agrupar ID Loja, Somar Valor Final

tabela_faturamento = tabela_vendas [["ID Loja", "Valor Final"]].groupby("ID Loja").sum()
display (tabela_faturamento)

#Sucesso! :)

#Calcular quantidade de produtos vendidos em cada loja
#Mostrar ID Loja e Quantidade
#Agrupar ID Loja, somar Quantidade

tabela_quantidade = tabela_vendas [["ID Loja", "Quantidade"]].groupby("ID Loja").sum()
display(tabela_quantidade)

#Calcular Ticket médio de cada loja
#Ticket médio: Valor final/ Quantidade
#Mostrar tabela ID Loja agrupada + Valor final agrupado, ou seja, tabela_faturamento dividido pela coluna quantidade

ticket_medio = tabela_faturamento["Valor Final"]/tabela_quantidade['Quantidade']
display (ticket_medio)

#Criando função:
#Enviar e-mails
def enviar_email(nome_da_loja, tabela):
  import smtplib
  import email.message

  server = smtplib.SMTP('smtp.gmail.com:587')  
  corpo_email = f"""
<p>Prezados,</p>
<p>Segue relatório de vendas,</p>
{tabela.to_html()}
<p>Qualquer dúvida estou à disposição.</p>
  """
    
  msg = email.message.Message()
  msg['Subject'] = f"Relatório de vendas - {nome_da_loja}"
    
  # Fazer antes (apenas na 1ª vez): Ativar Aplicativos não Seguros.
  # Gerenciar Conta Google -> Segurança -> Aplicativos não Seguros -> Habilitar
  # Caso mesmo assim dê o erro: smtplib.SMTPAuthenticationError: (534,
  # Você faz o login no seu e-mail e depois entra em: https://accounts.google.com/DisplayUnlockCaptcha
  msg['From'] = 'email1@gmail.com'
  msg['To'] = 'email2@gmail.com'
  password = "senha"
  msg.add_header('Content-Type', 'text/html')
  msg.set_payload(corpo_email )
    
  s = smtplib.SMTP('smtp.gmail.com: 587')
  s.starttls()
  # Login Credentials for sending the mail
  s.login(msg['From'], password)
  s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
  print('Email enviado')
  
  #Enviar e-mail
enviar_email("Diretoria", tabela_faturamento)
