import win32com.client as win32
from datetime import date

data_atual = date.today()

vendedores2 = []

vendedores = []

roberta = []
leandro = []
gaby = []

j = 0
for i in vendedores:
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)

    email.To = f"{vendedores[j][1]}" 
    email.Subject = "Fechamento mensal de vendas TECNO FOODS" #assunto
    email.HTMLBody = f"""
    <p>Olá {vendedores[j][0]}</p>

    <p>Segue o fechamento de suas vendas até o dia {data_atual}.</p>

    <p>Observação: Os dados da LMT possuem um mês de atraso.</p>

    <p>Abraço e Boas Vendas!</p>

    <p>Atenciosamente Gustavo Perez!</p>
    """
    anexo = f"C://Users/Tecnofoods/Desktop/Resumos/ResumoCoord/{vendedores[j][0]}.xlsx"
    email.Attachments.Add(anexo)

    email.Send()
    

    j = j + 1

outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)

email.To = f"{vendedores2[0][1]};{roberta[0][1]};{leandro[0][1]}" 
email.Subject = "Fechamento mensal de vendas TECNO FOODS" #assunto
email.HTMLBody = f"""
<p>Olá {roberta[0][0]} e {leandro[0][0]}</p>

<p>Segue o fechamento de vendas até o dia {data_atual}.</p>

<p>Observação: Os dados da LMT possuem um mês de atraso.</p>

<p>Abraço e Boas Vendas!</p>

<p>Atenciosamente Gustavo Perez!</p>
"""
anexo = "C://Users/"
email.Attachments.Add(anexo)

email.Send()


print("E-mail enviado!")


