import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4 #595,542

# olhar a base de dados/importar a base de dados
tabela = pd.read_csv("horas_extras.csv")
print(tabela)


# criar um relatório para cada item da base de dados
for linha in tabela.index:
  # criar 1 relatório
  nome = tabela.loc[linha, "Nome"]
  departamento = tabela.loc[linha, "Departamento"]
  horas_extras = tabela.loc[linha, "Horas Extras"]
  mes_ref = tabela.loc[linha, "Referência"]
  
  documento_pdf = canvas.Canvas(f"Relatório de Horas Extras {nome}.pdf", A4)
  
  # titulo
  #escolhe a fonte
  documento_pdf.setFont("Helvetica-Bold", 18)
  documento_pdf.drawString(100, 750, f"Relatório de Horas Extras {nome}")
  
  # escreve o texto
  documento_pdf.setFont("Helvetica", 12)
  distancia = 30
  documento_pdf.drawString(100, 700, f"Nome: {nome}")
  documento_pdf.drawString(100, 700 - 1 * distancia, f"Departamento: {departamento}")
  documento_pdf.drawString(100, 700 - 2 * distancia, f"Quantidade de Horas Extras: {horas_extras}")
  documento_pdf.drawString(100, 700 - 3 * distancia, f"Mês de Referência: {mes_ref}")
  documento_pdf.drawString(100, 700 - 4 * distancia, "Esse relatório foi gerado automaticamente com Python")
  
  
  # salvar os relatórios
  documento_pdf.save()
  
  
