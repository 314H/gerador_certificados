#!/usr/bin/python3

from xlrd import open_workbook,cellname
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib
import re

def gerar_certificado (nome_certificado, nome_arquivo, nome_certificado_template): 
    nome_do_arquivo_certificado = "certificado_" + nome_arquivo.replace(" ","_") + ".png"
    imagem_certificado = Image.open(nome_certificado_template)
    imagem_certificado_desenho = ImageDraw.Draw(imagem_certificado)
    fonte_certificado = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 200, encoding="unic")

    imagem_certificado_desenho.text((640,1000), nome_certificado, font=fonte_certificado, fill=(0,0,0))
    imagem_certificado.save(nome_do_arquivo_certificado, "PNG", resolution=100.0)
    return nome_do_arquivo_certificado

def enviar_email (email_de, email_para, email_assunto, email_corpo, nome_do_certificado):
    msg = MIMEMultipart()
    msg['From'] = email_de
    msg['To'] = email_para
    msg['Subject'] = email_assunto
    msg.attach(MIMEText(email_corpo, 'plain'))

    arquivo_anexo = open(nome_do_certificado, "rb")
    part = MIMEBase('image', 'png')
    part.set_payload(arquivo_anexo.read())
    part.add_header('Content-Disposition', 'attachment', filename=nome_do_certificado)
    encoders.encode_base64(part)
    msg.attach(part)
    email_texto = msg.as_string()

    objeto_conexao_smtp.sendmail(email_de, email_para, email_texto)
   
def conecta_smtp(email, senha, servidor_smtp, porta_smtp):
    while(True):
        email = email
        senha = senha
        servidor_smtp = servidor_smtp
        porta_smtp = porta_smtp
        try:
            objeto_conexao_smtp = smtplib.SMTP(servidor_smtp, porta_smtp)
            objeto_conexao_smtp.starttls()
            objeto_conexao_smtp.login(email, senha)
            print("Logado com sucesso.")
            break
        except smtplib.SMTPException:
            print("Algo deu errado. Confira os dados informados.") 
    return objeto_conexao_smtp # Retorna o objeto do servidor SMTP

def busca_parametros_no_arquivo():
  lista_parametros = []
  with open("parametros.txt","r") as parametros_txt:
    for linha in parametros_txt:
      lista_parametros.append(re.split("=",linha)[1].rstrip("\n"))
  return lista_parametros # Retorna lista com os parâmetros do arquivo

if __name__ == "__main__":

  email, senha, servidor_smtp, porta_smtp, email_titulo, email_corpo, planilha_participantes, foto_template_certificado = busca_parametros_no_arquivo()
  objeto_conexao_smtp = conecta_smtp(email, senha, servidor_smtp, porta_smtp)
      
  planilha_inscricoes = open_workbook(planilha_participantes) # Planilha Excel que será lida
  folha_planilha_inscricoes = planilha_inscricoes.sheet_by_index(0) # Qual "sheet" da sua planilha ler?

  for i in range(1,folha_planilha_inscricoes.nrows): # Inicia a leitura de toda a planilha
    aluno_nome_certificado = folha_planilha_inscricoes.row_values(i)[3] # Coluna 3
    aluno_nome_arquivo = folha_planilha_inscricoes.row_values(i)[2] # Coluna 2
    aluno_email = folha_planilha_inscricoes.row_values(i)[1] # Coluna 1

    nome_do_certificado = gerar_certificado(aluno_nome_certificado, aluno_nome_arquivo, foto_template_certificado) # Gera o certificado e armazena o nome gerado
    enviar_email(email, aluno_email, email_titulo, email_corpo, nome_do_certificado) # Chama a função de envio de e-mail enviando a variável gerada logo acima