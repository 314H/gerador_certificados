#!/usr/bin/python3

from xlrd import open_workbook,cellname
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path

def gerar_certificado (nome_certificado, nome_arquivo): 
    nome_do_arquivo_certificado = "certificado_" + nome_arquivo.replace(" ","_") + ".pdf"
    imagem_certificado = Image.open("DA_final.png")
    imagem_certificado_desenho = ImageDraw.Draw(imagem_certificado)
    fonte_certificado = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 200, encoding="unic")

    imagem_certificado_desenho.text((640,1000), nome_certificado, font=fonte_certificado, fill=(0,0,0))
    imagem_certificado = imagem_certificado.convert('RGB')
    imagem_certificado.save(nome_do_arquivo_certificado, "PDF", resolution=100.0)
    return nome_do_arquivo_certificado

def enviar_email (email_de, email_para, email_assunto, email_corpo, nome_do_certificado):
    msg = MIMEMultipart()
    msg['From'] = email_de
    msg['To'] = email_para
    msg['Subject'] = email_assunto
    msg.attach(MIMEText(email_corpo, 'plain'))
    
    arquivo_anexo = os.path.basename(nome_do_certificado)
    attachment = open(nome_do_certificado, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename=arquivo_anexo)

    msg.attach(part)
    email_texto = msg.as_string()

    server.sendmail(email_de, email_para, email_texto)

if __name__ == "__main__":
    planilha_inscricoes = open_workbook('inscricoes.xlsx')
    sheet = planilha_inscricoes.sheet_by_index(0)

    email = ""
    senha = ""
    email_mensagem = "Olá, tudo bem? Segue em anexo o certificado da palestra."
    email_assunto = "Certificado da Palestra XXXX"

    server = smtplib.SMTP("smtp.gmail.com", 587)
    server.starttls()
    server.login(email, senha)

    for i in range(1,sheet.nrows):
        aluno_nome_certificado = sheet.row_values(i)[3]
        aluno_nome_arquivo = sheet.row_values(i)[2]
        aluno_email = sheet.row_values(i)[1]

        nome_do_certificado = gerar_certificado(aluno_nome_certificado, aluno_nome_arquivo)
        enviar_email(email, aluno_email, email_assunto, email_mensagem, nome_do_certificado)