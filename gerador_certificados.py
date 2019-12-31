#!/usr/bin/env python3

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
import json

class Gerador_Certificados ():
    def __init__ (self):
        arquivo_credenciais = open("credenciais.json","r")
        dados_arquivo_credenciais = json.load(arquivo_credenciais)
        self.email = dados_arquivo_credenciais["e-mail"]
        self.senha = dados_arquivo_credenciais["senha"]
        self.servidor_smtp = dados_arquivo_credenciais["servidor_smtp"]
        self.porta_smtp = dados_arquivo_credenciais["porta_smtp"]
        self.email_titulo = dados_arquivo_credenciais["e-mail_titulo"]
        self.email_corpo = dados_arquivo_credenciais["e-mail_corpo"]
        self.planilha_participantes = dados_arquivo_credenciais["planilha_participantes"]
        self.foto_template_certificado = dados_arquivo_credenciais["foto_template_certificado"]
        print(self.email_titulo)
        arquivo_credenciais.close()

    def gerar_certificado (self, nome_certificado, nome_arquivo): 
        nome_do_arquivo_certificado = "certificado_" + nome_arquivo.replace(" ","_") + ".png"
        imagem_certificado = Image.open(self.foto_template_certificado)
        imagem_certificado_desenho = ImageDraw.Draw(imagem_certificado)
        fonte_certificado = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 200, encoding="unic")
        imagem_certificado_desenho.text((640,1000), nome_certificado, font=fonte_certificado, fill=(0,0,0))
        imagem_certificado.save(nome_do_arquivo_certificado, "PNG", resolution=100.0)
        return nome_do_arquivo_certificado

    def enviar_email (self, email_para, nome_do_certificado):
        msg = MIMEMultipart()
        msg['From'] = self.email
        msg['To'] = email_para
        msg['Subject'] = self.email_titulo
        msg.attach(MIMEText(self.email_corpo, 'plain'))
        arquivo_anexo = open(nome_do_certificado, "rb")
        part = MIMEBase('image', 'png')
        part.set_payload(arquivo_anexo.read())
        part.add_header('Content-Disposition', 'attachment', filename=nome_do_certificado)
        encoders.encode_base64(part)
        msg.attach(part)
        email_texto = msg.as_string()
        self.objeto_conexao_smtp.sendmail(self.email, email_para, email_texto)
    
    def conecta_smtp(self):
        while(True):
            try:
                self.objeto_conexao_smtp = smtplib.SMTP(self.servidor_smtp, self.porta_smtp)
                self.objeto_conexao_smtp.starttls()
                self.objeto_conexao_smtp.login(self.email, self.senha)
                print("Conectado com SMTP!")
                break
            except smtplib.SMTPException:
                print("Houve um problema com a conexão SMTP!")

    def ler_planilha_e_executar (self):
        self.conecta_smtp()
        planilha_inscricoes = open_workbook(self.planilha_participantes) # Planilha Excel que será lida
        folha_planilha_inscricoes = planilha_inscricoes.sheet_by_index(0) # Qual "sheet" da sua planilha ler?
        for linha_planilha in range(1,folha_planilha_inscricoes.nrows): # Inicia a leitura de toda a planilha
            aluno_nome_certificado = folha_planilha_inscricoes.row_values(linha_planilha)[3] # Coluna 3
            aluno_nome_arquivo = folha_planilha_inscricoes.row_values(linha_planilha)[2] # Coluna 2
            aluno_email = folha_planilha_inscricoes.row_values(linha_planilha)[1] # Coluna 1
            # Gera o certificado e armazena o nome gerado
            nome_do_certificado = self.gerar_certificado(aluno_nome_certificado, aluno_nome_arquivo) 
            self.enviar_email(aluno_email, nome_do_certificado) 

if __name__ == "__main__":
    gerador_certificados = Gerador_Certificados()
    gerador_certificados.ler_planilha_e_executar()