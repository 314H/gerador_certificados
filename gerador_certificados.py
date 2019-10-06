#!/usr/bin/python3

from xlrd import open_workbook,cellname
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from getpass import getpass
import smtplib

'''
  A função "gerar_certificado" recebe dois parâmetros, gera o certificado no formato PNG e retorna uma String com o nome do certificado gerado.
  
  [PARÂMETROS]
    1. nome_certificado - Esse parâmetro é o nome que o aluno deseja que aparece no certificado, por exemplo, "Mateus Müller"
    2. nome_arquivo - Essa é a String que usaremos no nome do arquivo, que será o nome completo, "Mateus Gabriel Müller"

  [VARIÁVEIS]
    1. nome_do_arquivo_certificado - Nome do arquivo depois de virar PNG. Note que ele pega a variável "nome_arquivo" e modifica. Ficará algo como "certificado_Mateus_Gabriel_Müller.png".
    2. imagem_certificado - Esse é objeto do tipo Image que irá conter o certificado que vamos escrever no formato PNG.
    3. fonte_certificado - Fonte usada para escrever no certificado.
    4. imagem_certificado_desenho - Objeto do tipo ImageDraw usado para desenhar na imagem.

  Depois de gerar o certificado com o método "save", é retornado o nome do arquivo, ou seja, a String "certificado_Mateus_Gabriel_Müller.png". Isso é feito para que depois enviemos o anexo correto por e-mail.
'''
def gerar_certificado (nome_certificado, nome_arquivo): 
    nome_do_arquivo_certificado = "certificado_" + nome_arquivo.replace(" ","_") + ".png"
    imagem_certificado = Image.open("DA_final.png")
    imagem_certificado_desenho = ImageDraw.Draw(imagem_certificado)
    fonte_certificado = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 200, encoding="unic")

    imagem_certificado_desenho.text((640,1000), nome_certificado, font=fonte_certificado, fill=(0,0,0))
    imagem_certificado.save(nome_do_arquivo_certificado, "PNG", resolution=100.0)
    return nome_do_arquivo_certificado

'''
  A função "enviar_email" recebe os parâmetros de envio de e-mail e realiza o envio.

  [PARÂMETROS]
    1. email_de - Este é o "from", ou seja, a partir de qual e-mail será enviado. Geralmente, será o e-mail do Diretóri Acadêmico.
    2. email_para - Para quem vamos enviar o e-mail? Esse será o e-mail do aluno na planilha de Excel.
    3. email_assunto - Qual o assunto que vai aparecer no cabeçalho do e-mail? Geralmente algo como "Certificado da Palestra XXXXX".
    4. email_corpo - Essa é a descrição do e-mail. Geralmente algo como "Olá, esse é o certificado da palestra XXXX".
    5. nome_do_certificado - Esse é o nome do certificado que será enviado em anexo. O valor retornado na função "gerar_certificado" é o mesmo do "nome_do_certificado".

  [VARIÁVEIS]
    1. msg - Objeto do tipo MIMEMultipart que vai criar a estrutura do e-mail
    2. arquivo_anexo - Vai armazenar o binário do certificado gerado
    3. part - Objeto do tipo MIMEBase que usamos para definir mais partes do header do e-mail, com os anexos
    4. email_texto - Resultado final quando convertemos todo o resto para String
'''    
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

'''
  A função conecta_smtp() é utilizada para realizar a conexão ao servidor de e-mail.

  [VARIÁVEIS]
    1. servidor_smtp - Armazena a string de servidor SMTP que o usuário digitou.
    2. porta_smtp - Armazena a porta do servidor SMTP que o usuário digitou.
    3. email - Recebe o e-mail que será utilizado para enviar os certificados.
    4. senha - Recebe a senha do e-mail que o usuário digitou.
'''    
    
def conecta_smtp():
    while(True):
        servidor_smtp = input("Informe o servidor SMTP: ")
        porta_smtp = int(input("Informe a porta do servidor SMTP: "))
        email = input("Informe o e-mail que enviará o certificado: ")
        senha = getpass("Informe a senha deste e-mail: ")
        try:
            objeto_conexao_smtp = smtplib.SMTP(servidor_smtp, porta_smtp) # Objeto SMTP
            objeto_conexao_smtp.starttls() # Inicia a conexão
            objeto_conexao_smtp.login(email, senha) # Loga com e-mail e senha
            print("Logado com sucesso.") # Retorno para o usuário.
            break # Se a conexão for bem sucedida, sai do laço
        except smtplib.SMTPException:
            print("Algo deu errado. Confira os dados informados.") # Retorno negativo para usuário

    return objeto_conexao_smtp # Retorna o objeto do servidor SMTP   
    
    
'''
  Esse é o método principal de execução onde é feita a leitura do arquivo .xlsx e a chamada das outras funções.

'''
if __name__ == "__main__":
  
    email = "" # Inicializando a variável para recuperar o e-mail de envio depois.
    objeto_conexao_smtp = conecta_smtp() # Variável objeto_conexao_smtp recebe o objeto SMTP
    
    planilha = input("Informe o nome da planilha de remetentes: ")
    
    planilha_inscricoes = open_workbook(planilha) # Planilha Excel que será lida
    folha_planilha_inscricoes = planilha_inscricoes.sheet_by_index(0) # Qual "sheet" da sua planilha ler?

    email_assunto = input("Informe o assunto do e-mail: ") # Assunto do e-mail
    email_mensagem = input("Informe a mensagem que irá no corpo do e-mail: ") # Tem como melhorar para a pessoas mandar algo personalizado, como um arquivo html

    for i in range(1,folha_planilha_inscricoes.nrows): # Inicia a leitura de toda a planilha
        aluno_nome_certificado = folha_planilha_inscricoes.row_values(i)[3] # Coluna 3
        aluno_nome_arquivo = folha_planilha_inscricoes.row_values(i)[2] # Coluna 2
        aluno_email = folha_planilha_inscricoes.row_values(i)[1] # Coluna 1

        nome_do_certificado = gerar_certificado(aluno_nome_certificado, aluno_nome_arquivo) # Gera o certificado e armazena o nome gerado
        enviar_email(email, aluno_email, email_assunto, email_mensagem, nome_do_certificado) # Chama a função de envio de e-mail enviando a variável gerada logo acima
