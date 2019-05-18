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

    server.sendmail(email_de, email_para, email_texto)

'''
  Esse é o método principal de execução onde é feita a leitura do arquivo .xlsx e a chamada das outras funções.

  [ONDE MEXER?]
    1. email - Altere a variável email abaixo para o seu e-mail do gmail ou qualquer outro
    2. senha - Coloque a senha deste email para que o script consiga acessa-lo
    3. email_mensagem - Qual será o corpo do e-mail?
    4. email_assunto - Qual o assunto do e-email?
    5. planilha_inscricoes - O nome padrão do arquivo que o script vai ler é "inscricoes.xlsx", mas você pode altera-lo.

'''
if __name__ == "__main__":
    planilha_inscricoes = open_workbook('inscricoes.xlsx') # Planilha Excel que será lida
    folha_planilha_inscricoes = planilha_inscricoes.sheet_by_index(0) # Qual "sheet" da sua planilha ler?

    email = "" # E-mail da conta que enviará o e-mail
    senha = "" # Senha da conta que enviará o e-mail
    email_mensagem = "Olá, tudo bem? Segue em anexo o certificado da palestra." # Corpo do e-mail
    email_assunto = "Certificado da Palestra XXXX" # Assunto do e-mail

    server = smtplib.SMTP("smtp.gmail.com", 587) # Servidor do seu provedor, nesse caso do gmail
    server.starttls()
    server.login(email, senha)

    for i in range(1,folha_planilha_inscricoes.nrows): # Inicia a leitura de toda a planilha
        aluno_nome_certificado = folha_planilha_inscricoes.row_values(i)[3] # Coluna 3
        aluno_nome_arquivo = folha_planilha_inscricoes.row_values(i)[2] # Coluna 2
        aluno_email = folha_planilha_inscricoes.row_values(i)[1] # Coluna 1

        nome_do_certificado = gerar_certificado(aluno_nome_certificado, aluno_nome_arquivo) # Gera o certificado e armazena o nome gerado
        enviar_email(email, aluno_email, email_assunto, email_mensagem, nome_do_certificado) # Chama a função de envio de e-mail enviando a variável gerada logo acima