# Gerador de Certificados para Diretórios Acadêmicos

Este foi um projeto desenvolvido para o Diretório Acadêmico de Enfermagem da Universidade Feevale, mas que decidi compartilhar o código-fonte para também ajudar outras pessoas. Note que, por enquanto, será necessário ter alguns conhecimentos de programação para pode utiliza-lo da melhor maneira. Sugiro que você chame algum amigo(a) que entenda o básico e tenho certeza que tudo vai dar certo! :)

## Como utilizar?

1. O primeiro passo é configurar a conta de e-mail do seu diretório para aceitar aplicativos menos seguros. Como a aplicação que desenvolvi (assim como qualquer outra desenvolvida por um mero mortal) não é "assinada digitalmente" ou que possui alguma conexão encriptada, será necessário habilitar aplicativos menos seguro. Clique aqui(https://myaccount.google.com/security) e desative a opção **Acesso a app menos seguro**.

2. Agora você precisa clonar este repositório para poder utilizar.

```
$ git clone https://github.com/mateusmuller/gerador_certificados
$ cd gerador_certificados
```

3. Show, agora você já tem os arquivos necessários. Agora você precisa abrir o arquivo **gerador_certificados.py** e editar as opções de conta (elas estão lá pela linha 85-90).

* email - E-mail da conta que enviará o e-mail
* senha - Senha da conta que enviará o e-mail
* email_mensagem - Mensagem no corpo do e-mail
* email_assunto - Assunto do e-mail

4. Coisas que você precisa se atentar: A imagem do certificado e o arquivo de Excel.

* A imagem do certificado sempre se chama **DA_final.png**, então renomeie o seu arquivo para o mesmo nome e coloque nessa pasta.
* O arquivo de Excel sempre se chama **inscricoes.xlsx**, então você também pode renomear. Outra coisa interessante é que sempre segue o mesmo padrão, onde a coluna 1 é e-mail, coluna 2 é o nome completo e a coluna 3 é o nome que deve aparecer no certificado.
* Na linha 34 é definido a posição do nome que será escrito (640,1000). Talvez você precise mudar isso, dependendo do tamanho do seu certificado.

5. Finalizado, basta executar o script com:

```
$ ./gerador_certificados.py
```

ou

```
$ python3 gerador_certificados.py
```

## Dicas

Sugiro usar o Google Forms para gerar o formulário de inscrição das palestras e depois só baixar o .xlsx e atribuir ao script.

## Licença

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
