# 📧 Envio Automático de E-mails com Python  

Este projeto é uma solução automatizada para envio de e-mails personalizados a partir de uma planilha Excel, facilitando processos como avaliações de desempenho, notificações e outras comunicações em massa.  

## 🚀 Funcionalidades  

✔️ Leitura automática de uma planilha Excel (`pandas`)  
✔️ Envio de e-mails individuais personalizados (`smtplib`)  
✔️ Autenticação segura via SMTP  
✔️ Registro de status de envio para evitar duplicações  
✔️ Confirmação de leitura e entrega  
✔️ Transformação do script em executável (.exe)  

## 📂 Estrutura do Projeto  
  
📦 notas_automaticas

 ┣ 📂 dist/ # Pasta que gerou o executável<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;┣ 📜 notas_desempenho.xlsx     # Exemplo de planilha (sem dados reais)<br>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;┗ 📜 script_teste.exe          # Executável<br>
 ┗ 📜 script_teste.py            # Código principal do programa<br> 


## 📊 Exemplo de Planilha  

O arquivo **`notas_desempenho.xlsx`** deve conter os seguintes campos:  

| Nome      | E-mail                | Nota  | Status de Envio | Data de Envio |
|-----------|-----------------------|------|----------------|---------------|
| João      | joao@email.com        | 9.5  | Enviado        | 13/03/2025    |
| Maria     | maria@email.com       | 8.7  | (em branco)    | (em branco)   |
| Carlos    | carlos@email.com      | 7.3  | Enviado        | 14/03/2025    |

Caso o **"Status de Envio"** esteja vazio, o e-mail será enviado normalmente. Se já estiver marcado como `"Enviado"`, o script perguntará se deseja reenviar. O campo **"Data de Envio"** será preenchido automaticamente.  

## 📧 Configuração SMTP
O programa utiliza o serviço SMTP do Expresso PE para envio dos e-mails.

Configuração Padrão: <br>
Servidor SMTP: smtps.expresso.pe.gov.br<br>
Porta: 587 (ou 465 dependendo da configuração de segurança)

O programa pedirá as credenciais de login antes de iniciar o envio.

⚠ Importante:
Caso a senha esteja incorreta, o programa exibirá uma mensagem de erro e pedirá para tentar novamente. Para aumentar a segurança, evite armazenar senhas no código.

## 🔄 Transformando em Executável
Para gerar um **`.exe`** que pode ser usado sem instalação de Python, utilize:

```bash
  pyinstaller --onefile --icon=icone.ico script_teste.py
```

Isso criará um arquivo .exe na pasta dist com ícone de prefererência, que pode ser movido para qualquer computador.
