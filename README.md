# Ponto.exe
A time card assistant

   Este projeto é uma aplicação desktop desenvolvida em Python que gera um lembrete e automatiza a comunicação diária sobre o status do ponto de trabalho.
A ferramenta foi criada para resolver um problema cotidiano, o esquecimento do ponto, muito corriqueiro mas que passa desaparcebido com facilidade, algo
que eu não imaginaria que causaria problemas mas que deve ser corrigido e automatizado, graças a este problema surge uma oportunidade de aprender e
aplicar o que foi aprendido.

O objetivo principal foi eliminar erros humanos, reduzir atrasos e agilizar a rotina operacional.

# Tecnologias utilizadas
    | Tecnologia  | Uso                                  | Version     |
    | ----------- | ------------------------------------ | ----------- |
    | Python      | Linguagem principal                  | 3.14        |
    | Tkinter     | Interface gráfica                    |             |
    | win32com    | Envio de e-mail via Outlook          |             |
    | APScheduler | Agenda diária                        | 3.11.1      |
    | PyInstaller | Empacotamento para .exe              |             |
    | pystray     | Ícone na bandeja                     | 0.19.5      |
    | PIL         | Ícone do app                         |             |
    | winreg      | Registro de inicialização no Windows |             |

# Funcionalidades
 Interface gráfica simples e intuitiva com Tkinter
 
 Botões com ações rápidas:
  Ainda em atendimento - para aviso pós horario<br>
  Finalizando - em caso de esquecimento do ponto esteja justificado em email o horario de saida<br>
  Mensagem personalizada - para informações que sejam relevantes como algum evento ou atividade pós horario de acordado.<br><br>

 Envio de e-mail automático utilizando Outlook com destinatários pré-configurados no código<br>
  Confirmação visual ao enviar.<br><br>

 Notificação automática diária às 16:45<br>
  Lembrete para enviar o e-mail apenas clicando em uma das  3 opções disponiveis.<br><br>

 Execução em background<br>
   Minimiza para bandeja do Windows<br><br>

 Inicia automaticamente com o Windows<br>
   <br>Build em .exe portátil via PyInstaller<br><br>

   # Build do Executável
   O projeto pode ser empacotado como um .exe portátil, que roda sem instalação.

Exemplo de comando utilizado:

    pyinstaller --noconfirm --windowed --onefile --icon=logo.ico Nome_Do_Arquivo.py
Após gerar, copie apenas o .exe para qualquer máquina da empresa.

# Estrutura do Projeto

      ponto_monitor/.venv
      │
      ├── ponto.py
      ├── logo.ico
      └── .venv/


