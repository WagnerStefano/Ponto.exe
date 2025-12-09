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
<br><br>
# passso a passo para executar o projeto
Dentro da pasta criada para o projeto devemos criar o .Venv
utilizando o comando 

      python -m venv .venv

após sua criação precisamos ativar com o comando

      .venv\Scripts\Activate.ps1

Se der erro de execução, rode a linha a seguir

      Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
      
após estas execuções você pode fazer um teste com 

      pip --version
que deve retornoar a versão do Python instalada e utilizad no projeto.

Com ambiente virtual ativado devemos Instalar as Dependências do projeto.
utilizando o comando
   
      pip install pywin32 
      apscheduler 
      pystray 
      pillow 
      pyinstaller
Obs: tkinter já vem no Python no Windows.

Verificar se está ok com <br>

      pip list

Com todas essas etapas executadas e validando o funcionamento do ambiente vamos a execução do projeto.

      python ponto.py

Com todo ambiente funcionando e nenhum erro no código podemos partir para 

# Gerando o .EXE

Aqui fica um aviso de que todas as dependencias devem estar instaladas dento o Venv para que haja a correta execução. <BR>
Para gerar o .exe podemos utilizar o comando 

      pyinstaller --noconfirm --windowed --onefile --icon=logo.ico ponto.py

 Que trará o diretório dist com o executavel <BR>
 Após esta execução basta transferir o .EXE para qualquer maquina que funcionará normalmente (Assim espero :D)
      <img width="801" height="749" alt="image" src="https://github.com/user-attachments/assets/b98d031e-5650-45bf-8e4f-0553dd1e9e2c" />
      <img width="252" height="173" alt="image" src="https://github.com/user-attachments/assets/7355b3c5-0858-46b5-a00e-4f7d9c76e5dd" />



Ainda existem pontos de acerto, melhorias a fazer e muitos testes mas no momento ainda estou desenvolvendo a ferramenta e amadurecendo a ideia.
