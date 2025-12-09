import tkinter as tk
from datetime import datetime
import threading
import sys
from apscheduler.schedulers.background import BackgroundScheduler
import pystray
from PIL import Image
import os
import winreg
import win32com.client as win32  # <-- Outlook

# ==========================
# CONFIGURAÇÕES
# ==========================

ICON_PATH = "logo.png"  # 32x32
ALERTA_HORA = "16:45"
REGISTRY_NAME = "AppAtendimento"

DESTINATARIOS = [
    "wagnerstefano98@gmail.com",
    "email2@empresa.com"
]
# ==========================


# ==========================
# FUNÇÃO ENVIAR E-MAIL (Outlook)
# ==========================

def enviar_email(assunto, texto):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    corpo = f"{texto}\n\nData e Hora: {agora}"

    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        mail.To = "; ".join(DESTINATARIOS)
        mail.Subject = assunto
        mail.Body = corpo

        mail.Send()

        mostrar_mensagem("E-mail enviado com sucesso!")

    except Exception as e:
        mostrar_mensagem("Erro ao enviar o e-mail!")
        print("Erro:", e)


# ==========================
# MENSAGEM VISUAL
# ==========================

def mostrar_mensagem(texto):
    popup = tk.Toplevel()
    popup.title("Informação")
    popup.geometry("250x110")

    label = tk.Label(popup, text=texto, font=("Arial", 11))
    label.pack(pady=10)

    btn = tk.Button(popup, text="OK", command=popup.destroy)
    btn.pack(pady=5)

    popup.lift()
    popup.attributes('-topmost', True)
    popup.after(100, lambda: popup.attributes('-topmost', True))


# ==========================
# AÇÕES DOS BOTÕES
# ==========================

def ainda_em_atendimento():
    enviar_email("Status: Ainda em atendimento", "O cliente ainda está em atendimento.")


def finalizando():
    enviar_email("Status: Finalizando", "O atendimento está sendo finalizado.")


def mensagem_custom():
    texto = entrada.get()
    if not texto.strip():
        mostrar_mensagem("Digite uma mensagem.")
        return

    enviar_email("Status: Atualização", texto)
    entrada.delete(0, tk.END)


# ==========================
# ALERTA DIÁRIO
# ==========================

def mostrar_alerta():
    agora = datetime.now().strftime("%H:%M")
    popup = tk.Toplevel()
    popup.title("Lembrete")
    popup.geometry("300x120")

    label = tk.Label(popup, text=f"São {agora}. \nNão esqueça de enviar o e-mail!", font=("Arial", 12))
    label.pack(pady=10)

    btn = tk.Button(popup, text="OK", command=popup.destroy)
    btn.pack(pady=5)

    popup.lift()
    popup.attributes('-topmost', True)
    popup.after(100, lambda: popup.attributes('-topmost', True))


scheduler = BackgroundScheduler()
hora, minuto = map(int, ALERTA_HORA.split(":"))
scheduler.add_job(mostrar_alerta, "cron", hour=hora, minute=minuto)
scheduler.start()


# ==========================
# BANDEJA DO SISTEMA
# ==========================

def criar_icone():
    image = Image.open(ICON_PATH)

    def restaurar(icon, item):
        icon.stop()
        janela.deiconify()

    menu = pystray.Menu(
        pystray.MenuItem("Restaurar", restaurar),
        pystray.MenuItem("Sair", lambda icon, item: sys.exit())
    )

    icon = pystray.Icon("app", image, "Atendimento", menu)
    icon.run()


def minimizar(event=None):
    janela.withdraw()
    threading.Thread(target=criar_icone, daemon=True).start()


# ==========================
# INICIAR COM WINDOWS
# ==========================

def adicionar_inicio_windows():
    path = os.path.abspath(sys.argv[0])

    chave = winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Run",
        0,
        winreg.KEY_SET_VALUE
    )

    winreg.SetValueEx(chave, REGISTRY_NAME, 0, winreg.REG_SZ, path)
    winreg.CloseKey(chave)


def verificar_inicio_windows():
    try:
        chave = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_READ
        )
        valor, tipo = winreg.QueryValueEx(chave, REGISTRY_NAME)
        winreg.CloseKey(chave)
    except FileNotFoundError:
        adicionar_inicio_windows()


# ==========================
# INTERFACE
# ==========================

janela = tk.Tk()
janela.title("Status Atendimento")

btn1 = tk.Button(janela, text="Ainda em atendimento", width=30, command=ainda_em_atendimento)
btn1.pack(pady=5)

btn2 = tk.Button(janela, text="Finalizando", width=30, command=finalizando)
btn2.pack(pady=5)

entrada = tk.Entry(janela, width=40)
entrada.pack(pady=5)

btn3 = tk.Button(janela, text="Enviar mensagem", width=30, command=mensagem_custom)
btn3.pack(pady=5)

janela.protocol("WM_DELETE_WINDOW", minimizar)

verificar_inicio_windows()

janela.mainloop()
