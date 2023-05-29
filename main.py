import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
import getpass
import csv
from email.message import EmailMessage
import ssl
import smtplib
import os


class Aplicacion(ttk.Frame):

    def __init__(self, parent):
        super().__init__(parent)
        self.mails_ = []
        self.cv_ = None
        self.cv_name = None
        """Labels"""
        style = ttk.Style()
        style.theme_use('alt')
        style.configure("Custom.TLabel",background="#adb5c1", foreground = '#313237',relief="flat")
        style.map('Custom.TLabel', background=[('active', '#adb5c1')])

        ttk.Label(parent, text="REMITENTE", font=("Arial", 11,"bold"),style="Custom.TLabel").place(x=20, y=20, width=95)
        ttk.Label(parent, text="PASSWORD", font=("Arial", 11,"bold"),style="Custom.TLabel").place(x=20, y=50, width=95)
        ttk.Label(parent, text="ASUNTO", font=("Arial", 11,"bold"),style="Custom.TLabel").place(x=20, y=80, width=95)  # Asunto
        ttk.Label(parent, text="SELECCIONE LISTA\nDE MAIL EN .CSV", font=("Arial", 12,"bold"),style="Custom.TLabel").place(x=20, y=120)
        ttk.Label(parent, text="REDACTE EL MAIL", font=("Arial", 11,"bold"),style="Custom.TLabel").place(x=95, y=180)

        """Entry"""
        style = ttk.Style()
        style.theme_use('alt')
        style.configure("Custom.TEntry",fieldbackground="#adb5c1", foreground = 'black',relief="flat")
        style.map('Custom.TEntry', fieldbackground=[('active', '#adb5c1')])

        self.etiqueta_mail_caja = ttk.Entry(parent, style="Custom.TEntry")
        self.etiqueta_mail_caja.place(x=150, y=20, width=200)  # Mail
        self.password_caja = ttk.Entry(parent,style="Custom.TEntry")  # Pass
        self.password_caja.place(x=150, y=50, width=200)
        self.asunto = ttk.Entry(parent,style="Custom.TEntry")  # Asunto
        self.asunto.place(x=150, y=80, width=200)

        """modifico el stilo de los botones"""

        style = ttk.Style()
        style.theme_use('alt')
        style.configure("Custom.TButton",background="#adb5c1",
                        foreground='#313237',
                        relief="flat",
                        font=("Arial",11,"bold"))
        style.map('Custom.TButton', background=[('active', '#76798f')])

        """botones"""
        self.boton_seleccion_mails = ttk.Button(parent, text="OPEN",
                                                command=self.openFile,
                                                style="Custom.TButton",
                                                )  # mails

        self.boton_seleccion_mails.place(x=270, y=125, width=80, height=30)
        self.cv = ttk.Button(parent, text="INS. CV", command=self.openCV,style="Custom.TButton")
        self.cv.place(x=350, y=220, width=70, height=30)  # CV
        self.enviar = ttk.Button(parent, text="ENVIAR", command=self.enviarMail,style="Custom.TButton")
        self.enviar.place(x=350, y=450,width=70, height=30)
        """Texto"""
        self.texto = tk.Text(parent)
        self.texto.config(background="#adb5c1",foreground = '#313237', relief="flat",font=("Arial",12,"bold"))
        self.texto.place(x=20, y=210, width=300, height=280)

    def openFile(self):
        user = getpass.getuser()
        filepath = tk.filedialog.askopenfilename(initialdir=f'C:/Users/{user}')
        with open(filepath, newline="", encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            self.mails_ = list(reader)
        self.boton_seleccion_mails.config(text=filepath)

    def openCV(self):
        user = getpass.getuser()
        filepath = tk.filedialog.askopenfilename(initialdir=f'C:/Users/{user}')
        self.cv_name = os.path.split(filepath)[1]
        with open(filepath, "rb") as content_file:
            self.cv_ = content_file.read()
        self.cv.config(text=filepath)

    def enviarMail(self):
        emisor = self.etiqueta_mail_caja.get()
        password = self.password_caja.get()
        asunto = self.asunto.get()
        cuerpo = self.texto.get("1.0", 'end-1c')
        """Crear el entorno del mail"""
        for mail in self.mails_:
            em = EmailMessage()
            em["From"] = emisor
            em["To"] = mail[0]
            em["Subject"] = asunto
            em.set_content(cuerpo)

            # Cargo el pdf
            em.add_attachment(self.cv_, maintype="application", subtype="pdf", filename=self.cv_name)

            # Creo el protocolo de envio
            contexto = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=contexto) as smtp:
                smtp.login(emisor, password)
                smtp.sendmail(emisor, mail[0], em.as_string())


if __name__ == '__main__':
    root = tk.Tk()  # aca empiezo la ventana principal
    root.geometry("430x500")
    app = Aplicacion(root)

    root.resizable(False, False) # Bloqueo la re dimension
    root.attributes("-toolwindow", True)

    root.title("AUTOMATED-MAIL")
    root.configure(bg="#4a4e53")

    root.mainloop()
