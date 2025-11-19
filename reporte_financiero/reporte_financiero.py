# reporte_financiero.py

import yfinance as yf
import pandas as pd
import smtplib
import os
import pdfkit
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
from dotenv import load_dotenv

import ttkbootstrap as ttk
from ttkbootstrap.constants import PRIMARY
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt

load_dotenv()

def obtener_datos():
    tickers = ["AAPL", "GOOGL", "MSFT"]
    resultados = []

    for ticker in tickers:
        datos = yf.Ticker(ticker).history(period="2d")
        if len(datos) < 2:
            continue

        cierre_hoy = datos['Close'].iloc[-1]
        cierre_ayer = datos['Close'].iloc[-2]
        cambio_pct = ((cierre_hoy - cierre_ayer) / cierre_ayer) * 100
        high = datos['High'].iloc[-1]
        low = datos['Low'].iloc[-1]
        rango_pct = ((high - low) / cierre_hoy) * 100

        resultados.append({
            "Ticker": ticker,
            "Precio Cierre": round(cierre_hoy, 2),
            "Cambio Diario (%)": round(cambio_pct, 2),
            "Rango Diario (%)": round(rango_pct, 2)
        })

    return pd.DataFrame(resultados)

def generar_pdf(reporte_html):
    opciones = {
        "encoding": "UTF-8",
        "page-size": "A4",
        "margin-top": "10mm",
        "margin-bottom": "10mm",
        "margin-left": "10mm",
        "margin-right": "10mm"
    }

    path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)

    with open("reporte.html", "w", encoding="utf-8") as f:
        f.write(reporte_html)

    pdfkit.from_file("reporte.html", "reporte.pdf", options=opciones, configuration=config)

def enviar_email(remitente, destinatario, asunto, cuerpo_html, tickers, panel):
    try:
        msg = MIMEMultipart("mixed")
        msg["Subject"] = asunto
        msg["From"] = remitente
        msg["To"] = destinatario

        cuerpo = MIMEText(cuerpo_html, "html")
        msg.attach(cuerpo)

        with open("reporte.pdf", "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="pdf")
            adjunto.add_header("Content-Disposition", "attachment", filename="reporte.pdf")
            msg.attach(adjunto)

        for ticker in tickers:
            nombre_archivo = f"grafico_{ticker}.pdf"
            if os.path.exists(nombre_archivo):
                with open(nombre_archivo, "rb") as f:
                    adjunto_grafico = MIMEApplication(f.read(), _subtype="pdf")
                    adjunto_grafico.add_header("Content-Disposition", "attachment", filename=nombre_archivo)
                    msg.attach(adjunto_grafico)

        with smtplib.SMTP("localhost") as server:
            server.sendmail(remitente, destinatario, msg.as_string())

        ttk.Label(panel, text="✅ Correo enviado correctamente", bootstyle="success").pack(pady=10)
    except Exception as e:
        ttk.Label(panel, text=f"❌ Error: {type(e).__name__} - {str(e)}", bootstyle="danger").pack(pady=10)

class PanelFinanciero:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Panel Financiero")
        self.root.geometry("1000x600")

        self.tickers = ["AAPL", "GOOGL", "MSFT"]
        self.df = obtener_datos()
        self.reporte_html = self.df.to_html(index=False)
        generar_pdf(self.reporte_html)

        self.menu = ttk.Frame(root, padding=10)
        self.menu.pack(side="left", fill="y")

        self.panel_dinamico = ttk.Frame(root, padding=20)
        self.panel_dinamico.pack(side="right", expand=True, fill="both")

        ttk.Label(self.menu, text="Menú", font=("Segoe UI", 14)).pack(pady=10)
        ttk.Button(self.menu, text="📊 Ver tabla", bootstyle=PRIMARY, command=self.mostrar_tabla).pack(pady=5, fill="x")
        for ticker in self.tickers:
            ttk.Button(self.menu, text=f"📈 Gráfico {ticker}", bootstyle=PRIMARY, command=lambda t=ticker: self.mostrar_grafico(t)).pack(pady=5, fill="x")
        ttk.Button(self.menu, text="📤 Enviar correo", bootstyle=PRIMARY, command=self.formulario_correo).pack(pady=10, fill="x")
        ttk.Button(self.menu, text="🚪 Salir", bootstyle="danger", command=root.quit).pack(pady=20, fill="x")

    def limpiar_panel(self):
        for widget in self.panel_dinamico.winfo_children():
            widget.destroy()

    def mostrar_tabla(self):
        self.limpiar_panel()
        ttk.Label(self.panel_dinamico, text="📊 Tabla de Datos Bursátiles", font=("Segoe UI", 14)).pack(pady=10)
        texto = ttk.Text(self.panel_dinamico, width=100, height=20)
        texto.insert("end", self.df.to_string(index=False))
        texto.pack()

    def mostrar_grafico(self, ticker):
        self.limpiar_panel()
        datos = yf.Ticker(ticker).history(period="1mo")
        cierre = datos["Close"]

        fig, ax = plt.subplots(figsize=(8, 4))
        ax.plot(cierre.index, cierre.values, marker="o", linestyle="-", color="blue")
        ax.set_title(f"Precio de Cierre - {ticker}")
        ax.set_xlabel("Fecha")
        ax.set_ylabel("USD")
        ax.grid(True)
        fig.tight_layout()
        fig.savefig(f"grafico_{ticker}.pdf")

        canvas = FigureCanvasTkAgg(fig, master=self.panel_dinamico)
        canvas.draw()
        canvas.get_tk_widget().pack()

    def formulario_correo(self):
        self.limpiar_panel()
        ttk.Label(self.panel_dinamico, text="📤 Enviar Reporte por Correo", font=("Segoe UI", 16, "bold")).pack(pady=10)

        marco = ttk.Frame(self.panel_dinamico, padding=10)
        marco.pack(pady=10, fill="both", expand=True)

        ttk.Label(marco, text="Correo Electronico:").pack(anchor="w", pady=(5, 0))
        entry_remitente = ttk.Entry(marco, width=50)
        entry_remitente.pack(pady=5)

        destinatario = os.getenv("EMAIL_TO", "")
        ttk.Label(marco, text="Correo destinatario:").pack(anchor="w", pady=(10, 0))
        ttk.Label(marco, text=destinatario, font=("Segoe UI", 10, "bold"), foreground="blue").pack(anchor="w", pady=5)

        ttk.Label(marco, text="Título del asunto:").pack(anchor="w", pady=(10, 0))
        entry_asunto = ttk.Entry(marco, width=50)
        entry_asunto.pack(pady=5)

        ttk.Label(marco, text="Asunto:").pack(anchor="w", pady=(10, 0))
        entry_comentario = ttk.Text(marco, width=50, height=6)
        entry_comentario.pack(pady=5)

        def ejecutar_envio():
            remitente = entry_remitente.get().strip()
            asunto = entry_asunto.get().strip()
            comentario = entry_comentario.get("1.0", "end").strip()

            if not remitente or not asunto or not comentario:
                ttk.Label(self.panel_dinamico, text="❌ Todos los campos son obligatorios", bootstyle="danger").pack(pady=5)
                return

            cuerpo_html = f"<h3>{asunto}</h3><p>{comentario}</p><hr>{self.reporte_html}"
            enviar_email(remitente, destinatario, asunto, cuerpo_html, self.tickers, self.panel_dinamico)

        ttk.Button(self.panel_dinamico, text="📨 Enviar ahora", bootstyle="success-outline", command=ejecutar_envio).pack(pady=15)

if __name__ == "__main__":
    app = ttk.Window(themename="flatly")
    PanelFinanciero(app)
    app
    app.mainloop()