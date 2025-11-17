# Importación de librerías estándar
import yfinance as yf  # Para obtener datos bursátiles
import pandas as pd  # Para manejar tablas de datos
import smtplib  # Para enviar correos
import os  # Para acceder a variables de entorno
import pdfkit  # Para convertir HTML a PDF
from email.mime.multipart import MIMEMultipart  # Estructura del correo
from email.mime.text import MIMEText  # Parte HTML del correo
from email.mime.application import MIMEApplication  # Adjuntar archivos
from datetime import datetime  # Obtener fecha actual
from dotenv import load_dotenv  # Cargar variables desde archivo .env

# Librerías para interfaz gráfica moderna
import ttkbootstrap as ttk  # Interfaz moderna basada en Bootstrap
from ttkbootstrap.constants import PRIMARY  # Color principal
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # Mostrar gráficos en Tkinter
import matplotlib.pyplot as plt  # Generar gráficos

# Cargar variables de entorno (correo, contraseña, destinatario)
load_dotenv()

# Función para obtener datos bursátiles de AAPL, GOOGL y MSFT
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

# Función para generar el PDF del reporte
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

# Función para enviar el correo con adjuntos
def enviar_email(reporte_html, tickers, panel):
    try:
        remitente = os.getenv("EMAIL_USER")
        password = os.getenv("EMAIL_PASS")
        destinatario = os.getenv("EMAIL_TO")

        if not remitente or not password or not destinatario:
            raise ValueError("Faltan datos en el archivo .env")

        msg = MIMEMultipart("mixed")
        fecha = datetime.now().strftime("%d/%m/%Y")
        msg["Subject"] = f"Reporte Bursátil Automático - {fecha}"
        msg["From"] = remitente
        msg["To"] = destinatario

        cuerpo = MIMEText(reporte_html, "html")
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

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(remitente, password)
            server.sendmail(remitente, destinatario, msg.as_string())

        ttk.Label(panel, text="✅ Correo enviado correctamente", bootstyle="success").pack(pady=10)
    except Exception as e:
        ttk.Label(panel, text=f"❌ Error: {type(e).__name__} - {str(e)}", bootstyle="danger").pack(pady=10)

# Clase principal de la aplicación
class PanelFinanciero:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Panel Financiero")
        self.root.geometry("1000x600")

        # Lista de empresas
        self.tickers = ["AAPL", "GOOGL", "MSFT"]
        self.df = obtener_datos()
        self.reporte_html = self.df.to_html(index=False)
        generar_pdf(self.reporte_html)

        # Panel de navegación lateral
        self.menu = ttk.Frame(root, padding=10)
        self.menu.pack(side="left", fill="y")

        # Panel dinámico donde cambia el contenido
        self.panel_dinamico = ttk.Frame(root, padding=20)
        self.panel_dinamico.pack(side="right", expand=True, fill="both")

        # Botones del menú
        ttk.Label(self.menu, text="Menú", font=("Segoe UI", 14)).pack(pady=10)
        ttk.Button(self.menu, text="📊 Ver tabla", bootstyle=PRIMARY, command=self.mostrar_tabla).pack(pady=5, fill="x")
        for ticker in self.tickers:
            ttk.Button(self.menu, text=f"📈 Gráfico {ticker}", bootstyle=PRIMARY, command=lambda t=ticker: self.mostrar_grafico(t)).pack(pady=5, fill="x")
        ttk.Button(self.menu, text="📤 Enviar correo", bootstyle=PRIMARY, command=self.enviar_correo).pack(pady=10, fill="x")
        ttk.Button(self.menu, text="🚪 Salir", bootstyle="danger", command=root.quit).pack(pady=20, fill="x")

    # Limpia el panel dinámico antes de mostrar nuevo contenido
    def limpiar_panel(self):
        for widget in self.panel_dinamico.winfo_children():
            widget.destroy()

    # Muestra la tabla de datos
    def mostrar_tabla(self):
        self.limpiar_panel()
        ttk.Label(self.panel_dinamico, text="📊 Tabla de Datos Bursátiles", font=("Segoe UI", 14)).pack(pady=10)
        texto = ttk.Text(self.panel_dinamico, width=100, height=20)
        texto.insert("end", self.df.to_string(index=False))
        texto.pack()

    # Muestra el gráfico de un ticker
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

    # Llama a la función de envío de correo
    def enviar_correo(self):
        self.limpiar_panel()
        ttk.Label(self.panel_dinamico, text="📤 Enviando correo...", font=("Segoe UI", 12)).pack(pady=10)
        enviar_email(self.reporte_html, self.tickers, self.panel_dinamico)

# Punto de entrada del programa
if __name__ == "__main__":
    app = ttk.Window(themename="flatly")  # Tema azul moderno
    PanelFinanciero(app)
    app.mainloop()