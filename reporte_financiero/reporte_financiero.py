# reporte_financiero.py
# Panel financiero con interfaz gráfica usando ttkbootstrap.
# Descarga datos bursátiles, genera reportes en PDF con estilo Bootstrap,
# muestra tablas y gráficos mejorados, y permite enviar todo por correo electrónico.

import yfinance as yf              # Librería para obtener datos financieros desde Yahoo Finance
import pandas as pd                # Manejo de datos tabulares con DataFrames
import smtplib                     # Envío de correos vía protocolo SMTP
import os                          # Acceso a variables de entorno y funciones del sistema
import pdfkit                      # Conversión de HTML a PDF usando wkhtmltopdf
from email.mime.multipart import MIMEMultipart       # Estructura de correo con múltiples partes
from email.mime.text import MIMEText                 # Parte de texto/HTML del correo
from email.mime.application import MIMEApplication   # Parte de adjuntos (PDF, etc.)
from dotenv import load_dotenv      # Carga variables desde archivo .env

import ttkbootstrap as ttk          # Interfaz gráfica moderna basada en Tkinter con estilos tipo Bootstrap
from ttkbootstrap.constants import PRIMARY           # Constante de estilo Bootstrap para botones
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg  # Integración de gráficos en Tkinter
import matplotlib.pyplot as plt     # Creación de gráficos con Matplotlib
import matplotlib.dates as mdates   # Formateo de fechas en el eje X de los gráficos

# Cargar variables de entorno desde archivo .env (EMAIL_USER, EMAIL_PASS, EMAIL_TO, etc.)
load_dotenv()


# -------------------------------
# Función para obtener datos bursátiles
# -------------------------------
def obtener_datos():
    tickers = ["AAPL", "GOOGL", "MSFT"]   # Lista de símbolos a consultar
    resultados = []                       # Lista para acumular resultados

    for ticker in tickers:                # Iterar cada símbolo
        datos = yf.Ticker(ticker).history(period="2d")   # Descargar histórico de 2 días
        if len(datos) < 2:                # Si no hay suficientes datos, saltar
            continue

        cierre_hoy = datos['Close'].iloc[-1]   # Precio de cierre del último día
        cierre_ayer = datos['Close'].iloc[-2]  # Precio de cierre del día anterior
        cambio_pct = ((cierre_hoy - cierre_ayer) / cierre_ayer) * 100  # Variación porcentual diaria

        high = datos['High'].iloc[-1]     # Máximo del último día
        low = datos['Low'].iloc[-1]       # Mínimo del último día
        rango_pct = ((high - low) / cierre_hoy) * 100   # Rango porcentual intradiario

        resultados.append({               # Agregar resultados redondeados a la lista
            "Ticker": ticker,
            "Precio Cierre": round(cierre_hoy, 2),
            "Cambio Diario (%)": round(cambio_pct, 2),
            "Rango Diario (%)": round(rango_pct, 2)
        })

    return pd.DataFrame(resultados)       # Convertir lista de dicts en DataFrame


# -------------------------------
# Función para generar PDF con estilo Bootstrap
# -------------------------------
def generar_pdf(df):
    # HTML con estilos Bootstrap para tabla
    reporte_html = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <link rel="stylesheet"
              href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css">
    </head>
    <body class="p-4">
        <h2 class="text-center mb-4">📊 Reporte Financiero</h2>
        {df.to_html(classes="table table-striped table-bordered", index=False)}
    </body>
    </html>
    """

    opciones = {                          # Opciones de formato del PDF
        "encoding": "UTF-8",
        "page-size": "A4",
        "margin-top": "10mm",
        "margin-bottom": "10mm",
        "margin-left": "10mm",
        "margin-right": "10mm"
    }

    path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"  # Ruta al ejecutable wkhtmltopdf
    config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)             # Configuración con la ruta

    with open("reporte.html", "w", encoding="utf-8") as f:   # Guardar HTML temporal
        f.write(reporte_html)

    pdfkit.from_file("reporte.html", "reporte.pdf", options=opciones, configuration=config)  # Generar PDF

    return reporte_html   # Devolver HTML para usar en correos


# -------------------------------
# Función para enviar correo con adjuntos
# -------------------------------
def enviar_email(asunto, cuerpo_html, tickers, panel):
    try:
        remitente = os.getenv("EMAIL_USER")    # Correo del remitente
        password = os.getenv("EMAIL_PASS")     # Contraseña o App Password
        destinatario = os.getenv("EMAIL_TO")   # Correo del destinatario

        msg = MIMEMultipart("mixed")           # Crear mensaje multipart
        msg["Subject"] = asunto                # Asunto
        msg["From"] = remitente                # Remitente
        msg["To"] = destinatario               # Destinatario

        cuerpo = MIMEText(cuerpo_html, "html") # Cuerpo del correo en formato HTML
        msg.attach(cuerpo)

        # Adjuntar PDF principal
        with open("reporte.pdf", "rb") as f:
            adjunto = MIMEApplication(f.read(), _subtype="pdf")
            adjunto.add_header("Content-Disposition", "attachment", filename="reporte.pdf")
            msg.attach(adjunto)

        # Adjuntar gráficos individuales si existen
        for ticker in tickers:
            archivo = f"grafico_{ticker}.pdf"
            if os.path.exists(archivo):
                with open(archivo, "rb") as g:
                    adj_g = MIMEApplication(g.read(), _subtype="pdf")
                    adj_g.add_header("Content-Disposition", "attachment", filename=archivo)
                    msg.attach(adj_g)

        # Conexión al servidor SMTP de Gmail
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()                  # Iniciar cifrado TLS
            server.login(remitente, password)  # Autenticación
            server.sendmail(remitente, destinatario, msg.as_string())  # Enviar mensaje

        ttk.Label(panel, text="✅ Correo enviado correctamente", bootstyle="success").pack(pady=10)

    except Exception as e:
        ttk.Label(panel, text=f"❌ Error: {type(e).__name__} - {str(e)}", bootstyle="danger").pack(pady=10)


# -------------------------------
# Clase principal del Panel Financiero
# -------------------------------
class PanelFinanciero:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Panel Financiero")     # Título de la ventana
        self.root.geometry("1000x600")             # Tamaño inicial

        self.tickers = ["AAPL", "GOOGL", "MSFT"]  # Lista de símbolos
        self.df = obtener_datos()                 # Obtener datos iniciales
        self.reporte_html = generar_pdf(self.df)  # Generar PDF con diseño mejorado

        # Crear paneles de menú y contenido dinámico
        self.menu = ttk.Frame(root, padding=10)
        self.menu.pack(side="left", fill="y")

        self.panel_dinamico = ttk.Frame(root, padding=20)
        self.panel_dinamico.pack(side="right", expand=True, fill="both")

        # Botones del menú
        ttk.Label(self.menu, text="Menú", font=("Segoe UI", 14)).pack(pady=10)
        ttk.Button(self.menu, text="📊 Ver tabla", bootstyle=PRIMARY, command=self.mostrar_tabla).pack(pady=5, fill="x")
        for ticker in self.tickers:
            ttk.Button(self.menu, text=f"📈 Gráfico {ticker}", bootstyle=PRIMARY,
                       command=lambda t=ticker: self.mostrar_grafico(t)).pack(pady=5, fill="x")
        ttk.Button(self.menu, text="📤 Enviar correo", bootstyle=PRIMARY, command=self.formulario_correo).pack(pady=10, fill="x")
        ttk.Button(self.menu, text="🚪 Salir", bootstyle="danger", command=root.quit).pack(pady=20, fill="x")

    def limpiar_panel(self):
        for widget in self.panel_dinamico.winfo_children():  # Iterar sobre widgets hijos
            widget.destroy()                                 # Destruir cada widget

    def mostrar_tabla(self):
        self.limpiar_panel()
        ttk.Label(self.panel_dinamico, text="📊 Tabla de Datos Bursátiles",
                  font=("Segoe UI", 14, "bold")).pack(pady=10)

        columnas = list(self.df.columns)  # Obtener nombres de columnas

                # Crear un Treeview estilizado para mostrar datos tabulares
        tree = ttk.Treeview(self.panel_dinamico, columns=columnas, show="headings", bootstyle="info")

        # Configurar encabezados y columnas (centradas y con ancho fijo)
        for col in columnas:
            tree.heading(col, text=col)             # Texto del encabezado
            tree.column(col, anchor="center", width=180)  # Alineación y ancho de la columna

        # Insertar filas del DataFrame en el Treeview
        for _, fila in self.df.iterrows():
            tree.insert("", "end", values=list(fila))  # Insertar cada fila como lista de valores

        # Scrollbar vertical con estilo
        scrollbar = ttk.Scrollbar(self.panel_dinamico, orient="vertical", command=tree.yview, bootstyle="round-success")
        tree.configure(yscrollcommand=scrollbar.set)  # Vincular scrollbar al Treeview

        # Empaquetar widgets para que ocupen el espacio de forma elegante
        tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.pack(side="right", fill="y")

    def mostrar_grafico(self, ticker):
        self.limpiar_panel()  # Limpiar contenido previo del panel

        datos = yf.Ticker(ticker).history(period="1mo")  # Descargar histórico de 1 mes
        cierre = datos["Close"]                          # Serie de precios de cierre

        fig, ax = plt.subplots(figsize=(9, 5))           # Crear figura y eje con tamaño más amplio
        ax.plot(cierre.index, cierre.values,
                marker="o", linestyle="-", color="royalblue", linewidth=2)  # Gráfico de líneas estilizado

        ax.set_title(f"Precio de Cierre - {ticker}", fontsize=14, fontweight="bold")  # Título del gráfico
        ax.set_xlabel("Fecha", fontsize=12)   # Etiqueta eje X
        ax.set_ylabel("USD", fontsize=12)     # Etiqueta eje Y
        ax.grid(True, linestyle="--", alpha=0.6)  # Cuadrícula con estilo más suave

        # Formatear fechas en eje X para que no se corten
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d-%b"))  # Mostrar día y mes
        plt.setp(ax.get_xticklabels(), rotation=45, ha="right")      # Rotar etiquetas para legibilidad

        fig.tight_layout()                               # Ajustar márgenes automáticamente
        fig.savefig(f"grafico_{ticker}.pdf")             # Guardar gráfico como PDF para adjuntar en correo

        # Incrustar la figura en el panel usando TkAgg
        canvas = FigureCanvasTkAgg(fig, master=self.panel_dinamico)
        canvas.draw()                                    # Renderizar la figura
        canvas.get_tk_widget().pack(fill="both", expand=True)  # Empaquetar el widget en el panel

    def formulario_correo(self):
        self.limpiar_panel()  # Limpiar contenido previo

        # Título del formulario
        ttk.Label(self.panel_dinamico, text="📤 Enviar Reporte por Correo",
                  font=("Segoe UI", 16, "bold")).pack(pady=10)

        # Marco contenedor con padding
        marco = ttk.Frame(self.panel_dinamico, padding=10)
        marco.pack(pady=10, fill="both", expand=True)

        # Mostrar destinatario predeterminado (desde .env)
        destinatario = os.getenv("EMAIL_TO", "")
        ttk.Label(marco, text="Correo destinatario:").pack(anchor="w", pady=(10, 0))
        ttk.Label(marco, text=destinatario, font=("Segoe UI", 10, "bold"),
                  foreground="blue").pack(anchor="w", pady=5)

        # Campo de asunto
        ttk.Label(marco, text="Título del asunto:").pack(anchor="w", pady=(10, 0))
        entry_asunto = ttk.Entry(marco, width=50)
        entry_asunto.pack(pady=5)

        # Campo de comentario
        ttk.Label(marco, text="Comentario:").pack(anchor="w", pady=(10, 0))
        entry_comentario = ttk.Text(marco, width=50, height=6)
        entry_comentario.pack(pady=5)

        # Acción de envío: validación y llamada a enviar_email
        def ejecutar_envio():
            asunto = entry_asunto.get().strip()                    # Obtener asunto limpio
            comentario = entry_comentario.get("1.0", "end").strip()# Obtener comentario limpio

            if not asunto or not comentario:                       # Validar campos obligatorios
                ttk.Label(self.panel_dinamico, text="❌ Todos los campos son obligatorios",
                          bootstyle="danger").pack(pady=5)
                return

            # Construir cuerpo HTML uniendo comentario y la tabla HTML del reporte
            cuerpo_html = f"<h3>{asunto}</h3><p>{comentario}</p><hr>{self.reporte_html}"
            enviar_email(asunto, cuerpo_html, self.tickers, self.panel_dinamico)  # Enviar correo con adjuntos

        # Botón para ejecutar el envío
        ttk.Button(self.panel_dinamico, text="📨 Enviar ahora",
                   bootstyle="success-outline", command=ejecutar_envio).pack(pady=15)


# -------------------------------
# Bloque principal de ejecución
# -------------------------------
if __name__ == "__main__":
    app = ttk.Window(themename="flatly")  # Crear ventana raíz con tema "flatly" (estilo Bootstrap)
    PanelFinanciero(app)                  # Instanciar el panel financiero sobre la ventana
    app.mainloop()                        # Iniciar el loop de eventos de la GUI