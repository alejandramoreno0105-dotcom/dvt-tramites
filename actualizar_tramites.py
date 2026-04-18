import pandas as pd
import json
import smtplib
import urllib.request
import urllib.parse
import io
import os
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

GMAIL_REMITENTE     = os.environ["GMAIL_REMITENTE"]
GMAIL_PASSWORD      = os.environ["GMAIL_PASSWORD"]
GMAIL_DESTINATARIOS = os.environ["GMAIL_DESTINATARIOS"].split(",")
CALLMEBOT_APIKEY    = os.environ["CALLMEBOT_APIKEY"]
CALLMEBOT_PHONE     = os.environ["CALLMEBOT_PHONE"]
GDRIVE_FILE_ID      = os.environ["GDRIVE_FILE_ID"]
GDRIVE_CREDENTIALS  = json.loads(os.environ["GDRIVE_CREDENTIALS"])
GITHUB_REPO         = os.environ.get("GITHUB_REPOSITORY", "usuario/dvt-tramites")
GITHUB_USUARIO      = GITHUB_REPO.split("/")[0]
GITHUB_REPO_NOMBRE  = GITHUB_REPO.split("/")[1]

DVT    = "FBIOyF - Dirección de Vinculación Tecnológica"
LIMITE = 5
HOY    = datetime.today()
LINK   = "https://" + GITHUB_USUARIO + ".github.io/" + GITHUB_REPO_NOMBRE + "/"

# ── Descargar Excel desde Google Drive ──────────────────────────────────────
print("Descargando Excel desde Google Drive...")
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

creds = service_account.Credentials.from_service_account_info(
    GDRIVE_CREDENTIALS,
    scopes=["https://www.googleapis.com/auth/drive.readonly"]
)
service = build("drive", "v3", credentials=creds, cache_discovery=False)
meta = service.files().get(fileId=GDRIVE_FILE_ID, fields="mimeType,name").execute()
mime = meta.get("mimeType", "")
print("Archivo: " + meta.get("name", "") + " (" + mime + ")")

fh = io.BytesIO()
if mime == "application/vnd.google-apps.spreadsheet":
    request = service.files().export_media(
        fileId=GDRIVE_FILE_ID,
        mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    request = service.files().get_media(fileId=GDRIVE_FILE_ID)

downloader = MediaIoBaseDownload(fh, request)
done = False
while not done:
    _, done = downloader.next_chunk()
fh.seek(0)
print("Excel descargado.")

# ── Procesar datos ───────────────────────────────────────────────────────────
df = pd.read_excel(fh)
df = df.drop_duplicates()
df["Fecha y hora Pase"] = pd.to_datetime(df["Fecha y hora Pase"], errors="coerce")
if "Título" in df.columns:
    df = df.rename(columns={"Título": "Titulo"})

df["dias"] = (HOY - df["Fecha y hora Pase"]).dt.days

# Trámites salidos desde DVT con estado "enviado"
salidas_dvt = df[df["Origen"] == DVT].copy()
enviados    = salidas_dvt[salidas_dvt["Estado"].str.lower().str.strip() == "enviado"].copy()
enviados    = enviados.sort_values("dias", ascending=False)

# Separar criticos (más de LIMITE días) y recientes
criticos  = enviados[enviados["dias"] > LIMITE]
recientes = enviados[enviados["dias"] <= LIMITE]

print("Enviados criticos: " + str(len(criticos)) + " | Enviados recientes: " + str(len(recientes)))

# ── Helpers HTML ─────────────────────────────────────────────────────────────
def dias_color(d):
    if d > 30:
        return "#7B241C"
    elif d > 14:
        return "#C0392B"
    elif d > LIMITE:
        return "#E67E22"
    else:
        return "#1E8449"

def card_enviado(row, es_critico):
    d     = int(row["dias"])
    c     = dias_color(d)
    borde = "#E74C3C" if es_critico else "#27AE60"
    if pd.notna(row["Fecha y hora Pase"]):
        fecha = row["Fecha y hora Pase"].strftime("%d/%m/%Y %H:%M")
    else:
        fecha = "-"
    tipo    = str(row.get("Tipo", ""))
    titulo  = str(row.get("Titulo", ""))
    exp     = str(row["Expediente"])
    destino = str(row["Destino"])

    alerta_badge = ""
    if es_critico:
        alerta_badge = '<span style="background:#FADBD8;color:#922B21;font-size:11px;padding:2px 8px;border-radius:20px;font-weight:600;">&#9888; ' + str(d) + ' dias sin confirmar</span>'
    else:
        alerta_badge = '<span style="background:#D5F5E3;color:#1E8449;font-size:11px;padding:2px 8px;border-radius:20px;">' + str(d) + ' dias</span>'

    html  = '<div style="background:#fff;border:0.5px solid #ddd;border-left:4px solid ' + borde + ';'
    html += 'border-radius:0 10px 10px 0;padding:14px 16px;margin-bottom:10px;">'

    # Cabecera: expediente + badge dias
    html += '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:6px;">'
    html += '<div style="font-size:12px;color:#888;font-weight:600;">' + exp + '</div>'
    html += alerta_badge
    html += '</div>'

    # Título
    html += '<div style="font-size:14px;color:#1a1a1a;font-weight:500;margin-bottom:8px;line-height:1.4;">' + titulo + '</div>'

    # Tipo + fecha
    html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;">'
    html += '<span style="background:#D6EAF8;color:#1A5276;font-size:11px;padding:2px 8px;border-radius:20px;">' + tipo + '</span>'
    html += '<span style="font-size:11px;color:#aaa;">Pase: ' + fecha + '</span>'
    html += '</div>'

    # Destino destacado
    html += '<div style="background:#F4F6F7;border-radius:8px;padding:8px 12px;display:flex;align-items:center;gap:8px;">'
    html += '<span style="font-size:11px;color:#888;">Actualmente en</span>'
    html += '<span style="font-size:13px;color:#1A5276;font-weight:600;">' + destino + '</span>'
    html += '</div>'

    html += '</div>'
    return html

cards_criticos  = "".join(card_enviado(r, True)  for _, r in criticos.iterrows())
cards_recientes = "".join(card_enviado(r, False) for _, r in recientes.iterrows())

# ── Generar HTML ─────────────────────────────────────────────────────────────
sin_criticos  = '<p style="font-size:13px;color:#aaa;padding:8px 0;">No hay tramites criticos esta semana.</p>'
sin_recientes = '<p style="font-size:13px;color:#aaa;padding:8px 0;">No hay tramites enviados recientemente.</p>'

html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DVT - Tramites Enviados</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F4F6F7;color:#1a1a1a}
.header{background:#1A5276;color:white;padding:18px 24px}
.header h1{font-size:17px;font-weight:600}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;padding:16px;max-width:720px;margin:0 auto}
.stat{background:#fff;border-radius:10px;border:0.5px solid #e0e0e0;padding:12px;text-align:center}
.stat-num{font-size:24px;font-weight:600}
.stat-label{font-size:11px;color:#888;margin-top:3px}
.danger{color:#C0392B}.warn{color:#E67E22}.ok{color:#1E8449}
.section{max-width:720px;margin:0 auto 24px;padding:0 16px}
.section-title{font-size:14px;font-weight:600;margin:16px 0 10px;padding-bottom:6px;border-bottom:0.5px solid #ddd;display:flex;align-items:center;gap:8px}
.badge-alert{font-size:11px;padding:2px 8px;border-radius:20px;background:#FADBD8;color:#922B21}
.badge-ok{font-size:11px;padding:2px 8px;border-radius:20px;background:#D5F5E3;color:#1E8449}
.leyenda{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:14px;font-size:11px;color:#666;align-items:center}
.dot{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:4px}
.footer{text-align:center;font-size:11px;color:#aaa;padding:20px;border-top:0.5px solid #e0e0e0;margin-top:8px}
</style>
</head>
<body>
<div class="header">
  <h1>Direccion de Vinculacion Tecnologica</h1>
  <p>Tramites con estado ENVIADO · Reporte del """ + HOY.strftime("%d/%m/%Y %H:%M") + """</p>
</div>

<div class="stats" style="padding-top:16px;">
  <div class="stat">
    <div class="stat-num danger">""" + str(len(criticos)) + """</div>
    <div class="stat-label">Criticos (mas de """ + str(LIMITE) + """ dias)</div>
  </div>
  <div class="stat">
    <div class="stat-num ok">""" + str(len(recientes)) + """</div>
    <div class="stat-label">En termino</div>
  </div>
  <div class="stat">
    <div class="stat-num warn">""" + str(len(enviados)) + """</div>
    <div class="stat-label">Total enviados</div>
  </div>
</div>

<div class="section">
  <div class="leyenda">
    <span><span class="dot" style="background:#E74C3C"></span>Critico: mas de """ + str(LIMITE) + """ dias sin confirmar</span>
    <span><span class="dot" style="background:#27AE60"></span>En termino</span>
  </div>

  <div class="section-title">
    Tramites criticos
    <span class="badge-alert">""" + str(len(criticos)) + """ sin confirmar</span>
  </div>
  """ + (cards_criticos if cards_criticos else sin_criticos) + """

  <div class="section-title" style="margin-top:24px;">
    Enviados recientemente
    <span class="badge-ok">""" + str(len(recientes)) + """ en termino</span>
  </div>
  """ + (cards_recientes if cards_recientes else sin_recientes) + """
</div>

<div class="footer">FBIOyF - UNR · Reporte automatico semanal · Lunes 10:00 AM</div>
</body>
</html>"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)
print("index.html generado.")

# ── WhatsApp ─────────────────────────────────────────────────────────────────
def enviar_whatsapp():
    print("Enviando WhatsApp...")
    lineas = "\n".join(
        "- " + str(r["Expediente"]) + " (" + str(int(r["dias"])) + "d) en: " + str(r["Destino"])[:35]
        for _, r in criticos.head(5).iterrows()
    )
    msg  = "DVT - Tramites Criticos " + HOY.strftime("%d/%m/%Y") + "\n\n"
    msg += str(len(criticos)) + " tramite(s) ENVIADOS hace mas de " + str(LIMITE) + " dias sin confirmar:\n\n"
    msg += lineas + "\n\n"
    msg += "Ver reporte completo: " + LINK
    params = urllib.parse.urlencode({"phone": CALLMEBOT_PHONE, "text": msg, "apikey": CALLMEBOT_APIKEY})
    req = urllib.request.Request("https://api.callmebot.com/whatsapp.php?" + params, headers={"User-Agent": "dvt"})
    with urllib.request.urlopen(req) as r:
        print("WhatsApp OK (" + str(r.status) + ")")

# ── Email ─────────────────────────────────────────────────────────────────────
def enviar_email():
    print("Enviando correo...")
    ths = "background:#1A5276;color:white;padding:9px 10px;text-align:left;font-size:12px;"
    tds = "padding:8px 10px;border-bottom:1px solid #eee;font-size:12px;vertical-align:top;"

    def fila_bg(d):
        if d > 30:
            return "#FDEDEC"
        elif d > 14:
            return "#FEF9E7"
        elif d > LIMITE:
            return "#FEF9E7"
        else:
            return "#F0FAF4"

    def hacer_tabla(rows):
        cols = ["Expediente", "Titulo", "Destino (donde esta)", "Dias sin confirmar"]
        ths_html = "".join('<th style="' + ths + '">' + c + "</th>" for c in cols)
        trs = ""
        for _, r in rows.iterrows():
            d   = int(r["dias"])
            bg  = fila_bg(d)
            c   = dias_color(d)
            trs += '<tr style="background:' + bg + '">'
            trs += '<td style="' + tds + 'font-weight:600;">' + str(r["Expediente"]) + "</td>"
            trs += '<td style="' + tds + '">' + str(r.get("Titulo",""))[:65] + "</td>"
            trs += '<td style="' + tds + 'color:#1A5276;font-weight:600;">' + str(r["Destino"]) + "</td>"
            trs += '<td style="' + tds + 'font-weight:700;color:' + c + ';">' + str(d) + " dias</td>"
            trs += "</tr>"
        return '<table style="width:100%;border-collapse:collapse;margin-bottom:8px;"><tr>' + ths_html + "</tr>" + trs + "</table>"

    bloque_criticos = ""
    if len(criticos) > 0:
        bloque_criticos  = '<h3 style="font-size:13px;margin:0 0 10px;color:#C0392B;">&#9888; Criticos — mas de ' + str(LIMITE) + ' dias sin confirmar</h3>'
        bloque_criticos += hacer_tabla(criticos)

    bloque_recientes = ""
    if len(recientes) > 0:
        bloque_recientes  = '<h3 style="font-size:13px;margin:20px 0 10px;color:#1E8449;">En termino</h3>'
        bloque_recientes += hacer_tabla(recientes)

    cuerpo  = '<div style="font-family:Arial,sans-serif;max-width:740px;margin:0 auto;">'
    cuerpo += '<div style="background:#1A5276;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">'
    cuerpo += '<h2 style="margin:0;font-size:16px;">DVT - Tramites con estado ENVIADO</h2>'
    cuerpo += '<p style="margin:5px 0 0;font-size:12px;opacity:.8;">'
    cuerpo += HOY.strftime("%d/%m/%Y") + " · " + str(len(criticos)) + " critico(s) · " + str(len(recientes)) + " en termino</p>"
    cuerpo += "</div>"
    cuerpo += '<div style="padding:16px 20px;background:#fff;border:1px solid #ddd;border-top:none;">'
    cuerpo += bloque_criticos + bloque_recientes
    cuerpo += '<div style="margin-top:20px;text-align:center;">'
    cuerpo += '<a href="' + LINK + '" style="background:#1A5276;color:white;padding:10px 28px;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600;">Ver reporte completo</a>'
    cuerpo += "</div></div>"
    cuerpo += '<div style="background:#f8f8f8;padding:10px 20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:11px;color:#aaa;text-align:center;">'
    cuerpo += "FBIOyF - UNR · Reporte automatico semanal · Lunes 10:00 AM</div></div>"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "DVT Tramites Enviados - " + str(len(criticos)) + " critico(s) - " + HOY.strftime("%d/%m/%Y")
    msg["From"]    = GMAIL_REMITENTE
    msg["To"]      = ", ".join(GMAIL_DESTINATARIOS)
    msg.attach(MIMEText(cuerpo, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(GMAIL_REMITENTE, GMAIL_PASSWORD)
        srv.sendmail(GMAIL_REMITENTE, GMAIL_DESTINATARIOS, msg.as_string())
    print("Correo enviado a: " + ", ".join(GMAIL_DESTINATARIOS))

# ── Main ─────────────────────────────────────────────────────────────────────
try:
    if len(enviados) > 0:
        enviar_whatsapp()
        enviar_email()
        print("\nListo. Reporte enviado.")
    else:
        print("\nNo hay tramites con estado enviado.")
except Exception as e:
    print("\nError: " + str(e))
    sys.exit(1)
