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

# ── Descargar Excel ──────────────────────────────────────────────────────────
print("Descargando Excel desde Google Drive...")
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

creds = service_account.Credentials.from_service_account_info(
    GDRIVE_CREDENTIALS, scopes=["https://www.googleapis.com/auth/drive.readonly"]
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

# Sección 1: Enviados con menos de 5 días (en término, salidos desde DVT)
enviados_recientes = df[
    (df["Origen"] == DVT) &
    (df["Estado"].str.lower().str.strip() == "enviado") &
    (df["dias"] <= LIMITE)
].sort_values("Fecha y hora Pase", ascending=False).copy()

# Sección 2: Trámites que están en DVT como destino, con máximo 30 días ahí
en_dvt = df[
    (df["Destino"] == DVT) &
    (df["dias"] <= 30)
].sort_values("dias", ascending=True).copy()

print("Enviados en termino: " + str(len(enviados_recientes)) + " | En DVT (hasta 30 dias): " + str(len(en_dvt)))

# ── Helpers ──────────────────────────────────────────────────────────────────
def fmt_fecha(val):
    if pd.notna(val):
        return val.strftime("%d/%m/%Y %H:%M")
    return "-"

def card_enviado(row):
    d      = int(row["dias"])
    fecha  = fmt_fecha(row["Fecha y hora Pase"])
    tipo   = str(row.get("Tipo", ""))
    titulo = str(row.get("Titulo", ""))
    exp    = str(row["Expediente"])
    dest   = str(row["Destino"])

    html  = '<div style="background:#fff;border:0.5px solid #ddd;border-left:4px solid #27AE60;border-radius:0 10px 10px 0;padding:14px 16px;margin-bottom:10px;">'
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">'
    html += '<span style="font-size:12px;color:#888;font-weight:600;">' + exp + '</span>'
    html += '<span style="background:#D5F5E3;color:#1E8449;font-size:11px;padding:2px 8px;border-radius:20px;">' + str(d) + ' dias</span>'
    html += '</div>'
    html += '<div style="font-size:14px;color:#1a1a1a;font-weight:500;margin-bottom:8px;line-height:1.4;">' + titulo + '</div>'
    html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;">'
    html += '<span style="background:#D6EAF8;color:#1A5276;font-size:11px;padding:2px 8px;border-radius:20px;">' + tipo + '</span>'
    html += '<span style="background:#F4F6F7;color:#555;font-size:11px;padding:2px 8px;border-radius:20px;">Pase: ' + fecha + '</span>'
    html += '</div>'
    html += '<div style="background:#EAF4FB;border-radius:8px;padding:8px 12px;display:flex;align-items:center;gap:8px;">'
    html += '<span style="font-size:11px;color:#888;">Enviado a</span>'
    html += '<span style="font-size:13px;color:#1A5276;font-weight:600;">' + dest + '</span>'
    html += '</div>'
    html += '</div>'
    return html

def card_en_dvt(row):
    d      = int(row["dias"])
    fecha  = fmt_fecha(row["Fecha y hora Pase"])
    tipo   = str(row.get("Tipo", ""))
    titulo = str(row.get("Titulo", ""))
    exp    = str(row["Expediente"])
    origen = str(row["Origen"])
    estado = str(row.get("Estado", "")) if not pd.isna(row.get("Estado")) else "sin estado"

    if d <= 7:
        color_dias = "#1E8449"
        bg_dias    = "#D5F5E3"
    elif d <= 15:
        color_dias = "#E67E22"
        bg_dias    = "#FDEBD0"
    else:
        color_dias = "#C0392B"
        bg_dias    = "#FADBD8"

    html  = '<div style="background:#fff;border:0.5px solid #ddd;border-left:4px solid #1A5276;border-radius:0 10px 10px 0;padding:14px 16px;margin-bottom:10px;">'
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px;">'
    html += '<span style="font-size:12px;color:#888;font-weight:600;">' + exp + '</span>'
    html += '<span style="background:' + bg_dias + ';color:' + color_dias + ';font-size:11px;padding:2px 8px;border-radius:20px;">' + str(d) + ' dias en DVT</span>'
    html += '</div>'
    html += '<div style="font-size:14px;color:#1a1a1a;font-weight:500;margin-bottom:8px;line-height:1.4;">' + titulo + '</div>'
    html += '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:10px;">'
    html += '<span style="background:#D6EAF8;color:#1A5276;font-size:11px;padding:2px 8px;border-radius:20px;">' + tipo + '</span>'
    html += '<span style="background:#F4F6F7;color:#555;font-size:11px;padding:2px 8px;border-radius:20px;">Pase: ' + fecha + '</span>'
    html += '<span style="background:#F4F6F7;color:#555;font-size:11px;padding:2px 8px;border-radius:20px;">' + estado + '</span>'
    html += '</div>'
    html += '<div style="background:#EBF5FB;border-radius:8px;padding:8px 12px;display:flex;align-items:center;gap:8px;">'
    html += '<span style="font-size:11px;color:#888;">Vino desde</span>'
    html += '<span style="font-size:13px;color:#1A5276;font-weight:600;">' + origen + '</span>'
    html += '</div>'
    html += '</div>'
    return html

cards_enviados = "".join(card_enviado(r) for _, r in enviados_recientes.iterrows())
cards_en_dvt   = "".join(card_en_dvt(r)  for _, r in en_dvt.iterrows())

sin_env  = '<p style="font-size:13px;color:#aaa;padding:8px 0;">No hay tramites enviados en termino esta semana.</p>'
sin_dvt  = '<p style="font-size:13px;color:#aaa;padding:8px 0;">No hay tramites nuevos en DVT este mes.</p>'

# ── HTML ─────────────────────────────────────────────────────────────────────
html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DVT - Reporte de Tramites</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F4F6F7;color:#1a1a1a}
.header{background:#1A5276;color:white;padding:18px 24px}
.header h1{font-size:17px;font-weight:600}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;padding:16px;max-width:720px;margin:0 auto}
@media(min-width:500px){.stats{grid-template-columns:repeat(2,1fr)}}
.stat{background:#fff;border-radius:10px;border:0.5px solid #e0e0e0;padding:12px;text-align:center}
.stat-num{font-size:24px;font-weight:600}
.stat-label{font-size:11px;color:#888;margin-top:3px}
.azul{color:#1A5276}.ok{color:#1E8449}
.section{max-width:720px;margin:0 auto 24px;padding:0 16px}
.section-title{font-size:14px;font-weight:600;margin:16px 0 10px;padding-bottom:6px;border-bottom:0.5px solid #ddd;display:flex;align-items:center;gap:8px}
.badge-azul{font-size:11px;padding:2px 8px;border-radius:20px;background:#D6EAF8;color:#1A5276}
.badge-ok{font-size:11px;padding:2px 8px;border-radius:20px;background:#D5F5E3;color:#1E8449}
.footer{text-align:center;font-size:11px;color:#aaa;padding:20px;border-top:0.5px solid #e0e0e0;margin-top:8px}
</style>
</head>
<body>
<div class="header">
  <h1>Direccion de Vinculacion Tecnologica - Reporte de Tramites</h1>
  <p>""" + HOY.strftime("%d/%m/%Y %H:%M") + """</p>
</div>

<div class="stats" style="padding-top:16px;">
  <div class="stat">
    <div class="stat-num ok">""" + str(len(enviados_recientes)) + """</div>
    <div class="stat-label">Enviados en termino (hasta """ + str(LIMITE) + """ dias)</div>
  </div>
  <div class="stat">
    <div class="stat-num azul">""" + str(len(en_dvt)) + """</div>
    <div class="stat-label">Ingresados a DVT (ultimo mes)</div>
  </div>
</div>

<div class="section">
  <div class="section-title">
    Enviados en termino
    <span class="badge-ok">""" + str(len(enviados_recientes)) + """ tramites</span>
  </div>
  """ + (cards_enviados if cards_enviados else sin_env) + """
</div>

<div class="section">
  <div class="section-title">
    Ingresados a DVT en el ultimo mes
    <span class="badge-azul">""" + str(len(en_dvt)) + """ tramites</span>
  </div>
  """ + (cards_en_dvt if cards_en_dvt else sin_dvt) + """
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
    lineas_env = "\n".join(
        "- " + str(r["Expediente"]) + " -> " + str(r["Destino"])[:35] + " (" + fmt_fecha(r["Fecha y hora Pase"]) + ")"
        for _, r in enviados_recientes.head(4).iterrows()
    )
    lineas_dvt = "\n".join(
        "- " + str(r["Expediente"]) + " desde " + str(r["Origen"])[:30] + " (" + str(int(r["dias"])) + "d)"
        for _, r in en_dvt.head(4).iterrows()
    )
    msg  = "DVT Reporte Semanal " + HOY.strftime("%d/%m/%Y") + "\n\n"
    if lineas_env:
        msg += "ENVIADOS EN TERMINO (" + str(len(enviados_recientes)) + "):\n" + lineas_env + "\n\n"
    if lineas_dvt:
        msg += "INGRESADOS A DVT ESTE MES (" + str(len(en_dvt)) + "):\n" + lineas_dvt + "\n\n"
    msg += "Ver reporte: " + LINK
    params = urllib.parse.urlencode({"phone": CALLMEBOT_PHONE, "text": msg, "apikey": CALLMEBOT_APIKEY})
    req = urllib.request.Request("https://api.callmebot.com/whatsapp.php?" + params, headers={"User-Agent": "dvt"})
    with urllib.request.urlopen(req) as r:
        print("WhatsApp OK (" + str(r.status) + ")")

# ── Email ─────────────────────────────────────────────────────────────────────
def enviar_email():
    print("Enviando correo...")
    ths = "background:#1A5276;color:white;padding:9px 10px;text-align:left;font-size:12px;"
    tds = "padding:8px 10px;border-bottom:1px solid #eee;font-size:12px;vertical-align:top;"

    def hacer_tabla_env(rows):
        cols = ["Expediente", "Titulo", "Fecha y hora Pase", "Enviado a", "Dias"]
        ths_html = "".join('<th style="' + ths + '">' + c + "</th>" for c in cols)
        trs = ""
        for _, r in rows.iterrows():
            trs += '<tr style="background:#F0FAF4;">'
            trs += '<td style="' + tds + 'font-weight:600;">' + str(r["Expediente"]) + "</td>"
            trs += '<td style="' + tds + '">' + str(r.get("Titulo",""))[:55] + "</td>"
            trs += '<td style="' + tds + '">' + fmt_fecha(r["Fecha y hora Pase"]) + "</td>"
            trs += '<td style="' + tds + 'color:#1A5276;font-weight:600;">' + str(r["Destino"]) + "</td>"
            trs += '<td style="' + tds + 'color:#1E8449;font-weight:700;">' + str(int(r["dias"])) + "d</td>"
            trs += "</tr>"
        return '<table style="width:100%;border-collapse:collapse;margin-bottom:8px;"><tr>' + ths_html + "</tr>" + trs + "</table>"

    def hacer_tabla_dvt(rows):
        cols = ["Expediente", "Titulo", "Fecha y hora Pase", "Vino desde", "Dias en DVT"]
        ths_html = "".join('<th style="' + ths + '">' + c + "</th>" for c in cols)
        trs = ""
        for _, r in rows.iterrows():
            d = int(r["dias"])
            bg = "#D5F5E3" if d <= 7 else ("#FEF9E7" if d <= 15 else "#FADBD8")
            c  = "#1E8449" if d <= 7 else ("#E67E22" if d <= 15 else "#C0392B")
            trs += '<tr style="background:#fff;">'
            trs += '<td style="' + tds + 'font-weight:600;">' + str(r["Expediente"]) + "</td>"
            trs += '<td style="' + tds + '">' + str(r.get("Titulo",""))[:55] + "</td>"
            trs += '<td style="' + tds + '">' + fmt_fecha(r["Fecha y hora Pase"]) + "</td>"
            trs += '<td style="' + tds + 'color:#1A5276;font-weight:600;">' + str(r["Origen"]) + "</td>"
            trs += '<td style="' + tds + 'font-weight:700;"><span style="background:' + bg + ';color:' + c + ';padding:2px 7px;border-radius:20px;">' + str(d) + "d</span></td>"
            trs += "</tr>"
        return '<table style="width:100%;border-collapse:collapse;margin-bottom:8px;"><tr>' + ths_html + "</tr>" + trs + "</table>"

    bloque_env = ""
    if len(enviados_recientes) > 0:
        bloque_env  = '<h3 style="font-size:13px;margin:0 0 10px;color:#1E8449;">Enviados en termino (' + str(len(enviados_recientes)) + ')</h3>'
        bloque_env += hacer_tabla_env(enviados_recientes)

    bloque_dvt = ""
    if len(en_dvt) > 0:
        bloque_dvt  = '<h3 style="font-size:13px;margin:20px 0 10px;color:#1A5276;">Ingresados a DVT en el ultimo mes (' + str(len(en_dvt)) + ')</h3>'
        bloque_dvt += hacer_tabla_dvt(en_dvt)

    cuerpo  = '<div style="font-family:Arial,sans-serif;max-width:760px;margin:0 auto;">'
    cuerpo += '<div style="background:#1A5276;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">'
    cuerpo += '<h2 style="margin:0;font-size:16px;">DVT - Reporte Semanal de Tramites</h2>'
    cuerpo += '<p style="margin:5px 0 0;font-size:12px;opacity:.8;">' + HOY.strftime("%d/%m/%Y") + '</p>'
    cuerpo += "</div>"
    cuerpo += '<div style="padding:16px 20px;background:#fff;border:1px solid #ddd;border-top:none;">'
    cuerpo += bloque_env + bloque_dvt
    cuerpo += '<div style="margin-top:20px;text-align:center;">'
    cuerpo += '<a href="' + LINK + '" style="background:#1A5276;color:white;padding:10px 28px;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600;">Ver reporte completo</a>'
    cuerpo += "</div></div>"
    cuerpo += '<div style="background:#f8f8f8;padding:10px 20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:11px;color:#aaa;text-align:center;">'
    cuerpo += "FBIOyF - UNR · Reporte automatico semanal · Lunes 10:00 AM</div></div>"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "DVT Reporte Semanal - " + HOY.strftime("%d/%m/%Y")
    msg["From"]    = GMAIL_REMITENTE
    msg["To"]      = ", ".join(GMAIL_DESTINATARIOS)
    msg.attach(MIMEText(cuerpo, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(GMAIL_REMITENTE, GMAIL_PASSWORD)
        srv.sendmail(GMAIL_REMITENTE, GMAIL_DESTINATARIOS, msg.as_string())
    print("Correo enviado a: " + ", ".join(GMAIL_DESTINATARIOS))

# ── Main ─────────────────────────────────────────────────────────────────────
try:
    enviar_whatsapp()
    enviar_email()
    print("\nListo. Reporte enviado.")
except Exception as e:
    print("\nError: " + str(e))
    sys.exit(1)
