import pandas as pd
import json
import smtplib
import urllib.request
import urllib.parse
import base64
import sys
import io
import os
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

df = pd.read_excel(fh)
df = df.drop_duplicates()
df["Fecha y hora Pase"] = pd.to_datetime(df["Fecha y hora Pase"], errors="coerce")

salidas  = df[df["Origen"] == DVT].copy()
entradas = df[df["Destino"] == DVT].copy()
salidas["dias"]  = (HOY - salidas["Fecha y hora Pase"]).dt.days
entradas["dias"] = (HOY - entradas["Fecha y hora Pase"]).dt.days

sal_alerta = salidas[salidas["dias"]   > LIMITE].sort_values("dias", ascending=False)
ent_alerta = entradas[entradas["dias"] > LIMITE].sort_values("dias", ascending=False)

print("Salidas criticas: " + str(len(sal_alerta)) + " | Entradas criticas: " + str(len(ent_alerta)))

def urgencia_color(d):
    if d > 30:
        return "#7B241C"
    elif d > 14:
        return "#C0392B"
    else:
        return "#E67E22"

def estado_badge(e):
    if pd.isna(e) or e is None:
        return '<span style="background:#f0f0f0;color:#888;font-size:11px;padding:2px 8px;border-radius:20px;">sin estado</span>'
    if str(e).lower() == "enviado":
        return '<span style="background:#FFF3CD;color:#856404;font-size:11px;padding:2px 8px;border-radius:20px;">enviado</span>'
    return '<span style="background:#D1ECE1;color:#155724;font-size:11px;padding:2px 8px;border-radius:20px;">confirmado</span>'

def card(row, flecha, lugar):
    d = int(row["dias"])
    c = urgencia_color(d)
    if pd.notna(row["Fecha y hora Pase"]):
        fecha = row["Fecha y hora Pase"].strftime("%d/%m/%Y %H:%M")
    else:
        fecha = "-"
    tipo = str(row.get("Tipo", ""))
    titulo = str(row["Titulo"])
    exp = str(row["Expediente"])
    badge = estado_badge(row.get("Estado"))
    html = '<div style="background:#fff;border:0.5px solid #ddd;border-left:4px solid ' + c + ';'
    html += 'border-radius:0 10px 10px 0;padding:12px 14px;margin-bottom:8px;'
    html += 'display:flex;justify-content:space-between;align-items:flex-start;">'
    html += '<div style="flex:1;min-width:0;">'
    html += '<div style="font-size:11px;color:#888;font-weight:600;margin-bottom:3px;">' + exp + '</div>'
    html += '<div style="font-size:13px;color:#1a1a1a;margin-bottom:6px;line-height:1.4;">' + titulo + '</div>'
    html += '<div style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:5px;">'
    html += '<span style="background:#D6EAF8;color:#1A5276;font-size:11px;padding:2px 8px;border-radius:20px;">' + tipo + '</span>'
    html += badge
    html += '<span style="font-size:11px;color:#aaa;">' + fecha + '</span>'
    html += '</div>'
    html += '<div style="font-size:12px;color:#555;">' + flecha + ' <strong style="color:#333;">' + lugar + '</strong></div>'
    html += '</div>'
    html += '<div style="text-align:center;min-width:52px;margin-left:12px;">'
    html += '<div style="font-size:22px;font-weight:600;color:' + c + ';line-height:1;">' + str(d) + '</div>'
    html += '<div style="font-size:10px;color:#aaa;">dias</div>'
    html += '</div></div>'
    return html

if "Título" in sal_alerta.columns:
    sal_alerta = sal_alerta.rename(columns={"Título": "Titulo"})
if "Título" in ent_alerta.columns:
    ent_alerta = ent_alerta.rename(columns={"Título": "Titulo"})
if "Título" in salidas.columns:
    salidas = salidas.rename(columns={"Título": "Titulo"})
if "Título" in entradas.columns:
    entradas = entradas.rename(columns={"Título": "Titulo"})

cards_s = "".join(card(r, "&#8594;", str(r["Destino"])) for _, r in sal_alerta.iterrows())
cards_e = "".join(card(r, "&#8592;", str(r["Origen"]))  for _, r in ent_alerta.iterrows())

en_termino_sal = len(salidas[salidas["dias"].between(1, LIMITE)])
en_termino_ent = len(entradas[entradas["dias"].between(1, LIMITE)])

no_sal = '<p style="font-size:13px;color:#aaa;padding:8px 0;">Sin salidas criticas.</p>'
no_ent = '<p style="font-size:13px;color:#aaa;padding:8px 0;">Sin entradas criticas.</p>'

html = """<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DVT - Tramites Criticos</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F4F6F7;color:#1a1a1a}
.header{background:#1A5276;color:white;padding:18px 24px}
.header h1{font-size:17px;font-weight:600}
.header p{font-size:12px;opacity:.7;margin-top:4px}
.stats{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;padding:16px;max-width:720px;margin:0 auto}
@media(min-width:500px){.stats{grid-template-columns:repeat(4,1fr)}}
.stat{background:#fff;border-radius:10px;border:0.5px solid #e0e0e0;padding:12px;text-align:center}
.stat-num{font-size:24px;font-weight:600}
.stat-label{font-size:11px;color:#888;margin-top:3px}
.danger{color:#C0392B}.warn{color:#E67E22}
.section{max-width:720px;margin:0 auto 24px;padding:0 16px}
.section-title{font-size:14px;font-weight:600;margin:16px 0 10px;padding-bottom:6px;border-bottom:0.5px solid #ddd;display:flex;align-items:center;gap:8px}
.badge{font-size:11px;padding:2px 8px;border-radius:20px;background:#FADBD8;color:#922B21}
.leyenda{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px}
.ley{display:flex;align-items:center;gap:5px;font-size:11px;color:#666}
.dot{width:10px;height:10px;border-radius:50%}
.footer{text-align:center;font-size:11px;color:#aaa;padding:20px;border-top:0.5px solid #e0e0e0;margin-top:8px}
</style>
</head>
<body>
<div class="header">
  <h1>Direccion de Vinculacion Tecnologica - Tramites Criticos</h1>
  <p>Reporte semanal · """ + HOY.strftime("%d/%m/%Y %H:%M") + """ · Mas de """ + str(LIMITE) + """ dias sin movimiento</p>
</div>
<div class="stats" style="padding-top:16px;">
  <div class="stat"><div class="stat-num danger">""" + str(len(sal_alerta)) + """</div><div class="stat-label">Salidas criticas</div></div>
  <div class="stat"><div class="stat-num danger">""" + str(len(ent_alerta)) + """</div><div class="stat-label">Entradas criticas</div></div>
  <div class="stat"><div class="stat-num warn">""" + str(en_termino_sal) + """</div><div class="stat-label">Salidas en termino</div></div>
  <div class="stat"><div class="stat-num warn">""" + str(en_termino_ent) + """</div><div class="stat-label">Entradas en termino</div></div>
</div>
<div class="section">
  <div class="leyenda">
    <div class="ley"><div class="dot" style="background:#7B241C"></div>Mas de 30 dias</div>
    <div class="ley"><div class="dot" style="background:#C0392B"></div>15-30 dias</div>
    <div class="ley"><div class="dot" style="background:#E67E22"></div>6-14 dias</div>
  </div>
  <div class="section-title">Salidas desde DVT <span class="badge">""" + str(len(sal_alerta)) + """ criticas</span></div>
  """ + (cards_s if cards_s else no_sal) + """
</div>
<div class="section">
  <div class="section-title">Ingresos a DVT <span class="badge">""" + str(len(ent_alerta)) + """ criticos</span></div>
  """ + (cards_e if cards_e else no_ent) + """
</div>
<div class="footer">FBIOyF - UNR · Reporte automatico · Lunes 10:00 AM</div>
</body>
</html>"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)
print("index.html generado.")

def enviar_whatsapp(total):
    print("Enviando WhatsApp...")
    ls = "\n".join("- " + str(r["Expediente"]) + " (" + str(int(r["dias"])) + "d) -> " + str(r["Destino"])[:40] for _, r in sal_alerta.head(4).iterrows())
    le = "\n".join("- " + str(r["Expediente"]) + " (" + str(int(r["dias"])) + "d) <- " + str(r["Origen"])[:40]  for _, r in ent_alerta.head(4).iterrows())
    msg = "DVT Reporte Semanal " + HOY.strftime("%d/%m/%Y") + "\n\n"
    msg += str(total) + " tramite(s) critico(s) (mas de " + str(LIMITE) + " dias)\n\n"
    if ls:
        msg += "SALIDAS:\n" + ls + "\n\n"
    if le:
        msg += "ENTRADAS:\n" + le + "\n\n"
    msg += "Dashboard: " + LINK
    params = urllib.parse.urlencode({"phone": CALLMEBOT_PHONE, "text": msg, "apikey": CALLMEBOT_APIKEY})
    req = urllib.request.Request("https://api.callmebot.com/whatsapp.php?" + params, headers={"User-Agent": "dvt"})
    with urllib.request.urlopen(req) as r:
        print("WhatsApp OK (" + str(r.status) + ")")

def enviar_email(total):
    print("Enviando correo...")
    ts  = "width:100%;border-collapse:collapse;font-size:13px;margin-bottom:8px;"
    ths = "background:#1A5276;color:white;padding:8px;text-align:left;"
    tds = "padding:7px 8px;border-bottom:1px solid #eee;"

    def fila_bg(d):
        if d > 30:
            return "#FDEDEC"
        elif d > 14:
            return "#FEF9E7"
        else:
            return "#fff"

    def hacer_fila(r, cols_data):
        bg = fila_bg(int(r["dias"]))
        celdas = "".join('<td style="' + tds + '">' + str(v) + "</td>" for v in cols_data(r))
        return '<tr style="background:' + bg + '">' + celdas + "</tr>"

    def hacer_tabla(rows, cols, fn):
        ths_html = "".join('<th style="' + ths + '">' + c + "</th>" for c in cols)
        trs_html = "".join(hacer_fila(r, fn) for _, r in rows.iterrows())
        return '<table style="' + ts + '"><tr>' + ths_html + "</tr>" + trs_html + "</table>"

    bloque_s = ""
    if len(sal_alerta) > 0:
        bloque_s = '<h3 style="font-size:13px;margin:20px 0 8px;color:#1A5276;">Salidas demoradas (' + str(len(sal_alerta)) + ')</h3>'
        bloque_s += hacer_tabla(sal_alerta, ["Expediente", "Titulo", "Destino", "Dias"],
            lambda r: [r["Expediente"], str(r["Titulo"])[:70], r["Destino"],
                       '<strong style="color:' + urgencia_color(int(r["dias"])) + '">' + str(int(r["dias"])) + " dias</strong>"])

    bloque_e = ""
    if len(ent_alerta) > 0:
        bloque_e = '<h3 style="font-size:13px;margin:20px 0 8px;color:#1A5276;">Entradas demoradas (' + str(len(ent_alerta)) + ')</h3>'
        bloque_e += hacer_tabla(ent_alerta, ["Expediente", "Titulo", "Origen", "Dias"],
            lambda r: [r["Expediente"], str(r["Titulo"])[:70], r["Origen"],
                       '<strong style="color:' + urgencia_color(int(r["dias"])) + '">' + str(int(r["dias"])) + " dias</strong>"])

    cuerpo  = '<div style="font-family:Arial,sans-serif;max-width:720px;margin:0 auto;">'
    cuerpo += '<div style="background:#1A5276;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">'
    cuerpo += '<h2 style="margin:0;font-size:16px;">DVT - Reporte Semanal de Tramites Criticos</h2>'
    cuerpo += '<p style="margin:4px 0 0;font-size:12px;opacity:.8;">Lunes ' + HOY.strftime("%d/%m/%Y") + " · " + str(total) + " tramite(s) con mas de " + str(LIMITE) + " dias sin movimiento</p>"
    cuerpo += "</div>"
    cuerpo += '<div style="padding:16px 20px;background:#fff;border:1px solid #ddd;border-top:none;">'
    cuerpo += bloque_s + bloque_e
    cuerpo += '<div style="margin-top:24px;text-align:center;">'
    cuerpo += '<a href="' + LINK + '" style="background:#1A5276;color:white;padding:10px 28px;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600;">Ver dashboard completo</a>'
    cuerpo += "</div></div>"
    cuerpo += '<div style="background:#f8f8f8;padding:10px 20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:11px;color:#aaa;text-align:center;">'
    cuerpo += "FBIOyF - UNR · Reporte automatico semanal · Lunes 10:00 AM</div></div>"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = "DVT Tramites Criticos - " + str(total) + " alerta(s) - " + HOY.strftime("%d/%m/%Y")
    msg["From"]    = GMAIL_REMITENTE
    msg["To"]      = ", ".join(GMAIL_DESTINATARIOS)
    msg.attach(MIMEText(cuerpo, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(GMAIL_REMITENTE, GMAIL_PASSWORD)
        srv.sendmail(GMAIL_REMITENTE, GMAIL_DESTINATARIOS, msg.as_string())
    print("Correo enviado a: " + ", ".join(GMAIL_DESTINATARIOS))

try:
    total = len(sal_alerta) + len(ent_alerta)
    if total > 0:
        enviar_whatsapp(total)
        enviar_email(total)
        print("\nListo. " + str(total) + " tramites criticos notificados.")
    else:
        print("\nDashboard actualizado. Sin tramites criticos.")
except Exception as e:
    print("\nError: " + str(e))
    sys.exit(1)
