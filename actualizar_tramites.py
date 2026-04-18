import pandas as pd
import json
import smtplib
import urllib.request
import urllib.parse
import base64
import sys
import os
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime

# ── Credenciales ─────────────────────────────────────────────────────────────
GMAIL_REMITENTE      = os.environ["GMAIL_REMITENTE"]
GMAIL_PASSWORD       = os.environ["GMAIL_PASSWORD"]
GMAIL_DESTINATARIOS  = os.environ["GMAIL_DESTINATARIOS"].split(",")
CALLMEBOT_APIKEY     = os.environ["CALLMEBOT_APIKEY"]
CALLMEBOT_PHONE      = os.environ["CALLMEBOT_PHONE"]
GDRIVE_FILE_ID       = os.environ["GDRIVE_FILE_ID"]
GDRIVE_CREDENTIALS   = json.loads(os.environ["GDRIVE_CREDENTIALS"])
GITHUB_REPO          = os.environ.get("GITHUB_REPOSITORY", "usuario/dvt-tramites")
GITHUB_USUARIO       = GITHUB_REPO.split("/")[0]
GITHUB_REPO_NOMBRE   = GITHUB_REPO.split("/")[1]

DVT    = "FBIOyF - Dirección de Vinculación Tecnológica"
LIMITE = 5
HOY    = datetime.today()
LINK   = f"https://{GITHUB_USUARIO}.github.io/{GITHUB_REPO_NOMBRE}/"

# ── Bajar Excel desde Google Drive ───────────────────────────────────────────
print("Descargando Excel desde Google Drive...")
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

creds = service_account.Credentials.from_service_account_info(
    GDRIVE_CREDENTIALS,
    scopes=["https://www.googleapis.com/auth/drive.readonly"]
)
service = build("drive", "v3", credentials=creds, cache_discovery=False)

# Detectar si es un Google Sheets o un Excel (.xlsx)
meta = service.files().get(fileId=GDRIVE_FILE_ID, fields="mimeType,name").execute()
mime = meta.get("mimeType", "")
print(f"Archivo encontrado: {meta.get('name')} ({mime})")

fh = io.BytesIO()
if mime == "application/vnd.google-apps.spreadsheet":
    # Es Google Sheets → exportar como xlsx
    request = service.files().export_media(
        fileId=GDRIVE_FILE_ID,
        mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    # Es un .xlsx subido directamente
    request = service.files().get_media(fileId=GDRIVE_FILE_ID)

downloader = MediaIoBaseDownload(fh, request)
done = False
while not done:
    _, done = downloader.next_chunk()

fh.seek(0)
print("Excel descargado correctamente.")

# ── Leer y limpiar datos ─────────────────────────────────────────────────────
df = pd.read_excel(fh)
df = df.drop_duplicates()
df["Fecha y hora Pase"] = pd.to_datetime(df["Fecha y hora Pase"], errors="coerce")

salidas  = df[df["Origen"] == DVT].copy()
entradas = df[df["Destino"] == DVT].copy()
salidas["dias"]  = (HOY - salidas["Fecha y hora Pase"]).dt.days
entradas["dias"] = (HOY - entradas["Fecha y hora Pase"]).dt.days

sal_alerta = salidas[salidas["dias"]   > LIMITE].sort_values("dias", ascending=False)
ent_alerta = entradas[entradas["dias"] > LIMITE].sort_values("dias", ascending=False)

print(f"Salidas críticas: {len(sal_alerta)} | Entradas críticas: {len(ent_alerta)}")

# ── HTML ─────────────────────────────────────────────────────────────────────
def urgencia_color(d):
    return "#7B241C" if d > 30 else ("#C0392B" if d > 14 else "#E67E22")

def estado_badge(e):
    if pd.isna(e) or e is None:
        return '<span style="background:#f0f0f0;color:#888;font-size:11px;padding:2px 8px;border-radius:20px;">sin estado</span>'
    if str(e).lower() == "enviado":
        return '<span style="background:#FFF3CD;color:#856404;font-size:11px;padding:2px 8px;border-radius:20px;">enviado</span>'
    return '<span style="background:#D1ECE1;color:#155724;font-size:11px;padding:2px 8px;border-radius:20px;">confirmado</span>'

def card(row, flecha, lugar):
    d = int(row["dias"])
    c = urgencia_color(d)
    fecha = row["Fecha y hora Pase"].strftime("%d/%m/%Y %H:%M") if pd.notna(row["Fecha y hora Pase"]) else "-"
    return f"""<div style="background:#fff;border:0.5px solid #ddd;border-left:4px solid {c};
border-radius:0 10px 10px 0;padding:12px 14px;margin-bottom:8px;
display:flex;justify-content:space-between;align-items:flex-start;">
<div style="flex:1;min-width:0;">
  <div style="font-size:11px;color:#888;font-weight:600;margin-bottom:3px;">{row['Expediente']}</div>
  <div style="font-size:13px;color:#1a1a1a;margin-bottom:6px;line-height:1.4;">{row['Título']}</div>
  <div style="display:flex;flex-wrap:wrap;gap:5px;margin-bottom:5px;">
    <span style="background:#D6EAF8;color:#1A5276;font-size:11px;padding:2px 8px;border-radius:20px;">{str(row.get('Tipo',''))}</span>
    {estado_badge(row.get('Estado'))}
    <span style="font-size:11px;color:#aaa;">{fecha}</span>
  </div>
  <div style="font-size:12px;color:#555;">{flecha} <strong style="color:#333;">{lugar}</strong></div>
</div>
<div style="text-align:center;min-width:52px;margin-left:12px;">
  <div style="font-size:22px;font-weight:600;color:{c};line-height:1;">{d}</div>
  <div style="font-size:10px;color:#aaa;">días</div>
</div>
</div>"""

cards_s = "".join(card(r, "→", r["Destino"]) for _, r in sal_alerta.iterrows())
cards_e = "".join(card(r, "←", r["Origen"])  for _, r in ent_alerta.iterrows())

html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DVT - Tramites Criticos</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#F4F6F7;color:#1a1a1a}}
.header{{background:#1A5276;color:white;padding:18px 24px}}
.header h1{{font-size:17px;font-weight:600}}
.header p{{font-size:12px;opacity:.7;margin-top:4px}}
.stats{{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;padding:16px;max-width:720px;margin:0 auto}}
@media(min-width:500px){{.stats{{grid-template-columns:repeat(4,1fr)}}}}
.stat{{background:#fff;border-radius:10px;border:0.5px solid #e0e0e0;padding:12px;text-align:center}}
.stat-num{{font-size:24px;font-weight:600}}
.stat-label{{font-size:11px;color:#888;margin-top:3px}}
.danger{{color:#C0392B}}.warn{{color:#E67E22}}
.section{{max-width:720px;margin:0 auto 24px;padding:0 16px}}
.section-title{{font-size:14px;font-weight:600;margin:16px 0 10px;padding-bottom:6px;border-bottom:0.5px solid #ddd;display:flex;align-items:center;gap:8px}}
.badge{{font-size:11px;padding:2px 8px;border-radius:20px;background:#FADBD8;color:#922B21}}
.leyenda{{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px}}
.ley{{display:flex;align-items:center;gap:5px;font-size:11px;color:#666}}
.dot{{width:10px;height:10px;border-radius:50%}}
.footer{{text-align:center;font-size:11px;color:#aaa;padding:20px;border-top:0.5px solid #e0e0e0;margin-top:8px}}
</style>
</head>
<body>
<div class="header">
  <h1>Dirección de Vinculación Tecnológica — Trámites Críticos</h1>
  <p>Reporte semanal · {HOY.strftime("%d/%m/%Y %H:%M")} · Más de {LIMITE} días sin movimiento</p>
</div>
<div class="stats" style="padding-top:16px;">
  <div class="stat"><div class="stat-num danger">{len(sal_alerta)}</div><div class="stat-label">Salidas críticas</div></div>
  <div class="stat"><div class="stat-num danger">{len(ent_alerta)}</div><div class="stat-label">Entradas críticas</div></div>
  <div class="stat"><div class="stat-num warn">{len(salidas[salidas["dias"].between(1,LIMITE)])}</div><div class="stat-label">Salidas en término</div></div>
  <div class="stat"><div class="stat-num warn">{len(entradas[entradas["dias"].between(1,LIMITE)])}</div><div class="stat-label">Entradas en término</div></div>
</div>
<div class="section">
  <div class="leyenda">
    <div class="ley"><div class="dot" style="background:#7B241C"></div>Más de 30 días</div>
    <div class="ley"><div class="dot" style="background:#C0392B"></div>15–30 días</div>
    <div class="ley"><div class="dot" style="background:#E67E22"></div>6–14 días</div>
  </div>
  <div class="section-title">Salidas desde DVT <span class="badge">{len(sal_alerta)} críticas</span></div>
  {cards_s or '<p style="font-size:13px;color:#aaa;padding:8px 0;">Sin salidas críticas.</p>'}
</div>
<div class="section">
  <div class="section-title">Ingresos a DVT <span class="badge">{len(ent_alerta)} críticos</span></div>
  {cards_e or '<p style="font-size:13px;color:#aaa;padding:8px 0;">Sin entradas críticas.</p>'}
</div>
<div class="footer">FBIOyF – UNR · Reporte automático · Lunes 10:00 AM</div>
</body>
</html>"""

with open("index.html", "w", encoding="utf-8") as f:
    f.write(html)
print("index.html generado.")

# ── WhatsApp ─────────────────────────────────────────────────────────────────
def enviar_whatsapp(total):
    print("Enviando WhatsApp...")
    ls = "\n".join(f"- {r['Expediente']} ({int(r['dias'])}d) -> {str(r['Destino'])[:40]}" for _, r in sal_alerta.head(4).iterrows())
    le = "\n".join(f"- {r['Expediente']} ({int(r['dias'])}d) <- {str(r['Origen'])[:40]}"  for _, r in ent_alerta.head(4).iterrows())
    msg = (f"DVT Reporte Semanal {HOY.strftime('%d/%m/%Y')}\n\n"
           f"{total} tramite(s) critico(s) (mas de {LIMITE} dias)\n\n"
           + (f"SALIDAS:\n{ls}\n\n" if ls else "")
           + (f"ENTRADAS:\n{le}\n\n" if le else "")
           + f"Dashboard: {LINK}")
    params = urllib.parse.urlencode({"phone": CALLMEBOT_PHONE, "text": msg, "apikey": CALLMEBOT_APIKEY})
    with urllib.request.urlopen(urllib.request.Request(
            f"https://api.callmebot.com/whatsapp.php?{params}",
            headers={"User-Agent": "dvt"})) as r:
        print(f"WhatsApp OK ({r.status})")

# ── Email ─────────────────────────────────────────────────────────────────────
def enviar_email(total):
    print("Enviando correo...")
    ts  = "width:100%;border-collapse:collapse;font-size:13px;margin-bottom:8px;"
    ths = "background:#1A5276;color:white;padding:8px;text-align:left;"
    tds = "padding:7px 8px;border-bottom:1px solid #eee;"
    bg  = lambda d: "#FDEDEC" if d > 30 else ("#FEF9E7" if d > 14 else "#fff")

    def tabla(rows, cols, fn):
        h = "".join(f'<th style="{ths}">{c}</th>' for c in cols)
        b = "".join(f'<tr style="background:{bg(int(r[\"dias\"]))}">'
                    + "".join(f'<td style="{tds}">{v}</td>' for v in fn(r))
                    + "</tr>" for _, r in rows.iterrows())
        return f'<table style="{ts}"><tr>{h}</tr>{b}</table>'

    bs = (f'<h3 style="font-size:13px;margin:20px 0 8px;color:#1A5276;">Salidas demoradas ({len(sal_alerta)})</h3>'
          + tabla(sal_alerta, ["Expediente","Título","Destino","Días"],
                  lambda r: [r["Expediente"], str(r["Título"])[:70], r["Destino"],
                             f'<strong style="color:{urgencia_color(int(r["dias"]))}">{int(r["dias"])} días</strong>'])
         ) if len(sal_alerta) > 0 else ""
    be = (f'<h3 style="font-size:13px;margin:20px 0 8px;color:#1A5276;">Entradas demoradas ({len(ent_alerta)})</h3>'
          + tabla(ent_alerta, ["Expediente","Título","Origen","Días"],
                  lambda r: [r["Expediente"], str(r["Título"])[:70], r["Origen"],
                             f'<strong style="color:{urgencia_color(int(r["dias"]))}">{int(r["dias"])} días</strong>'])
         ) if len(ent_alerta) > 0 else ""

    cuerpo = f"""<div style="font-family:Arial,sans-serif;max-width:720px;margin:0 auto;">
<div style="background:#1A5276;color:white;padding:16px 20px;border-radius:8px 8px 0 0;">
  <h2 style="margin:0;font-size:16px;">DVT — Reporte Semanal de Trámites Críticos</h2>
  <p style="margin:4px 0 0;font-size:12px;opacity:.8;">Lunes {HOY.strftime('%d/%m/%Y')} · {total} trámite(s) con más de {LIMITE} días sin movimiento</p>
</div>
<div style="padding:16px 20px;background:#fff;border:1px solid #ddd;border-top:none;">
  {bs}{be}
  <div style="margin-top:24px;text-align:center;">
    <a href="{LINK}" style="background:#1A5276;color:white;padding:10px 28px;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600;">Ver dashboard completo →</a>
  </div>
</div>
<div style="background:#f8f8f8;padding:10px 20px;border:1px solid #ddd;border-top:none;border-radius:0 0 8px 8px;font-size:11px;color:#aaa;text-align:center;">
  FBIOyF – UNR · Reporte automático semanal · Lunes 10:00 AM
</div>
</div>"""

    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"DVT Trámites Críticos — {total} alerta(s) — {HOY.strftime('%d/%m/%Y')}"
    msg["From"]    = GMAIL_REMITENTE
    msg["To"]      = ", ".join(GMAIL_DESTINATARIOS)
    msg.attach(MIMEText(cuerpo, "html"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as srv:
        srv.login(GMAIL_REMITENTE, GMAIL_PASSWORD)
        srv.sendmail(GMAIL_REMITENTE, GMAIL_DESTINATARIOS, msg.as_string())
    print(f"Correo enviado a: {', '.join(GMAIL_DESTINATARIOS)}")

# ── Main ─────────────────────────────────────────────────────────────────────
try:
    total = len(sal_alerta) + len(ent_alerta)
    if total > 0:
        enviar_whatsapp(total)
        enviar_email(total)
        print(f"\nListo. {total} trámites críticos notificados.")
    else:
        print("\nDashboard actualizado. Sin trámites críticos esta semana.")
except Exception as e:
    print(f"\nError: {e}")
    sys.exit(1)
