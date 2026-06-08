"""
generar_dashboard.py
────────────────────────────────────────────────────────────────
Oliver Sports · Dashboard de Cobranzas

LÓGICA DE MESES:
  Si hoy es JUNIO → mes cerrado = MAYO
  · Vencidas en MAYO          → pendiente (aviso → suspender)
  · Vencidas en ABRIL o antes → mora (mail legal)
  · Vencidas en JUNIO o más   → ignoradas

HOJAS EN EL SHEET:
  2026-05 PASO1  → snapshot al enviar avisos (1er día hábil)
  2026-05 PASO2  → snapshot al suspender (+3 días)
  (se crean automáticamente)

USO:
  source dashboard-env/bin/activate
  python3 generar_dashboard.py          ← detecta el paso solo
  python3 generar_dashboard.py --paso 1 ← forzar paso 1
  python3 generar_dashboard.py --paso 2 ← forzar paso 2
────────────────────────────────────────────────────────────────
"""

import json
import sys
import gspread
from pathlib import Path
from datetime import datetime, date
from google.oauth2.service_account import Credentials

# ════════════════════════════════════════════════════
#  ▼▼▼  EDITÁ ESTAS DOS LÍNEAS  ▼▼▼
SHEET_ID    = "1A4HcNz_eua54z5SqTGnJxQzgen94U1DZmHO4bWcgpKg"
CREDENTIALS = "credenciales.json"
#  ▲▲▲  SOLO ESTO HAY QUE TOCAR  ▲▲▲
# ════════════════════════════════════════════════════

TEMPLATE_HTML = "template.html"
OUTPUT_HTML   = "dashboard_cobranzas.html"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

MESES_ES = {
    1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril',
    5:'Mayo',  6:'Junio',   7:'Julio', 8:'Agosto',
    9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'
}

# Facturas a excluir de pendientes (ej: facturadas en el mes actual, no son cuotas adeudadas)
# Agregar números de factura que no correspondan al mes cerrado
EXCLUIR_NOS = set()


# ════════════════════════════════════════════════════
#  FECHAS
# ════════════════════════════════════════════════════

def mes_cerrado():
    hoy = date.today()
    if hoy.month == 1:
        return date(hoy.year - 1, 12, 1)
    return date(hoy.year, hoy.month - 1, 1)

def fin_mes(d):
    if d.month == 12:
        return date(d.year + 1, 1, 1)
    return date(d.year, d.month + 1, 1)

def nombre_hoja(mc, paso):
    return f"{mc.year}-{mc.month:02d} PASO{paso}"


# ════════════════════════════════════════════════════
#  CONEXIÓN
# ════════════════════════════════════════════════════

def conectar():
    creds_path = Path(CREDENTIALS)
    if not creds_path.exists():
        print(f"❌  No se encontró: {CREDENTIALS}")
        raise SystemExit(1)
    creds  = Credentials.from_service_account_file(str(creds_path), scopes=SCOPES)
    client = gspread.authorize(creds)
    try:
        sheet = client.open_by_key(SHEET_ID)
    except gspread.exceptions.SpreadsheetNotFound:
        print("❌  Sheet no encontrado. Verificá el SHEET_ID.")
        raise SystemExit(1)
    return sheet


# ════════════════════════════════════════════════════
#  PARSERS
# ════════════════════════════════════════════════════

def parse_float(val):
    if val is None: return 0.0
    s = str(val).strip().replace('€','').replace(' ','')
    if s in ('', '-', 'nan'): return 0.0
    if ',' in s and '.' in s:
        s = s.replace('.','').replace(',','.') if s.rfind(',') > s.rfind('.') else s.replace(',','')
    elif ',' in s:
        s = s.replace(',','.')
    try:    return float(s)
    except: return 0.0

def parsear_fecha(val):
    if val is None: return None
    if isinstance(val, datetime): return val.date()
    if isinstance(val, date):     return val
    s = str(val).strip()
    if s in ('', 'nan'): return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try: return datetime.strptime(s, fmt).date()
        except: continue
    return None


# ════════════════════════════════════════════════════
#  LEER HOJA PRINCIPAL (QuickBooks)
# ════════════════════════════════════════════════════

def leer_hoja_principal(sheet, mc):
    print("   Leyendo hoja principal...", end=" ", flush=True)
    ws   = sheet.get_worksheet(0)
    data = ws.get_all_values()

    header_row = 0
    for i, row in enumerate(data[:5]):
        if any('status' == c.strip().lower() for c in row):
            header_row = i
            break

    headers = [h.strip() for h in data[header_row]]
    rows    = data[header_row + 1:]
    print(f"{len(rows)} filas")

    def col(names):
        for n in names:
            for i, h in enumerate(headers):
                if h.lower() == n.lower():
                    return i
        return None

    i_no      = col(['No.','No','Number'])
    i_date    = col(['Date','Fecha'])
    i_cust    = col(['Customer','Cliente'])
    i_due     = col(['Due date','Due Date','Vencimiento'])
    i_balance = col(['Balance','Saldo'])
    i_email   = col(['Email','Mail'])
    i_status  = col(['Status','Estado'])

    if i_status is None:
        print(f"❌  No se encontró columna Status. Encabezados: {headers}")
        raise SystemExit(1)

    limite = fin_mes(mc)
    hoy    = date.today()
    result = []

    for row in rows:
        if not any(row): continue
        sv = row[i_status].strip() if i_status < len(row) else ''
        if sv not in ('overdue', 'open'): continue

        due = parsear_fecha(row[i_due] if i_due is not None and i_due < len(row) else '')
        if due is None or due >= limite: continue

        no = row[i_no].strip() if i_no is not None and i_no < len(row) else ''
        if no in EXCLUIR_NOS: continue  # excluir facturas que no corresponden al mes

        emit = parsear_fecha(row[i_date] if i_date is not None and i_date < len(row) else '')
        bal  = parse_float(row[i_balance] if i_balance is not None and i_balance < len(row) else 0)

        cat = 'pendiente' if (due.year == mc.year and due.month == mc.month) else 'mora'

        emails_raw = row[i_email] if i_email is not None and i_email < len(row) else ''
        emails = [
            e.strip() for e in emails_raw.split(',')
            if e.strip()
            and 'tryoliver'    not in e.lower()
            and 'oliversports' not in e.lower()
        ]

        result.append({
            'no':           no,
            'emission':     emit.strftime('%d/%m/%Y') if emit else '—',
            'customer':     row[i_cust].strip() if i_cust is not None and i_cust < len(row) else '',
            'due':          due.strftime('%d/%m/%Y'),
            'due_month':    due.strftime('%Y-%m'),
            'balance':      round(bal, 2),
            'email':        ', '.join(emails),
            'days_overdue': (hoy - due).days,
            'status':       cat,
        })

    result.sort(key=lambda x: (0 if x['status']=='mora' else 1, -x['days_overdue']))
    return result


# ════════════════════════════════════════════════════
#  SNAPSHOT → HOJA DEL SHEET
# ════════════════════════════════════════════════════

SNAPSHOT_HEADERS = [
    'No.', 'Cliente', 'Email', 'Vencimiento', 'Mes venc.',
    'Balance', 'Dias atraso', 'Tipo', 'Timestamp'
]

def guardar_snapshot(sheet, nombre, rows, paso):
    nombres_existentes = [ws.title for ws in sheet.worksheets()]
    if nombre in nombres_existentes:
        print(f"   ⚠  La hoja '{nombre}' ya existe — se sobreescribirá.")
        ws = sheet.worksheet(nombre)
        ws.clear()
    else:
        ws = sheet.add_worksheet(title=nombre, rows=len(rows)+10, cols=len(SNAPSHOT_HEADERS))
        print(f"   ✅ Hoja '{nombre}' creada.")

    ts   = datetime.now().strftime('%d/%m/%Y %H:%M')
    data = [SNAPSHOT_HEADERS]
    for r in rows:
        data.append([
            r['no'], r['customer'], r['email'], r['due'], r['due_month'],
            r['balance'], r['days_overdue'], r['status'], ts,
        ])
    ws.update('A1', data)

    try:
        ws.format('A1:I1', {
            "backgroundColor": {"red": 0.05, "green": 0.11, "blue": 0.18},
            "textFormat": {"bold": True, "foregroundColor": {"red":1,"green":1,"blue":1}},
        })
    except Exception:
        pass
    return ws


def leer_snapshot(sheet, nombre):
    """
    Lee una hoja del Sheet (PASO1 o PASO2 — formato QuickBooks exportado por vos).
    Formato esperado: fila 0 = título, fila 1 = headers, fila 2+ = datos
    Columnas: Date, Type, No., Customer, Memo, Due date, Balance, Email, Tax, Amount, Status
    Retorna None si la hoja no existe.
    """
    nombres = [ws.title for ws in sheet.worksheets()]
    if nombre not in nombres:
        return None

    ws   = sheet.worksheet(nombre)
    data = ws.get_all_values()
    if len(data) < 2:
        return []

    # Buscar fila de encabezados (puede estar en fila 0 o 1)
    header_idx = 0
    for i, row in enumerate(data[:5]):
        if any(c.strip().lower() in ('no.', 'status', 'customer') for c in row):
            header_idx = i
            break

    headers = [h.strip() for h in data[header_idx]]
    rows    = data[header_idx + 1:]

    def col(names):
        for n in names:
            for i, h in enumerate(headers):
                if h.strip().lower() == n.lower():
                    return i
        return None

    i_no     = col(['No.', 'No', 'Number'])
    i_cust   = col(['Customer', 'Cliente'])
    i_email  = col(['Email', 'Mail'])
    i_due    = col(['Due date', 'Due Date', 'Vencimiento'])
    i_bal    = col(['Balance', 'Saldo'])
    i_status = col(['Status', 'Estado'])
    i_date   = col(['Date', 'Fecha'])

    # Inferir mes del nombre de la hoja: "2026-05 PASO1" → 2026-05
    mc_hoja = nombre[:7]  # "2026-05"
    try:
        y, m = int(mc_hoja[:4]), int(mc_hoja[5:7])
        from datetime import date as _d
        limite = _d(y, m + 1, 1) if m < 12 else _d(y + 1, 1, 1)
    except Exception:
        limite = None

    today = __import__('datetime').date.today()
    result = []

    for row in rows:
        if not any(row): continue

        no  = row[i_no].strip()     if i_no  is not None and i_no  < len(row) else ''
        sv  = row[i_status].strip() if i_status is not None and i_status < len(row) else ''
        if not no or not sv: continue

        due = parsear_fecha(row[i_due] if i_due is not None and i_due < len(row) else '')
        bal = parse_float(row[i_bal] if i_bal is not None and i_bal < len(row) else 0)
        emit = parsear_fecha(row[i_date] if i_date is not None and i_date < len(row) else '')

        emails_raw = row[i_email] if i_email is not None and i_email < len(row) else ''
        emails = [e.strip() for e in emails_raw.split(',')
                  if e.strip()
                  and 'tryoliver'    not in e.lower()
                  and 'oliversports' not in e.lower()]

        # Para PASO1: solo overdue del mes cerrado y mora de meses anteriores
        # Para PASO2: guardamos todo (overdue Y paid) para poder comparar
        due_month = due.strftime('%Y-%m') if due else ''
        days_overdue = (today - due).days if due else 0

        # Clasificar categoría (solo para overdue)
        if sv == 'overdue':
            if due and due_month == mc_hoja:
                cat = 'pendiente'
            elif due and (limite is None or due < limite):
                cat = 'mora'
            else:
                continue  # overdue futuro → ignorar
        elif sv == 'paid':
            cat = 'paid'  # pagó entre PASO1 y PASO2
        else:
            continue  # open u otros → ignorar

        result.append({
            'no':           no,
            'customer':     row[i_cust].strip() if i_cust is not None and i_cust < len(row) else '',
            'email':        ', '.join(emails),
            'emission':     emit.strftime('%d/%m/%Y') if emit else '—',
            'due':          due.strftime('%d/%m/%Y') if due else '—',
            'due_month':    due_month,
            'balance':      round(bal, 2),
            'days_overdue': days_overdue,
            'status':       cat,        # pendiente / mora (para PASO1) — paid (para PASO2)
            'status_raw':   sv,         # overdue / paid (status original de QuickBooks)
        })

    return result


def detectar_paso(sheet, mc):
    nombres  = [ws.title for ws in sheet.worksheets()]
    n1, n2   = nombre_hoja(mc, 1), nombre_hoja(mc, 2)
    if not n1 in nombres: return 1
    if not n2 in nombres: return 2
    print(f"\n   ⚠  Tanto '{n1}' como '{n2}' ya existen.")
    print(f"   1 → Regenerar PASO1  |  2 → Regenerar PASO2")
    resp = input("   Ingresá 1 o 2: ").strip()
    return int(resp) if resp in ('1','2') else 1


# ════════════════════════════════════════════════════
#  CONSTRUIR HISTORIAL PRE-CARGADO
# ════════════════════════════════════════════════════

def construir_historial(rows_p1, rows_p2, mc, paso):
    """
    Construye el historial de eventos para inyectar en el HTML.
    PASO1: registra avisos (pendiente) y notif. legales (mora)
    PASO2: agrega las suspensiones (los que no pagaron)
    Siempre acumula: PASO2 incluye los eventos de PASO1 + los nuevos.
    """
    nombre_p1 = nombre_hoja(mc, 1)
    nombre_p2 = nombre_hoja(mc, 2)
    mes_envio = date.today().strftime('%Y-%m')  # mes en que se trabaja (ej: 2026-06)

    ts1 = date.today().strftime('%d/%m/%Y') + ' 09:00'
    ts2 = date.today().strftime('%d/%m/%Y') + ' 09:00'

    hist = []

    # Eventos de PASO1: avisos + mora
    p1_pendiente = [r for r in rows_p1 if r['status'] == 'pendiente']
    p1_mora      = [r for r in rows_p1 if r['status'] == 'mora']

    for r in p1_pendiente:
        hist.append({
            'ts': ts1, 'month': mes_envio,
            'due_month': r['due_month'], 'hoja': nombre_p1,
            'invoice': r['no'], 'customer': r['customer'],
            'type': 'aviso', 'balance': r['balance'],
            'from': 'support@oliversports.ai', 'to': r['email'],
        })
    for r in p1_mora:
        hist.append({
            'ts': ts1, 'month': mes_envio,
            'due_month': r['due_month'], 'hoja': nombre_p1,
            'invoice': r['no'], 'customer': r['customer'],
            'type': 'mora', 'balance': r['balance'],
            'from': 'legal@oliversports.ai', 'to': r['email'],
        })

    # Eventos de PASO2: suspensiones (solo si estamos en paso 2)
    if paso == 2 and rows_p2:
        p2_pendiente = [r for r in rows_p2 if r['status'] == 'pendiente']
        for r in p2_pendiente:
            hist.append({
                'ts': ts2, 'month': mes_envio,
                'due_month': r['due_month'], 'hoja': nombre_p2,
                'invoice': r['no'], 'customer': r['customer'],
                'type': 'suspend', 'balance': r['balance'],
                'from': 'support@oliversports.ai', 'to': r['email'],
            })

    return hist


# ════════════════════════════════════════════════════
#  GENERAR HTML
# ════════════════════════════════════════════════════

def generar_html(rows_tabla, rows_p1, rows_p2, mc, paso):
    template_path = Path(TEMPLATE_HTML)
    if not template_path.exists():
        print(f"❌  No se encontró {TEMPLATE_HTML}")
        raise SystemExit(1)
    with open(template_path, 'r', encoding='utf-8') as f:
        html = f.read()

    mora_rows      = [r for r in rows_tabla if r['status'] == 'mora']
    pendiente_rows = [r for r in rows_tabla if r['status'] == 'pendiente']
    total_bal      = sum(r['balance'] for r in rows_tabla)
    mes_label      = f"{MESES_ES[mc.month]} {mc.year}"
    hoja_actual    = nombre_hoja(mc, paso)

    # Avisos enviados = pendientes del PASO1 (siempre del PASO1, no cambia en PASO2)
    count_avisos = len([r for r in rows_p1 if r['status'] == 'pendiente'])

    # Estados pre-cargados: suspendidos = los que están en PASO2 como pendiente
    estados = {}
    if paso == 2 and rows_p2:
        for r in rows_p2:
            if r['status'] == 'pendiente':
                estados[r['no']] = 'suspendido'

    # Historial pre-cargado
    hist = construir_historial(rows_p1, rows_p2 or [], mc, paso)

    # Inyectar localStorage antes del script principal
    ls_key_estado = f"oliver_estado_{hoja_actual.replace(' ', '_')}"
    preload = f"""<script>
(function(){{
  localStorage.setItem('{ls_key_estado}', JSON.stringify({json.dumps(estados)}));
  localStorage.setItem('oliver_hist_v3', JSON.stringify({json.dumps(hist)}));
}})();
</script>"""
    html = html.replace('<script>', preload + '\n<script>', 1)

    hoy_es = datetime.today().strftime('%A %d de %B de %Y')

    html = html.replace('__DATA_JSON__',       json.dumps(rows_tabla, ensure_ascii=True))
    html = html.replace('__TODAY__',           hoy_es)
    html = html.replace('__MES_CERRADO__',     mes_label)
    html = html.replace('__PASO__',            str(paso))
    html = html.replace('__HOJA_ACTUAL__',     hoja_actual)
    html = html.replace('__COUNT_AVISOS__',    str(count_avisos))
    html = html.replace('__TOTAL_BALANCE__',   f"{total_bal:,.2f}")
    html = html.replace('__COUNT_TOTAL__',     str(len(rows_tabla)))
    html = html.replace('__COUNT_PENDIENTE__', str(len(pendiente_rows)))
    html = html.replace('__COUNT_MORA__',      str(len(mora_rows)))
    html = html.replace('__BAL_PENDIENTE__',   f"{sum(r['balance'] for r in pendiente_rows):,.2f}")
    html = html.replace('__BAL_MORA__',        f"{sum(r['balance'] for r in mora_rows):,.2f}")
    html = html.replace('__GENERATED_AT__',    datetime.now().strftime('%d/%m/%Y %H:%M'))
    return html


# ════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════

def main():
    print("─" * 56)
    print("  Oliver Sports · Dashboard de Cobranzas")
    print("─" * 56)

    paso_forzado = None
    for arg in sys.argv[1:]:
        if arg in ('--paso', '-p'):
            idx = sys.argv.index(arg)
            if idx + 1 < len(sys.argv):
                paso_forzado = int(sys.argv[idx + 1])
        elif arg.startswith('--paso='):
            paso_forzado = int(arg.split('=')[1])

    mc        = mes_cerrado()
    mes_label = f"{MESES_ES[mc.month]} {mc.year}"
    hoy       = date.today()

    print(f"\n📅 Hoy           : {hoy.strftime('%d/%m/%Y')}")
    print(f"   Mes de trabajo : {mes_label}")
    print(f"   Mora legal     : facturas vencidas antes de {mes_label}")

    print(f"\n🔗 Conectando a Google Sheets...")
    sheet = conectar()
    print(f"   Conectado: '{sheet.title}'")

    paso = paso_forzado if paso_forzado else detectar_paso(sheet, mc)
    nombre = nombre_hoja(mc, paso)
    print(f"\n📋 Ejecutando: {nombre}")
    print("─" * 56)

    rows_p1 = []
    rows_p2 = []

    # ── PASO 1 ──────────────────────────────────────────────
    if paso == 1:
        print(f"\n📥 PASO 1 — Leyendo facturas de {mes_label}...")
        rows_p1 = leer_hoja_principal(sheet, mc)

        mora_n      = sum(1 for r in rows_p1 if r['status'] == 'mora')
        pendiente_n = sum(1 for r in rows_p1 if r['status'] == 'pendiente')
        total_bal   = sum(r['balance'] for r in rows_p1)

        print(f"\n   Pendientes {mes_label} : {pendiente_n}  (enviar aviso hoy)")
        print(f"   Mora legal           : {mora_n}  (meses anteriores)")
        print(f"   Balance total        : €{total_bal:,.2f}")

        if not rows_p1:
            print("\n   ✅ No hay facturas pendientes.")
            return

        rows_tabla = rows_p1

    # ── PASO 2 ──────────────────────────────────────────────
    elif paso == 2:
        nombre_p1 = nombre_hoja(mc, 1)
        nombre_p2 = nombre_hoja(mc, 2)
        print(f"\n📥 PASO 2 — Comparando '{nombre_p1}' (overdue) vs '{nombre_p2}' (paid/overdue)...")

        # Leer PASO1 — snapshot de overdue al día 1
        rows_p1_snap = leer_snapshot(sheet, nombre_p1)
        if rows_p1_snap is None:
            print(f"❌  No existe la hoja '{nombre_p1}'.")
            raise SystemExit(1)
        rows_p1 = rows_p1_snap

        # Leer PASO2 — el mismo reporte unos días después (vos lo subís)
        rows_p2_snap = leer_snapshot(sheet, nombre_p2)
        if rows_p2_snap is None:
            print(f"❌  No existe la hoja '{nombre_p2}'.")
            print(f"   Subí el reporte actualizado de QuickBooks como '{nombre_p2}' y volvé a correr.")
            raise SystemExit(1)

        # Construir mapa: no_factura → status_raw (paid / overdue)
        # leer_snapshot ya incluye 'status_raw' con el valor original de QuickBooks
        status_p2 = {r['no']: r.get('status_raw', r['status']) for r in rows_p2_snap}

        # Comparar PASO1 vs PASO2:
        # Si en PASO2 el status es 'paid'    → pagó ✅ no suspender
        # Si en PASO2 sigue siendo 'overdue' → no pagó 🔴 suspender
        # Si no aparece en PASO2             → asumir que sigue overdue
        pagaron = []
        rows_p2 = []
        for r in rows_p1_snap:
            st = status_p2.get(r['no'], 'overdue')
            if st == 'paid':
                pagaron.append(r)
            else:
                rows_p2.append(r)

        mora_n      = sum(1 for r in rows_p2 if r['status'] == 'mora')
        pendiente_n = sum(1 for r in rows_p2 if r['status'] == 'pendiente')
        total_bal   = sum(r['balance'] for r in rows_p2)

        print(f"\n   Pagaron luego del aviso : {len(pagaron)} ✅")
        for r in pagaron:
            print(f"      ✓ {r['customer'][:40]} | {r['no']} | €{r['balance']:,.2f}")
        print(f"\n   Sin pagar → suspender  : {pendiente_n}")
        print(f"   Sin pagar → mora legal : {mora_n}")
        print(f"   Balance a ejecutar     : €{total_bal:,.2f}")

        if not rows_p2:
            print("\n   ✅ Todos pagaron. No hay suspensiones.")
            return

        rows_tabla = rows_p2

    # ── GENERAR HTML ──────────────────────────────────────
    print(f"\n🖊  Generando {OUTPUT_HTML}...")
    html = generar_html(rows_tabla, rows_p1, rows_p2, mc, paso)
    with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"   ✅ Guardado: {OUTPUT_HTML}")
    print(f"\n   Abrilo en Chrome para trabajar.")
    print("─" * 56)


if __name__ == '__main__':
    main()
