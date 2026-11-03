import sys
import os

try:
    import pdfplumber
    import pandas as pd
    import re
    import glob
    from datetime import datetime
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, FormulaRule
    from openpyxl.worksheet.properties import PageSetupProperties
    from openpyxl.worksheet.page import PageMargins
    import threading
    import subprocess
    import platform
    import queue
    import json
    import unicodedata
    import difflib
    from decimal import Decimal, ROUND_HALF_UP
    from openpyxl.chart import BarChart, LineChart, DoughnutChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.text import RichText
    from openpyxl.drawing.line import LineProperties
    from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
    from openpyxl.worksheet.hyperlink import Hyperlink
except ImportError as _e:
    if __name__ != "__main__":
        raise
    _msg = (f"Falta una librería: {_e}\n\n"
            f"Python usado: {sys.executable}\n\n"
            "Instálala abriendo CMD y escribiendo:\n"
            f'"{sys.executable}" -m pip install pdfplumber pandas openpyxl')
    try:
        import tkinter as _tk
        from tkinter import messagebox as _mb
        _r = _tk.Tk(); _r.withdraw()
        _mb.showerror("Falta una librería", _msg)
    except Exception:
        print(_msg)
        input("Pulsa Enter para salir...")
    sys.exit(1)

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:          # servidor web sin pantalla: solo se usa el motor, no la ventana
    tk = filedialog = messagebox = ttk = None

NUM_PREGUNTAS = 18
ANCHO_MAX_COL = 45
UMBRAL_VISITA_CORTA = 20   # minutos: las visitas por debajo de este tiempo se cuentan como 'cortas'


def redondear(x, nd=1):
    """Redondeo 'mitad hacia arriba' (igual que la función REDONDEAR de Excel)."""
    try:
        return float(Decimal(repr(round(float(x), 9))).quantize(Decimal(1).scaleb(-nd), rounding=ROUND_HALF_UP))
    except Exception:
        return x

NOMBRES_PREGUNTAS = [
    "P1. Acción TNA", "P2. Roll up/Kit TNA", "P3. Buena imagen", "P4. Uniforme autorizado",
    "P5. Intensidades LYO/EVO", "P6. Stand visible y limpio", "P7. Showcase completo",
    "P8. Uso de Tablet", "P9. Conocimiento de registro", "P10. Abordaje espontáneo",
    "P11. Pregunta si es fumador", "P12. Presenta Ploom", "P13. Rebate objeciones",
    "P14. Ofrece combustible", "P15. Guía en el trial", "P16. Ofrece varias intensidades",
    "P17. Entrega regalo", "P18. Despedida y Ploom Club"
]

# --- BASE DE DATOS DE ESTANCOS (relaciona el nombre del estanco del PDF con su dirección) ---
def _sin_tildes(s):
    s = unicodedata.normalize("NFKD", str(s))
    return "".join(ch for ch in s if not unicodedata.combining(ch))


def clave_estanco(s):
    """Clave de comparación: minúsculas, sin tildes, sin guiones, espacios ni signos.
    'Boadilla del Monte - 003' y 'BOADILLA DEL MONTE-003' dan la misma clave."""
    return re.sub(r"[^a-z0-9]", "", _sin_tildes(s).lower())


def _clave_numero(k):
    """Ignora los ceros a la izquierda del número final: 'madrid029' == 'madrid29'."""
    m = re.match(r"^(.*?)(\d+)$", k)
    return (m.group(1) + str(int(m.group(2)))) if m else k


def cargar_base_estancos(ruta):
    """Lee las columnas 'Nombre' y 'Dirección' (y 'Provincia', si existe) del Excel de la base de datos."""
    raw = pd.read_excel(ruta, header=None, dtype=str)

    def claves_fila(i):
        return [clave_estanco(x) if isinstance(x, str) else "" for x in raw.iloc[i].tolist()]

    fila_cab = next((i for i in range(min(40, len(raw)))
                     if "nombre" in claves_fila(i) and "direccion" in claves_fila(i)), None)
    if fila_cab is None:
        raise ValueError("no encuentro las columnas 'Nombre' y 'Dirección' en el Excel")
    claves = claves_fila(fila_cab)
    c_nom, c_dir = claves.index("nombre"), claves.index("direccion")
    c_prov = claves.index("provincia") if "provincia" in claves else None

    base = {'k1': {}, 'k2': {}, 'n': 0}
    for fila in raw.iloc[fila_cab + 1:].itertuples(index=False):
        nombre, direccion = fila[c_nom], fila[c_dir]
        if not isinstance(nombre, str) or not nombre.strip():
            continue
        k1 = clave_estanco(nombre)
        if not k1:
            continue
        prov = fila[c_prov] if c_prov is not None else None
        reg = {'nombre': " ".join(nombre.split()),
               'direccion': " ".join(direccion.split()) if isinstance(direccion, str) else "",
               'prov': clave_estanco(prov) if isinstance(prov, str) else ""}
        base['k1'].setdefault(k1, []).append(reg)
        base['k2'].setdefault(_clave_numero(k1), []).append(reg)
        base['n'] += 1
    base['claves'] = list(base['k1'].keys())
    return base


def buscar_direccion(base, estanco, provincia=""):
    """Devuelve (dirección, aviso). El aviso va vacío si todo fue bien."""
    if not estanco or estanco == "No detectado":
        return "No encontrada", "sin nombre de estanco para buscar la dirección"
    k = clave_estanco(estanco)
    candidatos = [k] + ([k[len("estanco"):]] if k.startswith("estanco") and len(k) > len("estanco") else [])

    for cand in candidatos:
        for tabla, clave in (('k1', cand), ('k2', _clave_numero(cand))):
            regs = base[tabla].get(clave)
            if not regs:
                continue
            distintos = list({(r['nombre'], r['direccion']): r for r in regs}.values())
            if len(distintos) > 1:
                pk = clave_estanco(provincia) if provincia else ""
                por_prov = [r for r in distintos if pk and r['prov'] == pk]
                if len(por_prov) == 1:
                    distintos = por_prov
            if len(distintos) == 1:
                return (distintos[0]['direccion'] or "Sin dirección en la base"), ""
            return " | ".join(r['direccion'] for r in distintos), \
                   f"nombre de estanco repetido en la base ({len(distintos)} direcciones): revisar"

    parecido = difflib.get_close_matches(k, base['claves'], n=1, cutoff=0.88)
    sugerencia = f" (¿'{base['k1'][parecido[0]][0]['nombre']}'?)" if parecido else ""
    return "No encontrada", f"estanco '{estanco}' no encontrado en la base de datos{sugerencia}"


def extraer_datos(pdf_path):
    datos = {}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_completo += (pagina.extract_text() or "") + "\n"
            
            patron_coord = r"([-+]?\s*\d+[.,]\d+)"
            match_lat = re.search(r"Latitud\s*" + patron_coord, texto_completo, re.IGNORECASE)
            match_lon = re.search(r"Longitud\s*" + patron_coord, texto_completo, re.IGNORECASE)
            
            def limpiar_coord(match):
                if not match:
                    return 0.0
                texto = match.group(1).replace(' ', '').replace(',', '.')
                try:
                    return float(texto)
                except:
                    return 0.0
            
            datos['Latitud'] = limpiar_coord(match_lat)
            datos['Longitud'] = limpiar_coord(match_lon)
            
            def campo(etiqueta, respaldo):
                m_ = re.search(r"(?m)^[ \t]*" + etiqueta + r"[ \t]+(.+?)[ \t]*$", texto_completo)
                if not m_:
                    m_ = re.search(respaldo, texto_completo)
                if not m_:
                    return "No detectado"
                valor = re.sub(r"[*?~]", "", m_.group(1)).strip().lstrip("=<>").strip()
                return valor or "No detectado"

            letras = r"([A-Za-z0-9áéíóúÁÉÍÓÚñÑüÜ ]+)"
            datos['Provincia'] = campo("Provincia", r"Provincia\s*" + letras)
            datos['Estanco'] = campo("Estanco", r"Estanco\s*" + letras)
            datos['Promotor'] = campo("Promotor", r"Promotor\s*" + letras)
            datos['Tipo_Visita'] = campo("Tipo de visita", r"Tipo de visita\s*" + letras)
            datos['Coordinador'] = campo("Coordinador", r"Coordinador\s*" + letras)

            match_fecha = re.search(r"Fecha de visita\s*(\d{2}/\d{2}/\d{4})", texto_completo)
            match_horas = re.search(r"Hora inicio / fin\s*(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})", texto_completo)
            match_puntuacion = re.search(r"Puntuación\s*(\d+)\s*/\s*10", texto_completo)
            
            match_no_decl = re.search(r"respuestas\s+NO\s*:\s*(\d+)", texto_completo, re.IGNORECASE)
            match_envio = re.search(r"Pulsó ENVIAR.{0,80}?(\d{2}/\d{2}/\d{4})\s+(\d{2}:\d{2}:\d{2})", texto_completo, re.DOTALL)
            if match_envio:
                datos['Hora_Envio'] = match_envio.group(2)
                datos['Fecha_Envio'] = match_envio.group(1)
            else:
                datos['Hora_Envio'] = "00:00:00"
                datos['Fecha_Envio'] = "No detectado"
            
            datos['Fecha'] = match_fecha.group(1) if match_fecha else "No detectado"
            
            if match_horas:
                datos['Hora_Inicio'] = match_horas.group(1)
                datos['Hora_Fin'] = match_horas.group(2)
            else:
                datos['Hora_Inicio'] = "00:00"
                datos['Hora_Fin'] = "00:00"

            try:
                datos['Nota_Visita_Coordinador'] = int(match_puntuacion.group(1)) if match_puntuacion else 0
            except:
                datos['Nota_Visita_Coordinador'] = 0

            match_checklist = re.search(r"Tipo de acción(.*?)Valoración de la visita", texto_completo, re.DOTALL | re.IGNORECASE)
            
            if match_checklist:
                bloque_checklist = match_checklist.group(1)
                respuestas = re.findall(r'(?<![\w/])(SÍ|Sí|sí|NO|No|N/A)(?![\w/])', bloque_checklist)

                respuestas_limpias = []
                for r in respuestas:
                    if r.upper() == "SÍ": respuestas_limpias.append("Sí")
                    elif r.upper() == "NO": respuestas_limpias.append("No")
                    else: respuestas_limpias.append("N/A")

                for i in range(1, NUM_PREGUNTAS + 1):
                    if i <= len(respuestas_limpias):
                        datos[f'P{i}'] = respuestas_limpias[i-1]
                    else:
                        datos[f'P{i}'] = "No detectado"
            else:
                for i in range(1, NUM_PREGUNTAS + 1):
                    datos[f'P{i}'] = "No detectado"

    except Exception as e:
        print(f"Error procesando {pdf_path}: {e}")
        return None

    avisos = []
    if datos.get('Fecha') == "No detectado": avisos.append("sin fecha")
    if datos.get('Latitud') == 0.0 or datos.get('Longitud') == 0.0: avisos.append("sin GPS")
    if datos.get('Hora_Inicio') == "00:00": avisos.append("sin hora inicio/fin")
    if datos.get('Fecha_Envio') == "No detectado": avisos.append("sin hora de envío")
    if not match_puntuacion: avisos.append("sin nota del coordinador")
    n_resp = sum(1 for i in range(1, NUM_PREGUNTAS + 1) if datos.get(f'P{i}') in ("Sí", "No", "N/A"))
    if n_resp != NUM_PREGUNTAS: avisos.append(f"checklist con {n_resp}/{NUM_PREGUNTAS} respuestas")
    if match_no_decl:
        n_no = sum(1 for i in range(1, NUM_PREGUNTAS + 1) if datos.get(f'P{i}') == "No")
        if n_no != int(match_no_decl.group(1)):
            avisos.append(f"NO leídos {n_no} pero el informe declara {match_no_decl.group(1)}")
    datos['_avisos'] = avisos
    return datos

def calcular_metricas(datos):
    try:
        hora_inicio = datetime.strptime(datos['Hora_Inicio'], "%H:%M")
        hora_envio = datetime.strptime(datos['Hora_Envio'], "%H:%M:%S")
        diff = (hora_envio - hora_inicio).total_seconds() / 60
        duracion = int(diff)
        if duracion < 0:
            duracion = 0
    except Exception:
        duracion = 0
    
    datos['Duracion_min'] = duracion

    total_si = sum(1 for i in range(1, 19) if datos.get(f'P{i}') == "Sí")
    total_no = sum(1 for i in range(1, 19) if datos.get(f'P{i}') == "No")
    datos['Total_Si'] = total_si
    datos['Total_No'] = total_no
    

    return datos

def _a_fecha(s):
    try:
        return datetime.strptime(str(s), "%d/%m/%Y")
    except Exception:
        return s


def _a_hora(s, formato):
    try:
        return datetime.strptime(str(s), formato).time()
    except Exception:
        return s


def _ignorar_errores_texto(archivo, num_hoja, rango):
    """Evita los avisos verdes de 'número almacenado como texto' en las coordenadas (son texto a propósito)."""
    import zipfile
    tmp = archivo + ".tmp"
    try:
        destino = f"xl/worksheets/sheet{num_hoja}.xml"
        with zipfile.ZipFile(archivo) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == destino:
                    txt = data.decode("utf-8")
                    if "<ignoredErrors" not in txt and txt.rstrip().endswith("</worksheet>"):
                        txt = txt.rstrip()[:-len("</worksheet>")] + \
                              f'<ignoredErrors><ignoredError sqref="{rango}" numberStoredAsText="1"/></ignoredErrors></worksheet>'
                        data = txt.encode("utf-8")
                zout.writestr(item, data)
        os.replace(tmp, archivo)
    except Exception:
        if os.path.exists(tmp):
            os.remove(tmp)


def crear_dashboard(archivo, df, usar_formulas=True):
    wb = load_workbook(archivo)
    ws_datos = wb['Datos']
    U = UMBRAL_VISITA_CORTA

    color_azul_oscuro = "1F4E78"
    color_azul_claro = "D9E1F2"
    color_verde = "375623"

    font_titulo = Font(name='Calibri', size=16, bold=True, color="FFFFFF")
    font_subtitulo = Font(name='Calibri', size=12, bold=True, color="FFFFFF")
    font_header = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
    font_header_claro = Font(name='Calibri', size=11, bold=True, color="1F4E78")
    font_bold = Font(name='Calibri', size=11, bold=True)

    fill_azul_oscuro = PatternFill(start_color=color_azul_oscuro, end_color=color_azul_oscuro, fill_type="solid")
    fill_azul_claro = PatternFill(start_color=color_azul_claro, end_color=color_azul_claro, fill_type="solid")
    fill_verde = PatternFill(start_color=color_verde, end_color=color_verde, fill_type="solid")
    rojo_suave = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    amarillo_suave = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    gris_suave = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")

    border_fino = Border(left=Side(style='thin', color='BFBFBF'), right=Side(style='thin', color='BFBFBF'),
                         top=Side(style='thin', color='BFBFBF'), bottom=Side(style='thin', color='BFBFBF'))
    align_center = Alignment(horizontal="center", vertical="center")
    align_center_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cols = [cell.value for cell in ws_datos[1]]
    n = ws_datos.max_row  # última fila con datos

    def L(nombre):
        return get_column_letter(cols.index(nombre) + 1)

    def rng(nombre):
        return f"Datos!${L(nombre)}$2:${L(nombre)}${n}"

    def F(formula, valor):
        """Fórmula si el Excel es recalculable; si no, el valor fijo."""
        return formula if usar_formulas else valor

    def semaforo_duracion(ws_, rango, fila_borde=None):
        bd = border_fino if fila_borde else None
        ws_.conditional_formatting.add(rango, CellIsRule(operator='lessThan', formula=['5'], fill=rojo_suave, border=bd, stopIfTrue=True))
        ws_.conditional_formatting.add(rango, CellIsRule(operator='lessThan', formula=['10'], fill=amarillo_suave, border=bd, stopIfTrue=True))
        if fila_borde:
            ws_.conditional_formatting.add(rango, FormulaRule(formula=[f'$B{fila_borde}<>""'], border=border_fino))

    def preparar_impresion(ws_, orientacion, titulo_filas=None):
        ws_.page_setup.orientation = orientacion
        ws_.page_setup.paperSize = ws_.PAPERSIZE_A4
        ws_.page_setup.fitToWidth = 1
        ws_.page_setup.fitToHeight = 0
        ws_.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws_.print_options.horizontalCentered = True
        ws_.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)
        ws_.oddFooter.center.text = "Página &P de &N"
        if titulo_filas:
            ws_.print_title_rows = titulo_filas

    # =====================  HOJA "DATOS"  =====================
    ws_datos.freeze_panes = "A2"
    for i, nombre in enumerate(cols, start=1):
        if nombre in df.columns and len(df) > 0:
            largo = max([len(str(nombre))] + [len(str(v)) for v in df[nombre]])
        else:
            largo = len(str(nombre))
        ws_datos.column_dimensions[get_column_letter(i)].width = min(largo + 3, ANCHO_MAX_COL)

    for cell in ws_datos[1]:
        cell.font = font_header
        cell.fill = fill_azul_oscuro
        cell.alignment = align_center
        cell.border = border_fino

    idx = {nombre: cols.index(nombre) + 1 for nombre in cols}
    cI, cE = L('Hora_Inicio'), L('Hora_Envio')
    p_ini, p_fin = L('P1'), L(f'P{NUM_PREGUNTAS}')

    for r in range(2, n + 1):
        for nombre in ('Latitud', 'Longitud'):          # coordenadas: texto con punto y 6 decimales
            c = ws_datos.cell(row=r, column=idx[nombre])
            c.number_format = '@'
            if isinstance(c.value, (int, float)):
                c.value = f"{c.value:.6f}".replace(',', '.')

        c = ws_datos.cell(row=r, column=idx['Fecha'])   # fecha y horas como valores reales
        c.value = _a_fecha(c.value)
        if isinstance(c.value, datetime):
            c.number_format = 'dd/mm/yyyy'
        c.alignment = align_center
        c = ws_datos.cell(row=r, column=idx['Hora_Inicio'])
        c.value = _a_hora(c.value, "%H:%M")
        c.number_format = 'hh:mm'
        c.alignment = align_center
        c = ws_datos.cell(row=r, column=idx['Hora_Envio'])
        c.value = _a_hora(c.value, "%H:%M:%S")
        c.number_format = 'hh:mm:ss'
        c.alignment = align_center

        ws_datos.cell(row=r, column=idx['Duracion_min']).alignment = align_center
        for i in range(1, NUM_PREGUNTAS + 1):
            c = ws_datos.cell(row=r, column=idx[f'P{i}'])
            c.alignment = align_center
            c.border = border_fino

        if usar_formulas:
            ws_datos.cell(row=r, column=idx['Duracion_min'],
                          value=f"=IFERROR(MAX(0,INT(ROUND(({cE}{r}-{cI}{r})*86400,0)/60)),0)")
            ws_datos.cell(row=r, column=idx['Total_Si'], value=f'=COUNTIF({p_ini}{r}:{p_fin}{r},"Sí")')
            ws_datos.cell(row=r, column=idx['Total_No'], value=f'=COUNTIF({p_ini}{r}:{p_fin}{r},"No")')

    # Columnas auxiliares ocultas:
    #   _Visible: 1 si la fila se ve con el filtro actual, 0 si está filtrada
    #   _Dia: 1 en la primera fila visible de cada pareja (coordinador, fecha) -> días trabajados
    col_vis, col_dia = get_column_letter(len(cols) + 1), get_column_letter(len(cols) + 2)
    if usar_formulas:
        ws_datos.cell(row=1, column=len(cols) + 1, value="_Visible")
        ws_datos.cell(row=1, column=len(cols) + 2, value="_Dia")
        cc, cf = L('Coordinador'), L('Fecha')
        for r in range(2, n + 1):
            ws_datos.cell(row=r, column=len(cols) + 1, value=f"=SUBTOTAL(103,A{r})")
            ws_datos.cell(row=r, column=len(cols) + 2,
                          value=(f"=IF({col_vis}{r}=1,IF(COUNTIFS(${cc}$2:{cc}{r},{cc}{r},${cf}$2:{cf}{r},{cf}{r},"
                                 f"${col_vis}$2:{col_vis}{r},1)=1,1,0),0)"))
        ws_datos.column_dimensions[col_vis].hidden = True
        ws_datos.column_dimensions[col_dia].hidden = True
    V = f"Datos!${col_vis}$2:${col_vis}${n}"
    DIA = f"Datos!${col_dia}$2:${col_dia}${n}"

    ws_datos.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{n}"
    semaforo_duracion(ws_datos, f"{L('Duracion_min')}2:{L('Duracion_min')}{n}")

    rango_p = f"{p_ini}2:{p_fin}{n}"
    ws_datos.conditional_formatting.add(rango_p, CellIsRule(
        operator='equal', formula=['"No"'], fill=rojo_suave, font=Font(color="9C0006", bold=True)))
    ws_datos.conditional_formatting.add(rango_p, CellIsRule(
        operator='equal', formula=['"No detectado"'], fill=gris_suave, font=Font(color="7F7F7F", italic=True)))
    ws_datos.conditional_formatting.add(rango_p, CellIsRule(
        operator='equal', formula=['"N/A"'], font=Font(color="7F7F7F")))
    preparar_impresion(ws_datos, 'landscape', titulo_filas='1:1')

    # =====================  HOJA "RESUMEN"  =====================
    ws = wb.create_sheet("Resumen")
    ws.sheet_view.showGridLines = False

    ws.merge_cells('B2:H2')
    ws['B2'] = "DASHBOARD DE AUDITORÍA DE VISITAS"
    ws['B2'].font = font_titulo
    ws['B2'].fill = fill_azul_oscuro
    ws['B2'].alignment = align_center
    ws.merge_cells('B3:H3')
    ws['B3'] = ("Los datos de este resumen siguen los filtros aplicados en la hoja Datos."
                if usar_formulas else "Resumen con valores fijos (no sigue los filtros de la hoja Datos).")
    ws['B3'].font = Font(name='Calibri', size=9, italic=True, color="7F7F7F")
    ws['B3'].alignment = align_center

    dur, si, no = rng('Duracion_min'), rng('Total_Si'), rng('Total_No')
    total_visitas = len(df)
    n_cortas = int((df['Duracion_min'] < U).sum())
    total_si, total_fallos = int(df['Total_Si'].sum()), int(df['Total_No'].sum())
    pares = df.groupby(['Coordinador', 'Fecha']).ngroups
    vis_dia = redondear(total_visitas / pares, 1) if pares else 0
    cumpl = redondear(total_si / (total_si + total_fallos) * 100, 1) if (total_si + total_fallos) else 0

    kpis = [
        ("Total Visitas", F(f"=SUBTOTAL(103,{rng('Fecha')})", total_visitas), "B4", '0'),
        ("Duración Media (min)", F(f"=IFERROR(ROUND(AVERAGEIFS({dur},{V},1),1),0)", redondear(df['Duracion_min'].mean(), 1)), "C4", '0.0'),
        (f"Visitas < {U} min", F(f'=COUNTIFS({dur},"<{U}",{V},1)', n_cortas), "D4", '0'),
        ("Total Fallos (NO)", F(f"=SUMIFS({no},{V},1)", total_fallos), "E4", '0'),
        ("Visitas por Coord. y Día", F(f"=IFERROR(ROUND(B5/SUM({DIA}),1),0)", vis_dia), "F4", '0.0'),
        ("Cumplimiento Checklist", F(f"=IFERROR(ROUND(SUMIFS({si},{V},1)/(SUMIFS({si},{V},1)+SUMIFS({no},{V},1))*100,1),0)", cumpl), "G4", '0.0"%"'),
    ]
    for titulo_k, valor, celda, fmt in kpis:
        ws[celda] = titulo_k
        ws[celda].font = font_bold
        ws[celda].alignment = align_center_wrap
        val_cell = celda.replace("4", "5")
        ws[val_cell] = valor
        ws[val_cell].number_format = fmt
        ws[val_cell].font = Font(size=14, bold=True, color=color_azul_oscuro)
        ws[val_cell].alignment = align_center
    ws.row_dimensions[4].height = 30

    # ---- Rendimiento por coordinador (solo aparecen los coordinadores visibles con el filtro) ----
    start_coord_row = 7
    ws.merge_cells(start_row=start_coord_row, start_column=2, end_row=start_coord_row, end_column=8)
    cell_coord = ws.cell(row=start_coord_row, column=2, value="RENDIMIENTO POR COORDINADOR")
    cell_coord.font = font_subtitulo
    cell_coord.fill = fill_verde
    cell_coord.alignment = align_center

    resumen_coord = df.groupby('Coordinador').agg(
        Visitas=('Fecha', 'count'),
        Dias=('Fecha', 'nunique'),
        Duracion_Media=('Duracion_min', 'mean'),
        Cortas=('Duracion_min', lambda x: int((x < U).sum())),
        Fallos=('Total_No', 'sum')
    ).reset_index()
    resumen_coord['Duracion_Media'] = resumen_coord['Duracion_Media'].map(lambda v: redondear(v, 1))
    resumen_coord['Vis_Dia'] = (resumen_coord['Visitas'] / resumen_coord['Dias']).map(lambda v: redondear(v, 1))
    resumen_coord = resumen_coord.sort_values(by=['Visitas', 'Coordinador'], ascending=[False, True],
                                              kind='stable').reset_index(drop=True)

    headers_coord = ["Coordinador", "Visitas", "Días con visitas", "Visitas por día",
                     "Duración Media (min)", f"Visitas < {U} min", "Total Fallos"]
    for i, header in enumerate(headers_coord):
        cell = ws.cell(row=start_coord_row + 1, column=2 + i, value=header)
        cell.font = font_header_claro
        cell.fill = fill_azul_claro
        cell.border = border_fino
        cell.alignment = align_center_wrap
    ws.row_dimensions[start_coord_row + 1].height = 30

    m = len(resumen_coord)
    start_data_coord = start_coord_row + 2
    end_data_coord = start_data_coord + m - 1
    rc = rng('Coordinador')

    if usar_formulas:
        # Bloque auxiliar oculto (J:L): visitas visibles de cada coordinador y clave de orden
        for i, row in resumen_coord.iterrows():
            ra = start_data_coord + i
            ws.cell(row=ra, column=10, value=row['Coordinador'])
            ws.cell(row=ra, column=11, value=f"=COUNTIFS({rc},J{ra},{V},1)")
            ws.cell(row=ra, column=12, value=f"=IF(K{ra}>0,K{ra}-{i + 1}/1000000,-1000-{i + 1})")
        keys = f"$L${start_data_coord}:$L${end_data_coord}"
        names = f"$J${start_data_coord}:$J${end_data_coord}"
        for k in range(m):
            r = start_data_coord + k
            ws.cell(row=r, column=2,
                    value=f'=IF(LARGE({keys},{k + 1})<-500,"",INDEX({names},MATCH(LARGE({keys},{k + 1}),{keys},0)))')
            formulas = [
                f'=IF($B{r}="","",COUNTIFS({rc},$B{r},{V},1))',
                f'=IF($B{r}="","",SUMIFS({DIA},{rc},$B{r}))',
                f'=IF($B{r}="","",IFERROR(ROUND(C{r}/D{r},1),0))',
                f'=IF($B{r}="","",IFERROR(ROUND(AVERAGEIFS({dur},{rc},$B{r},{V},1),1),0))',
                f'=IF($B{r}="","",COUNTIFS({rc},$B{r},{V},1,{dur},"<{U}"))',
                f'=IF($B{r}="","",SUMIFS({no},{rc},$B{r},{V},1))',
            ]
            for j, f in enumerate(formulas):
                c = ws.cell(row=r, column=3 + j, value=f)
                c.alignment = align_center
        for rango in (f"B{start_data_coord}:E{end_data_coord}", f"G{start_data_coord}:H{end_data_coord}"):
            ws.conditional_formatting.add(rango, FormulaRule(formula=[f'$B{start_data_coord}<>""'], border=border_fino))
        for col in ("J", "K", "L"):
            ws.column_dimensions[col].hidden = True
    else:
        for i, row in resumen_coord.iterrows():
            r = start_data_coord + i
            ws.cell(row=r, column=2, value=row['Coordinador']).border = border_fino
            valores = [int(row['Visitas']), int(row['Dias']), float(row['Vis_Dia']), float(row['Duracion_Media']),
                       int(row['Cortas']), int(row['Fallos'])]
            for j, v in enumerate(valores):
                c = ws.cell(row=r, column=3 + j, value=v)
                c.alignment = align_center
                c.border = border_fino

    semaforo_duracion(ws, f"F{start_data_coord}:F{end_data_coord}", start_data_coord if usar_formulas else None)

    for col, ancho in zip("ABCDEFGH", (3, 30, 14, 16, 16, 22, 18, 14)):
        ws.column_dimensions[col].width = ancho
    ws.print_area = f"A1:H{end_data_coord}"
    preparar_impresion(ws, 'portrait')

    # =====================  HOJA "DASHBOARD" (presentación gráfica)  =====================
    wd = wb.create_sheet("Dashboard", 0)
    wb.move_sheet("Resumen", offset=-1)          # orden final: Dashboard, Resumen, Datos
    wd.sheet_view.showGridLines = False
    wd.sheet_properties.tabColor = color_azul_oscuro
    ws.sheet_properties.tabColor = "2E75B6"
    ws_datos.sheet_properties.tabColor = "7F7F7F"

    # ---- Hoja auxiliar oculta con los datos de los gráficos ----
    wc = wb.create_sheet("_Calc")
    wc.sheet_state = 'hidden'
    nc = min(m, 15)                                 # máx. coordinadores en los gráficos
    for j, h in enumerate(["Coordinador", "Visitas", "Visitas por día", "Duración media (min)",
                           f"% visitas < {U} min", "Fallos por visita"], start=1):
        wc.cell(row=1, column=j, value=h)
    for k in range(nc):
        rr, rs = k + 2, start_data_coord + k
        if usar_formulas:
            wc.cell(row=rr, column=1, value=f"=Resumen!B{rs}")
            wc.cell(row=rr, column=2, value=f"=Resumen!C{rs}")
            wc.cell(row=rr, column=3, value=f"=Resumen!E{rs}")
            wc.cell(row=rr, column=4, value=f"=Resumen!F{rs}")
            wc.cell(row=rr, column=5, value=f'=IF(A{rr}="","",IF(B{rr}=0,0,ROUND(Resumen!G{rs}/B{rr}*100,1)))')
            wc.cell(row=rr, column=6, value=f'=IF(A{rr}="","",IF(B{rr}=0,0,ROUND(Resumen!H{rs}/B{rr},1)))')
        else:
            fc = resumen_coord.iloc[k]
            for j, v in enumerate([fc['Coordinador'], int(fc['Visitas']), float(fc['Vis_Dia']),
                                   float(fc['Duracion_Media']),
                                   redondear(fc['Cortas'] / fc['Visitas'] * 100, 1),
                                   redondear(fc['Fallos'] / fc['Visitas'], 1)], start=1):
                wc.cell(row=rr, column=j, value=v)

    fechas = list(df['Fecha'].unique())
    fechas.sort(key=lambda s: _a_fecha(s) if isinstance(_a_fecha(s), datetime) else datetime.max)
    fechas = fechas[-62:]                           # máx. 62 días en el gráfico de evolución
    rf = rng('Fecha')
    for j, h in enumerate(["Fecha", "Etiqueta", "Visitas", "Duración media (min)"], start=8):
        wc.cell(row=1, column=j, value=h)
    for k, f_txt in enumerate(fechas):
        rr = k + 2
        f_val = _a_fecha(f_txt)
        c_f = wc.cell(row=rr, column=8, value=f_val)
        if isinstance(f_val, datetime):
            c_f.number_format = 'dd/mm/yyyy'
        wc.cell(row=rr, column=9, value=f_val.strftime("%d/%m") if isinstance(f_val, datetime) else str(f_txt))
        sub_f = df[df['Fecha'] == f_txt]
        wc.cell(row=rr, column=10, value=F(f"=COUNTIFS({rf},H{rr},{V},1)", int(len(sub_f))))
        wc.cell(row=rr, column=11, value=F(
            f"=IF(J{rr}=0,NA(),ROUND(AVERAGEIFS({dur},{rf},H{rr},{V},1),1))",
            redondear(sub_f['Duracion_min'].mean(), 1)))
    nf = len(fechas)

    # ---- Cabecera ----
    for col in range(1, 15):
        wd.column_dimensions[get_column_letter(col)].width = 2 if col in (1, 14) else 12

    def boton(ws_, rango, texto, destino):
        if ":" in rango:
            ws_.merge_cells(rango)
        ancla = ws_[rango.split(":")[0]]
        ancla.value = texto
        ancla.font = Font(name='Calibri', size=10, bold=True, color="FFFFFF")
        ancla.fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
        ancla.alignment = align_center
        ancla.hyperlink = Hyperlink(ref=ancla.coordinate, location=destino, display=texto)
        ws_.row_dimensions[ancla.row].height = 22

    boton(wd, "J1:K1", "►  Resumen", "'Resumen'!A1")
    boton(wd, "L1:M1", "►  Datos", "'Datos'!A1")
    boton(ws, "B1", "►  Dashboard", "'Dashboard'!A1")
    boton(ws, "C1:D1", "►  Datos", "'Datos'!A1")

    wd.merge_cells('B2:M2')
    wd['B2'] = "DASHBOARD DE AUDITORÍA DE VISITAS  ·  CAMPAÑA PLOOM 2026"
    wd['B2'].font = Font(name='Calibri', size=18, bold=True, color="FFFFFF")
    wd['B2'].fill = fill_azul_oscuro
    wd['B2'].alignment = align_center
    wd.row_dimensions[2].height = 40
    wd.merge_cells('B3:M3')
    wd['B3'] = ("Los indicadores y gráficos siguen los filtros aplicados en la hoja Datos (coordinador, fecha, etc.)."
                if usar_formulas else "Informe con valores fijos (no sigue los filtros de la hoja Datos).")
    wd['B3'].font = Font(name='Calibri', size=9, italic=True, color="7F7F7F")
    wd['B3'].alignment = align_center
    wd.row_dimensions[4].height = 8

    # ---- Tarjetas de indicadores ----
    pct_cortas = redondear(n_cortas / total_visitas * 100, 1) if total_visitas else 0
    fallos_vis = redondear(total_fallos / total_visitas, 1) if total_visitas else 0
    tarjetas = [
        ("VISITAS AUDITADAS", "=Resumen!B5", '0', "2E75B6",
         F(f'="de "&COUNTA({rng("Fecha")})&" en total"', f"de {total_visitas} en total")),
        ("VISITAS POR COORD. Y DÍA", "=Resumen!F5", '0.0', "1F4E78", "media de visitas diarias"),
        ("DURACIÓN MEDIA (MIN)", "=Resumen!C5", '0.0', "548235", "de inicio a envío del informe"),
        (f"VISITAS < {U} MIN", "=Resumen!D5", '0', "C00000",
         F('=IF(Resumen!B5=0,"",ROUND(Resumen!D5/Resumen!B5*100,1)&"% de las visitas")', f"{pct_cortas}% de las visitas")),
        ("FALLOS EN CHECKLIST", "=Resumen!E5", '0', "ED7D31",
         F('=IF(Resumen!B5=0,"",ROUND(Resumen!E5/Resumen!B5,1)&" por visita")', f"{fallos_vis} por visita")),
        ("CUMPLIMIENTO CHECKLIST", "=Resumen!G5", '0.0"%"', "7030A0", "respuestas Sí sobre Sí + No"),
    ]
    fill_tarjeta = PatternFill(start_color="F3F6FA", end_color="F3F6FA", fill_type="solid")
    blanco = Side(style='thick', color="FFFFFF")
    for i, (etq, val, fmt, acento, sub) in enumerate(tarjetas):
        c1, c2 = 2 + 2 * i, 3 + 2 * i
        wd.merge_cells(start_row=5, start_column=c1, end_row=5, end_column=c2)
        wd.merge_cells(start_row=6, start_column=c1, end_row=7, end_column=c2)
        wd.merge_cells(start_row=8, start_column=c1, end_row=8, end_column=c2)
        for rr in range(5, 9):
            for cc_ in (c1, c2):
                cel = wd.cell(row=rr, column=cc_)
                cel.fill = fill_tarjeta
                cel.border = Border(left=blanco if cc_ == c1 else None, right=blanco if cc_ == c2 else None,
                                    top=Side(style='thick', color=acento) if rr == 5 else None)
        a_et, a_val, a_sub = wd.cell(row=5, column=c1), wd.cell(row=6, column=c1), wd.cell(row=8, column=c1)
        a_et.value = etq
        a_et.font = Font(name='Calibri', size=9, bold=True, color="595959")
        a_et.alignment = align_center
        a_val.value = val
        a_val.number_format = fmt
        a_val.font = Font(name='Calibri', size=28, bold=True, color=acento)
        a_val.alignment = align_center
        a_sub.value = sub
        a_sub.font = Font(name='Calibri', size=9, italic=True, color="7F7F7F")
        a_sub.alignment = align_center
    for rr, h in ((5, 20), (6, 22), (7, 22), (8, 18)):
        wd.row_dimensions[rr].height = h

    # Semáforo en el número grande de la duración media
    wd.conditional_formatting.add("F6:G7", CellIsRule(operator='lessThan', formula=['5'], font=Font(color="C00000"), stopIfTrue=True))
    wd.conditional_formatting.add("F6:G7", CellIsRule(operator='lessThan', formula=['10'], font=Font(color="BF8F00")))

    # ---- Gráficos ----
    def txt_props(sz, color="595959", bold=False):
        return CharacterProperties(sz=sz, b=bold, solidFill=color)

    def titulo(ch, texto):
        ch.title = texto
        cp = txt_props(1200, "1F4E78", True)
        ch.title.tx.rich.p[0].pPr = ParagraphProperties(defRPr=cp)
        for run in ch.title.tx.rich.p[0].r:
            run.rPr = cp
        ch.title.overlay = False

    def ejes(ch, ymax=None, ytitle=None):
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E3E3"))
        ch.x_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=txt_props(900)), endParaRPr=txt_props(900))])
        ch.y_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=txt_props(900)), endParaRPr=txt_props(900))])
        ch.y_axis.scaling.min = 0
        if ymax:
            ch.y_axis.scaling.max = ymax
        if ytitle:
            ch.y_axis.title = ytitle

    def etiquetas(obj, formato='0.0;-0.0;;'):
        dl = DataLabelList()
        dl.showVal = True
        dl.showSerName = dl.showCatName = dl.showLegendKey = dl.showPercent = False
        dl.numFmt = formato
        dl.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=txt_props(900, "404040", True)), endParaRPr=txt_props(900))])
        obj.dLbls = dl

    def barras(titulo_txt, col_val, color, formato='0.0;-0.0;;', ytitle=None):
        ch = BarChart()
        ch.type = "col"
        ch.gapWidth = 60
        ch.add_data(Reference(wc, min_col=col_val, min_row=1, max_row=1 + nc), titles_from_data=True)
        ch.set_categories(Reference(wc, min_col=1, min_row=2, max_row=1 + nc))
        ch.series[0].graphicalProperties = GraphicalProperties(solidFill=color)
        ch.series[0].graphicalProperties.line.solidFill = color
        etiquetas(ch.series[0], formato)
        ejes(ch, None, ytitle)
        titulo(ch, titulo_txt)
        ch.legend = None
        ch.width, ch.height = 13.8, 7.6
        return ch

    wd.add_chart(barras("Visitas realizadas por coordinador", 2, "2E75B6", '0;-0;;'), "B10")
    wd.add_chart(barras("Media de visitas por coordinador y día", 3, "1F4E78"), "H10")

    # Evolución diaria: visitas (barras) y duración media (línea)
    ev = BarChart()
    ev.type = "col"
    ev.gapWidth = 40
    ev.add_data(Reference(wc, min_col=10, min_row=1, max_row=1 + nf), titles_from_data=True)
    ev.set_categories(Reference(wc, min_col=9, min_row=2, max_row=1 + nf))
    ev.series[0].graphicalProperties = GraphicalProperties(solidFill="9DC3E6")
    ev.series[0].graphicalProperties.line.solidFill = "9DC3E6"
    ejes(ev, None, "Visitas")
    ln = LineChart()
    ln.add_data(Reference(wc, min_col=11, min_row=1, max_row=1 + nf), titles_from_data=True)
    ln.series[0].graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="548235", w=28575))
    ln.series[0].marker.symbol = "circle"
    ln.series[0].marker.size = 6
    ln.series[0].marker.graphicalProperties = GraphicalProperties(solidFill="548235")
    ln.series[0].smooth = False
    ln.y_axis.axId = 200
    ln.y_axis.scaling.min = 0
    ln.y_axis.crosses = "max"
    ln.y_axis.delete = False
    ln.y_axis.majorGridlines = None
    ln.y_axis.title = "Duración media (min)"
    ln.y_axis.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=txt_props(900)), endParaRPr=txt_props(900))])
    ev += ln
    ev.legend.position = 'b'
    titulo(ev, "Evolución diaria: visitas y duración media")
    ev.width, ev.height = 13.8, 7.6
    wd.add_chart(ev, "B26")

    wd.add_chart(barras("Duración media por coordinador (min)", 4, "548235"), "H26")
    wd.add_chart(barras(f"Visitas de menos de {U} min (% de las visitas del coordinador)", 5, "C00000",
                        '0.0"%";-0.0"%";;'), "B42")
    wd.add_chart(barras("Fallos del checklist por visita (media)", 6, "ED7D31"), "H42")

    wd.print_area = "A1:N58"
    preparar_impresion(wd, 'portrait')
    wd.page_setup.fitToHeight = 1

    wb.active = 0
    for hoja in wb.worksheets:
        hoja.sheet_view.tabSelected = (hoja.title == "Dashboard")

    wb.save(archivo)
    _ignorar_errores_texto(archivo, wb.sheetnames.index('Datos') + 1, f"{L('Latitud')}2:{L('Longitud')}{n}")


def procesar_pdf(ruta, base=None):
    """Lee un PDF y devuelve (registro, avisos, direccion_ok). registro es None si no se pudo leer."""
    datos = extraer_datos(ruta)
    if not datos:
        return None, [], False
    avisos = datos.pop('_avisos', [])
    direccion_ok = False
    if base is not None:
        direccion, aviso_dir = buscar_direccion(base, datos.get('Estanco'), datos.get('Provincia'))
        datos['Direccion'] = direccion
        direccion_ok = direccion not in ("No encontrada", "Sin dirección en la base")
        if aviso_dir:
            avisos.append(aviso_dir)
    return calcular_metricas(datos), avisos, direccion_ok


def generar_excel(lista, archivo_salida, usar_formulas=True):
    """Construye el Excel completo (Datos, Resumen y Dashboard) a partir de los registros de las visitas."""
    df = pd.DataFrame(lista)
    df['Fecha_dt'] = pd.to_datetime(df['Fecha'], format='%d/%m/%Y', errors='coerce')
    df = df.sort_values(by=['Fecha_dt', 'Coordinador'], ascending=[True, True]).reset_index(drop=True)
    df = df.drop(columns=['Fecha_dt'])
    cols_generales = ['Fecha', 'Provincia', 'Estanco', 'Direccion', 'Tipo_Visita', 'Coordinador', 'Promotor',
                      'Latitud', 'Longitud', 'Hora_Inicio', 'Hora_Envio', 'Duracion_min',
                      'Nota_Visita_Coordinador', 'Total_Si', 'Total_No']
    cols_preguntas = [f'P{i}' for i in range(1, 19)]
    df = df[[x for x in cols_generales + cols_preguntas if x in df.columns]]
    df.to_excel(archivo_salida, sheet_name='Datos', index=False)
    crear_dashboard(archivo_salida, df, usar_formulas)
    return df


# --- INTERFAZ GRÁFICA ---
COLOR_AZUL = "#1F4E78"
CONFIG_PATH = os.path.join(os.path.expanduser("~"), ".auditor_visitas.json")


def cargar_config():
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def guardar_config(cfg):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f)
    except Exception:
        pass


class AppAuditoria:
    def __init__(self, root):
        self.root = root
        self.root.title("Auditor de Visitas - Ploom 2026")
        self.root.geometry("760x690")
        self.root.minsize(640, 600)
        self.root.configure(bg="#F3F5F8")

        self.config = cargar_config()
        self.carpeta_seleccionada = ""
        self.archivo_excel = ""
        self.cola = queue.Queue()
        self.cancelar = threading.Event()
        self.trabajando = False

        self._estilos()
        self._construir()
        self.root.protocol("WM_DELETE_WINDOW", self.cerrar)

        ultima = self.config.get("ultima_carpeta", "")
        if ultima and os.path.isdir(ultima):
            self._fijar_carpeta(ultima)

    # ---------- construcción ----------
    def _estilos(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("Primary.TButton", font=("Segoe UI", 10, "bold"), foreground="white",
                     background="#2196F3", padding=8, borderwidth=0)
        st.map("Primary.TButton", background=[("active", "#1976D2"), ("disabled", "#B0BEC5")])
        st.configure("Ok.TButton", font=("Segoe UI", 10, "bold"), foreground="white",
                     background="#4CAF50", padding=8, borderwidth=0)
        st.map("Ok.TButton", background=[("active", "#388E3C"), ("disabled", "#B0BEC5")])
        st.configure("Warn.TButton", font=("Segoe UI", 10, "bold"), foreground="white",
                     background="#FF9800", padding=8, borderwidth=0)
        st.map("Warn.TButton", background=[("active", "#F57C00"), ("disabled", "#B0BEC5")])
        st.configure("Cancel.TButton", font=("Segoe UI", 10, "bold"), foreground="white",
                     background="#E53935", padding=8, borderwidth=0)
        st.map("Cancel.TButton", background=[("active", "#C62828"), ("disabled", "#B0BEC5")])
        st.configure("Azul.Horizontal.TProgressbar", troughcolor="#DDE3EA",
                     background=COLOR_AZUL, thickness=16)

    def _construir(self):
        cab = tk.Frame(self.root, bg=COLOR_AZUL)
        cab.pack(fill="x")
        tk.Label(cab, text="📊  AUDITOR DE INFORMES DE VISITA", bg=COLOR_AZUL, fg="white",
                 font=("Segoe UI", 15, "bold")).pack(anchor="w", padx=20, pady=(14, 0))
        tk.Label(cab, text="Campaña Ploom 2026 · Extracción, detección de incoherencias y dashboard en Excel",
                 bg=COLOR_AZUL, fg="#BBD2EA", font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=(0, 14))

        cuerpo = tk.Frame(self.root, bg="#F3F5F8")
        cuerpo.pack(fill="both", expand=True, padx=20, pady=14)

        # Paso 1
        f1 = tk.LabelFrame(cuerpo, text=" 1 · Carpeta con los PDFs ", bg="#F3F5F8",
                           font=("Segoe UI", 9, "bold"), fg=COLOR_AZUL, padx=10, pady=8)
        f1.pack(fill="x")
        self.btn_carpeta = ttk.Button(f1, text="Seleccionar carpeta…", style="Ok.TButton",
                                      command=self.seleccionar_carpeta)
        self.btn_carpeta.pack(side="left")
        self.lbl_ruta = tk.Label(f1, text="Ninguna carpeta seleccionada", fg="gray", bg="#F3F5F8",
                                 anchor="w", justify="left", wraplength=480, font=("Segoe UI", 9))
        self.lbl_ruta.pack(side="left", padx=12, fill="x", expand=True)

        # Base de datos de estancos (opcional)
        fb = tk.LabelFrame(cuerpo, text=" Base de datos de estancos (opcional: añade la dirección) ", bg="#F3F5F8",
                           font=("Segoe UI", 9, "bold"), fg=COLOR_AZUL, padx=10, pady=8)
        fb.pack(fill="x", pady=(10, 0))
        self.btn_bbdd = ttk.Button(fb, text="Seleccionar Excel…", style="Primary.TButton",
                                   command=self.seleccionar_bbdd)
        self.btn_bbdd.pack(side="left")
        self.btn_quitar_bbdd = ttk.Button(fb, text="✖", width=3, command=self.quitar_bbdd)
        self.btn_quitar_bbdd.pack(side="right")
        self.lbl_bbdd = tk.Label(fb, text="", bg="#F3F5F8", anchor="w", justify="left",
                                 wraplength=420, font=("Segoe UI", 9))
        self.lbl_bbdd.pack(side="left", padx=12, fill="x", expand=True)

        # Paso 2
        f2 = tk.LabelFrame(cuerpo, text=" 2 · Procesar ", bg="#F3F5F8",
                           font=("Segoe UI", 9, "bold"), fg=COLOR_AZUL, padx=10, pady=8)
        f2.pack(fill="x", pady=(10, 0))
        fila = tk.Frame(f2, bg="#F3F5F8")
        fila.pack(fill="x")
        self.btn_procesar = ttk.Button(fila, text="Procesar y generar Excel", style="Primary.TButton",
                                       command=self.ejecutar_proceso, state="disabled")
        self.btn_procesar.pack(side="left")
        self.btn_cancelar = ttk.Button(fila, text="Cancelar", style="Cancel.TButton",
                                       command=self.cancelar_proceso, state="disabled")
        self.btn_cancelar.pack(side="left", padx=8)
        self.lbl_estado = tk.Label(fila, text="Listo", bg="#F3F5F8", fg="#455A64",
                                   font=("Segoe UI", 9), anchor="e")
        self.lbl_estado.pack(side="right")

        self.var_formulas = tk.BooleanVar(value=self.config.get("formulas", True))
        tk.Checkbutton(f2, text="Excel con fórmulas (se recalcula y el Resumen sigue los filtros)", variable=self.var_formulas,
                       bg="#F3F5F8", activebackground="#F3F5F8", font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))
        tk.Label(f2, text="Sin fórmulas el Resumen y el Dashboard no siguen los filtros. Úsalo solo para abrir en el móvil o en Vista protegida.",
                 bg="#F3F5F8", fg="gray", font=("Segoe UI", 8)).pack(anchor="w")

        self.progress = ttk.Progressbar(f2, orient="horizontal", mode="determinate",
                                        style="Azul.Horizontal.TProgressbar")
        self.progress.pack(fill="x", pady=(10, 0))

        # Registro
        f3 = tk.LabelFrame(cuerpo, text=" Registro ", bg="#F3F5F8",
                           font=("Segoe UI", 9, "bold"), fg=COLOR_AZUL, padx=6, pady=6)
        f3.pack(fill="both", expand=True, pady=(10, 0))
        self.txt_log = tk.Text(f3, height=10, wrap="word", state="disabled", bg="white",
                               font=("Consolas", 9), relief="flat")
        sb = ttk.Scrollbar(f3, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt_log.pack(side="left", fill="both", expand=True)
        self.txt_log.tag_config("ok", foreground="#2E7D32")
        self.txt_log.tag_config("warn", foreground="#E65100")
        self.txt_log.tag_config("error", foreground="#C62828")
        self.txt_log.tag_config("info", foreground="#37474F")

        # Resultado
        f4 = tk.Frame(cuerpo, bg="#F3F5F8")
        f4.pack(fill="x", pady=(10, 0))
        self.btn_abrir = ttk.Button(f4, text="📂 Abrir Excel generado", style="Warn.TButton",
                                    command=self.abrir_excel, state="disabled")
        self.btn_abrir.pack(side="left")
        self.btn_abrir_carpeta = ttk.Button(f4, text="Abrir carpeta", style="Primary.TButton",
                                            command=self.abrir_carpeta, state="disabled")
        self.btn_abrir_carpeta.pack(side="left", padx=8)
        self._mostrar_bbdd()

    # ---------- utilidades ----------
    def log(self, texto, tag="info"):
        self.txt_log.config(state="normal")
        self.txt_log.insert("end", texto + "\n", tag)
        self.txt_log.see("end")
        self.txt_log.config(state="disabled")

    def _mostrar_bbdd(self):
        ruta = self.config.get("bbdd", "")
        if ruta and os.path.isfile(ruta):
            self.lbl_bbdd.config(text=f"📋 {os.path.basename(ruta)}", fg="black")
        else:
            self.lbl_bbdd.config(text="Sin base de datos: no se añadirá la columna Dirección", fg="gray")

    def seleccionar_bbdd(self):
        ruta = filedialog.askopenfilename(title="Selecciona el Excel con los estancos (Nombre y Dirección)",
                                          filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if ruta:
            self.config["bbdd"] = ruta
            guardar_config(self.config)
            self._mostrar_bbdd()

    def quitar_bbdd(self):
        self.config["bbdd"] = ""
        guardar_config(self.config)
        self._mostrar_bbdd()

    def _fijar_carpeta(self, carpeta):
        self.carpeta_seleccionada = carpeta
        self.lbl_ruta.config(text=f"📁 {carpeta}", fg="black")
        self.btn_procesar.config(state="normal")
        self.config["ultima_carpeta"] = carpeta
        guardar_config(self.config)

    def seleccionar_carpeta(self):
        carpeta = filedialog.askdirectory(title="Selecciona la carpeta donde están los PDFs",
                                          initialdir=self.config.get("ultima_carpeta") or None)
        if carpeta:
            self._fijar_carpeta(carpeta)

    def _bloquear(self, trabajando):
        self.trabajando = trabajando
        self.btn_carpeta.config(state="disabled" if trabajando else "normal")
        self.btn_procesar.config(state="disabled" if trabajando or not self.carpeta_seleccionada else "normal")
        self.btn_cancelar.config(state="normal" if trabajando else "disabled")
        self.btn_bbdd.config(state="disabled" if trabajando else "normal")
        self.btn_quitar_bbdd.config(state="disabled" if trabajando else "normal")

    # ---------- proceso ----------
    def ejecutar_proceso(self):
        self.cancelar.clear()
        self._formulas_run = bool(self.var_formulas.get())
        self.config["formulas"] = self._formulas_run
        guardar_config(self.config)
        self.txt_log.config(state="normal"); self.txt_log.delete("1.0", "end"); self.txt_log.config(state="disabled")
        self.btn_abrir.config(state="disabled")
        self.btn_abrir_carpeta.config(state="disabled")
        self.progress.config(mode="determinate", value=0)
        self._bloquear(True)
        threading.Thread(target=self.procesar_datos, daemon=True).start()
        self.root.after(80, self._leer_cola)

    def cancelar_proceso(self):
        self.cancelar.set()
        self.lbl_estado.config(text="Cancelando…")

    def _leer_cola(self):
        """Todas las actualizaciones de la ventana pasan por aquí (hilo principal)."""
        try:
            while True:
                msg = self.cola.get_nowait()
                tipo = msg[0]
                if tipo == "log":
                    self.log(msg[1], msg[2])
                elif tipo == "progreso":
                    _, i, total, nombre = msg
                    self.progress.config(maximum=total, value=i)
                    self.lbl_estado.config(text=f"{i}/{total} · {nombre}")
                elif tipo == "generando":
                    self.progress.config(mode="indeterminate"); self.progress.start(12)
                    self.lbl_estado.config(text="Generando Excel…")
                elif tipo == "fin":
                    self._finalizar(*msg[1:])
                    return
        except queue.Empty:
            pass
        self.root.after(80, self._leer_cola)

    def procesar_datos(self):
        c = self.cola
        try:
            archivos_pdf = sorted(f for f in glob.glob(os.path.join(self.carpeta_seleccionada, "*"))
                                  if f.lower().endswith(".pdf"))
            if not archivos_pdf:
                c.put(("fin", False, 0, 0, 0, "No se encontraron archivos PDF en la carpeta."))
                return
            c.put(("log", f"Se encontraron {len(archivos_pdf)} PDFs.", "info"))

            base = None
            ruta_bbdd = self.config.get("bbdd", "")
            if not (ruta_bbdd and os.path.isfile(ruta_bbdd)):
                candidatos = sorted(glob.glob(os.path.join(self.carpeta_seleccionada, "BBDD*.xls*")))
                ruta_bbdd = candidatos[0] if candidatos else ""
            if ruta_bbdd:
                try:
                    c.put(("log", "Cargando la base de datos de estancos…", "info"))
                    base = cargar_base_estancos(ruta_bbdd)
                    c.put(("log", f"✔ Base cargada: {base['n']} estancos ({os.path.basename(ruta_bbdd)})", "ok"))
                except Exception as e:
                    c.put(("log", f"⚠ No se pudo leer la base de datos ({e}). Se continúa sin direcciones.", "warn"))
            else:
                c.put(("log", "Sin base de datos de estancos: no se añadirá la dirección.", "info"))
            con_direccion = 0

            lista, con_avisos, fallidos = [], 0, 0
            for i, archivo in enumerate(archivos_pdf, 1):
                if self.cancelar.is_set():
                    c.put(("fin", False, 0, 0, 0, "Proceso cancelado."))
                    return
                nombre = os.path.basename(archivo)
                c.put(("progreso", i, len(archivos_pdf), nombre))
                registro, avisos, direccion_ok = procesar_pdf(archivo, base)
                if registro is None:
                    fallidos += 1
                    c.put(("log", f"✖ {nombre}: no se pudo leer", "error"))
                    continue
                con_direccion += 1 if direccion_ok else 0
                lista.append(registro)
                if avisos:
                    con_avisos += 1
                    c.put(("log", f"⚠ {nombre}: {', '.join(avisos)}", "warn"))

            if not lista:
                c.put(("fin", False, 0, fallidos, 0, "No se pudo extraer datos de ningún PDF."))
                return

            if base is not None:
                c.put(("log", f"Direcciones encontradas: {con_direccion} de {len(lista)} visitas.",
                       "ok" if con_direccion == len(lista) else "warn"))
            c.put(("generando",))
            archivo_salida = os.path.join(self.carpeta_seleccionada, "Auditoria_Visitas_Total.xlsx")
            if os.path.exists(archivo_salida):
                try:
                    with open(archivo_salida, 'a'):
                        pass
                except PermissionError:
                    c.put(("fin", False, 0, fallidos, 0,
                           "El archivo Excel está abierto. Ciérralo y vuelve a intentarlo."))
                    return

            generar_excel(lista, archivo_salida, self._formulas_run)
            self.archivo_excel = archivo_salida
            c.put(("fin", True, len(lista), fallidos, con_avisos, ""))
        except Exception as e:
            c.put(("log", f"Error crítico: {e}", "error"))
            c.put(("fin", False, 0, 0, 0, f"Ocurrió un error inesperado:\n{e}"))

    def _finalizar(self, exito, n, fallidos, con_avisos, mensaje):
        self.progress.stop()
        self.progress.config(mode="determinate", value=self.progress["maximum"] if exito else 0)
        self._bloquear(False)
        if exito:
            self.lbl_estado.config(text="Completado")
            self.log(f"✔ {n} visitas procesadas · {con_avisos} con avisos · {fallidos} ilegibles", "ok")
            self.log(f"✔ Excel guardado: {self.archivo_excel}", "ok")
            self.btn_abrir.config(state="normal")
            self.btn_abrir_carpeta.config(state="normal")
            extra = ""
            if con_avisos or fallidos:
                extra = f"\n\nRevisa el registro: {con_avisos} PDFs con avisos y {fallidos} ilegibles."
            messagebox.showinfo("¡Éxito!", f"Se procesaron {n} visitas.{extra}")
        else:
            self.lbl_estado.config(text="Detenido")
            self.log(f"✖ {mensaje}", "error")
            if "cancelado" not in mensaje.lower():
                messagebox.showerror("Error", mensaje)

    # ---------- abrir ----------
    def _abrir(self, ruta):
        if not ruta or not os.path.exists(ruta):
            messagebox.showwarning("No encontrado", f"No existe el archivo o carpeta:\n{ruta}")
            return
        try:
            if platform.system() == 'Windows':
                os.startfile(os.path.normpath(ruta))
            elif platform.system() == 'Darwin':
                subprocess.call(['open', ruta])
            else:
                subprocess.call(['xdg-open', ruta])
        except Exception as e:
            messagebox.showerror("No se pudo abrir", f"{e}\n\nAbre manualmente:\n{ruta}")

    def abrir_excel(self):
        self._abrir(self.archivo_excel)

    def abrir_carpeta(self):
        self._abrir(self.carpeta_seleccionada)

    def cerrar(self):
        if self.trabajando and not messagebox.askyesno("Salir", "Hay un proceso en marcha. ¿Cerrar de todos modos?"):
            return
        self.cancelar.set()
        self.root.destroy()


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = AppAuditoria(root)
        root.mainloop()
    except Exception:
        import traceback
        err = traceback.format_exc()
        log_path = os.path.join(os.path.expanduser("~"), "auditor_visitas_error.txt")
        try:
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(err)
        except Exception:
            pass
        try:
            r = tk.Tk(); r.withdraw()
            messagebox.showerror("Error al iniciar", f"{err}\n\nGuardado en: {log_path}")
        except Exception:
            print(err)
            input("Pulsa Enter para salir...")
