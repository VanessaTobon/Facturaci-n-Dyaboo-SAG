import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import os
from datetime import datetime, date
import zipfile

st.set_page_config(
    page_title="Facturación Dyaboo",
    page_icon="🛍️",
    layout="wide"
)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def clean_nit(val):
    if val is None:
        return None
    s = str(val).strip()
    s = re.sub(r'^CC\s*', '', s, flags=re.IGNORECASE)
    s = s.replace('.', '').replace(',', '').replace(' ', '')
    try:
        s = str(int(float(s)))
    except Exception:
        pass
    return s if s.isdigit() else s

def clean_phone(val):
    if val is None or str(val).strip() == '':
        return ''
    try:
        return str(int(float(str(val))))
    except Exception:
        return str(val).strip()

def title_case_address(s):
    if not s:
        return ''
    return str(s).title()

def to_upper(s):
    if not s:
        return ''
    return str(s).upper().strip()

def parse_color_talla_from_name(product_name):
    if not product_name:
        return '', ''
    parts = str(product_name).split(' / ')
    talla = parts[-1].strip() if len(parts) > 1 else ''
    before_slash = parts[0] if parts else str(product_name)
    dash_parts = before_slash.split(' - ')
    color = dash_parts[-1].strip() if len(dash_parts) > 1 else ''
    return color, talla

def parse_ref_from_sku(sku):
    if not sku or str(sku).strip() in ('', '0', 'nan'):
        return '', '', ''
    sku = str(sku).strip()
    tallas = ['XXL', 'XL', 'XS', 'U', 'S', 'M', 'L',
              '6', '8', '10', '12', '14', '16', '18', '20', '22',
              '6X', '8X', '10X', '12X', '14X', '16X']
    for t in sorted(tallas, key=len, reverse=True):
        pattern = rf'^(.*?)({re.escape(t)})(\w+)$'
        m = re.match(pattern, sku)
        if m:
            return m.group(1), m.group(2), m.group(3)
    return sku[:-2], '', sku[-2:]

def build_cod_articulo(ref):
    if not ref:
        return ''
    ref = str(ref).strip()
    if ref and ref[0].isdigit():
        return '20' + ref + 'D'
    return ref + 'D'

def get_vendedor(source, is_first_item):
    if is_first_item and str(source).lower() == 'web':
        return '202'
    return '302'

def get_flete_code(flete_val):
    try:
        v = float(flete_val)
    except Exception:
        return None, None
    if v == 15000:
        return '15', round(15000 / 1.19, 8)
    elif v == 10000:
        return '10', round(10000 / 1.19, 8)
    elif v > 0:
        return 'OTRO', v
    return None, None

# ─────────────────────────────────────────────
# LOAD REFERENCE DATA
# ─────────────────────────────────────────────

@st.cache_data(show_spinner=False)
def load_reference_data(plantilla_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(plantilla_bytes), read_only=True, data_only=True)

    ws = wb['Ciudades']
    ciudades = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        ciudad_raw = row[6]
        zona = row[8]
        if ciudad_raw:
            ciudades[str(ciudad_raw).upper().strip()] = str(zona).strip() if zona else ''

    ws = wb['Tabla de Colores']
    color_map = {}
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        name = row[0]
        code = row[1]
        if name:
            color_map[str(name).upper().strip()] = str(code).strip() if code else ''

    ws = wb['Colores']
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        name = row[1]
        code = row[2]
        if name and str(name).upper().strip() not in color_map:
            color_map[str(name).upper().strip()] = str(code).strip() if code else ''

    ws = wb['Tallas']
    tallas_validas = set()
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if row[1]:
            tallas_validas.add(str(row[1]).strip())

    ws = wb['Terceros']
    terceros = set()
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            continue
        if row[1]:
            try:
                terceros.add(str(int(float(str(row[1])))).strip())
            except Exception:
                pass

    wb.close()
    return ciudades, color_map, tallas_validas, terceros

# ─────────────────────────────────────────────
# PROCESS ORDERS
# ─────────────────────────────────────────────

def process_orders(df, ciudades, color_map, tallas_validas, terceros, fecha_facturacion):
    errors = []
    orders_processed = []
    new_clients = {}

    order_groups = {}
    for _, row in df.iterrows():
        name = str(row.get('Name', '')).strip()
        if name not in order_groups:
            order_groups[name] = []
        order_groups[name].append(row)

    for order_name, items in order_groups.items():
        first_row = items[0]

        billing_company = first_row.get('Billing Company', '')
        nit_raw = str(billing_company).strip() if pd.notna(billing_company) else ''
        nit = clean_nit(nit_raw)

        items = [r for r in items if str(r.get('Lineitem sku', '')).strip() not in ('0', '', 'nan')
                 and 'envío' not in str(r.get('Lineitem name', '')).lower()
                 and 'envio' not in str(r.get('Lineitem name', '')).lower()]
        if not items:
            continue

        billing_city = str(first_row.get('Billing City', '')).strip()
        city_key = billing_city.upper().strip()

        if city_key.isdigit():
            errors.append({
                'Pedido': order_name, 'NIT': nit,
                'Campo': 'Ciudad', 'Valor': billing_city,
                'Error': 'La ciudad parece una cédula. ¿Cuál es la ciudad real?',
                'Corrección': ''
            })
            ciudad_sag = ''
            zona_sag = ''
        else:
            found_zona = ciudades.get(city_key, None)
            if found_zona is None:
                for k, v in ciudades.items():
                    if city_key in k or k in city_key:
                        found_zona = v
                        city_key = k
                        break
            if found_zona is None:
                errors.append({
                    'Pedido': order_name, 'NIT': nit,
                    'Campo': 'Ciudad', 'Valor': billing_city,
                    'Error': f'Ciudad "{billing_city}" no encontrada en tabla SAG.',
                    'Corrección': ''
                })
                ciudad_sag = city_key
                zona_sag = ''
            else:
                ciudad_sag = city_key
                zona_sag = found_zona

        billing_name = str(first_row.get('Billing Name', '')).strip()
        email = str(first_row.get('Email', '')).strip() if pd.notna(first_row.get('Email')) else '(en blanco)'
        phone = clean_phone(first_row.get('Billing Phone'))
        address = title_case_address(first_row.get('Billing Street', ''))
        source = str(first_row.get('Source', '')).strip()
        payment_method = str(first_row.get('Payment Method', '')).strip()

        try:
            flete_val = float(str(first_row.get('Shipping', 0)).strip())
        except Exception:
            flete_val = 0

        if nit and nit not in terceros and nit not in new_clients:
            parts = billing_name.split()
            ap1 = ap2 = n1 = n2 = ''
            if len(parts) >= 4:
                ap1, ap2, n1, n2 = parts[0], parts[1], parts[2], parts[3]
            elif len(parts) == 3:
                ap1, ap2, n1 = parts[0], parts[1], parts[2]
            elif len(parts) == 2:
                ap1, n1 = parts[0], parts[1]
            elif len(parts) == 1:
                n1 = parts[0]

            new_clients[nit] = {
                'NIT': nit,
                'Nombre Shopify': billing_name,
                'Nombre SAG (Validado)': '',
                'Ciudad': billing_city,
                'Departamento': str(first_row.get('Billing Province Name', '')).strip(),
                'Dirección': address,
                'Teléfono': phone,
                'Email': email if email else '(en blanco)',
                'Zona': zona_sag,
                'Ciudad SAG': ciudad_sag,
                'Apellido1': to_upper(ap1),
                'Apellido2': to_upper(ap2),
                'Nombre1': to_upper(n1),
                'Nombre2': to_upper(n2),
                'Observaciones': '',
                '_excluido': False
            }

        flete_asignado = False
        for idx, item in enumerate(items):
            is_first = idx == 0
            sku = str(item.get('Lineitem sku', '')).strip()
            product_name = str(item.get('Lineitem name', '')).strip()

            color_str, talla_str = parse_color_talla_from_name(product_name)
            ref, talla_code_from_sku, color_code_from_sku = parse_ref_from_sku(sku)

            talla_final = talla_str if talla_str else talla_code_from_sku

            color_key = color_str.upper().strip()
            color_code = color_map.get(color_key, '')
            if not color_code:
                color_code = color_code_from_sku
                if not color_code:
                    errors.append({
                        'Pedido': order_name, 'NIT': nit,
                        'Campo': 'Color', 'Valor': color_str,
                        'Error': f'Color "{color_str}" no tiene código. SKU: {sku}',
                        'Corrección': ''
                    })

            cod_articulo = build_cod_articulo(ref)

            try:
                valor = float(str(item.get('Lineitem price', 0)).strip())
            except Exception:
                valor = 0
            try:
                base_valor = float(str(item.get('Lineitem compare at price', 0)).strip())
            except Exception:
                base_valor = valor

            valor_sin_iva = round(valor / 1.19, 8)
            base_sin_iva = round(base_valor / 1.19, 8) if base_valor else valor_sin_iva
            descuento = round(base_sin_iva - valor_sin_iva, 8) if base_sin_iva > valor_sin_iva else 0

            flete_code = None
            flete_valor_sin_iva = 0
            if is_first and flete_val > 0 and not flete_asignado:
                fc, fv = get_flete_code(flete_val)
                if fc == 'OTRO':
                    errors.append({
                        'Pedido': order_name, 'NIT': nit,
                        'Campo': 'Flete', 'Valor': flete_val,
                        'Error': f'Flete ${flete_val:,.0f} no es $10.000 ni $15.000. ¿Qué código SAG usar?',
                        'Corrección': ''
                    })
                flete_code = fc
                flete_valor_sin_iva = fv or 0
                flete_asignado = True

            vendedor = get_vendedor(source, is_first)
            observacion = f"Pedido Dyaboo Online {order_name} Método de Pago: {payment_method}"

            orders_processed.append({
                '_order_name': order_name,
                '_nit': nit,
                'CÓD. FUENTE': 'PW',
                'NÚMERO': '',
                'FECHA': fecha_facturacion.strftime('%d/%m/%Y'),
                'NIT': nit,
                'OBSERVACIONES': observacion,
                'COD.ARTICULO': cod_articulo,
                'COD. BODEGA': '09',
                'CANTIDAD': int(item.get('Lineitem quantity', 1)),
                'VALOR': valor_sin_iva,
                'IVA': 19,
                'DSCTO': round(descuento, 2) if descuento else 0,
                'NIT VENDEDOR': vendedor,
                'TALLA': talla_final,
                'COLOR': color_code,
                '_flete_code': flete_code,
                '_flete_valor_sin_iva': flete_valor_sin_iva,
                '_color_str': color_str,
                '_sku': sku,
            })

    return orders_processed, errors, new_clients


def assign_consecutivos(orders_processed, excluded_nits, start_num):
    order_map = {}
    counter = start_num
    result = []
    for line in orders_processed:
        nit = line['_nit']
        order = line['_order_name']
        if nit in excluded_nits:
            continue
        if order not in order_map:
            order_map[order] = counter
            counter += 1
        line = dict(line)
        line['NÚMERO'] = order_map[order]
        result.append(line)
    return result, counter - 1

# ─────────────────────────────────────────────
# BUILD OUTPUT FILES
# ─────────────────────────────────────────────

def build_clientes_a_revisar(new_clients):
    rows = []
    for nit, c in new_clients.items():
        rows.append({
            'Cédula': nit,
            'Nombre Shopify': c['Nombre Shopify'],
            'Departamento': c['Departamento'],
            'Ciudad': c['Ciudad'],
            'Dirección': c['Dirección'],
            'Teléfono': c['Teléfono'],
            'Nombre SAG (Validado)': c['Nombre SAG (Validado)'],
            'Observaciones': c['Observaciones'],
        })
    return pd.DataFrame(rows)


def build_clientes_a_crear(new_clients, excluded_nits, fecha):
    rows = []
    for nit, c in new_clients.items():
        if nit in excluded_nits:
            continue
        nombre_sag = c['Nombre SAG (Validado)'] or to_upper(c['Nombre Shopify'])
        rows.append({
            'NIT ': nit,
            ' DV ': '',
            ' CÓDIGO ALTERNO ': '',
            ' NOMBRE ': nombre_sag,
            ' NOMBRE ALTERNO ': '',
            ' TELÉFONO PPAL ': c['Teléfono'],
            ' TELÉFONO ALTERNO ': '',
            ' FAX ': '',
            ' DIRECCIÓN ': c['Dirección'],
            ' CIUDAD ': c['Ciudad SAG'],
            ' NATURALEZA ': 'Natural',
            ' TIPO DE TERCERO ': '',
            'E_MAIL ': c['Email'],
            ' PÁGINA WEB ': '',
            'ACT. COMERCIAL ': '05-VENTA ONLINE',
            ' FORMA DE PAGO NORMAL ': 'CONTADO',
            ' CUPO ': '',
            ' ACTIVO (S/N) ': 'S',
            ' AUTORRETENEDOR (S/N) ': 'N',
            ' OBSERVACIONES ': '',
            ' TIPO DE RETEFTE NORMAL ': 'VO - NINGUNA -CXC-',
            ' TIPO DE CONTACTO ': '',
            ' NOMBRE DEL CONTACTO ': '',
            ' PRECIO DE VENTA (0..4) ': 4,
            ' NIT VENDEDOR ': 302,
            ' ZONA ': c['Zona'],
            ' CLASE DE CLIENTE ': 'V',
            ' CODIGO INTERNO ': '',
            ' Tipo de Dcto: Nit(A)/Cédula(C)/Cédula de Extr. (E)/ Cédula del Tutor(T)/Número Único de Identificación Personal NUIP(U) ': 'C',
            ' IVA (S/N) ': 'S',
            ' Dia Cumpleaños(01, 02, 03,04, ...)': '',
            ' Mes Nacimiento (01,02,03..12) ': '',
            ' CATEGORIA ': '',
            ' DSCTO COMERCIAL ': '',
            ' DSCTO PP ': '',
            ' NOMBRE1 ': c['Nombre1'],
            ' NOMBRE2 ': c['Nombre2'],
            ' APELLIDO1 ': c['Apellido1'],
            ' APELLIDO2 ': c['Apellido2'],
            ' DIRECCION2 ': '',
            ' RUT  ': '',
        })
    return pd.DataFrame(rows)


def build_clientes_complemento(new_clients, excluded_nits, fecha):
    rows = []
    for nit, c in new_clients.items():
        if nit in excluded_nits:
            continue
        nombre_sag = c['Nombre SAG (Validado)'] or to_upper(c['Nombre Shopify'])
        rows.append({
            'Nit ': nit,
            ' Digito Verificacion ': '',
            ' Tipo de Documento(C:Cedula, A:Nit) ': 'C',
            ' Nombre ': nombre_sag,
            ' Nombre1 ': c['Nombre1'],
            ' Nombre2 ': c['Nombre2'],
            ' Apellido1 ': c['Apellido1'],
            ' Apellido2 ': c['Apellido2'],
            ' Dirección ': c['Dirección'],
            ' Ciudad ': c['Ciudad SAG'],
            ' Telefono ppal ': c['Teléfono'],
            '     Código postal ': '000000',
            ' Email1': c['Email'],
            ' Email2(FE) ': c['Email'],
            ' Genera FE(S/N) ': 'S',
            ' Quien Aprueba ': 'LA WEB',
            ' Fecha Aprueba ': fecha.strftime('%d/%m/%Y'),
            ' Responsabilidad fiscal': 'R-99-PN',
            ' Nit Vendedor ': 302,
            'Clase Regimen (N:persona natural, G:Gran Contribuyente, R:No Responsable de IVA, C:Responsable de IVA Régimen Común, F:Régimen SIMPLE)': 'N',
        })
    return pd.DataFrame(rows)


def build_pedidos_df(orders_final):
    rows = []
    for line in orders_final:
        rows.append({
            'CÓD. FUENTE ': line['CÓD. FUENTE'],
            ' NÚMERO ': line['NÚMERO'],
            ' FECHA ': line['FECHA'],
            ' NIT ': line['NIT'],
            ' OBSERVACIONES ': line['OBSERVACIONES'],
            ' COD.ARTICULO ': line['COD.ARTICULO'],
            ' COD. BODEGA ': line['COD. BODEGA'],
            ' CANTIDAD ': line['CANTIDAD'],
            ' VALOR ': line['VALOR'],
            ' IVA ': line['IVA'],
            ' DSCTO ': line['DSCTO'],
            ' NIT VENDEDOR ': line['NIT VENDEDOR'],
            'TALLA': line['TALLA'],
            'COLOR': line['COLOR'],
        })
        if line.get('_flete_code') and line['_flete_code'] not in (None, 'OTRO'):
            rows.append({
                'CÓD. FUENTE ': line['CÓD. FUENTE'],
                ' NÚMERO ': line['NÚMERO'],
                ' FECHA ': line['FECHA'],
                ' NIT ': line['NIT'],
                ' OBSERVACIONES ': line['OBSERVACIONES'],
                ' COD.ARTICULO ': line['_flete_code'],
                ' COD. BODEGA ': line['COD. BODEGA'],
                ' CANTIDAD ': 1,
                ' VALOR ': line['_flete_valor_sin_iva'],
                ' IVA ': 19,
                ' DSCTO ': 0,
                ' NIT VENDEDOR ': line['NIT VENDEDOR'],
                'TALLA': 'SURT',
                'COLOR': 'SURT',
            })
    return pd.DataFrame(rows)


def df_to_excel_bytes(df, sheet_name='Sheet1'):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return buf.getvalue()


def df_to_csv_bytes(df):
    return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')


# ─────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────

st.title("🛍️ Facturación Dyaboo — SAG")
st.caption("Convierte pedidos de Shopify al formato de carga SAG")

# ── Sidebar ──
with st.sidebar:
    st.header("⚙️ Configuración")
    st.subheader("Plantilla SAG")
    st.caption("Solo cárgala cuando actualices Terceros, Ciudades o Colores.")

    uploaded_plantilla = st.file_uploader("Subir Plantilla.xlsx", type=['xlsx'], key='plantilla_uploader')
    if uploaded_plantilla:
        st.session_state['plantilla_bytes'] = uploaded_plantilla.read()
        st.session_state['plantilla_date'] = datetime.now().strftime('%d/%m/%Y %H:%M')
        st.success("✅ Plantilla guardada.")

    if 'plantilla_bytes' in st.session_state:
        st.info(f"📋 Plantilla activa desde:\n{st.session_state.get('plantilla_date', '—')}")
    else:
        st.warning("Sin plantilla. Carga una para continuar.")

    st.divider()
    st.subheader("Parámetros del lote")
    consecutivo_inicio = st.number_input(
        "Consecutivo inicial (último lote + 1)",
        min_value=1, value=5388, step=1
    )
    fecha_facturacion = st.date_input("Fecha de facturación", value=date.today())

# ── Guard: need plantilla ──
if 'plantilla_bytes' not in st.session_state:
    st.info("👈 Primero carga la Plantilla SAG en el panel izquierdo.")
    st.stop()

# Load reference data
try:
    with st.spinner("Cargando tablas de referencia..."):
        ciudades, color_map, tallas_validas, terceros = load_reference_data(
            st.session_state['plantilla_bytes']
        )
except Exception as e:
    st.error(f"Error leyendo la plantilla: {e}")
    st.stop()

# ── Progress indicator ──
def step_badge(n, label, active=False, done=False):
    if done:
        return f"✅ **Paso {n} — {label}**"
    elif active:
        return f"🔵 **Paso {n} — {label}**"
    else:
        return f"⚪ Paso {n} — {label}"

# ════════════════════════════════════════
# PASO 1 — CARGAR PEDIDOS
# ════════════════════════════════════════
st.markdown("---")
st.subheader("📂 Paso 1 — Cargar pedidos de Shopify")

uploaded_orders = st.file_uploader(
    "Subir export de Shopify (.xlsx o .csv)",
    type=['xlsx', 'csv'],
    key='orders_uploader'
)

if not uploaded_orders:
    st.info("Sube el archivo de pedidos exportado desde Shopify para continuar.")
    st.stop()

try:
    if uploaded_orders.name.endswith('.csv'):
        df = pd.read_csv(uploaded_orders)
    else:
        df = pd.read_excel(uploaded_orders)
    df = df.where(pd.notna(df), None)
except Exception as e:
    st.error(f"Error leyendo el archivo de pedidos: {e}")
    st.stop()

num_pedidos = df['Name'].nunique() if 'Name' in df.columns else '?'
st.success(f"✅ Archivo cargado — **{num_pedidos} pedidos**, {len(df)} líneas")

# Process
with st.spinner("Analizando pedidos..."):
    orders_processed, errors, new_clients = process_orders(
        df, ciudades, color_map, tallas_validas, terceros, fecha_facturacion
    )

# ════════════════════════════════════════
# PASO 2 — REVISIÓN DE ERRORES
# ════════════════════════════════════════
st.markdown("---")
st.subheader("🔍 Paso 2 — Revisión de errores")

errores_pendientes = False
if errors:
    st.warning(f"⚠️ Se encontraron **{len(errors)} errores** que necesitan atención:")
    err_df = pd.DataFrame(errors)
    edited = st.data_editor(
        err_df,
        use_container_width=True,
        num_rows="fixed",
        column_config={"Corrección": st.column_config.TextColumn(width="large")}
    )
    errores_sin_corregir = edited[edited['Corrección'].isna() | (edited['Corrección'] == '')].shape[0]
    if errores_sin_corregir > 0:
        st.warning(f"Quedan **{errores_sin_corregir} errores sin corrección**. Puedes continuar, pero los pedidos con cédula incorrecta deben marcarse en el paso de clientes.")
    else:
        st.success("✅ Todos los errores tienen corrección anotada.")
else:
    st.success("✅ Sin errores detectados.")

# ════════════════════════════════════════
# PASO 3 — CLIENTES NUEVOS + VALIDACIÓN POLICÍA
# ════════════════════════════════════════
st.markdown("---")
st.subheader("👤 Paso 3 — Validación de clientes nuevos")

excluded_nits = set()
clientes_validados = False

if not new_clients:
    st.success("✅ Todos los clientes ya existen en el SAG. No hay clientes a crear.")
    clientes_validados = True
else:
    st.info(f"Se encontraron **{len(new_clients)} clientes nuevos** que deben validarse en la Policía Nacional antes de continuar.")

    # ── 3a: Descargar para revisar ──
    st.markdown("#### 3a. Descarga el archivo y valida las cédulas en la Policía")
    revisar_df = build_clientes_a_revisar(new_clients)

    col1, col2 = st.columns([2, 1])
    with col1:
        st.download_button(
            "⬇️ Descargar Clientes a Revisar",
            data=df_to_excel_bytes(revisar_df),
            file_name=f"Clientes_a_Revisar_{fecha_facturacion.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.caption("Llena la columna **Nombre SAG (Validado)** con el nombre correcto y escribe **INCORRECTO** en Observaciones para las cédulas que no pasen.")

    # ── 3b: Subir revisado ──
    st.markdown("#### 3b. Sube el archivo revisado")

    uploaded_revisado = st.file_uploader(
        "Subir Clientes_a_Revisar revisado (con nombres validados y observaciones)",
        type=['xlsx'],
        key='revisado_uploader'
    )

    if not uploaded_revisado:
        st.error("🚫 **Debes subir el archivo revisado de clientes antes de poder generar los pedidos.** No se puede facturar sin validar las cédulas en la Policía.")
        st.stop()

    # Process revisado
    rev_df = pd.read_excel(uploaded_revisado)
    incorrectos = []
    validos = []

    for _, row in rev_df.iterrows():
        nit = str(row.get('Cédula', '')).strip()
        obs = str(row.get('Observaciones', '')).strip().upper()
        nombre_validado = str(row.get('Nombre SAG (Validado)', '')).strip()

        if nit in new_clients:
            if 'INCORRECTO' in obs:
                excluded_nits.add(nit)
                new_clients[nit]['_excluido'] = True
                incorrectos.append(nit)
            else:
                if nombre_validado and nombre_validado.lower() != 'nan':
                    new_clients[nit]['Nombre SAG (Validado)'] = nombre_validado
                new_clients[nit]['Observaciones'] = obs
                validos.append(nit)

    # Summary
    col1, col2 = st.columns(2)
    with col1:
        if validos:
            st.success(f"✅ **{len(validos)} clientes válidos** para crear en SAG")
    with col2:
        if incorrectos:
            st.error(f"🚫 **{len(incorrectos)} clientes INCORRECTOS** — sus pedidos serán excluidos")
            for nit in incorrectos:
                nombre = new_clients[nit]['Nombre Shopify']
                st.caption(f"  • {nombre} ({nit})")

    # ── 3c: Confirmación policial obligatoria ──
    st.markdown("#### 3c. Confirmación de validación policial")

    st.warning("⚠️ **Importante:** Antes de continuar confirma que revisaste todas las cédulas en la Policía Nacional.")

    confirmacion = st.checkbox(
        "✅ Confirmo que validé todas las cédulas en la Policía Nacional y el archivo subido refleja los resultados de esa validación.",
        key='confirmacion_policia'
    )

    if not confirmacion:
        st.error("🚫 Debes confirmar la validación policial para poder generar los archivos finales.")
        st.stop()

    clientes_validados = True
    st.success("✅ Validación policial confirmada. Puedes continuar.")

# ════════════════════════════════════════
# PASO 4 — GENERAR ARCHIVOS
# ════════════════════════════════════════
st.markdown("---")
st.subheader("📦 Paso 4 — Generar archivos SAG")

if not clientes_validados:
    st.stop()

orders_final, ultimo_consecutivo = assign_consecutivos(
    orders_processed, excluded_nits, consecutivo_inicio
)

if not orders_final:
    st.warning("No hay pedidos válidos para generar. Todos fueron excluidos.")
    st.stop()

# Metrics
col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total pedidos", len(set(l['_order_name'] for l in orders_processed)))
with col2:
    st.metric("Pedidos válidos", len(set(l['_order_name'] for l in orders_final)))
with col3:
    st.metric("Pedidos excluidos", len(excluded_nits))
with col4:
    st.metric("Consecutivos", f"{consecutivo_inicio} → {ultimo_consecutivo}")

pedidos_df = build_pedidos_df(orders_final)

st.markdown("**Vista previa del plano de pedidos:**")
st.dataframe(pedidos_df.head(20), use_container_width=True)

# ── Build files ──
fecha_str = fecha_facturacion.strftime('%Y%m%d')
files_to_zip = {}

pedidos_xlsx = df_to_excel_bytes(pedidos_df, 'PlanoPedidoSAG_TallaColor')
pedidos_csv = df_to_csv_bytes(pedidos_df)
files_to_zip[f'Pedidos_a_Cargar_{fecha_str}.xlsx'] = pedidos_xlsx
files_to_zip[f'Pedidos_a_Cargar_{fecha_str}.csv'] = pedidos_csv

valid_new_clients = {k: v for k, v in new_clients.items() if k not in excluded_nits}
if valid_new_clients:
    crear_df = build_clientes_a_crear(new_clients, excluded_nits, fecha_facturacion)
    complemento_df = build_clientes_complemento(new_clients, excluded_nits, fecha_facturacion)
    crear_xlsx = df_to_excel_bytes(crear_df, 'Cliente02')
    complemento_xlsx = df_to_excel_bytes(complemento_df, 'Clientes04-Complemento')
    files_to_zip[f'Clientes_a_Crear_{fecha_str}.xlsx'] = crear_xlsx
    files_to_zip[f'Clientes_Complementario_{fecha_str}.xlsx'] = complemento_xlsx

# ── Downloads ──
st.markdown("### ⬇️ Descargar archivos")

col1, col2 = st.columns(2)
with col1:
    st.download_button("⬇️ Pedidos a Cargar (XLSX)", pedidos_xlsx,
                       f"Pedidos_a_Cargar_{fecha_str}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with col2:
    st.download_button("⬇️ Pedidos a Cargar (CSV)", pedidos_csv,
                       f"Pedidos_a_Cargar_{fecha_str}.csv", mime="text/csv")

if valid_new_clients:
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("⬇️ Clientes a Crear", crear_xlsx,
                           f"Clientes_a_Crear_{fecha_str}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with col2:
        st.download_button("⬇️ Clientes Complementario", complemento_xlsx,
                           f"Clientes_Complementario_{fecha_str}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()

zip_buf = io.BytesIO()
with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
    for fname, fbytes in files_to_zip.items():
        zf.writestr(fname, fbytes)

st.download_button(
    "📦 Descargar TODO en ZIP",
    data=zip_buf.getvalue(),
    file_name=f"Dyaboo_SAG_{fecha_str}.zip",
    mime="application/zip",
    type="primary",
    use_container_width=True
)

st.success(f"✅ Lote completado. Último consecutivo: **{ultimo_consecutivo}** — el próximo lote arranca en **{ultimo_consecutivo + 1}**")
