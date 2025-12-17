from sqlalchemy import create_engine
import pandas as pd
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
import os, json, time
from datetime import datetime, timedelta
import numpy as np
from dateutil.relativedelta import relativedelta

# ------------------------------------------------------------------------------
# 🔐 Google Sheets auth
# ------------------------------------------------------------------------------
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
cred_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(cred_dict, scope)
client = gspread.authorize(creds)

# ------------------------------------------------------------------------------
# 📦 PostgreSQL
# ------------------------------------------------------------------------------
usuario = os.environ.get("POSTGRES_USER", "inpro2021nubeuser")
contraseña = os.environ.get("POSTGRES_PASSWORD", "")
host = os.environ.get(
    "POSTGRES_HOST",
    "infraestructura-aurora-datawarehouse-instance-zxhlvevffc1c.cijt7auhxunw.us-east-1.rds.amazonaws.com"
)
port_str = (os.environ.get("POSTGRES_PORT") or "5432").strip()
puerto = int(port_str)
base = os.environ.get("POSTGRES_DB", "finnegansbi")

engine = create_engine(f"postgresql+psycopg2://{usuario}:{contraseña}@{host}:{puerto}/{base}")

# ------------------------------------------------------------------------------
# 🚀 Funciones genéricas con retry
# ------------------------------------------------------------------------------
def set_with_retry(worksheet, df, retries=3, wait=5):
    for i in range(1, retries + 1):
        try:
            set_with_dataframe(worksheet, df, include_index=False, resize=False)
            print("✅ Exportación completada.")
            return
        except Exception as e:
            print(f"⚠️ Intento {i}/{retries} falló: {e}")
            if i < retries:
                print(f"⏳ Reintentando en {wait} segundos...")
                time.sleep(wait)
            else:
                raise

def update_with_retry(worksheet, values, range_name, retries=3, wait=5):
    for i in range(1, retries + 1):
        try:
            worksheet.update(values=values, range_name=range_name)
            print("✅ Exportación sin encabezado completada.")
            return
        except Exception as e:
            print(f"⚠️ Intento {i}/{retries} falló: {e}")
            if i < retries:
                print(f"⏳ Reintentando en {wait} segundos...")
                time.sleep(wait)
            else:
                raise

def get_or_create_worksheet(spreadsheet, title, rows=1000, cols=26):
    try:
        return spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        print(f"🆕 Creando hoja '{title}' ({rows}x{cols})...")
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)

# ------------------------------------------------------------------------------
# 🧩 FUNCIÓN GENÉRICA EXPORTAR TABLA COMPLETA
#    Permite limpiar SOLO un rango (ej: "A:D") en vez de borrar toda la hoja
# ------------------------------------------------------------------------------
def exportar_tabla_completa(query_or_df, spreadsheet, hoja_nombre, columnas_decimal=[], clear_range=None, create_if_missing=False):
    if isinstance(query_or_df, str):
        df = pd.read_sql(query_or_df, engine)
    else:
        df = query_or_df.copy()  # evita side-effects (para reutilizar DF en churn/RFM)

    for col in columnas_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"{x:.2f}".replace(".", ",") if pd.notnull(x) else "")

    worksheet = get_or_create_worksheet(spreadsheet, hoja_nombre) if create_if_missing else spreadsheet.worksheet(hoja_nombre)

    # Si se pide un rango específico, solo borra ese rango; si no, limpia toda la hoja (comportamiento original) :contentReference[oaicite:2]{index=2}
    if clear_range:
        worksheet.batch_clear([clear_range])
    else:
        worksheet.clear()

    set_with_retry(worksheet, df)
    print(f"✅ Exportado: {hoja_nombre}" + (f" (limpieza: {clear_range})" if clear_range else ""))

# ------------------------------------------------------------------------------
# 💡 FUNCIÓN ESPECÍFICA PARA CORREGIR IMPORTES DE SALDOS (División por 10000)
# ------------------------------------------------------------------------------
def exportar_tabla_corregida(query_or_df, spreadsheet, hoja_nombre):
    if isinstance(query_or_df, str):
        df = pd.read_sql(query_or_df, engine)
    else:
        df = query_or_df.copy()

    columnas_a_corregir_y_dividir = [
        "importemonedatransaccion",
        "importemonedaprincipal",
        "importemonedasecundaria"
    ]

    for col in columnas_a_corregir_y_dividir:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(".", "", regex=False).str.replace(",", "", regex=False)
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col] / 10000.0
            df[col] = df[col].apply(
                lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notnull(x) else ""
            )

    worksheet = spreadsheet.worksheet(hoja_nombre)
    worksheet.clear()
    set_with_retry(worksheet, df)
    print(f"✅ Exportado CORREGIDO: {hoja_nombre}")

# ------------------------------------------------------------------------------
# 🧩 Exportar solo A2:Q sin encabezado
# ------------------------------------------------------------------------------
def exportar_libro_mayor(query, spreadsheet, hoja_nombre, columnas_decimal=[]):
    df = pd.read_sql(query, engine)
    for col in columnas_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"{x:.2f}".replace(".", ",") if pd.notnull(x) else "")
    df_recortado = df.iloc[:, :17]
    valores = df_recortado.values.tolist()
    worksheet = spreadsheet.worksheet(hoja_nombre)
    worksheet.batch_clear(["A2:Q"])
    update_with_retry(worksheet, values=valores, range_name="A2")
    print("✅ Exportado sin encabezado: Aux Libro Mayor")

# ------------------------------------------------------------------------------
# 📤 Exportar A2:J sin encabezado
# ------------------------------------------------------------------------------
def exportar_stock(query, spreadsheet, hoja_nombre, columnas_decimal=[]):
    df = pd.read_sql(query, engine)
    for col in columnas_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"{x:.2f}".replace(".", ",") if pd.notnull(x) else "")
    df_recortado = df.iloc[:, :10]
    valores = df_recortado.values.tolist()
    worksheet = spreadsheet.worksheet(hoja_nombre)
    worksheet.batch_clear(["A2:J"])
    update_with_retry(worksheet, values=valores, range_name="A2")
    print("✅ Exportado sin encabezado: Aux Stock")

# ------------------------------------------------------------------------------
# 📤 Exportar A2:J sin encabezado
# ------------------------------------------------------------------------------
def exportar_sumas_y_saldos(query, spreadsheet, hoja_nombre, columnas_decimal=[]):
    df = pd.read_sql(query, engine)
    for col in columnas_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"{x:.2f}".replace(".", ",") if pd.notnull(x) else "")

    df_recortado = df.iloc[:, :10]
    valores = df_recortado.values.tolist()
    worksheet = spreadsheet.worksheet(hoja_nombre)
    worksheet.batch_clear(["A2:J"])
    update_with_retry(worksheet, values=valores, range_name="A2")
    print("✅ Exportado sin encabezado: Aux Sumas y Saldos")

# ------------------------------------------------------------------------------
# Funciones para análisis de churn (OPT: reutiliza df_facturacion_full si se pasa)
# ------------------------------------------------------------------------------
def obtener_datos_facturacion(df_facturacion_full=None):
    """
    Obtiene los datos para churn.
    Mantiene la lógica original: solo ventas donde cuentanombre comience con 'Ventas Merc' :contentReference[oaicite:3]{index=3}
    OPT: si se provee df_facturacion_full (SELECT *), se filtra en memoria para evitar otra consulta.
    """
    if df_facturacion_full is None:
        query = """
SELECT
    clientecodigo,
    clientenombre,
    fechacomprobante,
    empresacodigo,
    empresanombre,
    cuentanombre
FROM public.inpro2021nube_facturacion
WHERE cuentanombre LIKE 'Ventas Merc%%'
ORDER BY clientecodigo, fechacomprobante
"""
        df = pd.read_sql(query, engine)
    else:
        cols = ["clientecodigo", "clientenombre", "fechacomprobante", "empresacodigo", "empresanombre", "cuentanombre"]
        df = df_facturacion_full[cols].copy()
        df = df[df["cuentanombre"].astype("string").str.startswith("Ventas Merc", na=False)]
        df = df.sort_values(["clientecodigo", "fechacomprobante"])

    df["fechacomprobante"] = pd.to_datetime(df["fechacomprobante"], errors="coerce")
    df = df.dropna(subset=["fechacomprobante"])

    print(f"Datos churn cargados: {len(df)} registros (solo ventas 'Ventas Merc')")
    if len(df) > 0:
        print(f"Rango de fechas: {df['fechacomprobante'].min()} a {df['fechacomprobante'].max()}")
        print(f"Total de clientes únicos: {df['clientecodigo'].nunique()}")
    return df

def obtener_ultima_compra_hasta_fecha(df, cliente, fecha_fin):
    compras_cliente = df[(df["clientecodigo"] == cliente) & (df["fechacomprobante"] <= fecha_fin)]
    if len(compras_cliente) == 0:
        return None
    return compras_cliente["fechacomprobante"].max()

def calcular_meses_desde_fecha(fecha_inicio, fecha_fin):
    delta = relativedelta(fecha_fin, fecha_inicio)
    return delta.years * 12 + delta.months

def calcular_status_mensual(df, cliente, primera_compra, mes_inicio, mes_fin, status_mes_anterior):
    """
    Lógica de churn:
    Si un cliente no compra durante 3 meses seguidos (<=3 meses),
    al 4to mes (después de 3 meses completos) se declara churn.
    """
    if pd.isna(primera_compra):
        return None

    ultima_compra_hasta_mes = obtener_ultima_compra_hasta_fecha(df, cliente, mes_fin)

    if primera_compra >= mes_inicio and primera_compra <= mes_fin:
        return "Nuevo"

    if ultima_compra_hasta_mes is None:
        if primera_compra < mes_inicio:
            meses_desde_primera = calcular_meses_desde_fecha(primera_compra, mes_fin)
            if meses_desde_primera > 3:
                fecha_churn_declaration = primera_compra + relativedelta(months=4)
                fecha_churn_declaration = (
                    fecha_churn_declaration.replace(day=1)
                    + relativedelta(months=1)
                    - timedelta(days=1)
                )

                if fecha_churn_declaration >= mes_inicio and fecha_churn_declaration <= mes_fin:
                    return "Churn del Mes"
                elif fecha_churn_declaration < mes_inicio:
                    if status_mes_anterior in ["Churn del Mes", "Churn Sostenido"]:
                        return "Churn Sostenido"
                    else:
                        return "Churn del Mes"
            elif meses_desde_primera <= 3:
                return "Cliente sin compra"
        return None

    tiene_compra_actual = ultima_compra_hasta_mes >= mes_inicio

    if tiene_compra_actual:
        ultima_compra_antes = obtener_ultima_compra_hasta_fecha(df, cliente, mes_inicio - timedelta(days=1))

        if ultima_compra_antes is not None:
            meses_sin_compra = calcular_meses_desde_fecha(ultima_compra_antes, mes_inicio)
            if meses_sin_compra > 3:
                return "Recuperado"
            else:
                return "Retenido"
        else:
            if primera_compra < mes_inicio:
                meses_desde_primera = calcular_meses_desde_fecha(primera_compra, mes_inicio)
                if meses_desde_primera > 3:
                    return "Recuperado"
            return "Retenido"

    meses_sin_compra = calcular_meses_desde_fecha(ultima_compra_hasta_mes, mes_fin)

    if meses_sin_compra <= 3:
        return "Cliente sin compra"

    fecha_churn_declaration = ultima_compra_hasta_mes + relativedelta(months=4)
    fecha_churn_declaration = (
        fecha_churn_declaration.replace(day=1)
        + relativedelta(months=1)
        - timedelta(days=1)
    )

    if status_mes_anterior in ["Churn del Mes", "Churn Sostenido"]:
        return "Churn Sostenido"

    if fecha_churn_declaration >= mes_inicio and fecha_churn_declaration <= mes_fin:
        return "Churn del Mes"
    elif fecha_churn_declaration < mes_inicio:
        return "Churn Sostenido"
    else:
        return "Churn del Mes"

def generar_fechas_mensuales(df):
    if len(df) == 0:
        return []

    fecha_min = df["fechacomprobante"].min()
    fecha_max = df["fechacomprobante"].max()

    inicio = fecha_min.replace(day=1)
    if fecha_max.month == 12:
        fin = fecha_max.replace(day=31)
    else:
        fin = (fecha_max.replace(day=1) + relativedelta(months=1)) - timedelta(days=1)

    fechas = []
    fecha_actual = inicio
    while fecha_actual <= fin:
        mes_inicio = fecha_actual.replace(day=1)
        if fecha_actual.month == 12:
            mes_fin = fecha_actual.replace(day=31)
        else:
            mes_fin = (fecha_actual + relativedelta(months=1)).replace(day=1) - timedelta(days=1)

        fechas.append((mes_inicio, mes_fin))
        fecha_actual = fecha_actual + relativedelta(months=1)

    return fechas

def crear_matriz_churn(df):
    clientes = df[["clientecodigo", "clientenombre"]].drop_duplicates()
    fechas_mensuales = generar_fechas_mensuales(df)

    print(f"Períodos a procesar: {len(fechas_mensuales)} meses")

    primera_compra_df = df.groupby("clientecodigo")["fechacomprobante"].min().reset_index()
    primera_compra_df.columns = ["clientecodigo", "primera_compra"]
    primera_compra_dict = dict(zip(primera_compra_df["clientecodigo"], primera_compra_df["primera_compra"]))

    resultados = []
    total_clientes = len(clientes)

    for idx, (_, cliente_row) in enumerate(clientes.iterrows(), 1):
        cliente = cliente_row["clientecodigo"]
        cliente_nombre = cliente_row["clientenombre"]
        primera_compra = primera_compra_dict.get(cliente)

        if idx % 100 == 0:
            print(f"Procesando cliente {idx}/{total_clientes}...")

        status_mes_anterior = None

        for mes_inicio, mes_fin in fechas_mensuales:
            if primera_compra is not None and mes_fin < primera_compra.replace(day=1):
                continue

            status = calcular_status_mensual(df, cliente, primera_compra, mes_inicio, mes_fin, status_mes_anterior)

            if status is not None:
                mes_str = mes_inicio.strftime("%Y-%m")
                resultados.append(
                    {
                        "ClienteCodigo": cliente,
                        "ClienteNombre": cliente_nombre,
                        "Mes": mes_str,
                        "Status": status,
                    }
                )

            status_mes_anterior = status

    return pd.DataFrame(resultados)

# ------------------------------------------------------------------------------
# ✅ RFM (NUEVO) - Reutiliza df_facturacion_full (misma extracción)
# ------------------------------------------------------------------------------
def qscore(series: pd.Series, q: int = 5, reverse: bool = False) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    q_eff = int(min(q, s.nunique(dropna=True))) if s.nunique(dropna=True) > 0 else 1

    if q_eff < 2:
        scores = pd.Series(np.ones(len(s), dtype=int), index=s.index)
    else:
        ranked = s.rank(method="first", ascending=True)
        raw = pd.qcut(ranked, q=q_eff, labels=False) + 1  # 1..q_eff
        # Reescala a 1..q si q_eff < q
        scores = np.ceil(raw.astype(float) * q / q_eff).astype(int)

    return (q + 1) - scores if reverse else scores

def calcular_rfm(df_facturacion_full: pd.DataFrame, scoring_quantiles: int = 5, cuenta_keyword: str = "venta") -> pd.DataFrame:
    """
    Reglas solicitadas:
    - Solo líneas con cuentanombre contiene 'Venta' (case-insensitive)
    - Excluir líneas con importemonedaprincipal < 0 (notas de crédito)
    - No usa empresacodigo para agrupar
    """
    required = [
        "clientecodigo", "clientenombre", "fechacomprobante",
        "comprobantenumero", "cuentanombre", "importemonedaprincipal"
    ]
    missing = [c for c in required if c not in df_facturacion_full.columns]
    if missing:
        raise ValueError(f"Faltan columnas para RFM en df_facturacion_full: {missing}")

    df = df_facturacion_full[required].copy()

    # Fecha
    df["fechacomprobante"] = pd.to_datetime(df["fechacomprobante"], errors="coerce")
    df = df.dropna(subset=["fechacomprobante"])

    # Monto numérico
    df["importemonedaprincipal"] = pd.to_numeric(df["importemonedaprincipal"], errors="coerce")
    df = df.dropna(subset=["importemonedaprincipal"])

    # 1) Solo cuentas con "Venta"
    mask_venta = df["cuentanombre"].astype("string").str.contains(cuenta_keyword, case=False, na=False)
    df = df[mask_venta].copy()

    # 2) Excluir negativos (notas de crédito)
    df = df[df["importemonedaprincipal"] >= 0].copy()

    if len(df) == 0:
        return pd.DataFrame(columns=[
            "clientecodigo", "clientenombre", "last_purchase", "recency_days",
            "frequency", "monetary", "R_score", "F_score", "M_score", "RFM", "RFM_sum", "segment"
        ])

    # Consolidación a nivel transacción: cliente + fecha + comprobante (rellenando nulos)
    df["_doc_filled"] = df["comprobantenumero"].astype("string")
    m = df["_doc_filled"].isna()
    df.loc[m, "_doc_filled"] = "SIN_COMPROBANTE_" + df.index[m].astype(str)

    tx = (
        df.groupby(["clientecodigo", "fechacomprobante", "_doc_filled"], as_index=False)
          .agg(
              clientenombre=("clientenombre", "first"),
              tx_amount=("importemonedaprincipal", "sum")
          )
          .rename(columns={"fechacomprobante": "fecha", "_doc_filled": "comprobante"})
    )

    # Fecha de referencia
    as_of_date = tx["fecha"].max() + pd.Timedelta(days=1)

    rfm = (
        tx.groupby(["clientecodigo"], as_index=False)
          .agg(
              clientenombre=("clientenombre", "first"),
              last_purchase=("fecha", "max"),
              frequency=("comprobante", "nunique"),
              monetary=("tx_amount", "sum"),
          )
    )

    rfm["recency_days"] = (as_of_date - rfm["last_purchase"]).dt.days

    q = int(scoring_quantiles)
    rfm["R_score"] = qscore(rfm["recency_days"], q=q, reverse=True)
    rfm["F_score"] = qscore(rfm["frequency"], q=q, reverse=False)
    rfm["M_score"] = qscore(rfm["monetary"], q=q, reverse=False)

    rfm["RFM"] = rfm["R_score"].astype(str) + rfm["F_score"].astype(str) + rfm["M_score"].astype(str)
    rfm["RFM_sum"] = rfm[["R_score", "F_score", "M_score"]].sum(axis=1)

    def rfm_segment(row) -> str:
        R, F, M = row["R_score"], row["F_score"], row["M_score"]
        if (R >= 4) and (F >= 4) and (M >= 4):
            return "Champions"
        if (R >= 4) and (F >= 3):
            return "Leales"
        if (R == 5) and (F == 1):
            return "Nuevos"
        if (R >= 4) and (F in [1, 2]):
            return "Promesas"
        if (R == 3) and (F >= 3):
            return "Atención"
        if (R == 2) and (F in [1, 2]):
            return "Por dormirse"
        if (R <= 2) and (F >= 3):
            return "En Riesgo"
        if (R == 1) and (F == 1):
            return "Perdidos"
        return "Otros"

    rfm["segment"] = rfm.apply(rfm_segment, axis=1)

    # Orden sugerido
    rfm = rfm.sort_values(["segment", "RFM_sum"], ascending=[True, False])

    return rfm

# ------------------------------------------------------------------------------
# CONFIGURACIÓN DE QUERYS ESPECÍFICAS
# ------------------------------------------------------------------------------
QUERY_SALDOS_CLIENTES_FILTRADOS = """
SELECT * FROM public.inpro2021nube_composicion_saldos_clientes_inprocil c
WHERE
    c.empresanombre = 'INPROCIL S.A.' AND
    c.cuentacontablecodigo IN ('ANT101', 'AAP301', 'DML101') AND
    c.clientenombre not like '%%BENVENUTO%%'  AND
    c.clientenombre not like '%%CONCEPCION%%' AND
    c.clientenombre not like '%%BUIATTI%%' AND
    c.clientenombre not like '%%CAMPUZANO HORACIO DAVID%%' AND
    c.clientenombre not like '%%CONTIN %%' AND
    c.clientenombre not like '%%COOPERATIVA DE TRABAJO%%' AND
    c.clientenombre not like '%%DOMVIL%%' AND
    c.clientenombre not like '%%GAS MOVIL%%' AND
    c.clientenombre not like '%%GNC PATAGONICA%%' AND
    c.clientenombre not like '%%GOMEZ FABIAN%%' AND
    c.clientenombre not like '%%GOMEZ GUSTAVO%%' AND
    c.clientenombre not like '%%PALLETIZATE%%' AND
    c.clientenombre not like '%%PAUSYG%%' AND
    c.clientenombre not like '%%POWER CHECK%%' AND
    c.clientenombre not like '%%RODRIGUEZ ALEJANDRO%%' AND
    c.clientenombre not like '%%VALSI GAS%%'
"""

QUERY_SALDOS_PROVEEDORES_FILTRADOS = """
select * from public.inpro2021nube_composicion_saldo_proveedores_inprocil c
"""

# ------------------------------------------------------------------------------
# EXPORTACIONES PRINCIPALES
# ------------------------------------------------------------------------------
SPREADSHEET_SALDOS_URL = os.environ.get(
    "SPREADSHEET_SALDOS_URL",
    # Default apuntando al sheet que indicaste (puedes seguir sobreescribiendo por env var si ya lo usas)
    "https://docs.google.com/spreadsheets/d/1oR_fdVCyn1cA8zwH4XgU5VK63cZaDC3I1i3-SWaUT20/edit"
)
SPREADSHEET_LIBRO_MAYOR_URL = os.environ.get(
    "SPREADSHEET_LIBRO_MAYOR_URL",
    "https://docs.google.com/spreadsheets/d/<ID_SHEET_LIBRO_MAYOR>/edit"
)
SPREADSHEET_STOCK_PUC_URL = os.environ.get(
    "SPREADSHEET_STOCK_PUC_URL",
    "https://docs.google.com/spreadsheets/d/<ID_SHEET_STOCK_PUC>/edit"
)

# ------------------------------------------------------------------------------
# 📁 Spreadsheet 1
# ------------------------------------------------------------------------------
saldos_sheet = client.open_by_url(SPREADSHEET_SALDOS_URL)

# 1. Saldos clientes filtrados
exportar_tabla_completa(
    QUERY_SALDOS_CLIENTES_FILTRADOS,
    saldos_sheet,
    "Base Saldos Clientes",
    ["importemonedatransaccion", "importemonedaprincipal", "importemonedasecundaria"]
)

# 2. Saldos proveedores
print("\nEjecutando exportación: Composicion Saldo Proveedores de INPROCIL S.A.")
exportar_tabla_completa(
    QUERY_SALDOS_PROVEEDORES_FILTRADOS,
    saldos_sheet,
    "Composicion Saldo Proveedores",
    ["importemonedatransaccion", "importemonedaprincipal", "importemonedasecundaria"]
)

# 3. Base Sumas y Saldos
exportar_tabla_completa(
    "SELECT * FROM public.inpro2021nube_sumas_y_saldos",
    saldos_sheet,
    "Base Sumas y Saldos",
    ["sumadebe", "sumahaber", "saldoacumulado"]
)

# 4. Stock comprometido
exportar_tabla_completa(
    "SELECT * FROM public.inpro2021nube_stock_comprometido",
    saldos_sheet,
    "Base Pendientes Entrega",
    ["cantidadpendiente"]
)

# 5. Facturación (OPT: se lee 1 vez y se reutiliza para churn y RFM)
print("\nCargando facturación completa desde DW...")
df_facturacion_full = pd.read_sql("SELECT * FROM public.inpro2021nube_facturacion", engine)
print(f"Facturación total cargada: {len(df_facturacion_full)} filas")

exportar_tabla_completa(
    df_facturacion_full,  # reutiliza DF (evita re-query)
    saldos_sheet,
    "Base Facturacion",
    [
        "preciomonedatransaccion",
        "importemonedatransaccion",
        "importemonedaprincipal",
        "importemonedasecundaria",
        "cotizacionmonedatransaccion",
        "cantidad",
    ],
)

# ------------------------------------------------------------------------------
# ✅ Análisis de churn y exportación
#    Mantiene el comportamiento original de limpiar solo A:D en "Analisis_Churn" :contentReference[oaicite:4]{index=4}
# ------------------------------------------------------------------------------
print("\nEjecutando análisis de churn...")
df_facturacion_churn = obtener_datos_facturacion(df_facturacion_full=df_facturacion_full)
matriz_churn = crear_matriz_churn(df_facturacion_churn)

exportar_tabla_completa(
    matriz_churn,
    saldos_sheet,
    "Analisis_Churn",
    [],
    clear_range="A:D"
)

# ------------------------------------------------------------------------------
# ✅ RFM (NUEVO) - output a hoja "RFM" en el mismo spreadsheet
# ------------------------------------------------------------------------------
print("\nCalculando RFM...")
df_rfm = calcular_rfm(df_facturacion_full=df_facturacion_full, scoring_quantiles=5, cuenta_keyword="venta")
print(f"RFM generado: {len(df_rfm)} clientes")

exportar_tabla_completa(
    df_rfm,
    saldos_sheet,
    "RFM",
    columnas_decimal=[],       # se exporta numérico; formateás en Sheets si querés
    clear_range=None,
    create_if_missing=True     # crea la hoja si no existe
)

# ------------------------------------------------------------------------------
# 📁 Spreadsheet 2
# ------------------------------------------------------------------------------
libro_mayor_sheet = client.open_by_url(SPREADSHEET_LIBRO_MAYOR_URL)

exportar_libro_mayor(
    "SELECT * FROM public.inpro2021nube_libro_mayor",
    libro_mayor_sheet,
    "Aux Libro Mayor",
    ["Debe", "Haber", "importemonedaprincipal", "Imp. operacion ppal.", "Imp. operacion sec.", "Tipo Cambio"],
)

exportar_stock(
    "SELECT * FROM public.inpro2021nube_stock_con_PUC",
    libro_mayor_sheet,
    "Aux Stock",
    ["Stock", "UltimoPrecioCompra"],
)

exportar_sumas_y_saldos(
    "SELECT * FROM public.inpro2021nube_sumas_y_saldos",
    libro_mayor_sheet,
    "Aux Sumas y Saldos",
    ["Debe", "Haber", "saldoperiodo", "saldo", "saldoinicial"],
)

# ------------------------------------------------------------------------------
# 📁 Spreadsheet 3
# ------------------------------------------------------------------------------
stock_con_puc_sheet = client.open_by_url(SPREADSHEET_STOCK_PUC_URL)

exportar_stock(
    "SELECT * FROM public.inpro2021nube_stock_con_PUC",
    stock_con_puc_sheet,
    "Aux Stock",
    ["Stock", "UltimoPrecioCompra"],
)
