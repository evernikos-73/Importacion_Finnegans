import os
os.makedirs(os.path.expanduser('~/output'), exist_ok=True)

script = '''\
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
# Autenticacion Google Sheets
# ------------------------------------------------------------------------------
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
cred_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
creds = ServiceAccountCredentials.from_json_keyfile_dict(cred_dict, scope)
client = gspread.authorize(creds)

# ------------------------------------------------------------------------------
# PostgreSQL
# OPT: pool_pre_ping=True evita conexiones TCP muertas en ejecuciones largas
# ------------------------------------------------------------------------------
usuario    = os.environ.get("POSTGRES_USER", "inpro2021nubeuser")
contrasena = os.environ.get("POSTGRES_PASSWORD", "")
host       = os.environ.get(
    "POSTGRES_HOST",
    "infraestructura-aurora-datawarehouse-instance-zxhlvevffc1c.cijt7auhxunw.us-east-1.rds.amazonaws.com"
)
port_str   = (os.environ.get("POSTGRES_PORT") or "5432").strip()
puerto     = int(port_str)
base       = os.environ.get("POSTGRES_DB", "finnegansbi")

engine = create_engine(
    f"postgresql+psycopg2://{usuario}:{contrasena}@{host}:{puerto}/{base}",
    pool_pre_ping=True
)

# ------------------------------------------------------------------------------
# OPT: Constante centralizada para filtro de cuentanombre de Ventas
# Un unico lugar para cambiar si cambia la nomenclatura contable
# ------------------------------------------------------------------------------
CUENTA_KEYWORD_VENTA = "venta"

# ------------------------------------------------------------------------------
# Funciones genericas con retry
# ------------------------------------------------------------------------------
def set_with_retry(worksheet, df, retries=3, wait=5):
    for i in range(1, retries + 1):
        try:
            set_with_dataframe(worksheet, df, include_index=False, resize=False)
            print("OK Exportacion completada.")
            return
        except Exception as e:
            print(f"Intento {i}/{retries} fallo: {e}")
            if i < retries:
                print(f"Reintentando en {wait} segundos...")
                time.sleep(wait)
            else:
                raise

def update_with_retry(worksheet, values, range_name, retries=3, wait=5):
    for i in range(1, retries + 1):
        try:
            worksheet.update(values=values, range_name=range_name)
            print("OK Exportacion sin encabezado completada.")
            return
        except Exception as e:
            print(f"Intento {i}/{retries} fallo: {e}")
            if i < retries:
                print(f"Reintentando en {wait} segundos...")
                time.sleep(wait)
            else:
                raise

def get_or_create_worksheet(spreadsheet, title, rows=1000, cols=26):
    try:
        return spreadsheet.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        print(f"Creando hoja \'{title}\' ({rows}x{cols})...")
        return spreadsheet.add_worksheet(title=title, rows=rows, cols=cols)

# ------------------------------------------------------------------------------
# FUNCION GENERICA: Exportar tabla completa
# ------------------------------------------------------------------------------
def exportar_tabla_completa(query_or_df, spreadsheet, hoja_nombre, columnas_decimal=[], clear_range=None, create_if_missing=False):
    if isinstance(query_or_df, str):
        df = pd.read_sql(query_or_df, engine)
    else:
        df = query_or_df.copy()

    for col in columnas_decimal:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].apply(lambda x: f"{x:.2f}".replace(".", ",") if pd.notnull(x) else "")

    worksheet = get_or_create_worksheet(spreadsheet, hoja_nombre) if create_if_missing else spreadsheet.worksheet(hoja_nombre)

    if clear_range:
        worksheet.batch_clear([clear_range])
    else:
        worksheet.clear()

    set_with_retry(worksheet, df)
    print(f"OK Exportado: {hoja_nombre}" + (f" (limpieza: {clear_range})" if clear_range else ""))

# ------------------------------------------------------------------------------
# FUNCION: Corregir importes de Saldos (Division por 10000)
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
    print(f"OK Exportado CORREGIDO: {hoja_nombre}")

# ------------------------------------------------------------------------------
# FUNCION: Exportar Libro Mayor A2:Q sin encabezado
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
    print("OK Exportado sin encabezado: Aux Libro Mayor")

# ------------------------------------------------------------------------------
# FUNCION: Exportar Stock A2:J sin encabezado
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
    print("OK Exportado sin encabezado: Aux Stock")

# ------------------------------------------------------------------------------
# FUNCION: Exportar Sumas y Saldos A2:J sin encabezado
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
    print("OK Exportado sin encabezado: Aux Sumas y Saldos")

# ------------------------------------------------------------------------------
# AUX: Mapeo de clientes agrupados (Grupo Economico)
# ------------------------------------------------------------------------------
def norm_name(x) -> str:
    if x is None:
        return ""
    s = str(x).strip().upper()
    s = " ".join(s.split())
    return s

def obtener_mapa_clientes_agrupados(spreadsheet, sheet_name="AUX_Agrup_Clientes") -> dict:
    try:
        ws = spreadsheet.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        print(f"No existe la hoja \'{sheet_name}\'. Se continuara sin agrupacion.")
        return {}

    values = ws.get("A:B")
    if not values or len(values) < 2:
        print(f"Hoja \'{sheet_name}\' vacia o sin datos. Se continuara sin agrupacion.")
        return {}

    start_idx = 0
    h0 = norm_name(values[0][0]) if len(values[0]) > 0 else ""
    h1 = norm_name(values[0][1]) if len(values[0]) > 1 else ""
    if ("CLIENTENOMBRE" in h0) and ("CLIENTE AGRUPADO" in h1):
        start_idx = 1

    mapa = {}
    for row in values[start_idx:]:
        if not row or len(row) < 2:
            continue
        orig = row[0]
        grp  = row[1]
        k    = norm_name(orig)
        if not k:
            continue
        grp_clean = str(grp).strip() if grp is not None else ""
        if grp_clean == "":
            continue
        mapa[k] = grp_clean

    print(f"OK Mapeo de clientes agrupados cargado: {len(mapa)} reglas")
    return mapa

def aplicar_agrupacion_cliente(df: pd.DataFrame, mapa: dict, source_col="clientenombre", target_col="cliente_agrupado") -> pd.DataFrame:
    if source_col not in df.columns:
        raise ValueError(f"No existe columna \'{source_col}\' en el dataframe para agrupar clientes.")

    if not mapa:
        df[target_col] = df[source_col].astype("string").fillna("").astype(str).str.strip()
        m = df[target_col] == ""
        df.loc[m, target_col] = df.loc[m, source_col].astype("string").fillna("").astype(str).str.strip()
        return df

    def map_fn(x):
        if pd.isna(x):
            return ""
        sx = str(x).strip()
        if sx == "":
            return ""
        return mapa.get(norm_name(sx), sx)

    df[target_col] = df[source_col].apply(map_fn).astype(str).str.strip()
    m = df[target_col].astype(str).str.strip() == ""
    df.loc[m, target_col] = df.loc[m, source_col].astype("string").fillna("").astype(str).str.strip()
    return df

# ------------------------------------------------------------------------------
# AUX: Tipo de cambio USD desde AUX!A:B
# ------------------------------------------------------------------------------
def parse_number_locale(val):
    if val is None:
        return None
    s = str(val).strip()
    if s == "":
        return None
    s = s.replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def obtener_tc_usd_desde_aux(spreadsheet, sheet_name="AUX", range_name="A:B") -> float:
    ws = spreadsheet.worksheet(sheet_name)
    values = ws.get(range_name)

    tc = None
    for row in values:
        if len(row) < 2:
            continue
        v = parse_number_locale(row[1])
        if v is not None:
            tc = v

    if tc is None or tc <= 0:
        raise RuntimeError("No se pudo obtener un tipo de cambio USD valido desde AUX!A:B (columna B).")
    return tc

# ------------------------------------------------------------------------------
# CHURN: Analisis de retencion por cliente agrupado
# ------------------------------------------------------------------------------
def obtener_datos_facturacion(df_facturacion_full=None, mapa_clientes=None):
    if df_facturacion_full is None:
        query = """
SELECT
    clientenombre,
    fechacomprobante,
    cuentanombre
FROM public.inpro2021nube_facturacion
WHERE cuentanombre LIKE \'Ventas Merc%%\'
ORDER BY clientenombre, fechacomprobante
"""
        df = pd.read_sql(query, engine)
    else:
        cols = ["clientenombre", "fechacomprobante", "cuentanombre"]
        df = df_facturacion_full[cols].copy()
        df = df[df["cuentanombre"].astype("string").str.startswith("Ventas Merc", na=False)]
        df = df.sort_values(["clientenombre", "fechacomprobante"])

    df["fechacomprobante"] = pd.to_datetime(df["fechacomprobante"], errors="coerce")
    df = df.dropna(subset=["fechacomprobante"])

    mapa_clientes = mapa_clientes or {}
    df = aplicar_agrupacion_cliente(df, mapa_clientes, source_col="clientenombre", target_col="cliente_agrupado")

    print(f"Datos churn cargados: {len(df)} registros")
    if len(df) > 0:
        print(f"Rango de fechas: {df[\'fechacomprobante\'].min()} a {df[\'fechacomprobante\'].max()}")
        print(f"Total de clientes agrupados unicos: {df[\'cliente_agrupado\'].nunique()}")
    return df

def obtener_ultima_compra_hasta_fecha(df, cliente_agrupado, fecha_fin):
    compras_cliente = df[(df["cliente_agrupado"] == cliente_agrupado) & (df["fechacomprobante"] <= fecha_fin)]
    if len(compras_cliente) == 0:
        return None
    return compras_cliente["fechacomprobante"].max()

def calcular_meses_desde_fecha(fecha_inicio, fecha_fin):
    delta = relativedelta(fecha_fin, fecha_inicio)
    return delta.years * 12 + delta.months

def calcular_status_mensual(df, cliente_agrupado, primera_compra, mes_inicio, mes_fin, status_mes_anterior):
    if pd.isna(primera_compra):
        return None

    ultima_compra_hasta_mes = obtener_ultima_compra_hasta_fecha(df, cliente_agrupado, mes_fin)

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
        ultima_compra_antes = obtener_ultima_compra_hasta_fecha(df, cliente_agrupado, mes_inicio - timedelta(days=1))
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
    clientes = df[["cliente_agrupado"]].drop_duplicates()
    fechas_mensuales = generar_fechas_mensuales(df)
    print(f"Periodos a procesar: {len(fechas_mensuales)} meses")

    primera_compra_df = df.groupby("cliente_agrupado")["fechacomprobante"].min().reset_index()
    primera_compra_df.columns = ["cliente_agrupado", "primera_compra"]
    primera_compra_dict = dict(zip(primera_compra_df["cliente_agrupado"], primera_compra_df["primera_compra"]))

    resultados = []
    total_clientes = len(clientes)

    for idx, (_, cliente_row) in enumerate(clientes.iterrows(), 1):
        cliente_agrupado = cliente_row["cliente_agrupado"]
        primera_compra = primera_compra_dict.get(cliente_agrupado)

        if idx % 100 == 0:
            print(f"Procesando cliente agrupado {idx}/{total_clientes}...")

        status_mes_anterior = None

        for mes_inicio, mes_fin in fechas_mensuales:
            if primera_compra is not None and mes_fin < primera_compra.replace(day=1):
                continue

            status = calcular_status_mensual(df, cliente_agrupado, primera_compra, mes_inicio, mes_fin, status_mes_anterior)

            if status is not None:
                mes_str = mes_inicio.strftime("%Y-%m")
                resultados.append({
                    "Cliente Agrupado": cliente_agrupado,
                    "ClienteNombre": cliente_agrupado,
                    "Mes": mes_str,
                    "Status": status,
                })

            status_mes_anterior = status

    return pd.DataFrame(resultados)

# ------------------------------------------------------------------------------
# RFM: Recencia, Frecuencia, Monetario
# ------------------------------------------------------------------------------
def qscore(series: pd.Series, q: int = 5, reverse: bool = False) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce")
    q_eff = int(min(q, s.nunique(dropna=True))) if s.nunique(dropna=True) > 0 else 1

    if q_eff < 2:
        scores = pd.Series(np.ones(len(s), dtype=int), index=s.index)
    else:
        ranked = s.rank(method="first", ascending=True)
        raw = pd.qcut(ranked, q=q_eff, labels=False) + 1
        scores = np.ceil(raw.astype(float) * q / q_eff).astype(int)

    return (q + 1) - scores if reverse else scores

def calcular_rfm(
    df_facturacion_full: pd.DataFrame,
    mapa_clientes: dict,
    usd_tc_ars_por_usd: float,
    scoring_quantiles: int = 5,
    cuenta_keyword: str = CUENTA_KEYWORD_VENTA
) -> pd.DataFrame:
    required = [
        "clientenombre", "fechacomprobante",
        "comprobantenumero", "cuentanombre", "importemonedaprincipal"
    ]
    missing = [c for c in required if c not in df_facturacion_full.columns]
    if missing:
        raise ValueError(f"Faltan columnas para RFM: {missing}")

    if usd_tc_ars_por_usd is None or usd_tc_ars_por_usd <= 0:
        raise ValueError("usd_tc_ars_por_usd invalido (debe ser > 0).")

    df = df_facturacion_full[required].copy()
    df["fechacomprobante"] = pd.to_datetime(df["fechacomprobante"], errors="coerce")
    df = df.dropna(subset=["fechacomprobante"])
    df["importemonedaprincipal"] = pd.to_numeric(df["importemonedaprincipal"], errors="coerce")
    df = df.dropna(subset=["importemonedaprincipal"])

    mask_venta = df["cuentanombre"].astype("string").str.contains(cuenta_keyword, case=False, na=False)
    df = df[mask_venta].copy()
    df = df[df["importemonedaprincipal"] >= 0].copy()

    if len(df) == 0:
        return pd.DataFrame(columns=[
            "cliente_agrupado", "clientenombre", "last_purchase", "recency_days",
            "frequency", "monetary_total_usd", "monetary_avg_usd",
            "R_score", "F_score", "M_score", "RFM", "RFM_sum", "segment",
            "last_purchase_amount_ars"
        ])

    df = aplicar_agrupacion_cliente(df, mapa_clientes or {}, source_col="clientenombre", target_col="cliente_agrupado")

    df["_doc_filled"] = df["comprobantenumero"].astype("string")
    m = df["_doc_filled"].isna()
    df.loc[m, "_doc_filled"] = "SIN_COMPROBANTE_" + df.index[m].astype(str)

    tx = (
        df.groupby(["cliente_agrupado", "fechacomprobante", "_doc_filled"], as_index=False)
          .agg(tx_amount_ars=("importemonedaprincipal", "sum"))
          .rename(columns={"fechacomprobante": "fecha", "_doc_filled": "comprobante"})
    )

    tx["tx_amount_usd"] = tx["tx_amount_ars"] / float(usd_tc_ars_por_usd)
    tx["mes"] = tx["fecha"].dt.to_period("M").astype(str)
    as_of_date = tx["fecha"].max() + pd.Timedelta(days=1)

    rfm = (
        tx.groupby(["cliente_agrupado"], as_index=False)
          .agg(
              last_purchase=("fecha", "max"),
              frequency=("mes", "nunique"),
              monetary_total_usd=("tx_amount_usd", "sum"),
          )
    )

    rfm["recency_days"] = (as_of_date - rfm["last_purchase"]).dt.days
    rfm["monetary_avg_usd"] = np.where(
        rfm["frequency"] > 0,
        rfm["monetary_total_usd"] / rfm["frequency"],
        0.0
    )

    tx_day_sum = (
        tx.groupby(["cliente_agrupado", "fecha"], as_index=False)
          .agg(last_purchase_amount_ars=("tx_amount_ars", "sum"))
    )
    rfm = rfm.merge(
        tx_day_sum,
        left_on=["cliente_agrupado", "last_purchase"],
        right_on=["cliente_agrupado", "fecha"],
        how="left"
    ).drop(columns=["fecha"])
    rfm["last_purchase_amount_ars"] = rfm["last_purchase_amount_ars"].fillna(0.0)

    q = int(scoring_quantiles)
    rfm["R_score"] = qscore(rfm["recency_days"],     q=q, reverse=True)
    rfm["F_score"] = qscore(rfm["frequency"],        q=q, reverse=False)
    rfm["M_score"] = qscore(rfm["monetary_avg_usd"], q=q, reverse=False)

    rfm["RFM"]     = rfm["R_score"].astype(str) + rfm["F_score"].astype(str) + rfm["M_score"].astype(str)
    rfm["RFM_sum"] = rfm[["R_score", "F_score", "M_score"]].sum(axis=1)

    def rfm_segment(row) -> str:
        R, F, M = row["R_score"], row["F_score"], row["M_score"]
        if (R >= 4) and (F >= 4) and (M >= 4): return "Campeones"
        if (R >= 4) and (F >= 3):              return "Leales"
        if (R == 5) and (F == 1):              return "Nuevos"
        if (R >= 4) and (F in [1, 2]):         return "Promesas"
        if (R == 3) and (F >= 3):              return "Atencion"
        if (R == 2) and (F in [1, 2]):         return "Por dormirse"
        if (R <= 2) and (F >= 3):              return "En Riesgo"
        if (R == 1) and (F == 1):              return "Perdidos"
        return "Otros"

    rfm["segment"]       = rfm.apply(rfm_segment, axis=1)
    rfm["clientenombre"] = rfm["cliente_agrupado"]
    rfm = rfm.sort_values(["segment", "RFM_sum"], ascending=[True, False])

    rfm = rfm[[
        "cliente_agrupado", "clientenombre", "last_purchase", "recency_days",
        "frequency", "monetary_total_usd", "monetary_avg_usd", "R_score",
        "F_score", "M_score", "RFM", "RFM_sum", "segment", "last_purchase_amount_ars"
    ]].copy()

    return rfm

# ------------------------------------------------------------------------------
# NUEVO: ABC MENSUAL POR CLIENTE
# Clasifica cada cliente en A/B/C para cada mes calendario.
# Solo considera lineas con cuentanombre que empieza con "venta" (case-insensitive).
# Reutiliza df_facturacion_full y aplicar_agrupacion_cliente para consistencia
# con Churn y RFM.
# ------------------------------------------------------------------------------
def calcular_abc_mensual(
    df_facturacion_full: pd.DataFrame,
    mapa_clientes: dict,
    umbral_a: float = 0.80,
    umbral_b: float = 0.95,
    cuenta_keyword: str = CUENTA_KEYWORD_VENTA
) -> pd.DataFrame:
    """
    Calcula la clasificacion ABC de cada cliente por mes.

    Columnas del resultado:
        periodo          : ano-mes (ej: 2024-01)
        cliente_agrupado : nombre agrupado del cliente
        facturacion_mes  : suma de importemonedaprincipal del mes (solo ventas)
        pct_sobre_total  : porcentaje sobre el total del mes (0-100)
        pct_acumulado    : porcentaje acumulado dentro del mes (0-100)
        clase_abc        : A / B / C segun umbral_a y umbral_b
    """
    required = ["clientenombre", "fechacomprobante", "cuentanombre", "importemonedaprincipal"]
    missing  = [c for c in required if c not in df_facturacion_full.columns]
    if missing:
        raise ValueError(f"Faltan columnas para ABC: {missing}")

    df = df_facturacion_full[required].copy()

    # 1. Filtrar solo lineas de Venta (cuentanombre comienza con "venta", case-insensitive)
    df = df[df["cuentanombre"].astype("string").str.lower().str.startswith(cuenta_keyword, na=False)].copy()

    # 2. Parsear fechas
    df["fechacomprobante"] = pd.to_datetime(df["fechacomprobante"], errors="coerce")
    df = df.dropna(subset=["fechacomprobante"])

    # 3. Convertir importes
    df["importemonedaprincipal"] = pd.to_numeric(df["importemonedaprincipal"], errors="coerce")
    df = df.dropna(subset=["importemonedaprincipal"])

    # 4. Aplicar agrupacion de clientes (mismo criterio que RFM y Churn)
    df = aplicar_agrupacion_cliente(df, mapa_clientes or {}, source_col="clientenombre", target_col="cliente_agrupado")

    # 5. Crear columna de periodo YYYY-MM
    df["periodo"] = df["fechacomprobante"].dt.to_period("M").astype(str)

    # 6. Sumar facturacion por cliente y periodo
    resumen = (
        df.groupby(["periodo", "cliente_agrupado"], as_index=False)["importemonedaprincipal"]
        .sum()
        .rename(columns={"importemonedaprincipal": "facturacion_mes"})
    )

    # 7. Excluir filas con facturacion neta <= 0 (devoluciones netas)
    resumen = resumen[resumen["facturacion_mes"] > 0].copy()

    # 8. Ordenar de mayor a menor dentro de cada mes (requisito del ABC)
    resumen = resumen.sort_values(["periodo", "facturacion_mes"], ascending=[True, False])

    # 9. Calcular % sobre total del mes y % acumulado (vectorizado con groupby+transform)
    total_mes               = resumen.groupby("periodo")["facturacion_mes"].transform("sum")
    resumen["pct_sobre_total"] = resumen["facturacion_mes"] / total_mes
    resumen["pct_acumulado"]   = resumen.groupby("periodo")["pct_sobre_total"].cumsum()

    # 10. Clasificar A / B / C
    resumen["clase_abc"] = np.where(
        resumen["pct_acumulado"] <= umbral_a, "A",
        np.where(resumen["pct_acumulado"] <= umbral_b, "B", "C")
    )

    # 11. Formatear porcentajes para legibilidad en Sheets
    resumen["pct_sobre_total"] = (resumen["pct_sobre_total"] * 100).round(2)
    resumen["pct_acumulado"]   = (resumen["pct_acumulado"]   * 100).round(2)

    print(f"OK ABC mensual generado: {len(resumen)} filas (cliente-periodo)")
    return resumen.reset_index(drop=True)

# ------------------------------------------------------------------------------
# QUERIES SQL
# ------------------------------------------------------------------------------
QUERY_SALDOS_CLIENTES_FILTRADOS = """
SELECT * FROM public.inpro2021nube_composicion_saldos_clientes_inprocil c
WHERE
    c.empresanombre = \'INPROCIL S.A.\'
AND
    c.cuentacontablecodigo IN (\'ANT101\', \'AAP301\', \'DML101\') AND
    c.clientenombre not like \'%%BENVENUTO%%\'  AND
    c.clientenombre not like \'%%CONCEPCION%%\' AND
    c.clientenombre not like \'%%BUIATTI%%\' AND
    c.clientenombre not like \'%%CAMPUZANO HORACIO DAVID%%\' AND
    c.clientenombre not like \'%%CONTIN %%\' AND
    c.clientenombre not like \'%%COOPERATIVA DE TRABAJO%%\' AND
    c.clientenombre not like \'%%DOMVIL%%\' AND
    c.clientenombre not like \'%%GAS MOVIL%%\' AND
    c.clientenombre not like \'%%GNC PATAGONICA%%\' AND
    c.clientenombre not like \'%%GOMEZ FABIAN%%\' AND
    c.clientenombre not like \'%%GOMEZ GUSTAVO%%\' AND
    c.clientenombre not like \'%%PALLETIZATE%%\' AND
    c.clientenombre not like \'%%PAUSYG%%\' AND
    c.clientenombre not like \'%%POWER CHECK%%\' AND
    c.clientenombre not like \'%%RODRIGUEZ ALEJANDRO%%\' AND
    c.clientenombre not like \'%%VALSI GAS%%\'
"""

QUERY_SALDOS_PROVEEDORES_FILTRADOS = """
select * from public.inpro2021nube_composicion_saldo_proveedores_inprocil c
"""

QUERY_STOCK_FILTRADO = """
SELECT * FROM public.inpro2021nube_stock_con_PUC
WHERE nombreproducto ILIKE \'Cilindro I%%\'
   OR nombreproducto ILIKE \'Corte%%\'
   OR nombreproducto ILIKE \'Conformado%%\'
   OR nombreproducto ILIKE \'Tratado%%\'
   OR nombreproducto ILIKE \'TUBO DE ACERO%%\'
   OR nombreproducto ILIKE \'%%EXPANDIDO%%\'
"""

QUERY_CONTROL_MRP = """
WITH BasePartes AS (
    SELECT
        ordendeproduccion,
        productoparteprod,
        NULLIF(NULLIF(TRIM(cantidadparteprod::text), \'\'), \'NULL\')::numeric AS cantidadparteprod,
        numerocomprobante,
        fecha
    FROM public.analisis_de_partes_de_produccion
    WHERE fecha::timestamp >= \'2026-01-01\'
      AND ordendeproduccion IS NOT NULL
      AND TRIM(ordendeproduccion::text) <> \'\'
      AND LOWER(productoparteprod::text) NOT LIKE \'%%scrap%%\'
      AND Empresa LIKE \'%%INPROCIL%%\'
),
ProduccionMensual AS (
    SELECT
        DATE_TRUNC(\'month\', fecha::timestamp) AS mes_produccion,
        productoparteprod,
        SUM(cantidadparteprod) AS total_cantidad_mes
    FROM BasePartes
    GROUP BY DATE_TRUNC(\'month\', fecha::timestamp), productoparteprod
),
Combinaciones AS (
    SELECT DISTINCT
        DATE_TRUNC(\'month\', p.fecha::timestamp) AS mes_produccion,
        c.fecha::timestamp AS fecha_exacta,
        p.productoparteprod::text AS productoparteprod,
        c.ordendeproduccion::text AS ordendeproduccion,
        p.numerocomprobante::text AS numerocomprobante_parte,
        c.numerocomprobante::text AS numerocomprobante_consumo,
        p.cantidadparteprod
    FROM public.analisis_de_consumos_de_produccion c
    INNER JOIN BasePartes p ON c.ordendeproduccion = p.ordendeproduccion
    WHERE c.fecha::timestamp >= \'2026-01-01\'
      AND c.ordendeproduccion IS NOT NULL
      AND TRIM(c.ordendeproduccion::text) <> \'\'
      AND c.Empresa LIKE \'%%INPROCIL%%\'
),
AbsorcionCalculada AS (
    SELECT
        a.fecha::timestamp AS fecha_absorcion,
        a.producto::text AS producto,
        a.cuentacontable::text AS cuentacontable,
        (SUM(NULLIF(NULLIF(TRIM(a.importepesos::text), \'\'), \'NULL\')::numeric) / NULLIF(pm.total_cantidad_mes, 0))::numeric AS importe_absorcion,
        pm.total_cantidad_mes
    FROM public.inpro2021nube_informe_absorcion_costos a
    INNER JOIN ProduccionMensual pm
        ON a.producto = pm.productoparteprod
        AND DATE_TRUNC(\'month\', a.fecha::timestamp) = pm.mes_produccion
    WHERE a.fecha::timestamp >= \'2026-01-01\'
      AND a.Empresa LIKE \'%%INPROCIL%%\'
    GROUP BY a.fecha::timestamp, a.producto, a.cuentacontable, pm.total_cantidad_mes
)
SELECT
    c.fecha::timestamp AS fecha,
    c.productoconsumoprod::text AS "Producto Consumido",
    NULLIF(NULLIF(TRIM(c.cantidadconsumoprod::text), \'\'), \'NULL\')::numeric AS cantidadconsumoprod,
    c.unidadconsumoprod::text AS unidadconsumoprod,
    NULLIF(NULLIF(TRIM(c.preciounitvalorizadoconsumoprod::text), \'\'), \'NULL\')::numeric AS preciounitvalorizadoconsumoprod,
    NULLIF(NULLIF(TRIM(c.importevalorizadoconsumoprod::text), \'\'), \'NULL\')::numeric AS importevalorizadoconsumoprod,
    (NULLIF(NULLIF(TRIM(c.importevalorizadoconsumoprod::text), \'\'), \'NULL\')::numeric / NULLIF(p.cantidadparteprod, 0))::numeric AS Importe,
    c.monedavalorizacionconsumoprod::text AS monedavalorizacionconsumoprod,
    p.cantidadparteprod AS cantidadparteprod,
    p.productoparteprod::text AS productoparteprod,
    c.ordendeproduccion::text AS ordendeproduccion,
    p.numerocomprobante::text AS numerocomprobante_parte,
    c.numerocomprobante::text AS numerocomprobante_consumo,
    pm.total_cantidad_mes::numeric AS "Total Producido",
    CASE
        WHEN p.productoparteprod ILIKE \'%%Corte%%\'      THEN \'Corte\'
        WHEN p.productoparteprod ILIKE \'%%Conformado%%\' THEN \'Conformado\'
        WHEN p.productoparteprod ILIKE \'%%Tratado%%\'    THEN \'Tratado\'
        WHEN p.productoparteprod ILIKE \'%%EXPANDIDO%%\'  THEN \'Expansion\'
        ELSE \'Producto Terminado\'
    END AS "Etapa Produccion",
    CASE
        WHEN c.productoconsumoprod ~* \'hora hombre\'    THEN \'Personal\'
        WHEN c.productoconsumoprod ~* \'energia|energía\' THEN \'Energia\'
        WHEN c.productoconsumoprod ~* \'cilindro|tubo\'  THEN \'Materia Prima\'
        WHEN c.productoconsumoprod ~* \'gas\'            THEN \'Gas\'
        ELSE \'Insumos de Produccion\'
    END AS "Categoria Insumo Produccion",
    SUBSTRING(p.productoparteprod::text FROM \'[0-9]{3}-[0-9]+\') AS "Codigo Producto",
    CASE
        WHEN p.productoparteprod ILIKE \'\%\% GA%%\' OR p.productoparteprod ILIKE \'\%\%-GA-%%\' THEN \'Gases del Aire\'
        WHEN p.productoparteprod ILIKE \'\%\% GN%%\' OR p.productoparteprod ILIKE \'\%\%-GN-%%\'
          OR p.productoparteprod ILIKE \'\%\% GNV%%\' OR p.productoparteprod ILIKE \'\%\%-GNV-%%\' THEN \'Gas Natural\'
        WHEN p.productoparteprod ILIKE \'%%TUBO EXPANDIDO%%\'   THEN \'Gas Natural\'
        WHEN p.productoparteprod ILIKE \'%%DIOXIDO DE CARBONO%%\'
          OR p.productoparteprod ILIKE \'%%DIOXIDO DE CARBONO%%\' THEN \'CO2\'
        ELSE \'Otros\'
    END AS "Familia de Gas"
FROM public.analisis_de_consumos_de_produccion c
INNER JOIN BasePartes p ON c.ordendeproduccion = p.ordendeproduccion
LEFT JOIN ProduccionMensual pm
    ON p.productoparteprod = pm.productoparteprod
    AND DATE_TRUNC(\'month\', p.fecha::timestamp) = pm.mes_produccion
WHERE c.fecha::timestamp >= \'2026-01-01\'
  AND c.ordendeproduccion IS NOT NULL
  AND TRIM(c.ordendeproduccion::text) <> \'\'
  AND c.Empresa LIKE \'%%INPROCIL%%\'

UNION ALL

SELECT
    comb.fecha_exacta AS fecha,
    ac.cuentacontable::text AS "Producto Consumido",
    NULL::numeric AS cantidadconsumoprod,
    NULL::text AS unidadconsumoprod,
    NULL::numeric AS preciounitvalorizadoconsumoprod,
    (ac.importe_absorcion * comb.cantidadparteprod)::numeric AS importevalorizadoconsumoprod,
    ac.importe_absorcion AS Importe,
    \'Pesos\'::text AS monedavalorizacionconsumoprod,
    comb.cantidadparteprod AS cantidadparteprod,
    ac.producto AS productoparteprod,
    comb.ordendeproduccion AS ordendeproduccion,
    comb.numerocomprobante_parte AS numerocomprobante_parte,
    comb.numerocomprobante_consumo AS numerocomprobante_consumo,
    ac.total_cantidad_mes::numeric AS "Total Producido",
    CASE
        WHEN ac.producto ILIKE \'%%Corte%%\'      THEN \'Corte\'
        WHEN ac.producto ILIKE \'%%Conformado%%\' THEN \'Conformado\'
        WHEN ac.producto ILIKE \'%%Tratado%%\'    THEN \'Tratado\'
        WHEN ac.producto ILIKE \'%%EXPANDIDO%%\'  THEN \'Expansion\'
        ELSE \'Producto Terminado\'
    END AS "Etapa Produccion",
    CASE
        WHEN ac.cuentacontable ~* \'hora hombre\'    THEN \'Personal\'
        WHEN ac.cuentacontable ~* \'energia|energía\' THEN \'Energia\'
        WHEN ac.cuentacontable ~* \'cilindro|tubo\'  THEN \'Materia Prima\'
        WHEN ac.cuentacontable ~* \'gas\'            THEN \'Gas\'
        ELSE \'Insumos de Produccion\'
    END AS "Categoria Insumo Produccion",
    SUBSTRING(ac.producto::text FROM \'[0-9]{3}-[0-9]+\') AS "Codigo Producto",
    CASE
        WHEN ac.producto ILIKE \'\%\% GA%%\' OR ac.producto ILIKE \'\%\%-GA-%%\' THEN \'Gases del Aire\'
        WHEN ac.producto ILIKE \'\%\% GN%%\' OR ac.producto ILIKE \'\%\%-GN-%%\'
          OR ac.producto ILIKE \'\%\% GNV%%\' OR ac.producto ILIKE \'\%\%-GNV-%%\' THEN \'Gas Natural\'
        WHEN ac.producto ILIKE \'%%TUBO EXPANDIDO%%\'   THEN \'Gas Natural\'
        WHEN ac.producto ILIKE \'%%DIOXIDO DE CARBONO%%\'
          OR ac.producto ILIKE \'%%DIOXIDO DE CARBONO%%\' THEN \'CO2\'
        ELSE \'Otros\'
    END AS "Familia de Gas"
FROM AbsorcionCalculada ac
INNER JOIN Combinaciones comb
    ON ac.producto = comb.productoparteprod
    AND DATE_TRUNC(\'month\', ac.fecha_absorcion) = comb.mes_produccion;
"""

# ------------------------------------------------------------------------------
# URLs de los Spreadsheets
# ------------------------------------------------------------------------------
SPREADSHEET_SALDOS_URL = os.environ.get(
    "SPREADSHEET_SALDOS_URL",
    "https://docs.google.com/spreadsheets/d/1oR_fdVCyn1cA8zwH4XgU5VK63cZaDC3I1i3-SWaUT20/edit"
)
SPREADSHEET_LIBRO_MAYOR_URL = os.environ.get(
    "SPREADSHEET_LIBRO_MAYOR_URL",
    "https://docs.google.com/spreadsheets/d/1e9BuGiiOx-GhokgsM37MAaUfddxLH30T-gtYu3UtfOA/edit"
)
SPREADSHEET_CMV_ID = "1e9BuGiiOx-GhokgsM37MAaUfddxLH30T-gtYu3UtfOA"

# ==============================================================================
# EJECUCION PRINCIPAL
# ==============================================================================
saldos_sheet = client.open_by_url(SPREADSHEET_SALDOS_URL)

# 0) Mapeo de clientes agrupados (impacta Churn, RFM y ABC)
mapa_clientes_agrupados = obtener_mapa_clientes_agrupados(saldos_sheet, sheet_name="AUX_Agrup_Clientes")

# 1. Saldos clientes filtrados
exportar_tabla_completa(
    QUERY_SALDOS_CLIENTES_FILTRADOS,
    saldos_sheet,
    "Base Saldos Clientes",
    ["importemonedatransaccion", "importemonedaprincipal", "importemonedasecundaria"]
)

# 2. Saldos proveedores
print("\\nEjecutando exportacion: Composicion Saldo Proveedores de INPROCIL S.A.")
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

# 5. Facturacion
# OPT: Carga UNICA de df_facturacion_full, reutilizado en Churn, RFM y ABC
print("\\nCargando facturacion completa desde DW...")
df_facturacion_full = pd.read_sql("SELECT * FROM public.inpro2021nube_facturacion", engine)
print(f"Facturacion total cargada: {len(df_facturacion_full)} filas")

exportar_tabla_completa(
    df_facturacion_full,
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

# 6. Analisis Control MRP
print("\\nEjecutando exportacion: Analisis Control MRP...")
exportar_tabla_completa(
    QUERY_CONTROL_MRP,
    saldos_sheet,
    "MRP",
    columnas_decimal=[
        "cantidadconsumoprod",
        "preciounitvalorizadoconsumoprod",
        "importevalorizadoconsumoprod",
        "Importe",
        "cantidadparteprod"
    ],
    create_if_missing=True
)

# 7. Churn
print("\\nEjecutando analisis de churn...")
df_facturacion_churn = obtener_datos_facturacion(
    df_facturacion_full=df_facturacion_full,
    mapa_clientes=mapa_clientes_agrupados
)
matriz_churn = crear_matriz_churn(df_facturacion_churn)

exportar_tabla_completa(
    matriz_churn,
    saldos_sheet,
    "Analisis_Churn",
    [],
    clear_range="A:D"
)

# 8. RFM
print("\\nCalculando RFM...")
usd_tc = obtener_tc_usd_desde_aux(saldos_sheet, sheet_name="AUX", range_name="A:B")
print(f"Tipo de cambio (ARS/USD) tomado de AUX: {usd_tc}")

df_rfm = calcular_rfm(
    df_facturacion_full=df_facturacion_full,
    mapa_clientes=mapa_clientes_agrupados,
    usd_tc_ars_por_usd=usd_tc,
    scoring_quantiles=5,
    cuenta_keyword=CUENTA_KEYWORD_VENTA
)
print(f"RFM generado: {len(df_rfm)} clientes agrupados")

exportar_tabla_completa(
    df_rfm,
    saldos_sheet,
    "RFM",
    columnas_decimal=[],
    clear_range="A:N",
    create_if_missing=True
)

# 9. NUEVO: ABC Mensual por Cliente
print("\\nCalculando ABC mensual por cliente...")
df_abc = calcular_abc_mensual(
    df_facturacion_full=df_facturacion_full,
    mapa_clientes=mapa_clientes_agrupados,
    umbral_a=0.80,
    umbral_b=0.95,
    cuenta_keyword=CUENTA_KEYWORD_VENTA
)

exportar_tabla_completa(
    df_abc,
    saldos_sheet,
    "ABC_Mensual",
    columnas_decimal=["facturacion_mes", "pct_sobre_total", "pct_acumulado"],
    clear_range="A:F",
    create_if_missing=True
)

# ------------------------------------------------------------------------------
# Spreadsheet 2: Libro Mayor, Stock, Sumas y Saldos
# ------------------------------------------------------------------------------
libro_mayor_sheet = client.open_by_url(SPREADSHEET_LIBRO_MAYOR_URL)

exportar_libro_mayor(
    "SELECT * FROM public.inpro2021nube_libro_mayor",
    libro_mayor_sheet,
    "Aux Libro Mayor",
    ["Debe", "Haber", "importemonedaprincipal", "Imp. operacion ppal.", "Imp. operacion sec.", "Tipo Cambio"],
)

exportar_stock(
    QUERY_STOCK_FILTRADO,
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
# Spreadsheet 3: CMV y Costos Partes de Produccion
# ------------------------------------------------------------------------------
cmv_sheet = client.open_by_key(SPREADSHEET_CMV_ID)

print("\\nEjecutando exportacion: CMV...")
df_cmv = pd.read_sql("SELECT * FROM public.inpro2021nube_cmv", engine)
df_cmv_recortado = df_cmv.iloc[:, :15]

exportar_tabla_completa(
    query_or_df=df_cmv_recortado,
    spreadsheet=cmv_sheet,
    hoja_nombre="AUX_CMV",
    columnas_decimal=["importe"],
    clear_range="A:O",
    create_if_missing=True
)

print("\\nEjecutando exportacion: Costos Partes de Produccion...")
df_costos_partes = pd.read_sql("SELECT * FROM public.inpro2021nube_costos_partes_de_produccion", engine)

exportar_tabla_completa(
    query_or_df=df_costos_partes,
    spreadsheet=cmv_sheet,
    hoja_nombre="Aux Costo Cilindros",
    columnas_decimal=["importe"],
    clear_range="A:O",
    create_if_missing=True
)

print("\\nPROCESO COMPLETO FINALIZADO")
'''

output_path = os.path.expanduser('~/output/script_total_con_abc.py')
with open(output_path, 'w', encoding='utf-8-sig') as f:
    f.write(script)

print(f"Archivo guardado: {output_path}")
print(f"Tamano: {len(script)} caracteres, {script.count(chr(10))} lineas")
