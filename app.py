from datetime import datetime, date
import sqlite3
import pandas as pd
import streamlit as st

# Configuración de página
st.set_page_config(
    page_title="LIFRAN - Control Operativo Portuario",
    page_icon="⚓",
    layout="wide",
)

DB_NAME = "lifran.db"
EXCEL_PATH = "LIFRAN.xlsx"


# Inicializar Base de Datos SQLite desde el Excel si no existe
@st.cache_resource
def init_db():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(
      "SELECT name FROM sqlite_master WHERE type='table' AND name='control';"
  )
  exists = cursor.fetchone()

  if not exists:
    xls = pd.ExcelFile(EXCEL_PATH)
    for sheet in xls.sheet_names:
      df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
      # Convertir todas las columnas de texto a mayúsculas
      for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.upper()
      df.columns = [c.upper() for c in df.columns]
      df.to_sql(sheet.lower(), conn, if_exists="replace", index=False)

  # Tabla de configuración de empresa para los recibos si no existe
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS empresa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT, rif TEXT, telefono TEXT, direccion TEXT, 
            pie_pagina TEXT
        )
    """)
  cursor.execute("SELECT COUNT(*) FROM empresa")
  if cursor.fetchone()[0] == 0:
    cursor.execute(
        "INSERT INTO empresa (nombre, rif, telefono, direccion, pie_pagina) "
        "VALUES (?, ?, ?, ?, ?)",
        (
            "SERVICIOS Y SUMINISTROS ARO, C.A.",
            "J-XXXXXXXX-X",
            "+58 414-5687324",
            "BARINAS, VENEZUELA",
            "GRACIAS POR CONFIAR EN NUESTROS SERVICIOS PORTUARIOS.",
        ),
    )

  conn.commit()
  conn.close()


init_db()


def run_query(query, params=()):
  conn = sqlite3.connect(DB_NAME)
  df = pd.read_sql_query(query, conn, params=params)
  conn.close()
  return df


def execute_query(query, params=()):
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()
  cursor.execute(query, params)
  conn.commit()
  conn.close()


# --- MENÚ LATERAL ---
st.sidebar.title("⚓ LIFRAN NAVEGACIÓN")
menu_opciones = [
    "CONTROL OPERATIVO",
    "CLIENTES",
    "LUGARES",
    "BUQUES",
    "MUELLES",
    "REPORTES Y RESUMEN",
    "CONFIGURACIÓN RECIBO WHATSAPP",
]
choice = st.sidebar.selectbox("SELECCIONE VISTA", menu_opciones)

st.title(f"MÓDULO: {choice}")

# -----------------------------------------------------------------------------
# 1. CONTROL OPERATIVO (ALQUILER DE BAÑOS EN USD)
# -----------------------------------------------------------------------------
if choice == "CONTROL OPERATIVO":
  st.subheader("REGISTRO Y CONTROL DE ALQUILER DE BAÑOS PORTÁTILES (USD)")

  with st.form("form_control", clear_on_submit=True):
    col1, col2 = st.columns(2)

    df_cli = run_query("SELECT IDCLIENTE, NOMBRE FROM cliente")
    df_lug = run_query("SELECT IDLUGAR, LUGAR FROM lugar")
    df_buq = run_query("SELECT IDBUQUE, NOMBREBUQUE FROM buque")
    df_mue = run_query("SELECT IDMUELLE, NUMEROMUELLE FROM muelle")

    with col1:
      cli_dict = (
          dict(zip(df_cli["NOMBRE"], df_cli["IDCLIENTE"]))
          if not df_cli.empty
          else {}
      )
      sel_cli = st.selectbox(
          "CLIENTE", options=list(cli_dict.keys()) if cli_dict else ["REGISTRE CLIENTES"]
      )

      lug_dict = (
          dict(zip(df_lug["LUGAR"], df_lug["IDLUGAR"]))
          if not df_lug.empty
          else {}
      )
      sel_lug = st.selectbox(
          "LUGAR / PUERTO", options=list(lug_dict.keys()) if lug_dict else ["REGISTRE LUGARES"]
      )

      buq_dict = (
          dict(zip(df_buq["NOMBREBUQUE"], df_buq["IDBUQUE"]))
          if not df_buq.empty
          else {}
      )
      sel_buq = st.selectbox(
          "BUQUE", options=list(buq_dict.keys()) if buq_dict else ["REGISTRE BUQUES"]
      )

    with col2:
      mue_dict = (
          dict(
              zip(
                  df_mue["NUMEROMUELLE"].astype(str), df_mue["IDMUELLE"]
              )
          )
          if not df_mue.empty
          else {}
      )
      sel_mue = st.selectbox(
          "MUELLE", options=list(mue_dict.keys()) if mue_dict else ["REGISTRE MUELLES"]
      )

      f_inicio = st.date_input("FECHA DE INICIO", date.today())
      f_terminacion = st.date_input("FECHA DE TERMINACIÓN", date.today())
      num_cabinas = st.number_input(
          "CANTIDAD DE BAÑOS / CABINAS", min_value=1, value=1
      )
      tarifa_usd = st.number_input(
          "TARIFA DIARIA POR BAÑO (USD $)", min_value=0.0, value=25.0
      )

    observaciones = st.text_area("OBSERVACIONES").upper()
    submitted = st.form_submit_button("GUARDAR NUEVO OPERATIVO")

    if submitted:
      # Generar ID único corto en mayúsculas
      import uuid

      id_control = str(uuid.uuid4())[:8].upper()
      c_id = cli_dict[sel_cli]
      l_id = lug_dict[sel_lug]
      b_id = buq_dict[sel_buq]
      m_id = mue_dict[sel_mue]

      execute_query(
          """INSERT INTO control (IDCONTROL, CLIENTE, LUGAR, BUQUE, MUELLE, FECHAINICIO, FECHATERMINACION, NUMEROCABINAS, OBSERVACIONES)
                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
          (
              id_control,
              c_id,
              l_id,
              b_id,
              m_id,
              str(f_inicio),
              str(f_terminacion),
              num_cabinas,
              observaciones,
          ),
      )
      st.success("¡OPERACIÓN REGISTRADA CON ÉXITO!")

  st.divider()
  st.subheader("HISTORIAL DE OPERACIONES Y ENVÍO DE RECIBO POR WHATSAPP")

  query_sql = """
        SELECT c.IDCONTROL, cl.NOMBRE as CLIENTE_NOMBRE, l.LUGAR as LUGAR_NOMBRE, 
               b.NOMBREBUQUE, m.NUMEROMUELLE, c.FECHAINICIO, c.FECHATERMINACION, 
               c.NUMEROCABINAS, c.OBSERVACIONES, cl.TELEFONO
        FROM control c
        LEFT JOIN cliente cl ON c.CLIENTE = cl.IDCLIENTE
        LEFT JOIN lugar l ON c.LUGAR = l.IDLUGAR
        LEFT JOIN buque b ON c.BUQUE = b.IDBUQUE
        LEFT JOIN muelle m ON c.MUELLE = m.IDMUELLE
    """
  df_ops = run_query(query_sql)

  if not df_ops.empty:
    # Formatear fechas a DD/MM/YY para visualización
    df_ops["FECHAINICIO_DT"] = pd.to_datetime(
        df_ops["FECHAINICIO"], errors="coerce"
    )
    df_ops["FECHATERMINACION_DT"] = pd.to_datetime(
        df_ops["FECHATERMINACION"], errors="coerce"
    )

    df_ops["DIAS"] = (
        df_ops["FECHATERMINACION_DT"] - df_ops["FECHAINICIO_DT"]
    ).dt.days + 1
    df_ops["DIAS"] = df_ops["DIAS"].fillna(1).apply(lambda x: max(1, x))
    df_ops["TARIFA_USD"] = 25.0  # Tarifa estándar editable o fija
    df_ops["TOTAL_USD"] = (
        df_ops["DIAS"] * df_ops["NUMEROCABINAS"].fillna(1) * df_ops["TARIFA_USD"]
    )

    # Mostrar dataframe con formato de fecha DD/MM/YY
    display_df = df_ops.copy()
    display_df["FECHAINICIO"] = display_df["FECHAINICIO_DT"].dt.strftime(
        "%d/%m/%y"
    )
    display_df["FECHATERMINACION"] = display_df[
        "FECHATERMINACION_DT"
    ].dt.strftime("%d/%m/%y")

    st.dataframe(
        display_df[
            [
                "IDCONTROL",
                "CLIENTE_NOMBRE",
                "LUGAR_NOMBRE",
                "NOMBREBUQUE",
                "NUMEROMUELLE",
                "FECHAINICIO",
                "FECHATERMINACION",
                "NUMEROCABINAS",
                "DIAS",
                "TOTAL_USD",
                "OBSERVACIONES",
            ]
        ]
    )

    st.markdown("### 📱 GENERAR Y ENVIAR RECIBO POR WHATSAPP")
    op_ids = df_ops["IDCONTROL"].tolist()
    sel_op = st.selectbox("SELECCIONE ID DE CONTROL PARA RECIBO", op_ids)

    if sel_op:
      row = df_ops[df_ops["IDCONTROL"] == sel_op].iloc[0]
      empresa = run_query("SELECT * FROM empresa LIMIT 1").iloc[0]

      f_ini_str = (
          row["FECHAINICIO_DT"].strftime("%d/%m/%y")
          if pd.notnull(row["FECHAINICIO_DT"])
          else "N/D"
      )
      f_fin_str = (
          row["FECHATERMINACION_DT"].strftime("%d/%m/%y")
          if pd.notnull(row["FECHATERMINACION_DT"])
          else "N/D"
      )

      recibo_msj = f"""*--- {empresa['NOMBRE']} ---*
RIF: {empresa['RIF']} | TLF: {empresa['TELEFONO']}
*NOTA DE ENTREGA / RECIBO DE ALQUILER*
---------------------------------------
*CLIENTE:* {row['CLIENTE_NOMBRE']}
*PUERTO / LUGAR:* {row['LUGAR_NOMBRE']}
*BUQUE:* {row['NOMBREBUQUE']} (MUELLE N° {row['NUMEROMUELLE']})
*PERÍODO:* {f_ini_str} AL {f_fin_str} ({int(row['DIAS'])} DÍAS)
*CANTIDAD DE BAÑOS:* {int(row['NUMEROCABINAS']) if pd.notnull(row['NUMEROCABINAS']) else 1}
*TARIFA DIARIA:* ${row['TARIFA_USD']:.2f} USD
---------------------------------------
*TOTAL A PAGAR: ${row['TOTAL_USD']:.2f} USD*
*OBSERVACIONES:* {row['OBSERVACIONES']}
---------------------------------------
{empresa['PIE_PAGINA']}"""

      st.text_area("VISTA PREVIA DEL RECIBO:", recibo_msj, height=220)

      telefono = (
          str(row["TELEFONO"]).strip().replace(" ", "").replace("+", "")
      )
      import urllib.parse

      whatsapp_link = (
          f"https://wa.me/{telefono}?text={urllib.parse.quote(recibo_msj)}"
      )

      st.markdown(
          f'<a href="{whatsapp_link}" target="_blank" style="display:inline-block;padding:12px 24px;background-color:#25D366;color:white;text-decoration:none;border-radius:6px;font-weight:bold;font-size:16px;">📲 ENVIAR RECIBO POR WHATSAPP</a>',
          unsafe_allow_html=True,
      )

# -----------------------------------------------------------------------------
# 2. CLIENTES
# -----------------------------------------------------------------------------
elif choice == "CLIENTES":
  st.subheader("ADMINISTRACIÓN DE CLIENTES")
  df_cli = run_query("SELECT * FROM cliente")
  ed_cli = st.data_editor(df_cli, num_rows="dynamic", use_container_width=True)
  if st.button("GUARDAR CAMBIOS EN CLIENTES"):
    for col in ed_cli.select_dtypes(include=["object"]).columns:
      ed_cli[col] = ed_cli[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_cli.to_sql("cliente", conn, if_exists="replace", index=False)
    conn.close()
    st.success("CLIENTES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 3. LUGARES
# -----------------------------------------------------------------------------
elif choice == "LUGARES":
  st.subheader("ADMINISTRACIÓN DE LUGARES / PUERTOS")
  df_lug = run_query("SELECT * FROM lugar")
  ed_lug = st.data_editor(df_lug, num_rows="dynamic", use_container_width=True)
  if st.button("GUARDAR CAMBIOS EN LUGARES"):
    for col in ed_lug.select_dtypes(include=["object"]).columns:
      ed_lug[col] = ed_lug[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_lug.to_sql("lugar", conn, if_exists="replace", index=False)
    conn.close()
    st.success("LUGARES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 4. BUQUES
# -----------------------------------------------------------------------------
elif choice == "BUQUES":
  st.subheader("ADMINISTRACIÓN DE BUQUES")
  df_buq = run_query("SELECT * FROM buque")
  ed_buq = st.data_editor(df_buq, num_rows="dynamic", use_container_width=True)
  if st.button("GUARDAR CAMBIOS EN BUQUES"):
    for col in ed_buq.select_dtypes(include=["object"]).columns:
      ed_buq[col] = ed_buq[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_buq.to_sql("buque", conn, if_exists="replace", index=False)
    conn.close()
    st.success("BUQUES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 5. MUELLES
# -----------------------------------------------------------------------------
elif choice == "MUELLES":
  st.subheader("ADMINISTRACIÓN DE MUELLES")
  df_mue = run_query("SELECT * FROM muelle")
  ed_mue = st.data_editor(df_mue, num_rows="dynamic", use_container_width=True)
  if st.button("GUARDAR CAMBIOS EN MUELLES"):
    for col in ed_mue.select_dtypes(include=["object"]).columns:
      ed_mue[col] = ed_mue[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_mue.to_sql("muelle", conn, if_exists="replace", index=False)
    conn.close()
    st.success("MUELLES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 6. REPORTES Y RESUMEN POR PERIODO
# -----------------------------------------------------------------------------
elif choice == "REPORTES Y RESUMEN":
  st.subheader("RESUMEN DE OPERACIONES Y ESTADÍSTICAS EN USD")

  query_sql = """
        SELECT c.IDCONTROL, cl.NOMBRE as CLIENTE_NOMBRE, l.LUGAR as LUGAR_NOMBRE, 
               b.NOMBREBUQUE, c.FECHAINICIO, c.FECHATERMINACION, c.NUMEROCABINAS
        FROM control c
        LEFT JOIN cliente cl ON c.CLIENTE = cl.IDCLIENTE
        LEFT JOIN lugar l ON c.LUGAR = l.IDLUGAR
        LEFT JOIN buque b ON c.BUQUE = b.IDBUQUE
    """
  df_rep = run_query(query_sql)

  if not df_rep.empty:
    df_rep["FECHAINICIO_DT"] = pd.to_datetime(
        df_rep["FECHAINICIO"], errors="coerce"
    )
    df_rep["FECHATERMINACION_DT"] = pd.to_datetime(
        df_rep["FECHATERMINACION"], errors="coerce"
    )
    df_rep["DIAS"] = (
        df_rep["FECHATERMINACION_DT"] - df_rep["FECHAINICIO_DT"]
    ).dt.days + 1
    df_rep["DIAS"] = df_rep["DIAS"].fillna(1).apply(lambda x: max(1, x))
    df_rep["TARIFA_USD"] = 25.0
    df_rep["TOTAL_USD"] = (
        df_rep["DIAS"] * df_rep["NUMEROCABINAS"].fillna(1) * df_rep["TARIFA_USD"]
    )

    c1, c2, c3 = st.columns(3)
    c1.metric(
        "INGRESOS TOTAles ACUMULADOS", f"${df_rep['TOTAL_USD'].sum():,.2f} USD"
    )
    c2.metric("OPERACIONES REGISTRADAS", len(df_rep))
    c3.metric(
        "PROMEDIO BAÑOS POR OPERACIÓN",
        f"{df_rep['NUMEROCABINAS'].mean():.1f}",
    )

    st.divider()
    st.markdown("### INGRESOS POR CLIENTE")
    res_cli = df_rep.groupby("CLIENTE_NOMBRE")["TOTAL_USD"].sum().reset_index()
    st.bar_chart(res_cli.set_index("CLIENTE_NOMBRE"))
    st.dataframe(res_cli)

    st.markdown("### INGRESOS POR PUERTO / LUGAR")
    res_lug = df_rep.groupby("LUGAR_NOMBRE")["TOTAL_USD"].sum().reset_index()
    st.dataframe(res_lug)
  else:
    st.info("NO HAY DATOS SUFICIENTES PARA LOS REPORTES.")

# -----------------------------------------------------------------------------
# 7. CONFIGURACIÓN DE RECIBO WHATSAPP
# -----------------------------------------------------------------------------
elif choice == "CONFIGURACIÓN RECIBO WHATSAPP":
  st.subheader("CONFIGURACIÓN DE DATOS FISCALES DE LA EMPRESA")
  st.markdown(
      "MODIFIQUE LOS DATOS DE LA EMPRESA QUE SALDRÁN REFLEJADOS EN EL RECIBO"
      " ENVIADO POR WHATSAPP."
  )

  df_emp = run_query("SELECT * FROM empresa LIMIT 1")
  if not df_emp.empty:
    row_e = df_emp.iloc[0]
    with st.form("form_empresa"):
        n_emp = st.text_input("NOMBRE DE LA EMPRESA", value=row_e["NOMBRE"])
        n_rif = st.text_input("RIF", value=row_e["RIF"])
        n_tlf = st.text_input("TELÉFONO DE CONTACTO", value=row_e["TELEFONO"])
        n_dir = st.text_input("DIRECCIÓN", value=row_e["DIRECCION"])
        n_pie = st.text_area(
            "PIE DE PÁGINA / MENSAJE FINAL", value=row_e["PIE_PAGINA"]
        )

        if st.form_submit_button("ACTUALIZAR DATOS DE EMPRESA"):
          execute_query(
              "UPDATE empresa SET NOMBRE=?, RIF=?, TELEFONO=?, DIRECCION=?, "
              "PIE_PAGINA=? WHERE id=?",
              (
                  n_emp.upper(),
                  n_rif.upper(),
                  n_tlf.upper(),
                  n_dir.upper(),
                  n_pie.upper(),
                  row_e["id"],
              ),
          )
          st.success("¡DATOS DE LA EMPRESA ACTUALIZADOS EXITOSAMENTE!")