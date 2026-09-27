from datetime import datetime, date
import sqlite3
import urllib.parse
import uuid
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


# Inicializar Base de Datos SQLite de forma robusta
@st.cache_resource
def init_db():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

  # 1. Asegurar tabla EMPRESA con esquema completo
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS empresa (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            NOMBRE TEXT, 
            RIF TEXT, 
            TELEFONO TEXT, 
            DIRECCION TEXT, 
            PIE_PAGINA TEXT
        )
    """)

  cursor.execute("SELECT COUNT(*) FROM empresa")
  if cursor.fetchone()[0] == 0:
    cursor.execute(
        "INSERT INTO empresa (NOMBRE, RIF, TELEFONO, DIRECCION, PIE_PAGINA) "
        "VALUES (?, ?, ?, ?, ?)",
        (
            "SERVICIOS Y SUMINISTROS ARO, C.A.",
            "J-XXXXXXXX-X",
            "+58 414-5687324",
            "BARINAS, VENEZUELA",
            "GRACIAS POR CONFIAR EN NUESTROS SERVICIOS PORTUARIOS.",
        ),
    )

  # 2. Cargar tablas desde Excel si no existen en SQLite
  cursor.execute(
      "SELECT name FROM sqlite_master WHERE type='table' AND name='control';"
  )
  exists = cursor.fetchone()

  if not exists:
    xls = pd.ExcelFile(EXCEL_PATH)
    for sheet in xls.sheet_names:
      df = pd.read_excel(EXCEL_PATH, sheet_name=sheet)
      for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].astype(str).str.upper()
      df.columns = [c.upper() for c in df.columns]
      df.to_sql(sheet.lower(), conn, if_exists="replace", index=False)

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
    "ELIMINAR REGISTROS",
    "REPORTES Y RESUMEN",
    "CONFIGURACIÓN RECIBO WHATSAPP",
]
choice = st.sidebar.selectbox("SELECCIONE VISTA", menu_opciones)

st.title(f"MÓDULO: {choice}")

# -----------------------------------------------------------------------------
# 1. CONTROL OPERATIVO
# -----------------------------------------------------------------------------
if choice == "CONTROL OPERATIVO":
  st.subheader("REGISTRO Y CONTROL DE ALQUILER DE BAÑOS PORTÁTILES (USD)")

  df_cli = run_query(
      "SELECT IDCLIENTE, NOMBRE, RIF, DIRECCION, TELEFONO FROM cliente"
  )
  df_lug = run_query("SELECT IDLUGAR, LUGAR FROM lugar")
  df_buq = run_query("SELECT IDBUQUE, NOMBREBUQUE FROM buque")
  df_mue = run_query("SELECT IDMUELLE, NUMEROMUELLE FROM muelle")

  with st.form("form_control"):
    col1, col2 = st.columns(2)

    with col1:
      st.markdown("### 🏢 CLIENTE")
      is_nuevo_cli = st.checkbox("➕ AGREGAR NUEVO CLIENTE")

      if is_nuevo_cli:
        nuevo_cli_input = st.text_input("NOMBRE O RAZÓN SOCIAL (NUEVO)").upper()
        nuevo_rif_input = st.text_input("RIF (NUEVO)").upper()
        nuevo_dir_input = st.text_input("DIRECCIÓN (NUEVA)").upper()
        nuevo_tlf_input = st.text_input("TELÉFONO (NUEVO)").upper()
        sel_cli = None
      else:
        lista_clientes = (
            df_cli["NOMBRE"].tolist() if not df_cli.empty else []
        )
        sel_cli = st.selectbox(
            "SELECCIONE CLIENTE EXISTENTE", options=lista_clientes
        )
        nuevo_cli_input, nuevo_rif_input, nuevo_dir_input, nuevo_tlf_input = (
            "",
            "",
            "",
            "",
        )

        if sel_cli and not df_cli.empty:
          cli_row = df_cli[df_cli["NOMBRE"] == sel_cli]
          if not cli_row.empty:
            r = cli_row.iloc[0]
            st.info(
                f"📌 **DATOS DEL CLIENTE:**\n\n"
                f"- **RIF:** {r.get('RIF', 'N/D')}\n"
                f"- **DIRECCIÓN:** {r.get('DIRECCION', 'N/D')}\n"
                f"- **TELÉFONO:** {r.get('TELÉFONO', r.get('TELEFONO', 'N/D'))}"
            )

      st.markdown("---")
      st.markdown("### 📍 LUGAR / PUERTO")
      is_nuevo_lug = st.checkbox("➕ AGREGAR NUEVO LUGAR")

      if is_nuevo_lug:
        nuevo_lug_input = st.text_input("NOMBRE DEL LUGAR / PUERTO (NUEVO)").upper()
        sel_lug = None
      else:
        lista_lugares = df_lug["LUGAR"].tolist() if not df_lug.empty else []
        sel_lug = st.selectbox(
            "SELECCIONE LUGAR EXISTENTE", options=lista_lugares
        )
        nuevo_lug_input = ""

      st.markdown("---")
      st.markdown("### 🚢 BUQUE")
      is_nuevo_buq = st.checkbox("➕ AGREGAR NUEVO BUQUE")

      if is_nuevo_buq:
        nuevo_buq_input = st.text_input("NOMBRE DEL BUQUE (NUEVO)").upper()
        sel_buq = None
      else:
        lista_buques = (
            df_buq["NOMBREBUQUE"].tolist() if not df_buq.empty else []
        )
        sel_buq = st.selectbox(
            "SELECCIONE BUQUE EXISTENTE", options=lista_buques
        )
        nuevo_buq_input = ""

    with col2:
      st.markdown("### 🏗️ MUELLE")
      is_nuevo_mue = st.checkbox("➕ AGREGAR NUEVO MUELLE")

      if is_nuevo_mue:
        nuevo_mue_input = st.text_input("NÚMERO DE MUELLE (NUEVO)").upper()
        sel_mue = None
      else:
        lista_muelles = (
            df_mue["NUMEROMUELLE"].astype(str).tolist()
            if not df_mue.empty
            else []
        )
        sel_mue = st.selectbox(
            "SELECCIONE MUELLE EXISTENTE", options=lista_muelles
        )
        nuevo_mue_input = ""

      st.markdown("---")
      st.markdown("### 📅 DATOS OPERATIVOS Y TARIFAS")
      # Fechas (Formato visual DD/MM/AA)
      f_inicio = st.date_input("FECHA DE INICIO (DD/MM/AA)", date.today())
      f_terminacion = st.date_input(
          "FECHA DE TERMINACIÓN (DD/MM/AA)", date.today()
      )
      num_cabinas = st.number_input(
          "CANTIDAD DE BAÑOS / CABINAS", min_value=1, value=1
      )
      tarifa_usd = st.number_input(
          "TARIFA DIARIA POR BAÑO (USD $)", min_value=0.0, value=25.0
      )

    observaciones = st.text_area("OBSERVACIONES").upper()
    submitted = st.form_submit_button("GUARDAR NUEVO OPERATIVO")

    if submitted:
      # 1. Procesar Cliente
      if is_nuevo_cli:
        if not nuevo_cli_input:
          st.error("EL NOMBRE DEL NUEVO CLIENTE ES OBLIGATORIO.")
          st.stop()
        c_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO cliente (IDCLIENTE, NOMBRE, RIF, DIRECCION, TELEFONO)"
            " VALUES (?, ?, ?, ?, ?)",
            (
                c_id,
                nuevo_cli_input,
                nuevo_rif_input,
                nuevo_dir_input,
                nuevo_tlf_input,
            ),
        )
      else:
        if not sel_cli:
          st.error("DEBE SELECCIONAR UN CLIENTE.")
          st.stop()
        c_id = df_cli[df_cli["NOMBRE"] == sel_cli]["IDCLIENTE"].values[0]

      # 2. Procesar Lugar
      if is_nuevo_lug:
        if not nuevo_lug_input:
          st.error("EL NOMBRE DEL NUEVO LUGAR ES OBLIGATORIO.")
          st.stop()
        l_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO lugar (IDLUGAR, LUGAR) VALUES (?, ?)",
            (l_id, nuevo_lug_input),
        )
      else:
        if not sel_lug:
          st.error("DEBE SELECCIONAR UN LUGAR.")
          st.stop()
        l_id = df_lug[df_lug["LUGAR"] == sel_lug]["IDLUGAR"].values[0]

      # 3. Procesar Buque
      if is_nuevo_buq:
        if not nuevo_buq_input:
          st.error("EL NOMBRE DEL NUEVO BUQUE ES OBLIGATORIO.")
          st.stop()
        b_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO buque (IDBUQUE, NOMBREBUQUE) VALUES (?, ?)",
            (b_id, nuevo_buq_input),
        )
      else:
        if not sel_buq:
          st.error("DEBE SELECCIONAR UN BUQUE.")
          st.stop()
        b_id = df_buq[df_buq["NOMBREBUQUE"] == sel_buq]["IDBUQUE"].values[0]

      # 4. Procesar Muelle
      if is_nuevo_mue:
        if not nuevo_mue_input:
          st.error("EL NÚMERO DEL NUEVO MUELLE ES OBLIGATORIO.")
          st.stop()
        m_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO muelle (IDMUELLE, NUMEROMUELLE) VALUES (?, ?)",
            (m_id, nuevo_mue_input),
        )
      else:
        if not sel_mue:
          st.error("DEBE SELECCIONAR UN MUELLE.")
          st.stop()
        m_id = df_mue[df_mue["NUMEROMUELLE"].astype(str) == sel_mue][
            "IDMUELLE"
        ].values[0]

      # 5. Guardar Operación
      id_control = str(uuid.uuid4())[:8].upper()
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
      st.success("¡OPERACIÓN REGISTRADA EXITOSAMENTE!")
      st.rerun()

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
    df_ops["TARIFA_USD"] = 25.0
    df_ops["TOTAL_USD"] = (
        df_ops["DIAS"] * df_ops["NUMEROCABINAS"].fillna(1) * df_ops["TARIFA_USD"]
    )

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
  st.subheader("ADMINISTRACIÓN Y REGISTRO DE CLIENTES")

  with st.expander("➕ AGREGAR NUEVO CLIENTE", expanded=False):
    with st.form("form_nuevo_cliente", clear_on_submit=True):
      c_nombre = st.text_input("NOMBRE O RAZÓN SOCIAL").upper()
      c_rif = st.text_input("RIF").upper()
      c_dir = st.text_input("DIRECCIÓN").upper()
      c_tlf = st.text_input("TELÉFONO").upper()
      if st.form_submit_button("GUARDAR CLIENTE"):
        if c_nombre and c_rif:
          new_id = str(uuid.uuid4())[:8].upper()
          execute_query(
              "INSERT INTO cliente (IDCLIENTE, NOMBRE, RIF, DIRECCION,"
              " TELEFONO) VALUES (?, ?, ?, ?, ?)",
              (new_id, c_nombre, c_rif, c_dir, c_tlf),
          )
          st.success("¡CLIENTE AGREGADO EXITOSAMENTE!")
          st.rerun()
        else:
          st.error("EL NOMBRE Y EL RIF SON OBLIGATORIOS.")

  st.markdown("### LISTADO GENERAL DE CLIENTES")
  df_cli = run_query("SELECT * FROM cliente")
  ed_cli = st.data_editor(df_cli, num_rows="dynamic", use_container_width=True)
  if st.button("ACTUALIZAR CAMBIOS EN TABLA CLIENTES"):
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
  st.subheader("ADMINISTRACIÓN Y REGISTRO DE LUGARES / PUERTOS")

  with st.expander("➕ AGREGAR NUEVO LUGAR", expanded=False):
    with st.form("form_nuevo_lugar", clear_on_submit=True):
      l_nombre = st.text_input("NOMBRE DEL LUGAR / PUERTO").upper()
      if st.form_submit_button("GUARDAR LUGAR"):
        if l_nombre:
          new_id = str(uuid.uuid4())[:8].upper()
          execute_query(
              "INSERT INTO lugar (IDLUGAR, LUGAR) VALUES (?, ?)",
              (new_id, l_nombre),
          )
          st.success("¡LUGAR AGREGADO EXITOSAMENTE!")
          st.rerun()
        else:
          st.error("EL NOMBRE DEL LUGAR ES OBLIGATORIO.")

  st.markdown("### LISTADO GENERAL DE LUGARES")
  df_lug = run_query("SELECT * FROM lugar")
  ed_lug = st.data_editor(df_lug, num_rows="dynamic", use_container_width=True)
  if st.button("ACTUALIZAR CAMBIOS EN TABLA LUGARES"):
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
  st.subheader("ADMINISTRACIÓN Y REGISTRO DE BUQUES")

  with st.expander("➕ AGREGAR NUEVO BUQUE", expanded=False):
    with st.form("form_nuevo_buque", clear_on_submit=True):
      b_nombre = st.text_input("NOMBRE DEL BUQUE").upper()
      if st.form_submit_button("GUARDAR BUQUE"):
        if b_nombre:
          new_id = str(uuid.uuid4())[:8].upper()
          execute_query(
              "INSERT INTO buque (IDBUQUE, NOMBREBUQUE) VALUES (?, ?)",
              (new_id, b_nombre),
          )
          st.success("¡BUQUE AGREGADO EXITOSAMENTE!")
          st.rerun()
        else:
          st.error("EL NOMBRE DEL BUQUE ES OBLIGATORIO.")

  st.markdown("### LISTADO GENERAL DE BUQUES")
  df_buq = run_query("SELECT * FROM buque")
  ed_buq = st.data_editor(df_buq, num_rows="dynamic", use_container_width=True)
  if st.button("ACTUALIZAR CAMBIOS EN TABLA BUQUES"):
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
  st.subheader("ADMINISTRACIÓN Y REGISTRO DE MUELLES")

  with st.expander("➕ AGREGAR NUEVO MUELLE", expanded=False):
    with st.form("form_nuevo_muelle", clear_on_submit=True):
      m_num = st.text_input("NÚMERO / IDENTIFICADOR DE MUELLE").upper()
      if st.form_submit_button("GUARDAR MUELLE"):
        if m_num:
          new_id = str(uuid.uuid4())[:8].upper()
          execute_query(
              "INSERT INTO muelle (IDMUELLE, NUMEROMUELLE) VALUES (?, ?)",
              (new_id, m_num),
          )
          st.success("¡MUELLE AGREGADO EXITOSAMENTE!")
          st.rerun()
        else:
          st.error("EL NÚMERO DE MUELLE ES OBLIGATORIO.")

  st.markdown("### LISTADO GENERAL DE MUELLES")
  df_mue = run_query("SELECT * FROM muelle")
  ed_mue = st.data_editor(df_mue, num_rows="dynamic", use_container_width=True)
  if st.button("ACTUALIZAR CAMBIOS EN TABLA MUELLES"):
    for col in ed_mue.select_dtypes(include=["object"]).columns:
      ed_mue[col] = ed_mue[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_mue.to_sql("muelle", conn, if_exists="replace", index=False)
    conn.close()
    st.success("MUELLES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 6. ELIMINAR REGISTROS
# -----------------------------------------------------------------------------
elif choice == "ELIMINAR REGISTROS":
  st.subheader("🗑️ MÓDULO DE ELIMINACIÓN DE REGISTROS")
  st.warning(
      "¡ATENCIÓN! LOS REGISTROS ELIMINADOS NO SE PUEDEN RECUPERAR."
  )

  tab_del1, tab_del2, tab_del3, tab_del4, tab_del5 = st.tabs([
      "🚀 Operaciones Control",
      "👥 Clientes",
      "📍 Lugares",
      "🚢 Buques",
      "🏗️ Muelles",
  ])

  with tab_del1:
    st.markdown("### Eliminar Registro de Operación / Alquiler")
    df_ops_del = run_query("""
            SELECT c.IDCONTROL, cl.NOMBRE as CLIENTE, l.LUGAR, b.NOMBREBUQUE, c.FECHAINICIO 
            FROM control c
            LEFT JOIN cliente cl ON c.CLIENTE = cl.IDCLIENTE
            LEFT JOIN lugar l ON c.LUGAR = l.IDLUGAR
            LEFT JOIN buque b ON c.BUQUE = b.IDBUQUE
        """)
    if not df_ops_del.empty:
      op_dict = dict(
          zip(
              df_ops_del["IDCONTROL"]
              + " - "
              + df_ops_del["CLIENTE"].fillna("N/D")
              + " ("
              + df_ops_del["FECHAINICIO"].fillna("N/D")
              + ")",
              df_ops_del["IDCONTROL"],
          )
      )
      sel_del_op = st.selectbox(
          "Seleccione Operación a Borrar", options=list(op_dict.keys())
      )
      if st.button("🗑️ ELIMINAR OPERACIÓN SELECCIONADA"):
        execute_query(
            "DELETE FROM control WHERE IDCONTROL = ?", (op_dict[sel_del_op],)
        )
        st.success("¡Operación eliminada con éxito!")
        st.rerun()
    else:
      st.info("No hay operaciones registradas.")

  with tab_del2:
    st.markdown("### Eliminar Cliente")
    df_cli_del = run_query("SELECT IDCLIENTE, NOMBRE, RIF FROM cliente")
    if not df_cli_del.empty:
      cli_dict = dict(
          zip(
              df_cli_del["NOMBRE"] + " (RIF: " + df_cli_del["RIF"] + ")",
              df_cli_del["IDCLIENTE"],
          )
      )
      sel_del_cli = st.selectbox(
          "Seleccione Cliente a Borrar", options=list(cli_dict.keys())
      )
      if st.button("🗑️ ELIMINAR CLIENTE SELECCIONADO"):
        execute_query(
            "DELETE FROM cliente WHERE IDCLIENTE = ?",
            (cli_dict[sel_del_cli],),
        )
        st.success("¡Cliente eliminado con éxito!")
        st.rerun()
    else:
      st.info("No hay clientes registrados.")

  with tab_del3:
    st.markdown("### Eliminar Lugar / Puerto")
    df_lug_del = run_query("SELECT IDLUGAR, LUGAR FROM lugar")
    if not df_lug_del.empty:
      lug_dict = dict(zip(df_lug_del["LUGAR"], df_lug_del["IDLUGAR"]))
      sel_del_lug = st.selectbox(
          "Seleccione Lugar a Borrar", options=list(lug_dict.keys())
      )
      if st.button("🗑️ ELIMINAR LUGAR SELECCIONADO"):
        execute_query(
            "DELETE FROM lugar WHERE IDLUGAR = ?", (lug_dict[sel_del_lug],)
        )
        st.success("¡Lugar eliminado con éxito!")
        st.rerun()
    else:
      st.info("No hay lugares registrados.")

  with tab_del4:
    st.markdown("### Eliminar Buque")
    df_buq_del = run_query("SELECT IDBUQUE, NOMBREBUQUE FROM buque")
    if not df_buq_del.empty:
      buq_dict = dict(zip(df_buq_del["NOMBREBUQUE"], df_buq_del["IDBUQUE"]))
      sel_del_buq = st.selectbox(
          "Seleccione Buque a Borrar", options=list(buq_dict.keys())
      )
      if st.button("🗑️ ELIMINAR BUQUE SELECCIONADO"):
        execute_query(
            "DELETE FROM buque WHERE IDBUQUE = ?", (buq_dict[sel_del_buq],)
        )
        st.success("¡Buque eliminado con éxito!")
        st.rerun()
    else:
      st.info("No hay buques registrados.")

  with tab_del5:
    st.markdown("### Eliminar Muelle")
    df_mue_del = run_query("SELECT IDMUELLE, NUMEROMUELLE FROM muelle")
    if not df_mue_del.empty:
      mue_dict = dict(
          zip(
              df_mue_del["NUMEROMUELLE"].astype(str), df_mue_del["IDMUELLE"]
          )
      )
      sel_del_mue = st.selectbox(
          "Seleccione Muelle a Borrar", options=list(mue_dict.keys())
      )
      if st.button("🗑️ ELIMINAR MUELLE SELECCIONADO"):
        execute_query(
            "DELETE FROM muelle WHERE IDMUELLE = ?", (mue_dict[sel_del_mue],)
        )
        st.success("¡Muelle eliminado con éxito!")
        st.rerun()
    else:
      st.info("No hay muelles registrados.")

# -----------------------------------------------------------------------------
# 7. REPORTES Y RESUMEN POR PERIODO
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
        "INGRESOS TOTALES ACUMULADOS", f"${df_rep['TOTAL_USD'].sum():,.2f} USD"
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
# 8. CONFIGURACIÓN DE RECIBO WHATSAPP
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
      n_emp = st.text_input("NOMBRE DE LA EMPRESA", value=str(row_e["NOMBRE"]))
      n_rif = st.text_input("RIF", value=str(row_e["RIF"]))
      n_tlf = st.text_input("TELÉFONO DE CONTACTO", value=str(row_e["TELEFONO"]))
      n_dir = st.text_input("DIRECCIÓN", value=str(row_e["DIRECCION"]))
      n_pie = st.text_area(
          "PIE DE PÁGINA / MENSAJE FINAL", value=str(row_e["PIE_PAGINA"])
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