from datetime import datetime, date
import sqlite3
import urllib.parse
import uuid
import random
import string
import pandas as pd
import streamlit as st
import os

# Configuración de página
st.set_page_config(
    page_title="LIFRAN - Control Operativo Portuario",
    page_icon="⚓",
    layout="wide",
)

DB_NAME = "lifran.db"
EXCEL_PATH = "LIFRAN.xlsx"


# Función para generar Número de Control (3 Letras + 3 Números)
def generar_id_control():
    letras = ''.join(random.choices(string.ascii_uppercase, k=3))
    numeros = ''.join(random.choices(string.digits, k=3))
    return f"{letras}{numeros}"


# Inicializar Base de Datos SQLite de forma robusta
@st.cache_resource
def init_db():
  conn = sqlite3.connect(DB_NAME)
  cursor = conn.cursor()

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

  # Actualizar tabla empresa para guardar la última tarifa usada
  try:
      cursor.execute("ALTER TABLE empresa ADD COLUMN TARIFA_DEFECTO REAL DEFAULT 25.0")
  except sqlite3.OperationalError:
      pass # La columna ya existe

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

  # Actualizar tabla control para que el historial respete la tarifa de su fecha
  try:
      cursor.execute("ALTER TABLE control ADD COLUMN TARIFA REAL DEFAULT 25.0")
  except sqlite3.OperationalError:
      pass # La columna ya existe

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


# --- MENÚ LATERAL Y LOGO ---
if os.path.exists("logo.jpeg"):
    st.sidebar.image("logo.jpeg", width=150) # Imagen pequeña en la esquina superior izquierda

st.sidebar.title("⚓ LIFRAN NAVEGACIÓN")
menu_opciones = [
    "CONTROL OPERATIVO",
    "CLIENTES",
    "LUGARES Y PUERTOS",
    "BUQUES",
    "MUELLES",
    "EDITAR / ELIMINAR REGISTROS",
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

  # Leer la última tarifa guardada
  df_emp = run_query("SELECT * FROM empresa LIMIT 1")
  tarifa_guardada = float(df_emp["TARIFA_DEFECTO"].iloc[0]) if "TARIFA_DEFECTO" in df_emp.columns and not pd.isna(df_emp["TARIFA_DEFECTO"].iloc[0]) else 25.0

  df_cli = run_query(
      "SELECT IDCLIENTE, NOMBRE, RIF, DIRECCION, TELEFONO FROM cliente"
  )
  df_lug = run_query("SELECT IDLUGAR, LUGAR FROM lugar")
  df_buq = run_query("SELECT IDBUQUE, NOMBREBUQUE FROM buque")
  df_mue = run_query("SELECT IDMUELLE, NUMEROMUELLE FROM muelle")

  col1, col2 = st.columns(2)

  with col1:
    # --- CLIENTE ---
    lista_clientes = df_cli["NOMBRE"].tolist() if not df_cli.empty else []
    lista_clientes.insert(0, "➕ [ AGREGAR NUEVO CLIENTE ]")
    sel_cli = st.selectbox("CLIENTE", options=lista_clientes)

    err_cli = st.empty()
    nuevo_cli_nombre, nuevo_cli_rif, nuevo_cli_dir, nuevo_cli_tlf = "", "", "", ""
    if sel_cli == "➕ [ AGREGAR NUEVO CLIENTE ]":
      st.markdown("📝 *Complete los datos del nuevo cliente:*")
      nuevo_cli_nombre = st.text_input("NOMBRE O RAZÓN SOCIAL").upper()
      nuevo_cli_rif = st.text_input("RIF").upper()
      nuevo_cli_dir = st.text_input("DIRECCIÓN").upper()
      nuevo_cli_tlf = st.text_input("TELÉFONO").upper()
    else:
      if not df_cli.empty:
        cli_row = df_cli[df_cli["NOMBRE"] == sel_cli].iloc[0]
        st.info(
            f"📌 **DATOS REGISTRADOS:**\n\n"
            f"- **RIF:** {cli_row.get('RIF', 'N/D')}\n"
            f"- **DIRECCIÓN:** {cli_row.get('DIRECCION', 'N/D')}\n"
            f"- **TELÉFONO:** {cli_row.get('TELEFONO', 'N/D')}"
        )

    st.markdown("---")
    # --- LUGAR / PUERTO ---
    lista_lugares = df_lug["LUGAR"].tolist() if not df_lug.empty else []
    lista_lugares.insert(0, "➕ [ AGREGAR NUEVO LUGAR ]")
    sel_lug = st.selectbox("LUGAR / PUERTO", options=lista_lugares)

    err_lug = st.empty()
    nuevo_lug_nombre = ""
    if sel_lug == "➕ [ AGREGAR NUEVO LUGAR ]":
      nuevo_lug_nombre = st.text_input("NOMBRE DEL LUGAR / PUERTO").upper()

    st.markdown("---")
    # --- BUQUE ---
    lista_buques = df_buq["NOMBREBUQUE"].tolist() if not df_buq.empty else []
    lista_buques.insert(0, "➕ [ AGREGAR NUEVO BUQUE ]")
    sel_buq = st.selectbox("BUQUE", options=lista_buques)

    err_buq = st.empty()
    nuevo_buq_nombre = ""
    if sel_buq == "➕ [ AGREGAR NUEVO BUQUE ]":
      nuevo_buq_nombre = st.text_input("NOMBRE DEL BUQUE").upper()

  with col2:
    # --- MUELLE ---
    lista_muelles = (
        df_mue["NUMEROMUELLE"].astype(str).tolist() if not df_mue.empty else []
    )
    lista_muelles.insert(0, "➕ [ AGREGAR NUEVO MUELLE ]")
    sel_mue = st.selectbox("MUELLE", options=lista_muelles)

    err_mue = st.empty()
    nuevo_mue_numero = ""
    if sel_mue == "➕ [ AGREGAR NUEVO MUELLE ]":
      nuevo_mue_numero = st.text_input("NÚMERO DE MUELLE").upper()

    st.markdown("---")
    # Fechas (Formato visual estricto DD/MM/YYYY)
    f_inicio = st.date_input(
        "FECHA DE INICIO (DD/MM/AA)",
        value=date.today(),
        format="DD/MM/YYYY",
    )
    f_terminacion = st.date_input(
        "FECHA DE TERMINACIÓN (DD/MM/AA)",
        value=date.today(),
        format="DD/MM/YYYY",
    )
    num_cabinas = st.number_input(
        "CANTIDAD DE BAÑOS / CABINAS", min_value=1, value=1
    )
    
    # Campo de tarifa trayendo el último valor guardado por defecto
    tarifa_usd = st.number_input(
        "TARIFA DIARIA POR BAÑO (USD $)", min_value=0.0, value=tarifa_guardada
    )

    # --- CÁLCULO DINÁMICO DEL TOTAL ---
    dias = (f_terminacion - f_inicio).days + 1
    dias = max(1, dias)  # Mínimo 1 día
    total_a_pagar = dias * num_cabinas * tarifa_usd
    
    st.success(
        f"💰 **MONTO TOTAL A PAGAR: ${total_a_pagar:,.2f} USD** \n\n"
        f"*(Cálculo: {dias} días × {num_cabinas} baños × ${tarifa_usd})*"
    )

  observaciones = st.text_area("OBSERVACIONES").upper()

  # Botón de guardado normal
  if st.button("🚀 GUARDAR NUEVO OPERATIVO", use_container_width=True):
    error = False

    # 1. Resolver Cliente
    if sel_cli == "➕ [ AGREGAR NUEVO CLIENTE ]":
      if not nuevo_cli_nombre:
        err_cli.error("⚠️ EL NOMBRE DEL CLIENTE ES OBLIGATORIO.")
        error = True
      else:
        c_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO cliente (IDCLIENTE, NOMBRE, RIF, DIRECCION, TELEFONO) VALUES (?, ?, ?, ?, ?)",
            (c_id, nuevo_cli_nombre, nuevo_cli_rif, nuevo_cli_dir, nuevo_cli_tlf)
        )
    else:
      c_id = df_cli[df_cli["NOMBRE"] == sel_cli]["IDCLIENTE"].values[0]

    # 2. Resolver Lugar
    if sel_lug == "➕ [ AGREGAR NUEVO LUGAR ]":
      if not nuevo_lug_nombre:
        err_lug.error("⚠️ EL NOMBRE DEL LUGAR ES OBLIGATORIO.")
        error = True
      else:
        l_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO lugar (IDLUGAR, LUGAR) VALUES (?, ?)", (l_id, nuevo_lug_nombre)
        )
    else:
      l_id = df_lug[df_lug["LUGAR"] == sel_lug]["IDLUGAR"].values[0]

    # 3. Resolver Buque
    if sel_buq == "➕ [ AGREGAR NUEVO BUQUE ]":
      if not nuevo_buq_nombre:
        err_buq.error("⚠️ EL NOMBRE DEL BUQUE ES OBLIGATORIO.")
        error = True
      else:
        b_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO buque (IDBUQUE, NOMBREBUQUE) VALUES (?, ?)", (b_id, nuevo_buq_nombre)
        )
    else:
      b_id = df_buq[df_buq["NOMBREBUQUE"] == sel_buq]["IDBUQUE"].values[0]

    # 4. Resolver Muelle
    if sel_mue == "➕ [ AGREGAR NUEVO MUELLE ]":
      if not nuevo_mue_numero:
        err_mue.error("⚠️ EL NÚMERO DEL MUELLE ES OBLIGATORIO.")
        error = True
      else:
        m_id = str(uuid.uuid4())[:8].upper()
        execute_query(
            "INSERT INTO muelle (IDMUELLE, NUMEROMUELLE) VALUES (?, ?)", (m_id, nuevo_mue_numero)
        )
    else:
      m_id = df_mue[df_mue["NUMEROMUELLE"].astype(str) == sel_mue]["IDMUELLE"].values[0]

    # 5. Guardar el Operativo y actualizar memoria de Tarifa
    if not error:
      # Actualizar la tarifa guardada en empresa
      execute_query("UPDATE empresa SET TARIFA_DEFECTO = ? WHERE id = ?", (tarifa_usd, df_emp["id"].iloc[0]))
      
      id_control = generar_id_control()
      execute_query(
          """INSERT INTO control (IDCONTROL, CLIENTE, LUGAR, BUQUE, MUELLE, FECHAINICIO, FECHATERMINACION, NUMEROCABINAS, OBSERVACIONES, TARIFA)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
          (
              id_control, c_id, l_id, b_id, m_id,
              str(f_inicio), str(f_terminacion), num_cabinas, observaciones, tarifa_usd
          ),
      )
      st.success("¡OPERACIÓN REGISTRADA EXITOSAMENTE CON TODOS LOS DATOS NUEVOS!")
      st.rerun()


  st.divider()
  st.subheader("HISTORIAL DE OPERACIONES Y ENVÍO DE RECIBO POR WHATSAPP")

  query_sql = """
        SELECT c.rowid, c.IDCONTROL, cl.NOMBRE as CLIENTE_NOMBRE, l.LUGAR as LUGAR_NOMBRE, 
               b.NOMBREBUQUE, m.NUMEROMUELLE, c.FECHAINICIO, c.FECHATERMINACION, 
               c.NUMEROCABINAS, c.OBSERVACIONES, cl.TELEFONO, c.TARIFA
        FROM control c
        LEFT JOIN cliente cl ON c.CLIENTE = cl.IDCLIENTE
        LEFT JOIN lugar l ON c.LUGAR = l.IDLUGAR
        LEFT JOIN buque b ON c.BUQUE = b.IDBUQUE
        LEFT JOIN muelle m ON c.MUELLE = m.IDMUELLE
        ORDER BY c.rowid DESC
    """
  df_ops = run_query(query_sql)

  if not df_ops.empty:
    df_ops["FECHAINICIO_DT"] = pd.to_datetime(df_ops["FECHAINICIO"], errors="coerce")
    df_ops["FECHATERMINACION_DT"] = pd.to_datetime(df_ops["FECHATERMINACION"], errors="coerce")

    df_ops["DIAS"] = (df_ops["FECHATERMINACION_DT"] - df_ops["FECHAINICIO_DT"]).dt.days + 1
    df_ops["DIAS"] = df_ops["DIAS"].fillna(1).apply(lambda x: max(1, x))
    
    # Asegurar que NUMEROCABINAS y TARIFA sean numéricos
    df_ops["NUMEROCABINAS"] = pd.to_numeric(df_ops["NUMEROCABINAS"], errors="coerce").fillna(1)
    df_ops["TARIFA"] = pd.to_numeric(df_ops["TARIFA"], errors="coerce")
    
    df_ops["TARIFA_USD"] = df_ops["TARIFA"].fillna(25.0)
    df_ops["TOTAL_USD"] = df_ops["DIAS"] * df_ops["NUMEROCABINAS"] * df_ops["TARIFA_USD"]

    display_df = df_ops.copy()
    display_df["FECHAINICIO"] = display_df["FECHAINICIO_DT"].dt.strftime("%d/%m/%Y")
    display_df["FECHATERMINACION"] = display_df["FECHATERMINACION_DT"].dt.strftime("%d/%m/%Y")
    
    # Renombrar columnas para la grilla
    display_df = display_df.rename(columns={
        "IDCONTROL": "Nº_CONTROL",
        "CLIENTE_NOMBRE": "CLIENTE",
        "LUGAR_NOMBRE": "PUERTO",
        "NOMBREBUQUE": "BUQUE",
        "NUMEROMUELLE": "MUELLE",
        "NUMEROCABINAS": "BAÑOS",
    })

    st.dataframe(
        display_df[[
            "Nº_CONTROL", "CLIENTE", "PUERTO", "BUQUE",
            "MUELLE", "FECHAINICIO", "FECHATERMINACION",
            "BAÑOS", "DIAS", "TARIFA_USD", "TOTAL_USD", "OBSERVACIONES"
        ]],
        hide_index=True
    )

    st.markdown("### 📱 GENERAR Y ENVIAR RECIBO POR WHATSAPP")
    op_dict = dict(zip(
        "Nº " + display_df["Nº_CONTROL"] + " | " + display_df["CLIENTE"].fillna('N/D') + " (" + display_df["FECHAINICIO"].fillna('N/D') + ")", 
        df_ops["IDCONTROL"] # Usamos el df original para el value
    ))
    sel_op_name = st.selectbox("SELECCIONE OPERACIÓN PARA RECIBO", list(op_dict.keys()))

    if sel_op_name:
      sel_op = op_dict[sel_op_name]
      row = df_ops[df_ops["IDCONTROL"] == sel_op].iloc[0]
      empresa = run_query("SELECT * FROM empresa LIMIT 1").iloc[0]

      f_ini_str = row["FECHAINICIO_DT"].strftime("%d/%m/%Y") if pd.notnull(row["FECHAINICIO_DT"]) else "N/D"
      f_fin_str = row["FECHATERMINACION_DT"].strftime("%d/%m/%Y") if pd.notnull(row["FECHATERMINACION_DT"]) else "N/D"
      
      dias = int(row['DIAS'])

      # Formato exacto solicitado en la imagen
      recibo_msj = f"""*--- {empresa['NOMBRE']} ---*
RIF: {empresa['RIF']}
TLF: {empresa['TELEFONO']}
*NOTA DE ENTREGA / RECIBO DE ALQUILER*
*CONTROL Nº:* {row['IDCONTROL']}
---------------------------------------
*CLIENTE:* {row['CLIENTE_NOMBRE']}
*PUERTO / LUGAR:* {row['LUGAR_NOMBRE']}
*BUQUE:* {row['NOMBREBUQUE']} (MUELLE N° {row['NUMEROMUELLE']})
*PERÍODO:* {f_ini_str} AL {f_fin_str} ({dias} DÍAS)"""

      st.text_area("VISTA PREVIA DEL RECIBO:", recibo_msj, height=200)
      telefono = str(row["TELEFONO"]).strip().replace(" ", "").replace("+", "")
      whatsapp_link = f"https://wa.me/{telefono}?text={urllib.parse.quote(recibo_msj)}"

      st.markdown(
          f'<a href="{whatsapp_link}" target="_blank" style="display:inline-block;padding:12px 24px;background-color:#25D366;color:white;text-decoration:none;border-radius:6px;font-weight:bold;font-size:16px;">📲 ENVIAR RECIBO POR WHATSAPP</a>',
          unsafe_allow_html=True,
      )

# -----------------------------------------------------------------------------
# 2. CLIENTES
# -----------------------------------------------------------------------------
elif choice == "CLIENTES":
  st.subheader("ADMINISTRACIÓN Y REGISTRO DE CLIENTES")
  df_cli = run_query("SELECT * FROM cliente")
  
  ed_cli = st.data_editor(
      df_cli, 
      num_rows="dynamic", 
      use_container_width=True,
      column_config={"IDCLIENTE": None},
      hide_index=True
  )
  if st.button("ACTUALIZAR CAMBIOS EN TABLA CLIENTES"):
    for col in ed_cli.select_dtypes(include=["object"]).columns:
      ed_cli[col] = ed_cli[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_cli.to_sql("cliente", conn, if_exists="replace", index=False)
    conn.close()
    st.success("CLIENTES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 3. LUGARES Y PUERTOS
# -----------------------------------------------------------------------------
elif choice == "LUGARES Y PUERTOS":
  st.subheader("ADMINISTRACIÓN Y REGISTRO DE LUGARES / PUERTOS")
  df_lug = run_query("SELECT * FROM lugar")
  
  ed_lug = st.data_editor(
      df_lug, 
      num_rows="dynamic", 
      use_container_width=True,
      column_config={"IDLUGAR": None},
      hide_index=True
  )
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
  df_buq = run_query("SELECT * FROM buque")
  
  ed_buq = st.data_editor(
      df_buq, 
      num_rows="dynamic", 
      use_container_width=True,
      column_config={"IDBUQUE": None},
      hide_index=True
  )
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
  df_mue = run_query("SELECT * FROM muelle")
  
  ed_mue = st.data_editor(
      df_mue, 
      num_rows="dynamic", 
      use_container_width=True,
      column_config={"IDMUELLE": None},
      hide_index=True
  )
  if st.button("ACTUALIZAR CAMBIOS EN TABLA MUELLES"):
    for col in ed_mue.select_dtypes(include=["object"]).columns:
      ed_mue[col] = ed_mue[col].astype(str).str.upper()
    conn = sqlite3.connect(DB_NAME)
    ed_mue.to_sql("muelle", conn, if_exists="replace", index=False)
    conn.close()
    st.success("MUELLES ACTUALIZADOS.")

# -----------------------------------------------------------------------------
# 6. EDITAR / ELIMINAR REGISTROS
# -----------------------------------------------------------------------------
elif choice == "EDITAR / ELIMINAR REGISTROS":
  st.subheader("✏️ MÓDULO DE EDICIÓN Y ELIMINACIÓN DE REGISTROS")
  st.markdown("Puedes editar directamente las celdas o eliminar filas enteras.")

  tab_del1, tab_del2, tab_del3, tab_del4, tab_del5 = st.tabs([
      "🚀 Operaciones Control", "👥 Clientes", "📍 Lugares", "🚢 Buques", "🏗️ Muelles"
  ])

  with tab_del1:
    st.markdown("### Editar / Eliminar Operación de Alquiler")
    df_ops_edit = run_query("SELECT * FROM control")
    if not df_ops_edit.empty:
      # Editor completo para la tabla control
      ed_ops = st.data_editor(
          df_ops_edit, 
          num_rows="dynamic", 
          use_container_width=True,
          hide_index=True
      )
      if st.button("💾 GUARDAR CAMBIOS EN OPERACIONES"):
        for col in ed_ops.select_dtypes(include=["object"]).columns:
          ed_ops[col] = ed_ops[col].astype(str).str.upper()
        conn = sqlite3.connect(DB_NAME)
        ed_ops.to_sql("control", conn, if_exists="replace", index=False)
        conn.close()
        st.success("¡Operaciones actualizadas con éxito!")
    else:
      st.info("No hay operaciones registradas.")

  with tab_del2:
    st.markdown("### Editar / Eliminar Cliente")
    df_cli_edit = run_query("SELECT * FROM cliente")
    if not df_cli_edit.empty:
      ed_cli = st.data_editor(df_cli_edit, num_rows="dynamic", use_container_width=True, hide_index=True)
      if st.button("💾 GUARDAR CAMBIOS EN CLIENTES"):
        for col in ed_cli.select_dtypes(include=["object"]).columns:
          ed_cli[col] = ed_cli[col].astype(str).str.upper()
        conn = sqlite3.connect(DB_NAME)
        ed_cli.to_sql("cliente", conn, if_exists="replace", index=False)
        conn.close()
        st.success("¡Clientes actualizados con éxito!")
    else:
      st.info("No hay clientes registrados.")

  with tab_del3:
    st.markdown("### Editar / Eliminar Lugar")
    df_lug_edit = run_query("SELECT * FROM lugar")
    if not df_lug_edit.empty:
      ed_lug = st.data_editor(df_lug_edit, num_rows="dynamic", use_container_width=True, hide_index=True)
      if st.button("💾 GUARDAR CAMBIOS EN LUGARES"):
        for col in ed_lug.select_dtypes(include=["object"]).columns:
          ed_lug[col] = ed_lug[col].astype(str).str.upper()
        conn = sqlite3.connect(DB_NAME)
        ed_lug.to_sql("lugar", conn, if_exists="replace", index=False)
        conn.close()
        st.success("¡Lugares actualizados con éxito!")
    else:
      st.info("No hay lugares registrados.")

  with tab_del4:
    st.markdown("### Editar / Eliminar Buque")
    df_buq_edit = run_query("SELECT * FROM buque")
    if not df_buq_edit.empty:
      ed_buq = st.data_editor(df_buq_edit, num_rows="dynamic", use_container_width=True, hide_index=True)
      if st.button("💾 GUARDAR CAMBIOS EN BUQUES"):
        for col in ed_buq.select_dtypes(include=["object"]).columns:
          ed_buq[col] = ed_buq[col].astype(str).str.upper()
        conn = sqlite3.connect(DB_NAME)
        ed_buq.to_sql("buque", conn, if_exists="replace", index=False)
        conn.close()
        st.success("¡Buques actualizados con éxito!")
    else:
      st.info("No hay buques registrados.")

  with tab_del5:
    st.markdown("### Editar / Eliminar Muelle")
    df_mue_edit = run_query("SELECT * FROM muelle")
    if not df_mue_edit.empty:
      ed_mue = st.data_editor(df_mue_edit, num_rows="dynamic", use_container_width=True, hide_index=True)
      if st.button("💾 GUARDAR CAMBIOS EN MUELLES"):
        for col in ed_mue.select_dtypes(include=["object"]).columns:
          ed_mue[col] = ed_mue[col].astype(str).str.upper()
        conn = sqlite3.connect(DB_NAME)
        ed_mue.to_sql("muelle", conn, if_exists="replace", index=False)
        conn.close()
        st.success("¡Muelles actualizados con éxito!")
    else:
      st.info("No hay muelles registrados.")

# -----------------------------------------------------------------------------
# 7. REPORTES Y RESUMEN POR PERIODO (CON FILTROS)
# -----------------------------------------------------------------------------
elif choice == "REPORTES Y RESUMEN":
  st.subheader("📊 REPORTES Y ESTADÍSTICAS AVANZADAS")
  query_sql = """
        SELECT c.IDCONTROL, cl.NOMBRE as CLIENTE_NOMBRE, l.LUGAR as LUGAR_NOMBRE, 
               b.NOMBREBUQUE, m.NUMEROMUELLE, c.FECHAINICIO, c.FECHATERMINACION, c.NUMEROCABINAS, c.TARIFA
        FROM control c
        LEFT JOIN cliente cl ON c.CLIENTE = cl.IDCLIENTE
        LEFT JOIN lugar l ON c.LUGAR = l.IDLUGAR
        LEFT JOIN buque b ON c.BUQUE = b.IDBUQUE
        LEFT JOIN muelle m ON c.MUELLE = m.IDMUELLE
    """
  df_rep = run_query(query_sql)
  
  if not df_rep.empty:
    df_rep["FECHAINICIO_DT"] = pd.to_datetime(df_rep["FECHAINICIO"], errors="coerce")
    df_rep["FECHATERMINACION_DT"] = pd.to_datetime(df_rep["FECHATERMINACION"], errors="coerce")
    df_rep["DIAS"] = (df_rep["FECHATERMINACION_DT"] - df_rep["FECHAINICIO_DT"]).dt.days + 1
    df_rep["DIAS"] = df_rep["DIAS"].fillna(1).apply(lambda x: max(1, x))
    
    df_rep["NUMEROCABINAS"] = pd.to_numeric(df_rep["NUMEROCABINAS"], errors="coerce").fillna(1)
    df_rep["TARIFA"] = pd.to_numeric(df_rep["TARIFA"], errors="coerce")
    
    df_rep["TARIFA_USD"] = df_rep["TARIFA"].fillna(25.0)
    df_rep["TOTAL_USD"] = df_rep["DIAS"] * df_rep["NUMEROCABINAS"] * df_rep["TARIFA_USD"]

    # ---- SECCIÓN DE FILTROS (SIDEBAR) ----
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 🔍 FILTROS DE REPORTE")
    
    # Filtro por Fechas
    min_date = df_rep["FECHAINICIO_DT"].min().date() if not pd.isna(df_rep["FECHAINICIO_DT"].min()) else date.today()
    max_date = df_rep["FECHATERMINACION_DT"].max().date() if not pd.isna(df_rep["FECHATERMINACION_DT"].max()) else date.today()
    
    rango_fechas = st.sidebar.date_input(
        "Rango de Fechas",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY"
    )

    # Filtros por Atributos (Multiselect)
    clientes_unicos = df_rep["CLIENTE_NOMBRE"].dropna().unique().tolist()
    filtro_clientes = st.sidebar.multiselect("Filtrar por Clientes", options=clientes_unicos, default=clientes_unicos)

    lugares_unicos = df_rep["LUGAR_NOMBRE"].dropna().unique().tolist()
    filtro_lugares = st.sidebar.multiselect("Filtrar por Lugares / Puertos", options=lugares_unicos, default=lugares_unicos)

    buques_unicos = df_rep["NOMBREBUQUE"].dropna().unique().tolist()
    filtro_buques = st.sidebar.multiselect("Filtrar por Buques", options=buques_unicos, default=buques_unicos)

    # --- APLICAR FILTROS AL DATAFRAME ---
    df_filtrado = df_rep.copy()
    
    if len(rango_fechas) == 2:
        f_ini, f_fin = rango_fechas
        df_filtrado = df_filtrado[
            (df_filtrado["FECHAINICIO_DT"].dt.date >= f_ini) & 
            (df_filtrado["FECHATERMINACION_DT"].dt.date <= f_fin)
        ]
        
    df_filtrado = df_filtrado[df_filtrado["CLIENTE_NOMBRE"].isin(filtro_clientes)]
    df_filtrado = df_filtrado[df_filtrado["LUGAR_NOMBRE"].isin(filtro_lugares)]
    df_filtrado = df_filtrado[df_filtrado["NOMBREBUQUE"].isin(filtro_buques)]

    if not df_filtrado.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("INGRESOS TOTALES (FILTRADOS)", f"${df_filtrado['TOTAL_USD'].sum():,.2f} USD")
        c2.metric("OPERACIONES (FILTRADAS)", len(df_filtrado))
        c3.metric("BAÑOS ALQUILADOS (TOTAL)", f"{int(df_filtrado['NUMEROCABINAS'].sum())}")
        
        st.divider()
        st.markdown("### 📈 INGRESOS POR CLIENTE (FILTRADO)")
        res_cli = df_filtrado.groupby("CLIENTE_NOMBRE")["TOTAL_USD"].sum().reset_index()
        st.bar_chart(res_cli.set_index("CLIENTE_NOMBRE"))
        
        col_t1, col_t2 = st.columns(2)
        with col_t1:
            st.markdown("### 🏢 DESGLOSE POR CLIENTE")
            st.dataframe(res_cli, hide_index=True, use_container_width=True)
            
        with col_t2:
            st.markdown("### 📍 DESGLOSE POR PUERTO / LUGAR")
            res_lug = df_filtrado.groupby("LUGAR_NOMBRE")["TOTAL_USD"].sum().reset_index()
            st.dataframe(res_lug, hide_index=True, use_container_width=True)
            
        st.markdown("### 📋 TABLA DETALLADA DE OPERACIONES FILTRADAS")
        df_mostrar = df_filtrado.copy()
        df_mostrar["FECHAINICIO"] = df_mostrar["FECHAINICIO_DT"].dt.strftime("%d/%m/%Y")
        df_mostrar["FECHATERMINACION"] = df_mostrar["FECHATERMINACION_DT"].dt.strftime("%d/%m/%Y")
        st.dataframe(
            df_mostrar[["IDCONTROL", "CLIENTE_NOMBRE", "LUGAR_NOMBRE", "NOMBREBUQUE", "FECHAINICIO", "FECHATERMINACION", "NUMEROCABINAS", "TOTAL_USD"]],
            hide_index=True, use_container_width=True
        )

    else:
        st.warning("⚠️ No hay datos que coincidan con los filtros seleccionados.")
  else:
    st.info("NO HAY DATOS SUFICIENTES PARA LOS REPORTES.")

# -----------------------------------------------------------------------------
# 8. CONFIGURACIÓN DE RECIBO WHATSAPP
# -----------------------------------------------------------------------------
elif choice == "CONFIGURACIÓN RECIBO WHATSAPP":
  st.subheader("CONFIGURACIÓN DE DATOS FISCALES DE LA EMPRESA")
  st.markdown("Modifique los datos de la empresa que saldrán reflejados en el recibo enviado por WhatsApp.")
  df_emp = run_query("SELECT * FROM empresa LIMIT 1")
  if not df_emp.empty:
    row_e = df_emp.iloc[0]
    with st.form("form_empresa"):
      n_emp = st.text_input("NOMBRE DE LA EMPRESA", value=str(row_e["NOMBRE"]))
      n_rif = st.text_input("RIF", value=str(row_e["RIF"]))
      n_tlf = st.text_input("TELÉFONO DE CONTACTO", value=str(row_e["TELEFONO"]))
      n_dir = st.text_input("DIRECCIÓN", value=str(row_e["DIRECCION"]))
      n_pie = st.text_area("PIE DE PÁGINA", value=str(row_e["PIE_PAGINA"]))

      if st.form_submit_button("ACTUALIZAR DATOS DE EMPRESA"):
        execute_query(
            "UPDATE empresa SET NOMBRE=?, RIF=?, TELEFONO=?, DIRECCION=?, PIE_PAGINA=? WHERE id=?",
            (n_emp.upper(), n_rif.upper(), n_tlf.upper(), n_dir.upper(), n_pie.upper(), row_e["id"]),
        )
        st.success("¡DATOS DE LA EMPRESA ACTUALIZADOS EXITOSAMENTE!")