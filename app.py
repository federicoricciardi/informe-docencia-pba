import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import ssl
import urllib.parse

# 1. Bypass SSL checks for environments behind corporate proxies/CA interception
ssl._create_default_https_context = ssl._create_unverified_context

# 2. Page Configuration
st.set_page_config(
    page_title="Monitoreo de Capacitación Docente - Ministerio de Salud PBA",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS institutional customization: Hide Streamlit default branding while keeping the sidebar control functional
st.markdown("""
    <style>
    /* Elimina el espacio blanco superior por defecto en Streamlit */
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 2rem !important;
    }
    /* Hacer el header transparente para que no tape el banner */
    [data-testid="stHeader"] {
        background: transparent !important;
        background-color: rgba(0, 0, 0, 0) !important;
    }
    /* Ocultar los botones de deploy y menú de tres puntos (acción) a la derecha */
    [data-testid="stHeaderActionElements"] {
        display: none !important;
    }
    /* Ocultar la línea decorativa superior */
    .stDecoration {
        display: none !important;
    }
    /* Ocultar el menú principal de opciones */
    #MainMenu {
        visibility: hidden !important;
    }
    /* Ocultar el pie de página de Streamlit */
    footer {
        visibility: hidden !important;
    }
    /* Ocultar el widget de estado */
    [data-testid="stStatusWidget"] {
        visibility: hidden !important;
    }
    </style>
""", unsafe_allow_html=True)

# 3. Google Drive Data URL (XLSX export format)
EXCEL_URL = "https://docs.google.com/spreadsheets/d/1P_GCMT92ytE2jJt0y_HW4S8zpGYAgDSVTm2e8K4DrYE/export?format=xlsx"

# 4. Token Mapping for the 12 Health Regions (Regiones Sanitarias PBA)
# The Excel sheet uses Roman numerals: 'I', 'II', ..., 'XII'
TOKENS_REGIONES = {
    "reg1-x8s9": "I",
    "reg2-m2k1": "II",
    "reg3-p9t4": "III",
    "reg4-v5r2": "IV",
    "reg5-k7w8": "V",
    "reg6-j3n5": "VI",
    "reg7-q2p8": "VII",
    "reg8-z6y4": "VIII",
    "reg9-x1a3": "IX",
    "reg10-s4d8": "X",
    "reg11-f2g6": "XI",
    "reg12-h8j9": "XII"
}

# Inverse mapping to get token for a given region (useful for the Director's link generator)
REGIONES_TOKENS = {v: k for k, v in TOKENS_REGIONES.items()}

# Helper to format region names beautifully
def format_region_name(region_val):
    if region_val in TOKENS_REGIONES.values():
        return f"Región Sanitaria {region_val}"
    return str(region_val)

# 5. Robust Column Normalization
def normalize_df_columns(df):
    mapping = {}
    for col in df.columns:
        col_str = str(col).strip()
        col_lower = col_str.lower()
        if 'regi' in col_lower or 'region' in col_lower:
            mapping[col] = 'Región'
        elif 'condic' in col_lower:
            mapping[col] = 'CONDICION'
        elif 'munic' in col_lower:
            mapping[col] = 'Municipio'
        elif 'curs' in col_lower and 'inscript' not in col_lower and 'aprob' not in col_lower:
            mapping[col] = 'Curso'
        elif 'correo' in col_lower or 'email' in col_lower or 'mail' in col_lower:
            mapping[col] = 'correo electrónico'
        elif 'nota' in col_lower:
            mapping[col] = 'Nota'
        elif 'apellido' in col_lower:
            mapping[col] = 'Apellido'
        elif 'nombre' in col_lower:
            mapping[col] = 'Nombre'
        elif 'dni' in col_lower:
            mapping[col] = 'DNI'
        else:
            mapping[col] = col_str
    return df.rename(columns=mapping)

# Helper to generate a robust unique identifier key per student
def add_unique_key_column(df):
    df_copy = df.copy()
    df_copy['DNI'] = df_copy['DNI'].fillna('').astype(str).str.strip()
    df_copy['correo electrónico'] = df_copy['correo electrónico'].fillna('').astype(str).str.strip()
    df_copy['Apellido'] = df_copy['Apellido'].fillna('').astype(str).str.strip()
    df_copy['Nombre'] = df_copy['Nombre'].fillna('').astype(str).str.strip()
    
    unique_keys = []
    for idx, row in df_copy.iterrows():
        dni = row['DNI']
        email = row['correo electrónico']
        name = f"{row['Apellido']}, {row['Nombre']}".strip()
        
        if dni and dni.lower() not in ['nan', 'null', '']:
            unique_keys.append(f"DNI_{dni}")
        elif email and email.lower() not in ['nan', 'null', '']:
            unique_keys.append(f"EMAIL_{email}")
        else:
            unique_keys.append(f"NAME_{name}")
            
    df_copy['UniqueKey'] = unique_keys
    return df_copy

# 6. Data Loading & Caching
@st.cache_data(ttl=600)  # Caches data for 10 minutes
def load_and_prepare_data(url):
    try:
        # Load sheets
        df_cursantes = pd.read_excel(url, sheet_name="CURSOS 2026")
        df_resumen = pd.read_excel(url, sheet_name="RESUMEN")
        
        # Normalize columns
        df_cursantes = normalize_df_columns(df_cursantes)
        df_resumen = normalize_df_columns(df_resumen)
        
        # Standardize empty values in 'Nota' to 'En curso'
        if 'Nota' in df_cursantes.columns:
            df_cursantes['Nota'] = df_cursantes['Nota'].fillna('En curso').astype(str).str.strip()
            df_cursantes['Nota'] = df_cursantes['Nota'].apply(
                lambda x: 'En curso' if str(x).strip() == '' or str(x).lower() in ['nan', 'none', '<na>', 'nat'] else str(x).strip()
            )
            
        return df_cursantes, df_resumen, None
    except Exception as e:
        return None, None, str(e)

# Helper to calculate metrics based on unique students
def compute_unique_metrics(df):
    if len(df) == 0:
        return 0, 0, 0, 0.0
        
    df_copy = add_unique_key_column(df)
    
    # Ensure Nota is clean
    df_copy['Nota'] = df_copy['Nota'].fillna('En curso').astype(str).str.strip()
    df_copy['Nota'] = df_copy['Nota'].apply(
        lambda x: 'En curso' if str(x).strip() == '' or str(x).lower() in ['nan', 'none', '<na>', 'nat'] else str(x).strip()
    )
    
    # Priority value for Nota: En curso (1), Aprobado (2), Desaprobado (3), No inició (4)
    # We prioritize 'En curso' so that active students are counted as currently cursando.
    def get_status_val(nota):
        n = str(nota).upper().strip()
        if 'NO INICIO' in n or 'NO INICIÓ' in n or 'ABANDONO' in n or 'ABANDONÓ' in n:
            return 4
        elif 'DESAPROBADO' in n:
            return 3
        elif 'APROBADO' in n:
            return 2
        else:
            return 1  # 'En curso'
            
    df_copy['status_val'] = df_copy['Nota'].apply(get_status_val)
    
    # For each unique student, get the minimum status_val (prioritizing En curso, then Aprobado, then Desaprobado, then No inició)
    df_student_status = df_copy.groupby('UniqueKey')['status_val'].min().reset_index()
    
    total_inscriptos = len(df_student_status)
    total_cursando = len(df_student_status[df_student_status['status_val'] == 1])
    total_aprobados = len(df_student_status[df_student_status['status_val'] == 2])
    
    tasa_aprobacion = (total_aprobados / total_inscriptos * 100) if total_inscriptos > 0 else 0.0
    
    return total_inscriptos, total_aprobados, total_cursando, tasa_aprobacion

# Helper to aggregate nominal data avoiding duplicates
def get_grouped_nominal_df(df):
    if len(df) == 0:
        return pd.DataFrame(columns=['Apellido', 'Nombre', 'DNI', 'Región', 'Municipio', 'correo electrónico', 'Curso', 'Nota'])
        
    df_copy = add_unique_key_column(df)
    
    # Ensure Nota has clean "En curso" for any empty/missing values
    df_copy['Nota'] = df_copy['Nota'].fillna('En curso').astype(str).str.strip()
    df_copy['Nota'] = df_copy['Nota'].apply(
        lambda x: 'En curso' if str(x).strip() == '' or str(x).lower() in ['nan', 'none', '<na>', 'nat'] else str(x).strip()
    )
    
    # Standard columns to aggregate
    agg_funcs = {}
    for col in ['Apellido', 'Nombre', 'DNI', 'Región', 'Municipio', 'correo electrónico']:
        if col in df_copy.columns:
            agg_funcs[col] = 'first'
            
    df_personal = df_copy.groupby('UniqueKey').agg(agg_funcs).reset_index()
    
    # Group courses and notes
    def group_courses(g):
        pairs = []
        for c, n in zip(g['Curso'], g['Nota']):
            c_val = str(c).strip() if pd.notna(c) else ""
            n_val = str(n).strip() if pd.notna(n) else "En curso"
            if n_val == "" or n_val.lower() in ['nan', 'none', '<na>', 'nat']:
                n_val = "En curso"
            if c_val:
                pairs.append((c_val, n_val))
        
        # Deduplicate pairs
        unique_pairs = []
        seen = set()
        for c_val, n_val in pairs:
            if (c_val, n_val) not in seen:
                seen.add((c_val, n_val))
                unique_pairs.append((c_val, n_val))
                
        # Sort by course name
        unique_pairs = sorted(unique_pairs, key=lambda x: x[0])
        
        if not unique_pairs:
            return pd.Series({'Curso': '', 'Nota': ''})
            
        courses = [p[0] for p in unique_pairs]
        notes = [p[1] for p in unique_pairs]
        
        return pd.Series({
            'Curso': ', '.join(courses),
            'Nota': ', '.join(notes)
        })
        
    df_courses = df_copy.groupby('UniqueKey').apply(group_courses)
    if isinstance(df_courses, pd.DataFrame):
        df_courses = df_courses.reset_index()
        
    df_grouped = pd.merge(df_personal, df_courses, on='UniqueKey').reset_index(drop=True)
    
    # Remove the UniqueKey
    if 'UniqueKey' in df_grouped.columns:
        df_grouped = df_grouped.drop(columns=['UniqueKey'])
        
    return df_grouped

# Load data
df_cursantes_raw, df_resumen, load_error = load_and_prepare_data(EXCEL_URL)

# 7. Render Header PBA Banner & CSS Styling
import os

# Detect if the banner image is present (supporting double extensions like .png.jpg)
banner_path = "banner_ministerio.png"
if not os.path.exists(banner_path):
    if os.path.exists("banner_ministerio.png.jpg"):
        banner_path = "banner_ministerio.png.jpg"

if os.path.exists(banner_path):
    st.image(banner_path, use_container_width=True)
else:
    # Fallback to HTML banner if the image file is not found
    st.markdown(
        """
        <div style="background-color: #0F238C; color: white; padding: 15px 25px; border-radius: 8px; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
            <div>
                <h2 style="margin: 0; font-size: 24px; font-weight: bold; color: white; letter-spacing: 0.5px;">GOBIERNO DE LA PROVINCIA DE BUENOS AIRES</h2>
                <p style="margin: 0; font-size: 14px; opacity: 0.9; color: #F0F8FF;">Ministerio de Salud | Dirección de Capacitación y Desarrollo de Carrera</p>
            </div>
            <div style="text-align: right; min-width: 200px; padding-top: 5px;">
                <span style="background-color: #0695D6; color: white; padding: 6px 12px; border-radius: 20px; font-weight: bold; font-size: 13px;">
                    🎓 Monitoreo de Capacitación Docente
                </span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

# 8. Security Routing (Token Handling)
def get_query_token():
    try:
        # Streamlit 1.30+ API
        return st.query_params.get("token")
    except Exception:
        try:
            # Fallback for older versions
            return st.experimental_get_query_params().get("token", [None])[0]
        except Exception:
            return None

token_param = get_query_token()

if load_error:
    st.error(f"⚠️ Error al conectar con la base de datos de Google Drive: {load_error}")
    st.info("Por favor, verifica el enlace para compartir o la conexión de red.")
else:
    # Applying the strict condition filter: Only INSCRIPTO DE LA RED are shown
    # (Except the raw RESUMEN sheet which already has pre-calculated general data)
    df_cursantes_filtered = df_cursantes_raw[df_cursantes_raw['CONDICION'] == 'INSCRIPTO DE LA RED'].copy()

    # Determine Access Mode
    if token_param is not None:
        # User supplied a token
        if token_param in TOKENS_REGIONES:
            locked_region = TOKENS_REGIONES[token_param]
            is_restricted = True
        else:
            # Invalid token: block access
            st.markdown(
                """
                <div style="background-color: #F8D7DA; color: #721C24; padding: 30px; border-radius: 8px; border: 1px solid #F5C6CB; margin-top: 50px; text-align: center; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
                    <h3 style="margin-top: 0; color: #721C24;">🚫 Acceso Denegado</h3>
                    <p style="font-size: 16px; margin-bottom: 0;">El token de acceso provisto en la URL es inválido o ha expirado. Por favor, solicite un enlace de acceso válido al Administrador del Sistema.</p>
                </div>
                """,
                unsafe_allow_html=True
            )
            st.stop()
    else:
        # No token: Full Director Panel
        locked_region = None
        is_restricted = False

    # 9. Main Application Views
    if is_restricted:
        # --- VIEW: REGIONAL RESTRICTED VIEW (COORDINATOR) ---
        st.markdown(f"### 📍 Panel de Control Regional — **{format_region_name(locked_region)}**")
        st.caption("Esta vista contiene únicamente información de su región sanitaria y está restringida según el protocolo de confidencialidad.")

        # Sidebar Course Filter for Coordinator
        st.sidebar.header("🔍 Filtros de Región")
        st.sidebar.write(f"📍 **{format_region_name(locked_region)}**")
        
        # Get courses taken by students of this region
        df_region_full = df_cursantes_filtered[df_cursantes_filtered['Región'] == locked_region]
        course_list = sorted(list(df_region_full['Curso'].dropna().unique()))
        selected_course = st.sidebar.selectbox(
            "Seleccionar Curso",
            ["Todos"] + course_list
        )

        # Filter database exclusively for this region
        df_region = df_region_full.copy()
        if selected_course != "Todos":
            df_region = df_region[df_region['Curso'] == selected_course]

        # Calculate metrics for the locked region using unique student counts
        total_inscriptos, total_aprobados, total_cursando, tasa_aprobacion = compute_unique_metrics(df_region)

        # KPI Metrics Cards (Custom Beautiful HTML Row)
        st.markdown(
            f"""
            <div style="display: flex; gap: 20px; margin-bottom: 25px; flex-wrap: wrap;">
                <div style="flex: 1; min-width: 220px; background-color: #F0F8FF; border-left: 5px solid #0695D6; padding: 15px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.04);">
                    <div style="font-size: 14px; color: #555555; font-weight: bold; text-transform: uppercase;">Inscriptos de la Red</div>
                    <div style="font-size: 28px; font-weight: bold; color: #262730; margin-top: 5px;">{total_inscriptos}</div>
                    <div style="font-size: 12px; color: #777777; margin-top: 5px;">Personal activo único registrado</div>
                </div>
                <div style="flex: 1; min-width: 220px; background-color: #E6F4EA; border-left: 5px solid #24A148; padding: 15px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.04);">
                    <div style="font-size: 14px; color: #555555; font-weight: bold; text-transform: uppercase;">Aprobados de la Red</div>
                    <div style="font-size: 28px; font-weight: bold; color: #24A148; margin-top: 5px;">{total_aprobados}</div>
                    <div style="font-size: 12px; color: #777777; margin-top: 5px;">Alumnos únicos aprobados</div>
                </div>
                <div style="flex: 1; min-width: 220px; background-color: #FDF2E2; border-left: 5px solid #E6A23C; padding: 15px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.04);">
                    <div style="font-size: 14px; color: #555555; font-weight: bold; text-transform: uppercase;">Cursando actualmente</div>
                    <div style="font-size: 28px; font-weight: bold; color: #E6A23C; margin-top: 5px;">{total_cursando}</div>
                    <div style="font-size: 12px; color: #777777; margin-top: 5px;">En curso (sin aprobaciones aún)</div>
                </div>
                <div style="flex: 1; min-width: 220px; background-color: #F3E8FF; border-left: 5px solid #8E44AD; padding: 15px; border-radius: 6px; box-shadow: 0 2px 4px rgba(0,0,0,0.04);">
                    <div style="font-size: 14px; color: #555555; font-weight: bold; text-transform: uppercase;">Tasa de Aprobación</div>
                    <div style="font-size: 28px; font-weight: bold; color: #8E44AD; margin-top: 5px;">{tasa_aprobacion:.1f}%</div>
                    <div style="font-size: 12px; color: #777777; margin-top: 5px;">Porcentaje de alumnos aprobados</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("📚 Distribución por Curso")
            if total_inscriptos > 0:
                # Group by course, avoiding double counting within the same course
                df_region_chart = add_unique_key_column(df_region)
                df_region_chart = df_region_chart.drop_duplicates(subset=['UniqueKey', 'Curso'])
                df_course = df_region_chart.groupby('Curso').size().reset_index(name='Inscriptos')
                fig_course = px.bar(
                    df_course,
                    x='Curso',
                    y='Inscriptos',
                    text='Inscriptos',
                    color_discrete_sequence=['#0695D6'],
                    labels={'Inscriptos': 'Inscriptos de la Red'}
                )
                fig_course.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=20, r=20, t=20, b=20),
                    height=300
                )
                st.plotly_chart(fig_course, use_container_width=True)
            else:
                st.info("No hay datos de cursos disponibles para esta región.")

        with col2:
            st.subheader("🏙️ Cursantes por Municipio")
            if total_inscriptos > 0:
                # Group by municipality, counting unique students
                df_region_chart = add_unique_key_column(df_region)
                df_region_chart = df_region_chart.drop_duplicates(subset=['UniqueKey'])
                df_muni = df_region_chart.groupby('Municipio').size().reset_index(name='Inscriptos').sort_values(by='Inscriptos', ascending=False)
                fig_muni = px.bar(
                    df_muni.head(10), # Top 10 municipalities
                    y='Municipio',
                    x='Inscriptos',
                    text='Inscriptos',
                    orientation='h',
                    color_discrete_sequence=['#0F238C'],
                    labels={'Inscriptos': 'Inscriptos de la Red'}
                )
                fig_muni.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=20, r=20, t=20, b=20),
                    height=300
                )
                fig_muni.update_yaxes(autorange="reversed")
                st.plotly_chart(fig_muni, use_container_width=True)
            else:
                st.info("No hay datos de municipios disponibles para esta región.")

        st.subheader("📋 Listado Nominal de Cursantes (Red de la Región)")
        st.caption("Filtros aplicados automáticamente: Solo alumnos de la RED sanitaria de la región seleccionada (agrupados por cursante único).")
        
        # Search filter
        search_term = st.text_input("🔍 Buscar por DNI, Apellido o Nombre", "")
        
        # Group nominal data to avoid duplicates
        df_display = get_grouped_nominal_df(df_region)
        
        if search_term:
            search_term = search_term.lower()
            df_display = df_display[
                df_display['Nombre'].astype(str).str.lower().str.contains(search_term) |
                df_display['Apellido'].astype(str).str.lower().str.contains(search_term) |
                df_display['DNI'].astype(str).str.lower().str.contains(search_term)
            ]

        columns_to_show = ['Apellido', 'Nombre', 'DNI', 'Región', 'Municipio', 'correo electrónico', 'Curso', 'Nota']
        valid_cols = [c for c in columns_to_show if c in df_display.columns]
        
        st.dataframe(df_display[valid_cols], use_container_width=True, hide_index=True)

        st.markdown("---")
        # Support button
        st.caption("Si detecta alguna inconsistencia en el padrón de alumnos, por favor contáctese con la Mesa de Ayuda de la Dirección Central.")
        central_mail_subject = urllib.parse.quote(f"Consulta Soporte Técnico - Región {locked_region}")
        central_mail_link = f"mailto:soporte.capacitacion@ms.gba.gov.ar?subject={central_mail_subject}"
        st.markdown(f'<a href="{central_mail_link}"><button style="background-color: transparent; border: 1px solid #0695D6; color: #0695D6; padding: 8px 15px; border-radius: 4px; font-weight: bold; cursor: pointer;">✉️ Reportar Inconsistencia</button></a>', unsafe_allow_html=True)

    else:
        # --- VIEW: FULL PANEL (DIRECTOR VIEW) ---
        st.markdown("### 📊 Panel de información")

        # Create sidebar filters
        st.sidebar.header("🔍 Filtros Generales")
        
        # Region Filter
        region_list = sorted(list(df_cursantes_filtered['Región'].dropna().unique()))
        # Filter regions list to keep only Roman numerals corresponding to the 12 regions
        clean_regions = [r for r in region_list if r in TOKENS_REGIONES.values()]
        
        selected_region = st.sidebar.selectbox(
            "Seleccionar Región Sanitaria",
            ["Todas"] + clean_regions,
            format_func=lambda x: "Todas las Regiones" if x == "Todas" else format_region_name(x)
        )
        
        # Course Filter
        course_list = sorted(list(df_cursantes_filtered['Curso'].dropna().unique()))
        selected_course = st.sidebar.selectbox(
            "Seleccionar Curso",
            ["Todos"] + course_list
        )

        # Apply filtering based on selection to the RAW (unfiltered by condition) dataset
        df_dir_general = df_cursantes_raw.copy()
        if selected_region != "Todas":
            df_dir_general = df_dir_general[df_dir_general['Región'] == selected_region]
        if selected_course != "Todos":
            df_dir_general = df_dir_general[df_dir_general['Curso'] == selected_course]

        # Strictly filter the General dataset for Red candidates
        df_dir_red = df_dir_general[df_dir_general['CONDICION'] == 'INSCRIPTO DE LA RED'].copy()

        # Calculate metrics for Bloque 1: General (All candidates) using unique student counts
        gen_inscriptos, gen_aprobados, gen_cursando, gen_tasa = compute_unique_metrics(df_dir_general)

        # Calculate metrics for Bloque 2: Red (Network candidates only) using unique student counts
        red_inscriptos, red_aprobados, red_cursando, red_tasa = compute_unique_metrics(df_dir_red)

        # Set df_filtered = df_dir_red for all other charts and lists (nominal list needs the filtered red candidates)
        df_filtered = df_dir_red

        # Render double block metrics (side-by-side using st.columns(2))
        col_gen, col_red = st.columns(2)

        with col_gen:
            st.markdown(
                f"""
                <div style="background-color: #F8F9FA; padding: 20px; border-radius: 8px; border: 1px solid #E2E8F0; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">
                    <h4 style="margin-top: 0; color: #262730; border-bottom: 2px solid #E2E8F0; padding-bottom: 8px; font-weight: bold;">Matrícula General</h4>
                    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-top: 15px;">
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #94A3B8;">
                            <div style="font-size: 11px; color: #64748B; font-weight: bold; text-transform: uppercase;">Inscriptos</div>
                            <div style="font-size: 24px; font-weight: bold; color: #1E293B; margin-top: 2px;">{gen_inscriptos}</div>
                        </div>
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #E6A23C;">
                            <div style="font-size: 11px; color: #64748B; font-weight: bold; text-transform: uppercase;">Cursando Actualmente</div>
                            <div style="font-size: 24px; font-weight: bold; color: #E6A23C; margin-top: 2px;">{gen_cursando}</div>
                        </div>
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #24A148;">
                            <div style="font-size: 11px; color: #64748B; font-weight: bold; text-transform: uppercase;">Aprobados</div>
                            <div style="font-size: 24px; font-weight: bold; color: #24A148; margin-top: 2px;">{gen_aprobados}</div>
                        </div>
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #8E44AD;">
                            <div style="font-size: 11px; color: #64748B; font-weight: bold; text-transform: uppercase;">Tasa de Aprobación</div>
                            <div style="font-size: 24px; font-weight: bold; color: #8E44AD; margin-top: 2px;">{gen_tasa:.1f}%</div>
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with col_red:
            st.markdown(
                f"""
                <div style="background-color: #F0F8FF; padding: 20px; border-radius: 8px; border: 1px solid #BEE3F8; margin-bottom: 25px; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">
                    <h4 style="margin-top: 0; color: #006699; border-bottom: 2px solid #BEE3F8; padding-bottom: 8px; font-weight: bold;">Inscriptos de la Red</h4>
                    <div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 15px; margin-top: 15px;">
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #0695D6;">
                            <div style="font-size: 11px; color: #006699; font-weight: bold; text-transform: uppercase;">Inscriptos</div>
                            <div style="font-size: 24px; font-weight: bold; color: #0695D6; margin-top: 2px;">{red_inscriptos}</div>
                        </div>
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #E6A23C;">
                            <div style="font-size: 11px; color: #006699; font-weight: bold; text-transform: uppercase;">Cursando Actualmente</div>
                            <div style="font-size: 24px; font-weight: bold; color: #E6A23C; margin-top: 2px;">{red_cursando}</div>
                        </div>
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #24A148;">
                            <div style="font-size: 11px; color: #006699; font-weight: bold; text-transform: uppercase;">Aprobados</div>
                            <div style="font-size: 24px; font-weight: bold; color: #24A148; margin-top: 2px;">{red_aprobados}</div>
                        </div>
                        <div style="background-color: white; padding: 12px; border-radius: 6px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); border-left: 4px solid #8E44AD;">
                            <div style="font-size: 11px; color: #006699; font-weight: bold; text-transform: uppercase;">Tasa de Aprobación</div>
                            <div style="font-size: 24px; font-weight: bold; color: #8E44AD; margin-top: 2px;">{red_tasa:.1f}%</div>
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

        # Tabs for Director
        tab_dash, tab_resumen, tab_nominal, tab_enlaces = st.tabs([
            "📊 Dashboard de Métricas",
            "📈 Resumen de Cursos",
            "🔍 Registro Nominal de Cursantes",
            "🔑 Gestión de Enlaces Regionales"
        ])

        with tab_dash:
            col1, col2 = st.columns(2)

            with col1:
                st.subheader("📍 Inscriptos de la Red por Región Sanitaria")
                # Group by region for cleaned regions only
                df_reg_counts = df_cursantes_filtered[df_cursantes_filtered['Región'].isin(clean_regions)].groupby('Región').size().reset_index(name='Inscriptos')
                # Sort in order of Roman numerals
                roman_order = {r: i for i, r in enumerate(clean_regions)}
                df_reg_counts['order'] = df_reg_counts['Región'].map(roman_order)
                df_reg_counts = df_reg_counts.sort_values(by='order')

                fig_reg = px.bar(
                    df_reg_counts,
                    x='Región',
                    y='Inscriptos',
                    text='Inscriptos',
                    color_discrete_sequence=['#0695D6'],
                    labels={'Inscriptos': 'Inscriptos de la Red', 'Región': 'Región Sanitaria'}
                )
                fig_reg.update_layout(
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)',
                    margin=dict(l=20, r=20, t=20, b=20),
                    height=320
                )
                st.plotly_chart(fig_reg, use_container_width=True)

            with col2:
                st.subheader("📚 Aprobados vs Cursando por Curso")
                if len(df_filtered) > 0:
                    df_course_aprob = df_filtered.groupby(['Curso', 'Nota'].copy(), dropna=False).size().reset_index(name='Cantidad')
                    
                    def get_estado_label(nota):
                        n = str(nota).upper().strip()
                        if 'APROBADO' in n:
                            return 'Aprobado'
                        elif 'DESAPROBADO' in n:
                            return 'Desaprobado'
                        elif 'NO INICIO' in n or 'NO INICIÓ' in n or 'ABANDONO' in n or 'ABANDONÓ' in n:
                            return 'No inició / Abandono'
                        else:
                            return 'Cursando'
                            
                    df_course_aprob['Estado'] = df_course_aprob['Nota'].apply(get_estado_label)
                    
                    fig_course_aprob = px.bar(
                        df_course_aprob,
                        x='Curso',
                        y='Cantidad',
                        color='Estado',
                        barmode='group',
                        color_discrete_map={
                            'Aprobado': '#24A148', 
                            'Desaprobado': '#DA1E28', 
                            'No inició / Abandono': '#94A3B8', 
                            'Cursando': '#E6A23C'
                        }
                    )
                    fig_course_aprob.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        margin=dict(l=20, r=20, t=20, b=20),
                        height=320
                    )
                    st.plotly_chart(fig_course_aprob, use_container_width=True)
                else:
                    st.info("No hay datos para graficar.")

            # Email Send Section
            st.markdown("---")
            st.subheader("📧 Envío de Reporte Semanal a Referentes")
            
            # Select box to compose email for a specific region
            col_mail_reg, col_mail_btn = st.columns([2, 2])
            with col_mail_reg:
                mail_region = st.selectbox(
                    "Seleccione la Región Sanitaria para enviar el mail de reporte:",
                    clean_regions,
                    format_func=format_region_name
                )
            
            # Compute stats for selected mail region
            df_mail_reg = df_cursantes_filtered[df_cursantes_filtered['Región'] == mail_region]
            mail_inscriptos = len(df_mail_reg)
            mail_aprobados = len(df_mail_reg[df_mail_reg['Nota'].astype(str).str.upper() == 'APROBADO'])
            mail_tasa = (mail_aprobados / mail_inscriptos * 100) if mail_inscriptos > 0 else 0.0
            
            # Generate the URL access link for this region
            region_token = REGIONES_TOKENS.get(mail_region, "")
            
            # Construct mailto link
            destinatario = f"referente.region{mail_region}@ms.gba.gov.ar"
            subject = urllib.parse.quote(f"Reporte de Avance - Capacitación Docente - Región Sanitaria {mail_region}")
            
            # Access link structure (reconstructed safely)
            acceso_link = f"https://monitoreo-capacitacion-docente.streamlit.app/?token={region_token}" # standard deployment URL format
            
            body = urllib.parse.quote(
                f"Estimado/a Referente de la Región Sanitaria {mail_region},\n\n"
                f"Le hacemos llegar el reporte actualizado sobre el estado de la capacitación docente en su región (Personal de la Red de Salud):\n\n"
                f"📈 Métricas de la Región Sanitaria {mail_region}:\n"
                f"  - Inscriptos de la Red: {mail_inscriptos}\n"
                f"  - Aprobados de la Red: {mail_aprobados}\n"
                f"  - Tasa de Aprobación: {mail_tasa:.1f}%\n\n"
                f"🔑 Enlace de Acceso a su Panel Regional Restringido:\n"
                f"Para consultar el listado nominal de cursantes y realizar el seguimiento en tiempo real, ingrese al siguiente enlace:\n"
                f"{acceso_link}\n\n"
                f"Atentamente,\n"
                f"Dirección de Capacitación y Desarrollo de Carrera\n"
                f"Ministerio de Salud de la Provincia de Buenos Aires"
            )
            
            mailto_link = f"mailto:{destinatario}?subject={subject}&body={body}"
            
            with col_mail_btn:
                st.write("")
                st.write("")
                st.markdown(
                    f'<a href="{mailto_link}" target="_blank" style="text-decoration:none;"><button style="background-color:#0695D6; color:white; border:none; padding:10px 20px; border-radius:5px; font-weight:bold; cursor:pointer; width:100%; box-shadow:0 2px 4px rgba(0,0,0,0.1);">📧 Redactar Correo para Región {mail_region}</button></a>', 
                    unsafe_allow_html=True
                )
            
        with tab_resumen:
            st.subheader("📈 Cuadro Resumen (Pestaña 'RESUMEN')")
            st.caption("Pantallazo general de inscripciones totales y aprobados consolidados por curso (incluye inscriptos de la red y externos).")
            
            # Clean and display the resumen sheet dataframe
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)

        with tab_nominal:
            st.subheader("🔍 Padrón de Cursantes de la Red")
            st.caption(f"Visualización actual: {selected_region if selected_region == 'Todas' else format_region_name(selected_region)} | Curso: {selected_course}")
            st.caption("Excluye registros que no correspondan a la condición 'INSCRIPTO DE LA RED'.")

            search_query = st.text_input("🔍 Filtrar por DNI, Nombre o Apellido:", key="search_director")
            
            # Group the nominal data to avoid duplicates
            df_nominal_display = get_grouped_nominal_df(df_filtered)

            if search_query:
                search_query = search_query.lower()
                df_nominal_display = df_nominal_display[
                    df_nominal_display['Nombre'].astype(str).str.lower().str.contains(search_query) |
                    df_nominal_display['Apellido'].astype(str).str.lower().str.contains(search_query) |
                    df_nominal_display['DNI'].astype(str).str.lower().str.contains(search_query)
                ]

            cols_nominal = ['Apellido', 'Nombre', 'DNI', 'Región', 'Municipio', 'correo electrónico', 'Curso', 'Nota']
            valid_cols_nominal = [c for c in cols_nominal if c in df_nominal_display.columns]

            st.dataframe(df_nominal_display[valid_cols_nominal], use_container_width=True, hide_index=True)

            # Export Button (Download as CSV)
            csv_data = df_nominal_display[valid_cols_nominal].to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Descargar Listado Nominal (CSV)",
                data=csv_data,
                file_name="padron_capacitacion_pba.csv",
                mime="text/csv"
            )

        with tab_enlaces:
            st.subheader("🔑 Enlaces de Acceso Regional Restringido")
            st.caption("Copie y comparta los siguientes enlaces únicos con los respectivos Coordinadores Regionales para darles acceso exclusivo a la información de su Región Sanitaria.")

            enlaces_rows = []
            # We reconstruct based on current query parameters host, defaulting to localhost if running locally
            # In Streamlit, we can't easily fetch browser hostname reliably, so we show the path structure
            for reg in clean_regions:
                tok = REGIONES_TOKENS.get(reg, "")
                # Create a sample list
                enlaces_rows.append({
                    "Región Sanitaria": format_region_name(reg),
                    "Token Seguro": tok,
                    "Enlace de Acceso": f"/?token={tok}"
                })
            
            df_enlaces = pd.DataFrame(enlaces_rows)
            st.dataframe(df_enlaces, use_container_width=True, hide_index=True)
            st.info("💡 Consejo: Al copiar el Enlace de Acceso, anteponga la URL en la que está publicada la app (ej: https://monitoreo-capacitacion.streamlit.app/?token=reg6-j3n5).")
