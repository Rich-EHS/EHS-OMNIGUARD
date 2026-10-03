import streamlit as st
import pandas as pd
import io
import json
import google.generativeai as genai
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==========================================
# CONFIGURACIÓN DE LA PÁGINA
# ==========================================
st.set_page_config(page_title="EHS OmniGuard AI", layout="wide", page_icon="⛑️")

# ==========================================
# MOTORES MATEMÁTICOS DE EVALUACIÓN
# ==========================================
def calcular_riesgo_inicial(p, s, f):
    score = p * s * f
    if score >= 60: return "INTOLERABLE", "🔴", score, "#FF4B4B"
    elif score >= 27: return "ALTO", "🟠", score, "#FF9800"
    elif score >= 10: return "MEDIO", "🟡", score, "#FFD166"
    else: return "ACEPTABLE", "🟢", score, "#06D6A0"

def calcular_riesgo_residual(p, g):
    score = p * g
    if score >= 15: return "INTOLERABLE", "🔴", score, "#FF4B4B"
    elif score >= 8: return "ALTO", "🟠", score, "#FF9800"
    elif score >= 4: return "MEDIO", "🟡", score, "#FFD166"
    else: return "ACEPTABLE", "🟢", score, "#06D6A0"

# ==========================================
# INTERFAZ DE USUARIO (UI)
# ==========================================
col_logo, col_title = st.columns([1, 10])
with col_logo: st.markdown("<h1 style='text-align: center; font-size: 3rem;'>⛑️</h1>", unsafe_allow_html=True)
with col_title:
    st.title("EHS OmniGuard AI")
    st.markdown("Plataforma Universal de Análisis de Riesgos impulsada por Gemini")

with st.sidebar:
    st.markdown("### ⚙️ Configuración del Motor")
    # El usuario ingresa su API Key aquí
    api_key = st.text_input("🔑 Tu API Key de Google Gemini:", type="password", help="Obtén tu clave gratuita en aistudio.google.com")
    
    st.markdown("### 👤 Perfil Activo")
    # Perfil dinámico
    nombre_supervisor = st.text_input("Nombre / Iniciales del Supervisor:", value="RMCJ")
    st.info(f"**Supervisor:** {nombre_supervisor}\n\n**Módulo:** Análisis Dinámico IA")
    st.divider()
    st.markdown("### 💡 Jerarquía de Controles")
    st.markdown("1. **Eliminación**\n2. **Sustitución**\n3. **Ingeniería**\n4. **Administrativos**\n5. **EPP**")

st.subheader("1. Parámetros del Proceso (Cualquier Sector)")
actividad_input = st.text_area("Describe la actividad con el mayor detalle posible:", height=68, placeholder="Ej. Cambio de válvulas en línea de gas presurizada, soldadura en altura, manejo de sosa cáustica...")

if "datos_analisis" not in st.session_state:
    st.session_state.datos_analisis = None

if st.button("🚀 Ejecutar IA Evaluadora", type="primary"):
    if not api_key:
        st.error("⚠️ Por favor, ingresa tu API Key de Gemini en el panel lateral izquierdo para usar el motor de IA.")
    elif len(actividad_input) < 10:
        st.warning("⚠️ Describe la actividad con un poco más de detalle.")
    else:
        with st.spinner("Gemini está analizando tu proceso y cruzando datos con la normatividad mexicana (STPS, SEMARNAT, ASEA)..."):
            try:
                # Conexión real con Gemini
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-3.8-flash')
                
              
                # INGENIERÍA DE PROMPT NIVEL DIRECTOR EHS (30 AÑOS DE EXPERIENCIA)
                prompt = f"""
                Actúa como un Director Corporativo de EHS (Seguridad, Salud y Medio Ambiente) con 30 años de experiencia en la industria de alto riesgo en México y auditor certificado ISO 45001.
                Tu tarea es realizar un Análisis de Seguridad en el Trabajo (AST) impecable, exhaustivo y estrictamente apegado a la legalidad mexicana (STPS, ASEA, SEMARNAT).
                
                Proceso a evaluar: {actividad_input}
                
                REGLAS DE ORO INQUEBRANTABLES (CUMPLIMIENTO LEGAL Y TÉCNICO):
                1. TAXONOMÍA ESTRICTA: El 'Peligro' es la fuente. El 'Riesgo' es el evento específico.
                2. GRANULARIDAD ATÓMICA (PROHIBIDO AGRUPAR): Cada evento de riesgo debe ser individual. NUNCA mezcles caída de objetos con atrapamientos o cortes en un mismo punto. Si hay 5 riesgos derivados de un peligro, crea 5 elementos separados en el JSON.
                3. CONTROLES MILIMÉTRICOS Y LÓGICOS: El control debe mitigar EXCLUSIVAMENTE el riesgo evaluado. Prohibido proponer mamparas de soldadura para caída de objetos.
                4. JERARQUÍA PURA: 
                   - Ingeniería: Guardas, barreras físicas, válvulas, rodapiés, extracción localizada. (NO pongas capacitación aquí).
                   - Administrativos: Procedimientos LOTO, permisos de trabajo, DC-3, señalización, acordonamiento, bandereros.
                   - EPP: Especificar el tipo normativo (Ej. "Casco Clase G", "Guantes anticorte Nivel 5", "Arnés de cuerpo completo con línea de vida de doble gancho").
                5. FUNDAMENTO LEGAL: Cita la Norma Oficial Mexicana (NOM) exacta y vigente aplicable a ESE riesgo específico.

                Debes responder ÚNICAMENTE con un objeto JSON válido con esta estructura exacta, sin texto adicional:
                [
                    {{
                        "id_act": "1.",
                        "actividad": "Nombre de la fase específica del proceso",
                        "peligros": [
                            {{
                                "id_peligro": "1.1",
                                "desc_peligro": "[TIPO DE PELIGRO] Descripción precisa (Ej. [FÍSICO] Cargas suspendidas por maniobra con grúa de 20 tons)",
                                "riesgos": [
                                    {{
                                        "id_riesgo": "1.1.1",
                                        "desc_riesgo": "Evento específico y su consecuencia (Ej. Aplastamiento de personal en piso por falla de estrobo. Consecuencia: Fatalidad).",
                                        "normativas": [
                                            {{"codigo": "NOM-006-STPS-2023", "descripcion": "Manejo y almacenamiento de materiales - Condiciones de seguridad."}}
                                        ],
                                        "controles": "Ingeniería: Uso de estrobos certificados, pestillo de seguridad en gancho, válvulas check en pistones de grúa.\\nAdministrativos: Plan de izaje crítico, acordonamiento del radio de giro a 360°, rigger certificado.\\nEPP: Casco de seguridad, botas con puntera de protección."
                                    }},
                                    {{
                                        "id_riesgo": "1.1.2",
                                        "desc_riesgo": "Atrapamiento de extremidades en puntos de pellizco al guiar la carga.",
                                        "normativas": [
                                            {{"codigo": "NOM-004-STPS-1999", "descripcion": "Sistemas de protección y dispositivos de seguridad en la maquinaria."}}
                                        ],
                                        "controles": "Ingeniería: N/A (Control enfocado en herramientas).\\nAdministrativos: Prohibido guiar la carga con las manos; uso obligatorio de vientos (cuerdas guía) de 3 metros de longitud.\\nEPP: Guantes de impacto."
                                    }}
                                ]
                            }}
                        ]
                    }}
                ]
                Genera al menos 2 peligros principales. Para cada peligro, desarrolla un MÍNIMO DE 3 RIESGOS ESPECÍFICOS INDIVIDUALES con controles operacionales lógicos, viables y profesionales.
                """
                
                response = model.generate_content(prompt)
                
                # Limpiar la respuesta para asegurar que sea JSON puro
                json_string = response.text.strip().replace('```json', '').replace('```', '')
                st.session_state.datos_analisis = json.loads(json_string)
                st.rerun()
                
            except Exception as e:
                st.error(f"Error de conexión o formato de la IA: {e}")

# 2. MÓDULO DE VALIDACIÓN Y EVALUACIÓN
if st.session_state.datos_analisis:
    st.success("✅ Jerarquía estructurada por Gemini. Procede a evaluar las variables matemáticas.")
    resultados_finales = []

    for idx_act, act in enumerate(st.session_state.datos_analisis):
        st.markdown(f"### Actividad {act['id_act']} {act['actividad']}")
        
        for idx_peligro, peligro in enumerate(act['peligros']):
            for idx_riesgo, riesgo in enumerate(peligro['riesgos']):
                
                with st.expander(f"⚠️ Peligro {peligro['id_peligro']}: {peligro['desc_peligro']}", expanded=True):
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        st.markdown(f"**🔴 Riesgo Detectado {riesgo['id_riesgo']}:** {riesgo['desc_riesgo']}")
                        st.text_area(f"Controles Operacionales {riesgo['id_riesgo']}:", value=riesgo['controles'], height=120, key=f"ctrl_{idx_act}_{idx_peligro}_{idx_riesgo}")
                        
                    with col2:
                        st.markdown("**📚 Normatividad Mexicana Aplicable:**")
                        for norma in riesgo['normativas']:
                            # Ahora Gemini nos da el código y la descripción dinámicamente para el Tooltip
                            st.markdown(f"- **{norma['codigo']}** ℹ️", help=norma['descripcion'])

                    st.divider()
                    
                    # EVALUACIÓN MATEMÁTICA
                    st.markdown("#### ⚖️ Evaluación de Riesgos")
                    col_p1, col_s1, col_f1, col_res1, col_space, col_p2, col_g2, col_res2 = st.columns([1,1,1,1.5, 0.5, 1,1,1.5])
                    
                    with col_p1: p_ini = st.number_input("Probabilidad", 1, 5, 4, key=f"p1_{idx_act}_{idx_peligro}_{idx_riesgo}")
                    with col_s1: s_ini = st.number_input("Severidad", 1, 5, 4, key=f"s1_{idx_act}_{idx_peligro}_{idx_riesgo}")
                    with col_f1: f_ini = st.number_input("Frecuencia", 1, 5, 3, key=f"f1_{idx_act}_{idx_peligro}_{idx_riesgo}")
                    with col_res1:
                        nivel_i, icono_i, score_i, color_i = calcular_riesgo_inicial(p_ini, s_ini, f_ini)
                        st.markdown("**Riesgo Inicial (P×S×F)**")
                        st.markdown(f"<h4 style='text-align: center; color: {color_i}; margin-top:0;'>{icono_i} {nivel_i} ({score_i})</h4>", unsafe_allow_html=True)

                    with col_p2: p_res = st.number_input("Prob. c/Ctrl", 1, 5, 1, key=f"p2_{idx_act}_{idx_peligro}_{idx_riesgo}")
                    with col_g2: g_res = st.number_input("Grav. c/Ctrl", 1, 5, 4, key=f"g2_{idx_act}_{idx_peligro}_{idx_riesgo}")
                    with col_res2:
                        nivel_r, icono_r, score_r, color_r = calcular_riesgo_residual(p_res, g_res)
                        st.markdown("**Riesgo Residual (P×G)**")
                        st.markdown(f"<h4 style='text-align: center; color: {color_r}; margin-top:0;'>{icono_r} {nivel_r} ({score_r})</h4>", unsafe_allow_html=True)
                    
                    # Almacenamiento estructurado para Excel
                    normas_str = ", ".join([n['codigo'] for n in riesgo['normativas']])
                    resultados_finales.append({
                        "Supervisor": nombre_supervisor,
                        "Actividad": f"{act['id_act']} {act['actividad']}",
                        "Peligro Asociado": f"{peligro['id_peligro']} {peligro['desc_peligro']}",
                        "Riesgo Detectado": f"{riesgo['id_riesgo']} {riesgo['desc_riesgo']}",
                        "Normativa (México)": normas_str,
                        "Controles por Jerarquía": st.session_state[f"ctrl_{idx_act}_{idx_peligro}_{idx_riesgo}"],
                        "P (Ini)": p_ini, "S (Ini)": s_ini, "F (Ini)": f_ini,
                        "Nivel Inicial": f"{nivel_i} ({score_i})",
                        "P (Res)": p_res, "G (Res)": g_res,
                        "Nivel Residual": f"{nivel_r} ({score_r})"
                    })

    # ==========================================
    # EXPORTACIÓN A EXCEL
    # ==========================================
    st.subheader("3. Exportación de Matriz")
    if st.button("📊 Descargar Matriz Validada (Excel)"):
        df_export = pd.DataFrame(resultados_finales)
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        df_export.to_excel(writer, index=False, sheet_name='Matriz_OmniGuard')
        
        worksheet = writer.sheets['Matriz_OmniGuard']
        header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        
        for col_num in range(1, len(df_export.columns) + 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.fill = header_fill
            cell.font = header_font
            worksheet.column_dimensions[cell.column_letter].width = 30
            
        writer.close()
        output.seek(0)
        
        st.download_button(
            label="⬇️ Descargar Archivo",
            data=output,
            file_name=f"Matriz_EHS_{nombre_supervisor}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )