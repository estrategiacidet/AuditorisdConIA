from __future__ import annotations

import os
from pathlib import Path
from typing import Iterable

import openpyxl
import pypdf
import streamlit as st
from docx import Document
from google import genai
from google.genai import types
from dotenv import load_dotenv


BASE_DIR = Path(__file__).resolve().parent
EXCEL_TEMPLATE = BASE_DIR / "lista de verificación.xlsx"

load_dotenv(BASE_DIR / ".env")

DEFAULT_API_KEY = os.getenv("GEMINI_API_KEY", "")


INSTRUCCIONES_A1 = """
Eres un Consultor Senior Experto en Diagnóstico y Contexto Organizacional de CIDET. Tu único objetivo es recopilar, investigar y estructurar la información general y estratégica de la empresa asignada a partir de los documentos provistos y de su sitio web oficial.

Este diagnóstico NO es una auditoría formal; es una fase de reconocimiento preliminar para alimentar la revisión documental de un plan de auditoría.

LINEAMIENTOS ESTRICTOS DE COMPORTAMIENTO Y CALIDAD (CONTROL DE SESGOS):
1. Objetividad Absoluta: Está prohibido hacer juicios de valor, críticas, o declarar hallazgos, conformidades o incumplimientos normativos. Limítate a reportar datos objetivos de las fuentes.
2. Restricción de Evidencia: Bástate exclusivamente en los datos explícitos del documento de Cámara de Comercio, el Informe Guía suministrado y la URL web oficial que se te proporciona. Si un dato solicitado en la estructura no se encuentra en estas fuentes (por ejemplo, el organigrama detallado o el comité de SST), debes escribir textualmente: "Información no disponible en los documentos de entrada preliminares; se deberá validar en la auditoría de campo". Queda prohibido inventar o asumir datos.
3. Uso de la URL: Si se incluye una URL en los datos de entrada, navega en ella para extraer y validar la misión, visión, pilares estratégicos, presencia o menciones al sistema de gestión de la empresa.
4. Tono Corporativo: Utiliza una redacción neutral, técnica, clara y estrictamente profesional, mimetizando el lenguaje formal e institucional de CIDET.

PROCESO DE INVESTIGACIÓN:
FASE 0: INGESTA Y REVISIÓN DOCUMENTAL Y DIGITAL
- Analizar el documento de constitución (Cámara de Comercio).
- Inspeccionar el contenido de la página web provista (si está disponible).
- Identificar cualquier referencia al organigrama, estructura o líderes de la empresa.

FASE 1: EXTRACCIÓN DE CONTEXTO
- Identificar organización, NIT, representantes legales, dirección, objeto social y códigos CIIU.
- Determinar las actividades principales y fortalezas estratégicas declaradas por la empresa en sus canales oficiales.
- Identificar menciones a actores clave, comités de SST (COPASST) o vigencia de acreditaciones bajo ISO 45001.

FASE 2: CONSOLIDACIÓN DE INFORMACIÓN
Estructurar los datos extraídos en los siguientes ejes antes de redactar: Tipo de empresa, años de trayectoria, actividades principales, actores clave, estado preliminar de SST y fortalezas estratégicas.

FASE 3: GENERACIÓN DE INFORME (ESTRUCTURA OBLIGATORIA DE SALIDA)
Debes estructurar y entregar el informe siguiendo exactamente este formato de salida. No agregues introducciones ni conclusiones fuera de los bloques delimitados:

[INICIO_CONTEXTO]

# INFORME CONTEXTO ORGANIZACIONAL
## Empresa: [Insertar Nombre de la Empresa completo y NIT si está disponible]
## Fecha de Generación: 24 de junio de 2026

### 1. RESUMEN EJECUTIVO
[Redacta un resumen analítico y fluido de un máximo de dos párrafos que sintetice la naturaleza de la empresa analizada, su propósito principal y el alcance del reconocimiento preliminar realizado para CIDET].

### 2. INFORMACIÓN GENERAL Y MARCO OPERATIVO
- **Razón Social Completa:** [Nombre]
- **NIT / Identificación Legal:** [Número]
- **Actividad Económica Principal (CIIU):** [Indicar actividad y código si aparece]
- **Años de Trayectoria / Constitución:** [Fecha o años desde su constitución legal]
- **Representante Legal / Actores Clave:** [Nombres de cargos directivos identificados]
- **Estructura Organizacional preliminar:** [Describir brevemente cómo se organiza según el texto/web o reportar si no está disponible]

### 3. CONTEXTO ESTRATÉGICO Y COMPONENTE EN SST
- **Actividades Principales del Negocio:** [Detallar las operaciones críticas de la organización recopiladas de los documentos y su sitio web]
- **Fortalezas Estratégicas Identificadas:** [Mencionar las ventajas operativas, misión, visión o ventajas competitivas explícitas en las fuentes]
- **Comité de SST / Estado ISO 45001:** [Reportar lo hallado en los documentos o en el sitio web sobre comités o certificaciones de seguridad, o declarar su ausencia documental para validación en campo]

[FIN_CONTEXTO]
"""


def normalize_path(raw_value: str) -> Path:
    cleaned = raw_value.strip().strip('"').strip("'")
    return Path(cleaned).expanduser().resolve()


def iter_pdf_text(file_paths: Iterable[Path]) -> str:
    chunks: list[str] = []
    for pdf_path in file_paths:
        try:
            reader = pypdf.PdfReader(str(pdf_path))
            for page in reader.pages:
                text = page.extract_text() or ""
                if text:
                    chunks.append(text)
        except Exception:
            continue
    return "\n".join(chunks)


def read_context_files(base_dir: Path) -> tuple[str, str]:
    texto_camara = ""
    texto_guia = ""
    pdfs = [p for p in base_dir.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    for pdf in pdfs:
        name = pdf.name.lower()
        if "camara" in name:
            texto_camara += iter_pdf_text([pdf]) + "\n"
        elif "informe" in name:
            texto_guia += iter_pdf_text([pdf]) + "\n"
    return texto_camara.strip(), texto_guia.strip()


def read_excel_numerals(template_path: Path) -> list[str]:
    workbook = openpyxl.load_workbook(template_path, data_only=True)
    try:
        sheet = workbook["lista de verificación"] if "lista de verificación" in workbook.sheetnames else workbook.active
        values: list[str] = []
        for row in range(2, 149):
            cell_value = sheet.cell(row=row, column=4).value
            if cell_value not in (None, ""):
                values.append(str(cell_value).strip())
        return values
    finally:
        workbook.close()


def read_evidence_text(folder: Path) -> str:
    if not folder.exists():
        return ""
    pdfs = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == ".pdf"]
    return iter_pdf_text(pdfs)


def save_docx_report(title: str, body: str, output_path: Path) -> None:
    doc = Document()
    doc.add_heading(title, level=1)
    doc.add_paragraph(body)
    doc.save(str(output_path))


st.set_page_config(page_title="CIDET - Asistente de Auditoría IA", page_icon="⚡", layout="wide")

st.title("⚡ CIDET - Asistente de Auditoría Asistida con IA")
st.markdown("Optimización y automatización de diagnóstico organizacional y evaluación ISO 45001")
st.write("---")

with st.sidebar:
    st.header("🔑 Configuración de Seguridad")
    api_key_usuario = st.text_input(
        "Gemini API Key",
        value=DEFAULT_API_KEY,
        type="password",
        help="También puedes definir la variable de entorno GEMINI_API_KEY.",
    )
    st.write("---")
    st.markdown("**Desarrollado para el Líder de Estrategia e Innovación - CIDET**")

st.header("📁 1. Configuración del Proyecto del Cliente")

ruta_base = st.text_input(
    "Pega aquí la ruta de la auditoría de la empresa que vas a procesar hoy:",
    placeholder=r"C:\Users\usuario\OneDrive - CIDET\...",
)

url_empresa = st.text_input(
    "Pega aquí el enlace de la página web de la empresa (opcional):",
    placeholder="https://www.empresa.com",
)

if ruta_base:
    ruta_base_limpia = normalize_path(ruta_base)
    ruta_evidencias = ruta_base_limpia / "Evidencias"

    if not ruta_base_limpia.exists():
        st.error(f"No se encontró la carpeta indicada: `{ruta_base_limpia}`")
        st.stop()

    st.success("📍 Conectado exitosamente a la carpeta del cliente.")

    tab1, tab2 = st.tabs(["🕵️ Agente 1: Investigador de Contexto", "📋 Agente 2 & 3: Lector de Estructura e Informe Word"])

    with tab1:
        st.subheader("Análisis de Cámara de Comercio, Sitio Web e Informe Guía")

        if st.button("🚀 Ejecutar Agente Investigador", key="btn_agente1"):
            if not api_key_usuario.strip():
                st.error("Debes configurar tu Gemini API Key para ejecutar esta fase.")
            else:
                with st.spinner("El Agente 1 está analizando los documentos..."):
                    try:
                        client = genai.Client(api_key=api_key_usuario.strip())
                        texto_camara, texto_guia = read_context_files(ruta_base_limpia)

                        paquete = (
                            f"[CÁMARA DE COMERCIO]:\n{texto_camara}\n\n"
                            f"[EJEMPLO INFORME]:\n{texto_guia}"
                        )
                        if url_empresa.strip():
                            paquete += f"\n\n[SITIO WEB OFICIAL A ANALIZAR]: {url_empresa.strip()}"

                        response = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=paquete,
                            config=types.GenerateContentConfig(
                                system_instruction=INSTRUCCIONES_A1,
                                temperature=0.2,
                            ),
                        )

                        ruta_out_word = ruta_base_limpia / "Contexto_Organizacional.docx"
                        ruta_out_txt = ruta_base_limpia / "Contexto_Organizacional.txt"

                        save_docx_report(
                            "Reporte de Contexto Organizacional y Pre-Diagnóstico",
                            response.text or "",
                            ruta_out_word,
                        )
                        ruta_out_txt.write_text(response.text or "", encoding="utf-8")

                        st.success("✅ ¡Documento Word de Contexto generado con éxito!")
                        st.text_area("Vista previa del Contexto:", response.text or "", height=300)
                    except Exception as exc:
                        st.error(f"Ocurrió un error con el Agente 1: {exc}")

    with tab2:
        st.subheader("Generación de Informe Word usando la Matriz de Excel como Insumo")
        st.write("Esta fase leerá la matriz de Excel para extraer los numerales obligatorios y redactará el informe final en Word.")

        if st.button("🔍 Generar Informe Basado en la Matriz", key="btn_agente2"):
            if not api_key_usuario.strip():
                st.error("Debes configurar tu Gemini API Key para ejecutar esta fase.")
            else:
                archivo_contexto = ruta_base_limpia / "Contexto_Organizacional.txt"
                if not archivo_contexto.exists():
                    st.error("No se encontró el Contexto Organizacional. Debes ejecutar primero el Agente 1.")
                elif not EXCEL_TEMPLATE.exists():
                    st.error(f"No se encontró el archivo `lista de verificación.xlsx` en: {EXCEL_TEMPLATE}")
                else:
                    with st.spinner("Leyendo la estructura del Excel y analizando evidencias..."):
                        try:
                            client = genai.Client(api_key=api_key_usuario.strip())

                            contexto_previo = archivo_contexto.read_text(encoding="utf-8")
                            listado_numerales_excel = read_excel_numerals(EXCEL_TEMPLATE)
                            texto_numerales_insumo = ", ".join(listado_numerales_excel)
                            texto_evidencias = read_evidence_text(ruta_evidencias)

                            prompt_informe_completo = f"""
Actúas como un Lead Auditor Senior de Sistemas de Gestión en CIDET. Tu labor es generar el informe final en Word.

INSUMO CLAVE 1 - Contexto Organizacional (Agente 1):
{contexto_previo}

INSUMO CLAVE 2 - Estructura y Numerales Obligatorios leídos del Excel de verificación:
Los siguientes son los numerales exactos que contiene la lista de chequeo corporativa y que debes evaluar basándote en la evidencia:
{texto_numerales_insumo}

INSUMO CLAVE 3 - Evidencias Físicas del Cliente:
{texto_evidencias}

Cruza minuciosamente las evidencias con los numerales provistos del Excel bajo los lineamientos de la norma NTC-ISO 45001:2018, el Decreto 1072 de 2015 y la Resolución 0312 de 2019.
Redacta de forma extensa, rigurosa y formal el INFORME DE AUDITORÍA CORPORATIVO siguiendo exactamente la estructura institucional de CIDET:

INFORME DE AUDITORÍA DE SISTEMA DE GESTIÓN
ISO 45001:2018 - CIDET

1. DATOS GENERALES
Documento: FR-SIG-0010 V3
Fecha del informe: 24 de junio de 2026
Objetivo: Evaluar el grado de cumplimiento del sistema de gestión de seguridad y salud en el trabajo, conforme a lo establecido en la norma NTC-ISO 45001:2018 y la legislación colombiana aplicable.
Alcance: Revisión documental del Sistema de Gestión de la Seguridad y Salud en el Trabajo aplicado a los procesos y centros de trabajo definidos en el alcance del cliente bajo el ecosistema digital.
Requisito: NTC-ISO 45001:2018, Decreto 1072 de 2015, Resolución 0312 de 2019 y normatividad legal vigente en SST aplicable.

3. FORTALEZAS Y ASPECTOS POR MEJORAR
FORTALEZAS
[Genera una lista numerada con puntos fuertes basados exclusivamente en los hallazgos conformes de las evidencias].

ASPECTOS POR MEJORAR
[Lista numerada de oportunidades de mejora detectadas].

4. VERIFICACIÓN DE LA EFICACIA DE LAS ACCIONES RESULTANTES DE NO CONFORMIDADES PREVIAS
NUMERAL | DESCRIPCIÓN | CONCLUSIONES
"En las evidencias analizadas no se reportan estados de auditorías externas o evaluaciones previas a este ejercicio".

5. NO CONFORMIDADES DETECTADAS (ANÁLISIS CLÍNICO DOCUMENTAL)
[Por cada brecha detectada en los numerales del Excel, genera el siguiente bloque riguroso]:
NO CONFORMIDAD: [Mayor o Menor]
Descripción: [Hecho claro + Evidencia de soporte + Incumplimiento Normativo de la ISO 45001 o legislación colombiana].
Requisito aplicable: NTC-ISO 45001:2018
Numeral exacto: [Cruzar con el numeral correspondiente del Excel]

6. RIESGOS IDENTIFICADOS PARA SER CONSIDERADOS EN EL SIGUIENTE PROGRAMA DE AUDITORÍAS
[Detalla los riesgos operativos, documentales o legales identificados].

7. CONCLUSIÓN DEL EJERCICIO DE AUDITORÍA INTERNA
[Dictamen integral sobre el estado de madurez, ciclo PHVA y exposición legal en Colombia reflejando las brechas halladas].

FIRMA AUDITOR LÍDER
"""

                            informe_res = client.models.generate_content(
                                model="gemini-2.5-flash",
                                contents=prompt_informe_completo,
                                config=types.GenerateContentConfig(temperature=0.2),
                            )

                            ruta_word_final = ruta_base_limpia / "Informe_Final_Auditoria_ISO45001.docx"
                            save_docx_report(
                                "Informe Ejecutivo Final de Auditoría ISO 45001 - CIDET",
                                informe_res.text or "",
                                ruta_word_final,
                            )

                            st.success("🎉 ¡Informe Ejecutivo en Word generado con éxito!")
                            st.markdown(f"💾 **Archivo guardado en:** `{ruta_word_final}`")
                            st.subheader("📝 Vista Previa del Informe Redactado:")
                            st.text_area("Contenido:", informe_res.text or "", height=400)
                        except Exception as exc:
                            st.error(f"Ocurrió un error en el procesamiento combinado: {exc}")
else:
    st.info("Ingresa la ruta de la carpeta del cliente para comenzar.")
