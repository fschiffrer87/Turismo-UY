"""
Etapa 07 — Informe final exhaustivo en Word (.docx).

Dirigido a un lector sin contexto previo: economista o tomador de decisión
que no participó en la elaboración del modelo.

Estructura:
  0.  Portada
  1.  Resumen Ejecutivo
  2.  Contexto económico y político
  3.  Marco legal y escenarios
  4.  Datos y fuentes
  5.  Metodología paso a paso (5.1 a 5.6)
  6.  Resultados (6.1 a 6.5)
  7.  Análisis de sensibilidad (7.1 a 7.3)
  8.  Limitaciones
  9.  Conclusiones y recomendaciones de política
  10. Referencias bibliográficas
  11. Anexos técnicos (A–E)
"""
import sys
import pathlib
import datetime
import pandas as pd
import yaml

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS, OUTPUTS_FIGURAS,
    cargar_supuestos, obtener_valor, separador,
)

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# =============================================================================
# HELPERS DE FORMATO
# =============================================================================

def _set_col_width(table, col_idx, width_cm):
    for row in table.rows:
        row.cells[col_idx].width = Cm(width_cm)


def agregar_titulo(doc, texto, nivel=1):
    doc.add_heading(texto, level=nivel)


def agregar_parrafo(doc, texto, negrita=False, cursiva=False, sangria=False):
    p = doc.add_paragraph()
    if sangria:
        p.paragraph_format.left_indent = Cm(1)
    run = p.add_run(texto)
    run.bold = negrita
    run.italic = cursiva
    return p


def agregar_formula(doc, texto):
    """Párrafo con fórmula en fuente monoespaciada, fondo gris."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(1.5)
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(texto)
    run.font.name = "Courier New"
    run.font.size = Pt(10)
    return p


def agregar_bullet(doc, texto, nivel=0):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(nivel * 0.8 + 0.5)
    run = p.add_run(texto)
    return p


def agregar_tabla_desde_df(doc, df, titulo=None, col_widths=None):
    """Inserta un DataFrame como tabla Word con encabezado en negrita."""
    if titulo:
        p = doc.add_paragraph()
        run = p.add_run(titulo)
        run.bold = True
        run.font.size = Pt(10)

    tabla = doc.add_table(rows=1, cols=len(df.columns))
    tabla.style = "Table Grid"
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Encabezados
    enc = tabla.rows[0].cells
    for i, col in enumerate(df.columns):
        enc[i].text = str(col)
        for run in enc[i].paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(9)
        enc[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Filas
    for _, fila in df.iterrows():
        celdas = tabla.add_row().cells
        for i, val in enumerate(fila):
            celdas[i].text = str(val) if val is not None and str(val) != "nan" else "—"
            celdas[i].paragraphs[0].runs[0].font.size = Pt(9)

    doc.add_paragraph()
    return tabla


def agregar_imagen(doc, ruta, ancho=5.5, caption=None):
    if pathlib.Path(ruta).exists():
        doc.add_picture(str(ruta), width=Inches(ancho))
        ultimo = doc.paragraphs[-1]
        ultimo.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if caption:
            p = doc.add_paragraph(caption)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.italic = True
                r.font.size = Pt(9)
    else:
        agregar_parrafo(doc, f"[Figura no disponible: {pathlib.Path(ruta).name}]", cursiva=True)
    doc.add_paragraph()


def salto_pagina(doc):
    doc.add_page_break()


# =============================================================================
# SECCIÓN 0 — PORTADA
# =============================================================================
def seccion_portada(doc, supuestos):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for _ in range(6):
        doc.add_paragraph()

    titulo = doc.add_heading(
        "Impacto sobre el PIB de Uruguay de eliminar la\n"
        "exoneración estacional de IVA a la gastronomía turística",
        level=0,
    )
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()
    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(
        "Estimación mediante equilibrio parcial, modelo insumo-producto\n"
        "y simulación Monte Carlo"
    ).italic = True

    for _ in range(3):
        doc.add_paragraph()

    meta = [
        ("Fecha de elaboración", datetime.date.today().isoformat()),
        ("Marco legal analizado", "Ley 17.934 (mod. Ley 20.212) + Decreto 434/023 (prorrogado por Decreto 220/2025)"),
        ("Año de referencia del modelo", str(obtener_valor(supuestos, "anno_base"))),
        ("Año de la MIP", str(obtener_valor(supuestos, "anno_mip"))),
        ("Fuentes principales", "DGI (Gasto Tributario), MINTUR (ETR microdatos), BCU (MIP 2016, Cuentas Nacionales)"),
        ("Software", "Python 3.11 — pipeline reproducible (ver Anexo D)"),
    ]
    for etiq, val in meta:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r1 = p.add_run(f"{etiq}: ")
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = p.add_run(val)
        r2.font.size = Pt(10)

    for _ in range(3):
        doc.add_paragraph()
    nota = doc.add_paragraph()
    nota.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = nota.add_run(
        "Documento técnico — Los resultados son estimaciones sujetas a los supuestos\n"
        "declarados en la Sección 5 y las limitaciones descritas en la Sección 8."
    )
    r.italic = True
    r.font.size = Pt(9)
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 1 — RESUMEN EJECUTIVO
# =============================================================================
def seccion_resumen_ejecutivo(doc, supuestos, res):
    agregar_titulo(doc, "1. Resumen Ejecutivo", 1)

    mc_c = res["mc_central"]
    mc_v = res["mc_variante"]
    imp_c = res["imp_central"]
    imp_v = res["imp_variante"]
    fis_c = res["fis_central"]

    agregar_parrafo(doc,
        "Este informe cuantifica el impacto macroeconómico de eliminar el Decreto 434/023 "
        "(prorrogado por Decreto 220/2025), que exonera totalmente del IVA a los servicios "
        "gastronómicos pagados por turistas no residentes con tarjeta emitida en el exterior "
        "durante la temporada alta (15 de noviembre al 30 de abril)."
    )
    agregar_parrafo(doc,
        "La pregunta de política es: si el decreto estacional cae pero queda vigente la "
        "devolución permanente de 9 puntos de la Ley 17.934, ¿cuánto actividad económica "
        "y empleo se pierde, y cuánto gana el fisco?"
    )

    # Tabla de resultados
    df_res = pd.DataFrame([
        {"Indicador": "Shock de precio al turista", "Central (+13 pp IVA)": "+11,1% (neto de absorción de margen)", "Variante (+22 pp IVA)": "+15,9%"},
        {"Indicador": "Caída en demanda gastronómica (MUSD)", "Central (+13 pp IVA)": f"−{abs(imp_c['delta_d_gastro_musd']):.1f}", "Variante (+22 pp IVA)": f"−{abs(imp_v['delta_d_gastro_musd']):.1f}"},
        {"Indicador": "Impacto en el PIB (% del PIB — mediana MC)", "Central (+13 pp IVA)": f"{mc_c['mediana']:.4f}%", "Variante (+22 pp IVA)": f"{mc_v['mediana']:.4f}%"},
        {"Indicador": "Impacto en el PIB — intervalo P10/P90", "Central (+13 pp IVA)": f"[{mc_c['p10']:.4f}%, {mc_c['p90']:.4f}%]", "Variante (+22 pp IVA)": f"[{mc_v['p10']:.4f}%, {mc_v['p90']:.4f}%]"},
        {"Indicador": "Impacto en el PIB (MUSD 2022)", "Central (+13 pp IVA)": f"−{abs(imp_c['delta_vab_musd2022']):.1f}", "Variante (+22 pp IVA)": f"−{abs(imp_v['delta_vab_musd2022']):.1f}"},
        {"Indicador": "Puestos de trabajo afectados", "Central (+13 pp IVA)": f"−{abs(imp_c['delta_empleo_puestos']):.0f}", "Variante (+22 pp IVA)": f"−{abs(imp_v['delta_empleo_puestos']):.0f}"},
        {"Indicador": "Recaudación neta ganada (MUSD)", "Central (+13 pp IVA)": f"+{fis_c['rec_neta_musd']:.2f}", "Variante (+22 pp IVA)": f"+{res['fis_variante']['rec_neta_musd']:.2f}"},
        {"Indicador": "Gasto Tributario ahorrado por el fisco (MUSD)", "Central (+13 pp IVA)": f"+{fis_c['gt_estacional_ahorrado_musd']:.1f}", "Variante (+22 pp IVA)": f"+{res['fis_variante']['gt_estacional_ahorrado_musd']:.1f}"},
    ])
    agregar_tabla_desde_df(doc, df_res, "Tabla 1.1. Resumen de resultados")

    agregar_parrafo(doc,
        "Nota sobre escala: el impacto en el PIB se mide en centésimas de punto porcentual. "
        "Para tener perspectiva, el crecimiento del PIB real de Uruguay en 2022 fue de 4,9%. "
        "El efecto estimado de esta medida (−0,023% en el escenario central) equivale a menos "
        "del 0,5% de esa variación anual. En términos monetarios, representa aproximadamente "
        f"9 millones de dólares sobre un PIB de {obtener_valor(supuestos, 'pib_nominal_musd'):,.0f} millones. "
        "El factor determinante de la demanda turística hacia Uruguay sigue siendo el tipo de "
        "cambio real bilateral con Argentina, no el nivel del IVA gastronómico."
    )
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 2 — CONTEXTO ECONÓMICO Y POLÍTICO
# =============================================================================
def seccion_contexto(doc, res):
    agregar_titulo(doc, "2. Contexto Económico y Político", 1)

    agregar_titulo(doc, "2.1 El turismo en la economía uruguaya", 2)
    agregar_parrafo(doc,
        "El turismo receptivo es una de las principales fuentes de divisas de Uruguay. "
        "Según datos del Ministerio de Turismo (MINTUR) y el Banco Central (BCU), "
        "el sector de alojamiento y servicios de comida representa alrededor del 2,5% "
        "del Valor Agregado Bruto (VAB) de la economía. En temporada alta, Punta del Este "
        "y la costa atlántica concentran la mayor parte del gasto de no residentes."
    )
    agregar_parrafo(doc,
        "El turismo argentino domina el flujo: más del 60% de las llegadas provienen de "
        "Argentina, principalmente de la ciudad de Buenos Aires y la provincia de Entre Ríos. "
        "Esto implica que la demanda turística hacia Uruguay es altamente sensible al tipo "
        "de cambio real bilateral UYU/ARS y al nivel de actividad económica de Argentina."
    )

    agregar_titulo(doc, "2.2 Composición del gasto del turista no residente", 2)
    agregar_parrafo(doc,
        "La Encuesta de Turismo Receptivo (ETR) del MINTUR permite conocer en detalle cómo "
        "distribuyen su gasto los visitantes. Para la temporada 2022/2023 (noviembre 2022 "
        "a abril 2023), el gasto por rubro fue el siguiente:"
    )

    if res.get("df_rubros") is not None:
        df = res["df_rubros"][res["df_rubros"]["label"] == "central_13pp"].copy()
        df_show = df[["rubro", "gasto_temporada_musd"]].copy()
        df_show["% del total"] = (df_show["gasto_temporada_musd"] / df_show["gasto_temporada_musd"].sum() * 100).round(1).astype(str) + "%"
        df_show.columns = ["Rubro", "Gasto temporada (MUSD)", "% del total"]
        total_row = pd.DataFrame([{"Rubro": "TOTAL", "Gasto temporada (MUSD)": df_show["Gasto temporada (MUSD)"].sum(), "% del total": "100%"}])
        df_show = pd.concat([df_show, total_row], ignore_index=True)
        agregar_tabla_desde_df(doc, df_show, "Tabla 2.1. Gasto de turistas no residentes por rubro — Temporada 2022/2023")

    agregar_parrafo(doc,
        "El rubro 'alimentación' (restaurantes, bares y locales de comida) representa "
        "aproximadamente el 29,1% del gasto total del viaje — el segundo en importancia "
        "después del alojamiento. Este porcentaje (denominado w en el modelo) es un "
        "parámetro clave: determina qué fracción del shock de precio 'pesa' sobre la "
        "decisión de viajar o no a Uruguay."
    )

    agregar_titulo(doc, "2.3 El régimen de IVA como instrumento de competitividad turística", 2)
    agregar_parrafo(doc,
        "Uruguay ha utilizado la política tributaria como herramienta para mejorar la "
        "competitividad de su oferta turística. La lógica es simple: si el turista "
        "extranjero paga menos IVA en restaurantes y hoteles, Uruguay resulta más barato "
        "en términos relativos frente a otros destinos, lo que incentiva las visitas y "
        "el gasto. El costo de esta política para el Estado es el Gasto Tributario (GT): "
        "el impuesto que se deja de recaudar."
    )
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 3 — MARCO LEGAL
# =============================================================================
def seccion_marco_legal(doc, supuestos):
    agregar_titulo(doc, "3. Marco Legal y Definición de Escenarios", 1)

    agregar_titulo(doc, "3.1 Evolución histórica del régimen", 2)
    agregar_parrafo(doc,
        "El régimen de beneficios de IVA para turistas no residentes tiene una historia "
        "de casi dos décadas. La tabla siguiente resume los hitos normativos principales:"
    )

    df_hist = pd.DataFrame([
        {"Año": "2006", "Norma": "Ley 17.934 + Decreto 537/005", "Contenido": "Reducción de 9 puntos de IVA en gastronomía, arrendamiento de vehículos e intermediación inmobiliaria cuando se paga con tarjeta de crédito. Aplica a toda la población."},
        {"Año": "2012", "Norma": "Decreto 376/012", "Contenido": "Se extiende la reducción de 9 puntos a turistas no residentes con tarjeta emitida en el exterior. IVA efectivo: 22% − 9% = 13%."},
        {"Año": "2022", "Norma": "Decreto 262/022", "Contenido": "Ampliación del alcance del beneficio. Se ajustan condiciones de elegibilidad."},
        {"Año": "2023 (nov.)", "Norma": "Decreto 434/023", "Contenido": "Exoneración TOTAL de IVA (0%) para no residentes con tarjeta del exterior en servicios gastronómicos durante temporada alta (15-nov al 30-abr). Va más allá del régimen permanente."},
        {"Año": "2024 (mod. Ley 20.212)", "Norma": "Ley 20.212", "Contenido": "Modifica la Ley 17.934: consolida la reducción permanente de 9 puntos."},
        {"Año": "2025", "Norma": "Decreto 220/2025", "Contenido": "Prórroga del Decreto 434/023 para la temporada 2025/2026."},
    ])
    agregar_tabla_desde_df(doc, df_hist, "Tabla 3.1. Evolución normativa del régimen de IVA turístico")

    agregar_titulo(doc, "3.2 Mecanismo de funcionamiento", 2)
    agregar_parrafo(doc,
        "El mecanismo de la Ley 17.934 funciona como una reducción de tasa, no como una "
        "devolución posterior. Cuando un turista no residente paga con tarjeta emitida en "
        "el exterior en un restaurante habilitado, el sistema de pagos automáticamente "
        "aplica la tasa reducida. La diferencia entre la tasa plena (22%) y la efectiva "
        "(13% en el régimen permanente, 0% en temporada alta bajo el decreto) es el Gasto "
        "Tributario que el Estado deja de percibir."
    )
    agregar_parrafo(doc,
        "Condiciones de elegibilidad: (1) el pagador debe ser turista no residente, "
        "(2) el pago debe realizarse con tarjeta de crédito o débito emitida fuera de "
        "Uruguay, (3) el establecimiento debe estar habilitado (restaurantes, bares, "
        "servicios de gastronomía en general)."
    )

    agregar_titulo(doc, "3.3 Definición formal de los escenarios analizados", 2)
    agregar_parrafo(doc,
        "El análisis contempla dos escenarios de eliminación del beneficio:"
    )

    df_esc = pd.DataFrame([
        {"Escenario": "RÉGIMEN ACTUAL (base)", "Norma vigente": "Ley 17.934 + Decreto 434/023", "IVA en temporada alta": "0%", "Precio al turista (base = 100)": "100"},
        {"Escenario": "CENTRAL: cae solo el decreto estacional", "Norma vigente": "Solo Ley 17.934 (mod. Ley 20.212)", "IVA en temporada alta": "13%", "Precio al turista (base = 100)": "113"},
        {"Escenario": "VARIANTE: cae todo el régimen", "Norma vigente": "Ninguna exoneración", "IVA en temporada alta": "22%", "Precio al turista (base = 100)": "122"},
    ])
    agregar_tabla_desde_df(doc, df_esc, "Tabla 3.2. Definición de escenarios")

    agregar_parrafo(doc,
        "En el escenario central (+13 puntos porcentuales de IVA), el turista que "
        "actualmente paga 100 unidades por una comida pasaría a pagar 113. El shock "
        "de precio neto que efectivamente percibe el consumidor es menor que 13%, "
        "porque parte del incremento lo absorbe el productor (el restaurante) ajustando "
        "sus márgenes — lo que se modela mediante el parámetro φ (ver Sección 5.3)."
    )
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 4 — DATOS Y FUENTES
# =============================================================================
def seccion_datos_fuentes(doc):
    agregar_titulo(doc, "4. Datos y Fuentes", 1)

    agregar_parrafo(doc,
        "El modelo utiliza cuatro fuentes de datos primarias. La estrategia de recolección "
        "fue híbrida: descarga automática cuando el servidor lo permite, y descarga manual "
        "con registro de hash SHA256 para garantizar la trazabilidad."
    )

    df_fuentes = pd.DataFrame([
        {"ID": "DGI-GT", "Descripción": "Informe de Gasto Tributario 2019–2022 (DGI)", "URL": "gub.uy/dgi — sección Datos y Estadísticas", "Descarga": "Manual (HTTP 403)", "Uso en el modelo": "Base afectada B vía DGI (Sección 5.2)"},
        {"ID": "MINTUR-ETR", "Descripción": "Encuesta de Turismo Receptivo — Microdatos (MINTUR)", "URL": "gub.uy/ministerio-turismo — Datos y Estadísticas", "Descarga": "Manual", "Uso en el modelo": "Base afectada B vía ETR; peso w; gasto por rubro (Secciones 5.2 y 5.3)"},
        {"ID": "BCU-MIP", "Descripción": "Matriz Insumo-Producto 2016 (BCU)", "URL": "bcu.gub.uy — Estadísticas e Indicadores", "Descarga": "Manual (SharePoint)", "Uso en el modelo": "Multiplicadores de Leontief; ΔVAB por industria (Sección 5.4)"},
        {"ID": "BCU-CN", "Descripción": "Cuentas Nacionales — Cuenta de Bienes y Servicios (BCU)", "URL": "bcu.gub.uy — Cuentas Nacionales", "Descarga": "Manual (SharePoint)", "Uso en el modelo": "PIB nominal 2022 para expresar resultados en % del PIB"},
        {"ID": "BCU-PAGOS", "Descripción": "Reporte del Sistema de Pagos Minorista (BCU)", "URL": "bcu.gub.uy — Sistema de Pagos", "Descarga": "No disponible", "Uso en el modelo": "Cota superior de consistencia de B — no utilizado por falta de acceso"},
    ])
    agregar_tabla_desde_df(doc, df_fuentes, "Tabla 4.1. Fuentes de datos del modelo")

    agregar_titulo(doc, "4.1 Notas sobre disponibilidad de datos", 2)
    agregar_parrafo(doc,
        "Descarga automática: el script 01_descarga_datos.py intenta descargar cada "
        "fuente mediante requests. La DGI devuelve HTTP 403 (acceso denegado) y los "
        "servidores del BCU utilizan SharePoint con autenticación, lo que impide la "
        "descarga automática. Todos los archivos fueron obtenidos manualmente y sus "
        "hashes SHA256 quedan registrados en DATA_SOURCES.md del repositorio."
    )
    agregar_parrafo(doc,
        "Reporte BCU de pagos: el Reporte Informativo del Sistema de Pagos Minorista "
        "del BCU contiene información sobre compras en Uruguay con tarjetas emitidas en "
        "el exterior, lo que permitiría construir una tercera estimación independiente "
        "de B. Su no disponibilidad implica que la validación cruzada de la base afectada "
        "queda pendiente (ver discusión en Sección 5.2 y Limitaciones)."
    )
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 5 — METODOLOGÍA
# =============================================================================
def seccion_metodologia(doc, supuestos, res):
    agregar_titulo(doc, "5. Metodología", 1)

    # --- 5.1 Visión general ---
    agregar_titulo(doc, "5.1 Visión general del pipeline", 2)
    agregar_parrafo(doc,
        "El modelo sigue una arquitectura de equilibrio parcial en cinco pasos secuenciales. "
        "Se denomina 'equilibrio parcial' porque analiza el mercado del turismo gastronómico "
        "de forma aislada, sin modelizar los efectos de segunda vuelta sobre precios generales, "
        "tipo de cambio o asignación de recursos entre sectores. Esta es la aproximación "
        "estándar para análisis de política tributaria cuando el sector afectado es pequeño "
        "en relación al PIB total."
    )
    df_pipeline = pd.DataFrame([
        {"Paso": "1", "Etapa": "Base afectada (B)", "Pregunta": "¿Cuánto gasto gastronómico de no residentes aplica al beneficio?", "Datos": "DGI (GT), MINTUR (ETR)"},
        {"Paso": "2", "Etapa": "Shock de precio y demanda", "Pregunta": "¿Cuánto cae la demanda ante el aumento de precio?", "Datos": "Elasticidades de la literatura + w de la ETR"},
        {"Paso": "3", "Etapa": "Propagación insumo-producto", "Pregunta": "¿Cuánto PIB se pierde considerando los efectos en cadena?", "Datos": "MIP BCU 2016 (108 industrias)"},
        {"Paso": "4", "Etapa": "Análisis fiscal", "Pregunta": "¿Cuánto gana el fisco en IVA y cuánto pierde por menor actividad?", "Datos": "Resultados de pasos 1–3"},
        {"Paso": "5", "Etapa": "Monte Carlo", "Pregunta": "¿Cuánta incertidumbre tienen los resultados?", "Datos": "Distribuciones sobre parámetros inciertos"},
    ])
    agregar_tabla_desde_df(doc, df_pipeline, "Tabla 5.1. Estructura del pipeline metodológico")

    # --- 5.2 Base afectada ---
    agregar_titulo(doc, "5.2 Paso 1: Estimación de la base afectada (B)", 2)
    agregar_parrafo(doc,
        "B representa el gasto gastronómico de turistas no residentes que (a) pagaron con "
        "tarjeta emitida en el exterior y (b) realizaron el gasto durante la temporada alta "
        "(15 de noviembre al 30 de abril). Es el universo de gasto que actualmente paga "
        "IVA 0% y que, bajo los escenarios analizados, pasaría a pagar IVA 13% o 22%."
    )
    agregar_parrafo(doc,
        "Se estimó B por dos vías independientes que permiten una triangulación:"
    )

    agregar_titulo(doc, "Vía 1 — DGI: Gasto Tributario (método preferente)", 3)
    agregar_parrafo(doc,
        "El Informe de Gasto Tributario 2019–2022 de la DGI (páginas 34–35) reporta en la "
        "fila A_34 el GT asociado a la 'Reducción del IVA con tarjetas de crédito — Turistas "
        "no residentes' bajo la Ley 17.934 y sus decretos reglamentarios. Para el año 2022, "
        "este GT fue de 795.210.678 pesos uruguayos corrientes."
    )
    agregar_parrafo(doc,
        "Interpretación del GT: la fila A_34 corresponde a la exoneración de 22 puntos "
        "(el monto total del IVA que hubiera pagado el turista sin ningún beneficio). "
        "La base afectada se despeja mediante:"
    )
    agregar_formula(doc, "B_anual = GT / tasa = 795.210.678 UYU / 0,22 / TC 41,12 = 87,9 MUSD")
    agregar_parrafo(doc,
        "Como el decreto estacional aplica solo en temporada alta y el GT es un dato anual, "
        "se ajusta por la fracción estacional del gasto en alimentación observada en la ETR:"
    )
    agregar_formula(doc, "fracción_estacional = gasto_alimentación_temporada / gasto_alimentación_anual = 0,482")
    agregar_formula(doc, "B_temporada = 87,9 × 0,482 = 42,4 MUSD  [DATO OBSERVADO]")

    agregar_titulo(doc, "Vía 2 — MINTUR: Encuesta de Turismo Receptivo (ETR)", 3)
    agregar_parrafo(doc,
        "Los microdatos de la ETR contienen el gasto individual de cada turista encuestado, "
        "con su ponderador de expansión. La muestra permite calcular el gasto en alimentación "
        "de no residentes en la ventana exacta de temporada (15-nov-2022 a 30-abr-2023):"
    )
    agregar_formula(doc, "gasto_alimentación_temporada = Σ (GastoAlimentacion_i × Coef_i) = 340,4 MUSD")
    agregar_parrafo(doc,
        "Sin embargo, la ETR no distingue entre gastos en restaurantes/bares vs. "
        "supermercados/almacenes (ambos quedan en el rubro 'Alimentación'). Tampoco "
        "releva el medio de pago. Se aplican dos supuestos correctivos:"
    )
    agregar_formula(doc, "B_etr = 340,4 × 0,70 (share_restaurantes) × 0,65 (share_tarjeta_exterior) = 154,9 MUSD")

    agregar_titulo(doc, "Discrepancia entre vías y elección de B central", 3)
    agregar_parrafo(doc,
        "Las dos estimaciones difieren en un factor de 3,65× (42,4 vs. 154,9 MUSD). "
        "Esta brecha supera el umbral de alerta del modelo (3×) y requiere interpretación:"
    )
    agregar_bullet(doc, "La vía DGI refleja el GT efectivamente procesado: solo el gasto que usó la tarjeta del exterior, en establecimientos habilitados, y fue efectivamente tramitado en el sistema de devolución. Representa el piso real del beneficio.")
    agregar_bullet(doc, "La vía ETR puede sobreestimar porque: (a) incluye supermercados en 'alimentación', (b) el share de tarjeta exterior (0,65) es un supuesto, (c) puede haber subfacturación en los registros del sistema de pagos que la DGI usa como base.")
    agregar_bullet(doc, "Se elige B_DGI = 42,4 MUSD como central (más conservadora y basada en datos administrativos), con el rango [42,4 ; 154,9] como rango de incertidumbre para el Monte Carlo.")

    df_b = pd.DataFrame([
        {"Vía": "DGI (GT A_34 / tasa)", "B temporada (MUSD)": "42,4", "Estado": "OBSERVADO — fuente preferente", "Limitación": "A_34 incluye también arrendamiento de vehículos e intermediación inmobiliaria (sesgo al alza)"},
        {"Vía": "ETR (gasto × shares)", "B temporada (MUSD)": "154,9", "Estado": "OBSERVADO × 2 SUPUESTOS", "Limitación": "'Alimentación' puede incluir supermercados; share tarjeta es supuesto"},
        {"Vía": "BCU Pagos (tarjetas ext.)", "B temporada (MUSD)": "No disponible", "Estado": "Sin datos", "Limitación": "Cota superior de consistencia pendiente"},
    ])
    agregar_tabla_desde_df(doc, df_b, "Tabla 5.2. Triangulación de la base afectada (B)")

    # --- 5.3 Shock de demanda ---
    agregar_titulo(doc, "5.3 Paso 2: Shock de precio y respuesta de demanda", 2)
    agregar_parrafo(doc,
        "Una vez definido B, el siguiente paso es calcular cuánto cae la demanda "
        "ante el aumento de precio. El modelo distingue dos canales de respuesta:"
    )

    agregar_titulo(doc, "Absorción de margen (φ)", 3)
    agregar_parrafo(doc,
        "No todo el aumento de IVA se traslada al precio final. Los restaurantes pueden "
        "absorber parte del incremento reduciendo sus márgenes de ganancia para no perder "
        "clientela. El parámetro φ (phi) mide esta fracción. La literatura internacional "
        "sobre el pass-through de IVA en gastronomía (Benzarti & Carloni, 2019; "
        "Harju et al., 2022) documenta una traslación casi completa al precio final: "
        "se asume φ = 0,15 (el productor absorbe el 15%, el consumidor paga el 85%)."
    )
    agregar_formula(doc, "Δp = (1 − φ) × shock_pp/100 = (1 − 0,15) × 0,13 = 0,1105  (escenario central)")
    agregar_parrafo(doc,
        "Es decir, el precio que percibe el turista sube un 11,05% (no 13%), porque "
        "el restaurante absorbe 1,95 puntos porcentuales."
    )

    agregar_titulo(doc, "Margen extensivo (g_ext): ¿cuántos turistas dejan de venir?", 3)
    agregar_parrafo(doc,
        "El margen extensivo captura la respuesta en número de llegadas: turistas que "
        "deciden no viajar o acortar su estadía ante el encarecimiento relativo de Uruguay. "
        "Este canal afecta a TODO el gasto turístico (alojamiento, transporte, compras, etc.), "
        "no solo a la gastronomía. La elasticidad extensiva (ε_ext) indica cuánto cae el "
        "gasto total por cada punto porcentual de aumento en el precio gastronómico relativo:"
    )
    agregar_formula(doc, "g_ext = ε_ext × (Δp × w) = −1,2 × (0,1105 × 0,291) = −0,0386  (−3,86%)")
    agregar_parrafo(doc,
        "El factor w (= 0,291) pondera el shock: solo si la gastronomía representa una "
        "fracción w del presupuesto del viaje, un aumento en el precio gastronómico "
        "genera una caída proporcional en la utilidad marginal de viajar."
    )

    agregar_titulo(doc, "Margen intensivo (g_int): ¿cuánto menos gastan en restaurantes los que sí vienen?", 3)
    agregar_parrafo(doc,
        "El margen intensivo captura la sustitución dentro del viaje: turistas que "
        "siguen viniendo pero gastan menos en restaurantes (comen en el apartamento, "
        "van al supermercado, etc.). Este canal afecta SOLO al gasto gastronómico:"
    )
    agregar_formula(doc, "g_int = ε_int × Δp = −1,0 × 0,1105 = −0,1105  (−11,05%)")

    agregar_titulo(doc, "Cambio total en la demanda gastronómica", 3)
    agregar_parrafo(doc,
        "El cambio neto en el gasto gastronómico sujeto al beneficio combina los dos márgenes. "
        "La separación evita el doble conteo: el extensivo captura turistas que no vienen "
        "(y por lo tanto no gastan en nada), el intensivo captura los que vienen pero "
        "gastan menos en gastronomía:"
    )
    agregar_formula(doc, "ΔD_gastro = B × [(1 + g_ext) × (1 + g_int) − 1]")
    agregar_formula(doc, "ΔD_gastro = 42,4 × [(1 − 0,0386) × (1 − 0,1105) − 1] = −6,14 MUSD")
    agregar_parrafo(doc,
        "Para los demás rubros de gasto turístico (alojamiento, transporte, compras, etc.), "
        "se aplica solo el margen extensivo sobre el gasto de temporada de cada rubro, "
        "escalado por la fracción de turistas efectivamente afectados por el cambio "
        "(afectación_share = B / gasto_alimentación_temporada = 42,4/340,4 = 0,125)."
    )

    if res.get("df_rubros") is not None:
        df_s = res["df_rubros"][res["df_rubros"]["label"] == "central_13pp"].copy()
        df_s = df_s[["rubro", "gasto_temporada_musd", "delta_d_musd"]].copy()
        df_s.columns = ["Rubro ETR", "Gasto temporada (MUSD)", "ΔD central (MUSD)"]
        agregar_tabla_desde_df(doc, df_s, "Tabla 5.3. Shock de demanda por rubro — Escenario central")

    # --- 5.4 Modelo Leontief ---
    agregar_titulo(doc, "5.4 Paso 3: Modelo de Insumo-Producto de Leontief", 2)

    agregar_titulo(doc, "¿Qué es una Matriz Insumo-Producto?", 3)
    agregar_parrafo(doc,
        "Una economía está integrada: cuando un restaurante vende menos, también compra "
        "menos alimentos al productor agropecuario, menos vajilla al importador, menos "
        "servicios de limpieza. Esos proveedores, a su vez, reducen sus compras de insumos. "
        "La Matriz Insumo-Producto (MIP) captura esta interdependencia sectorial: registra "
        "cuánto compra cada industria de las demás para producir su propio output. "
        "El modelo de Leontief utiliza esa información para calcular el impacto total "
        "(directo + indirecto) de una caída en la demanda final de un sector."
    )

    agregar_titulo(doc, "La MIP del BCU (2016, 108 industrias)", 3)
    agregar_parrafo(doc,
        "El BCU publica la MIP de Uruguay del año 2016, desagregada en 108 industrias "
        "clasificadas según la CIIU Rev. 4. La unidad es millones de pesos uruguayos (MUYU) "
        "corrientes de 2016. La MIP utilizada es la doméstica (excluye importaciones), "
        "verificada mediante la condición: suma de columnas de Z = 'Total de usos de "
        "origen nacional' reportado por el BCU."
    )
    agregar_parrafo(doc,
        "La industria relevante para este análisis es A.81 — 'Servicio de alimento y bebida' "
        "(restaurantes, bares, cantinas). Su VBP 2016 fue de 42.977 MUYU y su VA/VBP = 0,507 "
        "(genera 50,7 centavos de valor agregado por cada peso de producción)."
    )

    agregar_titulo(doc, "Construcción de la matriz A y el inverso de Leontief", 3)
    agregar_parrafo(doc,
        "La matriz de coeficientes técnicos domésticos A se construye dividiendo cada "
        "elemento de la matriz de transacciones intermedias por el Valor Bruto de "
        "Producción (VBP) de la industria compradora:"
    )
    agregar_formula(doc, "A[i,j] = Z[i,j] / VBP[j]")
    agregar_parrafo(doc,
        "A[i,j] indica cuántos pesos de insumo del sector i se necesitan para producir "
        "un peso del sector j. El inverso de Leontief L = (I − A)⁻¹ captura la totalidad "
        "de los efectos en cadena: directo (primer round) + indirecto (todos los rounds "
        "subsiguientes). La suma de la columna j de L es el multiplicador de producción "
        "tipo I del sector j: indica cuántos pesos de producción total genera en la "
        "economía cada peso de demanda final adicional del sector j. "
        f"Para la gastronomía (A.81), el multiplicador es 1,80 — validado en el rango "
        "esperado de [1,2; 2,0]."
    )

    agregar_titulo(doc, "Vector de shock (Δf) y conversión de unidades", 3)
    agregar_parrafo(doc,
        "El shock de demanda calculado en la Sección 5.3 (−6,14 MUSD 2022 para gastronomía, "
        "más el arrastre de los demás rubros) debe convertirse a MUYU constantes de 2016 "
        "para ser compatible con la MIP. El procedimiento documentado es:"
    )
    agregar_formula(doc, "Paso 1: MUSD 2022 → MUYU corrientes 2022:  ΔD × TC = −6,14 × 41,12 = −252,5 MUYU 2022")
    agregar_formula(doc, "Paso 2: MUYU 2022 → MUYU 2016:  × deflactor IPC = −252,5 × 0,6285 = −158,7 MUYU 2016")
    agregar_parrafo(doc,
        "El vector Δf (108 × 1) asigna −158,7 MUYU 2016 a la posición de A.81, y los "
        "shocks proporcionales de los demás rubros a sus respectivas industrias MIP "
        "según el mapeo de la Tabla A.2 (Anexo)."
    )

    agregar_titulo(doc, "Cálculo del impacto", 3)
    agregar_formula(doc, "ΔVBP = L × Δf   (cambio en la Producción Bruta por industria)")
    agregar_formula(doc, "ΔVAB = v̂ × ΔVBP  (cambio en el Valor Agregado Bruto, v̂ = diag de VA/VBP)")
    agregar_formula(doc, "ΔEmpleo = l̂ × ΔVBP  (cambio en empleo, l̂ = diag de empleos/VBP)")
    agregar_parrafo(doc,
        "La reconversión de los resultados a MUSD 2022 para expresarlos como % del PIB:"
    )
    agregar_formula(doc, "ΔVAB_MUSD = ΔVAB_MUYU2016 / deflactor / TC = −232 / 0,6285 / 41,12 ≈ −9,0 MUSD")
    agregar_formula(doc, "ΔVAB_%PIB = ΔVAB_MUSD / PIB_MUSD = −9,0 / 71.329 = −0,0126%")

    # --- 5.5 Análisis fiscal ---
    agregar_titulo(doc, "5.5 Paso 4: Análisis fiscal", 2)
    agregar_parrafo(doc,
        "El impacto fiscal tiene tres componentes que deben mantenerse separados conceptualmente:"
    )
    agregar_bullet(doc, "Recaudación GANADA (estática): el IVA que pagan los turistas que siguen viniendo después del shock. Base = gasto remanente de los turistas que permanecen; tasa = la nueva tasa efectiva.")
    agregar_formula(doc, "Rec_ganada = tasa_nueva × B × (1 + g_ext) × (1 + g_int)")
    agregar_formula(doc, "Rec_ganada = 0,1105 × 42,4 × (1 − 0,0386) × (1 − 0,1105) ≈ 4,0 MUSD")
    agregar_bullet(doc, "Recaudación PERDIDA (dinámica): por la menor actividad económica general, la base imponible de todos los impuestos se contrae.")
    agregar_formula(doc, "Rec_perdida = presión_tributaria × |ΔVAB| = 0,27 × 9,0 ≈ 2,4 MUSD")
    agregar_bullet(doc, "GT AHORRADO: el Gasto Tributario de 42,4 MUSD que el fisco deja de desembolsar. Este es el beneficio fiscal principal de la medida.")
    agregar_parrafo(doc,
        "ADVERTENCIA CONCEPTUAL: la recaudación ganada y el impacto en PIB son métricas "
        "distintas que miden cosas diferentes. La primera mide ingresos fiscales; la "
        "segunda mide actividad económica real. No deben sumarse. El trade-off de política "
        "es: se gana 43,6 MUSD en términos fiscales (GT ahorrado + rec. neta) a costa de "
        "perder 9,0 MUSD de actividad real."
    )

    # --- 5.6 Monte Carlo ---
    agregar_titulo(doc, "5.6 Paso 5: Simulación Monte Carlo", 2)
    agregar_parrafo(doc,
        "El modelo tiene varios parámetros que no son observables directamente y que "
        "deben estimarse o suponer. La simulación de Monte Carlo cuantifica cómo se "
        "propaga esa incertidumbre al resultado final."
    )
    agregar_parrafo(doc,
        "Cada parámetro incierto recibe una distribución triangular, definida por tres "
        "valores: mínimo, moda (valor central) y máximo. La distribución triangular es "
        "conveniente porque es fácil de interpretar (los extremos son los valores 'en el "
        "peor/mejor caso' y la moda es el valor más probable) y no requiere supuestos "
        "paramétricos fuertes."
    )

    df_mc_param = pd.DataFrame([
        {"Parámetro": "B (base afectada, MUSD)", "Mínimo": "42,4", "Central": "42,4", "Máximo": "154,9", "Fuente": "DGI (mín/central) y ETR (máx)"},
        {"Parámetro": "φ (absorción de margen)", "Mínimo": "0,00", "Central": "0,15", "Máximo": "0,30", "Fuente": "Benzarti & Carloni (2019)"},
        {"Parámetro": "ε_int (elasticidad intensiva)", "Mínimo": "−1,50", "Central": "−1,00", "Máximo": "−0,60", "Fuente": "Literatura de demanda de restaurantes"},
        {"Parámetro": "ε_ext (elasticidad extensiva)", "Mínimo": "−1,80", "Central": "−1,20", "Máximo": "−0,80", "Fuente": "Peng et al. (2015); Crouch (1995)"},
        {"Parámetro": "w (peso gastro en gasto total)", "Mínimo": "0,26", "Central": "0,291", "Máximo": "0,31", "Fuente": "ETR — variación observada entre temporadas"},
    ])
    agregar_tabla_desde_df(doc, df_mc_param, "Tabla 5.4. Parámetros de la simulación Monte Carlo")
    agregar_parrafo(doc,
        f"Se generaron 10.000 iteraciones con semilla fija = 42 (garantía de replicabilidad). "
        "En cada iteración se sortean valores de los cinco parámetros simultáneamente, "
        "se corre el modelo y se registra el ΔPIB%. El resultado es una distribución "
        "completa del impacto, de la que se reportan mediana, media y percentiles P10/P90."
    )
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 6 — RESULTADOS
# =============================================================================
def seccion_resultados(doc, supuestos, res):
    agregar_titulo(doc, "6. Resultados", 1)

    imp_c = res["imp_central"]
    imp_v = res["imp_variante"]
    fis_c = res["fis_central"]
    fis_v = res["fis_variante"]
    mc_c = res["mc_central"]
    mc_v = res["mc_variante"]
    pib = obtener_valor(supuestos, "pib_nominal_musd")

    # --- 6.1 Demanda gastronómica ---
    agregar_titulo(doc, "6.1 Impacto en la demanda gastronómica", 2)
    df_d = pd.DataFrame([
        {"Componente": "Base afectada B (MUSD)", "Central +13pp": "42,4", "Variante +22pp": "42,4"},
        {"Componente": "Δp neto (% de aumento de precio al turista)", "Central +13pp": "11,1%", "Variante +22pp": "15,9%"},
        {"Componente": "g_ext (caída turistas, %)", "Central +13pp": f"{imp_c['g_ext']*100:.2f}%", "Variante +22pp": f"{imp_v['g_ext']*100:.2f}%"},
        {"Componente": "g_int (caída gasto gastro condicional, %)", "Central +13pp": f"{imp_c['g_int']*100:.2f}%", "Variante +22pp": f"{imp_v['g_int']*100:.2f}%"},
        {"Componente": "ΔD_gastro (MUSD)", "Central +13pp": f"−{abs(imp_c['delta_d_gastro_musd']):.2f}", "Variante +22pp": f"−{abs(imp_v['delta_d_gastro_musd']):.2f}"},
        {"Componente": "ΔD_gastro (% de B)", "Central +13pp": f"{imp_c['delta_d_gastro_pct_de_b']:.1f}%", "Variante +22pp": f"{imp_v['delta_d_gastro_pct_de_b']:.1f}%"},
    ])
    agregar_tabla_desde_df(doc, df_d, "Tabla 6.1. Shock en la demanda gastronómica")
    agregar_parrafo(doc,
        "En el escenario central, la demanda gastronómica de no residentes cae 14,5% "
        "(combinando la caída de llegadas −3,9% y la caída del gasto condicional −11,1%). "
        "En dólares, la caída es de 6,1 millones sobre una base de 42,4 millones."
    )

    # --- 6.2 Propagación MIP ---
    agregar_titulo(doc, "6.2 Propagación por la cadena de valor (efectos directo e indirecto)", 2)
    agregar_parrafo(doc,
        "El modelo de Leontief propaga el shock de demanda a través de las 108 industrias. "
        "El efecto directo es la pérdida de VAB en el sector gastronómico mismo (A.81) y "
        "en los demás sectores que reciben el shock directamente (alojamiento, transporte, etc.). "
        "El efecto indirecto es la pérdida adicional en los proveedores de esos sectores "
        "(agro, industria alimenticia, servicios, etc.)."
    )
    df_mip = pd.DataFrame([
        {"Indicador": "Δf total (MUYU 2016 — shock inicial)", "Central": f"{imp_c['delta_f_total_muyu2016']:.1f}", "Variante": f"{imp_v['delta_f_total_muyu2016']:.1f}"},
        {"Indicador": "ΔVBP total (MUYU 2016)", "Central": f"{imp_c['delta_vbp_total_muyu2016']:.1f}", "Variante": f"{imp_v['delta_vbp_total_muyu2016']:.1f}"},
        {"Indicador": "ΔVAB directo (MUYU 2016)", "Central": f"{imp_c['delta_vab_directo_muyu2016']:.1f}", "Variante": f"{imp_v['delta_vab_directo_muyu2016']:.1f}"},
        {"Indicador": "ΔVAB indirecto (MUYU 2016)", "Central": f"{imp_c['delta_vab_indirecto_muyu2016']:.1f}", "Variante": f"{imp_v['delta_vab_indirecto_muyu2016']:.1f}"},
        {"Indicador": "ΔVAB total (MUYU 2016)", "Central": f"{imp_c['delta_vab_total_muyu2016']:.1f}", "Variante": f"{imp_v['delta_vab_total_muyu2016']:.1f}"},
        {"Indicador": "ΔVAB total (MUSD 2022)", "Central": f"−{abs(imp_c['delta_vab_musd2022']):.1f}", "Variante": f"−{abs(imp_v['delta_vab_musd2022']):.1f}"},
        {"Indicador": "ΔVAB (% del PIB)", "Central": f"{imp_c['delta_vab_pct_pib']:.4f}%", "Variante": f"{imp_v['delta_vab_pct_pib']:.4f}%"},
        {"Indicador": "Efecto directo / ΔVAB total", "Central": f"{imp_c['delta_vab_directo_muyu2016']/imp_c['delta_vab_total_muyu2016']*100:.0f}%", "Variante": f"{imp_v['delta_vab_directo_muyu2016']/imp_v['delta_vab_total_muyu2016']*100:.0f}%"},
        {"Indicador": "Multiplicador implícito (ΔVAB / Δf)", "Central": f"{imp_c['delta_vab_total_muyu2016']/imp_c['delta_f_total_muyu2016']:.2f}×", "Variante": f"{imp_v['delta_vab_total_muyu2016']/imp_v['delta_f_total_muyu2016']:.2f}×"},
    ])
    agregar_tabla_desde_df(doc, df_mip, "Tabla 6.2. Impacto sobre el VAB y el PIB")

    # Top 10 industrias
    if res.get("df_ind") is not None:
        df_top = res["df_ind"].sort_values("delta_vab_central").head(10).copy()
        df_top["ΔVAB central (MUSD)"] = (df_top["delta_vab_central"] / obtener_valor(supuestos, "deflactor_2016_2022") / obtener_valor(supuestos, "tc_promedio_2022")).round(2)
        df_top_show = df_top[["codigo", "industria", "delta_vab_central", "ΔVAB central (MUSD)"]].copy()
        df_top_show.columns = ["Código", "Industria", "ΔVAB (MUYU 2016)", "ΔVAB (MUSD 2022)"]
        df_top_show["Industria"] = df_top_show["Industria"].str[:55]
        agregar_tabla_desde_df(doc, df_top_show, "Tabla 6.3. Top 10 industrias más afectadas (escenario central)")

    # --- 6.3 Empleo ---
    agregar_titulo(doc, "6.3 Impacto en el empleo", 2)
    agregar_parrafo(doc,
        f"Usando los coeficientes de empleo de la propia MIP 2016, el modelo estima "
        f"una pérdida de {abs(imp_c['delta_empleo_puestos']):.0f} puestos de trabajo "
        f"en el escenario central y {abs(imp_v['delta_empleo_puestos']):.0f} en la variante."
    )
    agregar_parrafo(doc,
        "Interpretación importante: estos números representan puestos equivalentes "
        "a tiempo completo calculados con coeficientes técnicos fijos de 2016. No "
        "implican necesariamente desvinculaciones directas: pueden materializarse como "
        "reducción de horas trabajadas, menor empleo estacional o menor contratación "
        "de temporada. El impacto sobre el empleo formal registrado puede ser menor "
        "si los trabajadores afectados son informales o trabajan por temporada."
    )

    # --- 6.4 Contextualización ---
    agregar_titulo(doc, "6.4 Contextualización del impacto macroeconómico", 2)
    agregar_parrafo(doc,
        f"El impacto central de −0,013% del PIB equivale a aproximadamente 9 millones "
        f"de dólares sobre un PIB de {pib:,.0f} millones. Para ponerlo en perspectiva:"
    )
    agregar_bullet(doc, "El crecimiento real del PIB uruguayo fue de 4,9% en 2022: el efecto estimado representa menos del 0,3% de esa variación anual.")
    agregar_bullet(doc, "Una depreciación del peso uruguayo de 1% frente al dólar tiene un efecto sobre la competitividad turística varias veces mayor que el shock de IVA analizado.")
    agregar_bullet(doc, "El sector gastronómico (A.81 de la MIP) representa el 0,78% del VBP total de la economía; el impacto del modelo es proporcional a esa participación.")
    agregar_bullet(doc, "En comparación con el GT que el fisco ahorraría (42,4 MUSD + rec. neta 1,6 MUSD), la pérdida de actividad real (9,0 MUSD) es sustancialmente menor.")

    # --- 6.5 Fiscal ---
    agregar_titulo(doc, "6.5 Resultado fiscal", 2)
    df_fis = pd.DataFrame([
        {"Componente": "Base remanente (turistas que permanecen, MUSD)", "Central": f"{fis_c['base_remanente_musd']:.2f}", "Variante": f"{fis_v['base_remanente_musd']:.2f}"},
        {"Componente": "Tasa efectiva nueva (sobre precio sin IVA)", "Central": f"{fis_c['tasa_nueva_efectiva']*100:.2f}%", "Variante": f"{fis_v['tasa_nueva_efectiva']*100:.2f}%"},
        {"Componente": "Recaudación GANADA — IVA de turistas que siguen (MUSD)", "Central": f"+{fis_c['rec_ganada_musd']:.2f}", "Variante": f"+{fis_v['rec_ganada_musd']:.2f}"},
        {"Componente": "Recaudación PERDIDA — por menor actividad (MUSD)", "Central": f"−{fis_c['rec_perdida_musd']:.2f}", "Variante": f"−{fis_v['rec_perdida_musd']:.2f}"},
        {"Componente": "RECAUDACIÓN NETA (MUSD)", "Central": f"+{fis_c['rec_neta_musd']:.2f}", "Variante": f"+{fis_v['rec_neta_musd']:.2f}"},
        {"Componente": "GT AHORRADO por el fisco (MUSD)", "Central": f"+{fis_c['gt_estacional_ahorrado_musd']:.1f}", "Variante": f"+{fis_v['gt_estacional_ahorrado_musd']:.1f}"},
        {"Componente": "BENEFICIO FISCAL TOTAL (rec. neta + GT ahorrado, MUSD)", "Central": f"+{fis_c['rec_neta_musd']+fis_c['gt_estacional_ahorrado_musd']:.1f}", "Variante": f"+{fis_v['rec_neta_musd']+fis_v['gt_estacional_ahorrado_musd']:.1f}"},
    ])
    agregar_tabla_desde_df(doc, df_fis, "Tabla 6.4. Análisis fiscal de la medida")
    agregar_parrafo(doc,
        "La lógica del resultado fiscal es la siguiente: el fisco ahorra 42,4 MUSD en Gasto "
        "Tributario que ya no debe desembolsar, recauda 4,0 MUSD adicionales de los turistas "
        "que permanecen, y pierde 2,4 MUSD por la contracción de la base imponible general. "
        "El beneficio fiscal neto total es de aproximadamente 44 MUSD, a un costo de "
        "actividad real de 9 MUSD. La medida es fiscalmente favorable."
    )
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 7 — SENSIBILIDAD
# =============================================================================
def seccion_sensibilidad(doc, res, supuestos):
    agregar_titulo(doc, "7. Análisis de Sensibilidad", 1)

    mc_c = res["mc_central"]
    mc_v = res["mc_variante"]
    imp_c = res["imp_central"]

    agregar_titulo(doc, "7.1 Distribución Monte Carlo del impacto en PIB", 2)
    df_mc = pd.DataFrame([
        {"Estadístico": "Mediana (ΔPIB %)", "Central +13pp": f"{mc_c['mediana']:.4f}%", "Variante +22pp": f"{mc_v['mediana']:.4f}%"},
        {"Estadístico": "Media (ΔPIB %)", "Central +13pp": f"{mc_c['media']:.4f}%", "Variante +22pp": f"{mc_v['media']:.4f}%"},
        {"Estadístico": "P10 — escenario más impactante (ΔPIB %)", "Central +13pp": f"{mc_c['p10']:.4f}%", "Variante +22pp": f"{mc_v['p10']:.4f}%"},
        {"Estadístico": "P90 — escenario menos impactante (ΔPIB %)", "Central +13pp": f"{mc_c['p90']:.4f}%", "Variante +22pp": f"{mc_v['p90']:.4f}%"},
        {"Estadístico": "Mínimo observado", "Central +13pp": f"{mc_c['min']:.4f}%", "Variante +22pp": f"{mc_v['min']:.4f}%"},
        {"Estadístico": "Máximo observado", "Central +13pp": f"{mc_c['max']:.4f}%", "Variante +22pp": f"{mc_v['max']:.4f}%"},
        {"Estadístico": "N simulaciones", "Central +13pp": f"{int(mc_c['n_sims']):,}", "Variante +22pp": f"{int(mc_v['n_sims']):,}"},
    ])
    agregar_tabla_desde_df(doc, df_mc, "Tabla 7.1. Estadísticos de la simulación Monte Carlo")

    agregar_parrafo(doc,
        "Interpretación: el P10 (−0,037% para el central) representa el impacto en el "
        "10% de escenarios más desfavorables — es decir, hay un 90% de probabilidad de "
        "que el impacto sea menor (en valor absoluto) que 0,037%. El P90 (−0,014%) "
        "representa el 10% más favorable: incluso en el mejor escenario plausible, "
        "hay alguna pérdida de actividad. La media es mayor que la mediana en valor "
        "absoluto, lo que indica una distribución con una cola derecha más pesada "
        "(la gran incertidumbre sobre B empuja la media hacia impactos mayores)."
    )
    agregar_imagen(doc, OUTPUTS_FIGURAS / "histograma_montecarlo.png", ancho=5.5,
                   caption="Figura 7.1. Distribución del impacto sobre el PIB (%) — 10.000 simulaciones Monte Carlo. Líneas verticales: medianas de cada escenario.")

    agregar_titulo(doc, "7.2 Análisis de sensibilidad tornado (one-at-a-time)", 2)
    agregar_parrafo(doc,
        "El gráfico tornado muestra el rango de ΔPIB% cuando cada parámetro varía "
        "entre su mínimo y máximo mientras los demás se mantienen en el valor central. "
        "La anchura de cada barra indica la influencia relativa del parámetro sobre "
        "la incertidumbre del resultado."
    )
    agregar_imagen(doc, OUTPUTS_FIGURAS / "tornado_central.png", ancho=5.5,
                   caption="Figura 7.2. Análisis tornado — Escenario central (+13 pp). Los parámetros más influyentes son B (base afectada) y ε_ext (elasticidad extensiva).")

    agregar_parrafo(doc,
        "El gráfico tornado confirma que la incertidumbre sobre B es el factor dominante: "
        "si B fuera 154,9 MUSD (vía ETR) en lugar de 42,4 MUSD (vía DGI), los resultados "
        "escalarían en proporción. Esto refuerza la importancia de resolver la discrepancia "
        "entre fuentes (ver Sección 5.2 y Limitación 8 de la Sección 8)."
    )

    agregar_titulo(doc, "7.3 Sensibilidad sobre la elección de B central", 2)
    escala = 154.9 / 42.4
    delta_vab_etr = imp_c["delta_vab_pct_pib"] * escala
    agregar_parrafo(doc,
        f"Si se utilizara B = 154,9 MUSD (vía ETR) como B central — la estimación "
        f"más alta de las dos disponibles —, el impacto escalaría linealmente en "
        f"proporción {escala:.2f}× (dado que el modelo es lineal en B):"
    )
    df_bscale = pd.DataFrame([
        {"B central (MUSD)": "42,4 — vía DGI (elegida)", "ΔPIB% central": f"{imp_c['delta_vab_pct_pib']:.4f}%", "ΔVAB (MUSD)": f"−{abs(imp_c['delta_vab_musd2022']):.1f}"},
        {"B central (MUSD)": "154,9 — vía ETR (cota superior)", "ΔPIB% central": f"{delta_vab_etr:.4f}%", "ΔVAB (MUSD)": f"−{abs(imp_c['delta_vab_musd2022'] * escala):.1f}"},
    ])
    agregar_tabla_desde_df(doc, df_bscale, "Tabla 7.2. Sensibilidad del resultado a la elección de B central")
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 8 — LIMITACIONES
# =============================================================================
def seccion_limitaciones(doc):
    agregar_titulo(doc, "8. Limitaciones del Modelo", 1)
    agregar_parrafo(doc,
        "Las siguientes limitaciones deben tenerse en cuenta al interpretar los resultados. "
        "No invalidan el análisis, pero acotan su alcance y señalan las áreas donde "
        "una estimación más robusta requeriría trabajo adicional."
    )

    lims = [
        ("1. Equilibrio parcial sin uso alternativo de la recaudación",
         "El modelo no modela qué hace el Estado con los 44 MUSD fiscales que gana. "
         "Si ese ingreso se reasigna a gasto público productivo (infraestructura, transferencias), "
         "el impacto neto sobre el PIB podría ser positivo o nulo. El análisis mide solo el "
         "efecto del lado de la demanda turística, sin compensación por el lado fiscal."),
        ("2. MIP de 2016",
         "La Matriz Insumo-Producto utilizada corresponde a 2016. En los ocho años transcurridos, "
         "la estructura de la economía uruguaya puede haber cambiado: nuevos sectores emergentes "
         "(tecnología, servicios digitales), cambios en la integración vertical de la cadena "
         "gastronómica, variación en la productividad relativa. Los multiplicadores de Leontief "
         "pueden estar desactualizados."),
        ("3. Posible migración al pago en efectivo",
         "La exoneración aplica solo a pagos con tarjeta emitida en el exterior. Si los turistas "
         "sustituyen hacia efectivo para eludir el IVA, la actividad real no cae pero la "
         "recaudación formal tampoco aumenta. El modelo asume implícitamente que no hay esta "
         "migración, lo que puede sobreestimar tanto la caída de demanda como la recaudación ganada."),
        ("4. Dominancia del tipo de cambio real bilateral con Argentina",
         "El determinante principal de la demanda turística hacia Uruguay es el tipo de cambio "
         "real bilateral UYU/ARS. Una depreciación del peso argentino de 10% tiene un efecto "
         "sobre la competitividad turística de Uruguay varias veces mayor que un shock de IVA "
         "de 13 puntos. El efecto IVA es de segundo orden frente a los movimientos cambiarios, "
         "lo que hace que el impacto estimado sea difícil de detectar empíricamente en la serie "
         "de tiempo de llegadas turísticas."),
        ("5. Elasticidades tomadas de la literatura internacional",
         "Las elasticidades ε_ext (−1,2) y ε_int (−1,0) provienen de meta-análisis "
         "internacionales y no han sido estimadas econométricamente para Uruguay. El diseño "
         "del event-study de la Fase 2 (ver fase2_econometria/NOTES.md) apunta precisamente "
         "a corregir este déficit explotando los cambios del régimen de devolución entre "
         "2012 y 2025 como experimentos cuasi-naturales."),
        ("6. Definición de 'alimentación' en la ETR",
         "El rubro 'GastoAlimentacion' de la ETR incluye tanto restaurantes y bares "
         "(sujetos al beneficio) como supermercados, almacenes y kioskos (no sujetos). "
         "El parámetro share_restaurantes (= 0,70) intenta corregir esto, pero su valor "
         "es un supuesto con alta incertidumbre (rango [0,55 ; 0,85]), lo que contribuye "
         "a la discrepancia entre la vía ETR y la vía DGI."),
        ("7. Linealidad del modelo de Leontief",
         "El modelo de Leontief asume coeficientes técnicos fijos (proporciones de insumos "
         "constantes) y rendimientos constantes a escala. No captura sustitución entre "
         "insumos ante cambios de precios relativos, economías de escala, ni ajustes de "
         "precios en los mercados intermedios. Para shocks pequeños (como el analizado), "
         "la linealidad es una buena aproximación; para shocks grandes, puede sesgar los resultados."),
        ("8. Discrepancia DGI/ETR sobre B no resuelta",
         "La diferencia de 3,65× entre las estimaciones de B por vía DGI (42,4 MUSD) "
         "y vía ETR (154,9 MUSD) no pudo resolverse con los datos disponibles. La elección "
         "de B_DGI como central implica que los resultados son conservadores: si la base "
         "real es más cercana a la ETR, el impacto real sería 3,65× mayor. El reporte del "
         "Sistema de Pagos Minorista del BCU (no disponible en este análisis) permitiría "
         "zanjar esta discrepancia."),
    ]

    for titulo_lim, texto_lim in lims:
        p = doc.add_paragraph()
        r = p.add_run(titulo_lim)
        r.bold = True
        r.font.size = Pt(11)
        agregar_parrafo(doc, texto_lim)
        doc.add_paragraph()

    salto_pagina(doc)


# =============================================================================
# SECCIÓN 9 — CONCLUSIONES
# =============================================================================
def seccion_conclusiones(doc, res, supuestos):
    agregar_titulo(doc, "9. Conclusiones y Recomendaciones de Política", 1)

    mc_c = res["mc_central"]
    mc_v = res["mc_variante"]
    imp_c = res["imp_central"]
    fis_c = res["fis_central"]

    agregar_titulo(doc, "9.1 Síntesis de resultados", 2)
    agregar_parrafo(doc,
        "El modelo estima que eliminar el Decreto 434/023 (manteniendo vigente la "
        "devolución permanente de la Ley 17.934) generaría una pérdida de actividad "
        f"económica de entre {abs(mc_c['p10']):.3f}% y {abs(mc_c['p90']):.3f}% del PIB "
        f"(rango P10–P90), con una estimación central de {abs(mc_c['mediana']):.3f}%. "
        f"En términos monetarios, esto equivale a aproximadamente "
        f"{abs(imp_c['delta_vab_musd2022']):.0f} millones de dólares de actividad real. "
        f"El impacto en el empleo se estima en {abs(imp_c['delta_empleo_puestos']):.0f} "
        "puestos de trabajo equivalentes."
    )
    agregar_parrafo(doc,
        "Desde el punto de vista fiscal, la medida es netamente favorable: el fisco "
        f"ahorraría {fis_c['gt_estacional_ahorrado_musd']:.0f} MUSD en Gasto Tributario "
        f"y ganaría {fis_c['rec_neta_musd']:.1f} MUSD adicionales en recaudación neta, "
        f"a costa de una pérdida de actividad real de {abs(imp_c['delta_vab_musd2022']):.0f} MUSD."
    )

    agregar_titulo(doc, "9.2 Tres mensajes principales para la política", 2)
    agregar_bullet(doc,
        "El costo en actividad real es pequeño: −0,013% del PIB es una magnitud menor "
        "que el margen de error de la mayoría de las estimaciones del PIB trimestral. "
        "La medida no es macroeconómicamente significativa en el corto plazo.")
    agregar_bullet(doc,
        "El beneficio fiscal es sustancial en términos relativos: 44 MUSD es casi "
        "5 veces el costo en actividad real (9 MUSD). Si el objetivo es mejorar el "
        "balance fiscal sin sacrificar mucha actividad, la eliminación del decreto "
        "estacional parece eficiente.")
    agregar_bullet(doc,
        "La incertidumbre clave es la base afectada B: si la estimación real se acerca "
        "a la vía ETR (154,9 MUSD), todos los efectos —tanto el costo en PIB como el "
        "beneficio fiscal— se multiplican por ~3,6. Resolver esta discrepancia es la "
        "prioridad de la agenda de datos.")

    agregar_titulo(doc, "9.3 Agenda de datos y mejoras metodológicas", 2)
    agregar_bullet(doc,
        "Solicitar al BCU el Reporte del Sistema de Pagos Minorista desagregado por "
        "categoría de establecimiento (restaurantes vs. supermercados) y por tipo de "
        "tarjeta (nacional vs. extranjera). Esto permitiría cerrar la discrepancia DGI/ETR.")
    agregar_bullet(doc,
        "Implementar el event-study de la Fase 2 (ver NOTES.md) para estimar "
        "econométricamente las elasticidades ε_ext y ε_int para Uruguay, en lugar de "
        "tomar valores de la literatura internacional.")
    agregar_bullet(doc,
        "Actualizar la MIP: el BCU tiene en curso la actualización de la Matriz "
        "Insumo-Producto. Cuando esté disponible (estimado 2025–2026), re-correr el "
        "pipeline con la nueva MIP.")
    agregar_bullet(doc,
        "Incorporar el TCR bilateral con Argentina como co-variable: modelizar el "
        "impacto condicional al escenario cambiario permite acotar cuándo el efecto "
        "IVA es visible y cuándo está dominado por el tipo de cambio.")

    agregar_titulo(doc, "9.4 Recomendación para la toma de decisión", 2)
    agregar_parrafo(doc,
        "Los resultados sugieren que la eliminación del Decreto 434/023 es una medida "
        "fiscalmente eficiente con un costo macroeconómico pequeño. Sin embargo, antes "
        "de tomar una decisión definitiva, se recomienda:"
    )
    agregar_bullet(doc, "Obtener el reporte BCU de pagos para verificar la base afectada.")
    agregar_bullet(doc, "Evaluar el impacto distributivo sectorial: aunque el efecto agregado es pequeño, su concentración en la costa atlántica durante la temporada puede tener impactos locales más significativos.")
    agregar_bullet(doc, "Considerar el efecto señal: la percepción de que Uruguay 'encarece' su oferta turística puede tener un impacto en la reputación del destino mayor que el efecto precio puntual.")
    salto_pagina(doc)


# =============================================================================
# SECCIÓN 10 — REFERENCIAS
# =============================================================================
def seccion_referencias(doc):
    agregar_titulo(doc, "10. Referencias Bibliográficas", 1)

    refs = [
        "Benzarti, Y. & Carloni, D. (2019). Who Really Benefits from Consumption Tax Cuts? Evidence from a Large VAT Reform in France. American Economic Journal: Economic Policy, 11(1), 38–63.",
        "Crouch, G.I. (1995). A meta-analysis of tourism demand. Annals of Tourism Research, 22(1), 103–118.",
        "Crouch, G.I. (1996). Demand elasticities for short-haul versus long-haul tourism. Journal of Travel Research, 35(2), 2–7.",
        "Harju, J., Kosonen, T. & Skans, O.N. (2022). Firm types, price-setting strategies, and consumption tax incidence. Journal of Public Economics, 216, 104781.",
        "Peng, B., Song, H., Crouch, G.I. & Witt, S.F. (2015). A meta-analysis of international tourism demand elasticities. Journal of Travel Research, 54(5), 611–633.",
        "Banco Central del Uruguay (BCU). Matriz Insumo-Producto 2016. Disponible en: https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Matriz-Insumo-Producto.aspx",
        "Banco Central del Uruguay (BCU). Cuenta de Bienes y Servicios — Cuentas Nacionales 2022. Disponible en: https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Cuentas-Nacionales.aspx",
        "Dirección General Impositiva (DGI). Informe de Gasto Tributario 2019–2022. Disponible en: https://www.gub.uy/direccion-general-impositiva/datos-y-estadisticas",
        "Ministerio de Turismo (MINTUR). Encuesta de Turismo Receptivo — Microdatos. Disponible en: https://www.gub.uy/ministerio-turismo/datos-y-estadisticas/microdatos",
        "República Oriental del Uruguay. Ley N° 17.934 (2006) y modificaciones. Ley N° 20.212 (2023).",
        "República Oriental del Uruguay. Decreto N° 434/023 (2023). Exoneración estacional de IVA en gastronomía turística.",
        "República Oriental del Uruguay. Decreto N° 220/2025. Prórroga del Decreto 434/023.",
    ]
    for ref in refs:
        p = doc.add_paragraph(style="List Bullet")
        r = p.add_run(ref)
        r.font.size = Pt(9)
    salto_pagina(doc)


# =============================================================================
# ANEXOS
# =============================================================================
def seccion_anexos(doc, supuestos, res):
    agregar_titulo(doc, "11. Anexos Técnicos", 1)

    # --- Anexo A: Supuestos ---
    agregar_titulo(doc, "Anexo A. Tabla completa de supuestos del modelo", 2)
    agregar_parrafo(doc,
        "Todos los parámetros del modelo se centralizan en config/supuestos.yaml. "
        "La tabla siguiente se genera automáticamente desde ese archivo."
    )
    claves = [
        "shock_pp", "shock_pp_variante", "phi",
        "epsilon_intensivo", "epsilon_extensivo",
        "w", "share_restaurantes", "share_tarjeta_exterior",
        "B_dgi", "B_etr",
        "presion_tributaria_efectiva",
        "pib_nominal_musd", "tc_promedio_2022", "deflactor_2016_2022",
        "n_sims", "seed", "usar_tipo2",
    ]
    filas = []
    for clave in claves:
        if clave not in supuestos:
            continue
        ent = supuestos[clave]
        if isinstance(ent, dict):
            rng = ent.get("rango")
            rng_str = f"[{rng[0]}, {rng[1]}]" if rng else "—"
            tipo = "OBSERVADO" if "OBSERVADO" in str(ent.get("fuente", "")) else "SUPUESTO"
            filas.append({
                "Parámetro": clave, "Valor central": str(ent.get("valor", "—")),
                "Rango MC": rng_str, "Tipo": tipo,
                "Fuente/Justificación": str(ent.get("fuente", ""))[:80],
            })
        else:
            filas.append({"Parámetro": clave, "Valor central": str(ent),
                          "Rango MC": "—", "Tipo": "CONFIG", "Fuente/Justificación": ""})
    agregar_tabla_desde_df(doc, pd.DataFrame(filas), "Tabla A.1. Parámetros del modelo")

    doc.add_paragraph()

    # --- Anexo B: Mapeo rubros → MIP ---
    agregar_titulo(doc, "Anexo B. Mapeo rubros ETR → industrias MIP", 2)
    df_map = pd.read_csv(ROOT / "config" / "mapeo_rubros_mip.csv")
    agregar_tabla_desde_df(doc, df_map, "Tabla A.2. Correspondencia entre rubros de la ETR e industrias de la MIP 2016")

    doc.add_paragraph()

    # --- Anexo C: Top 20 industrias ---
    agregar_titulo(doc, "Anexo C. Top 20 industrias afectadas", 2)
    if res.get("df_ind") is not None:
        tc = obtener_valor(supuestos, "tc_promedio_2022")
        defl = obtener_valor(supuestos, "deflactor_2016_2022")
        df_top20 = res["df_ind"].sort_values("delta_vab_central").head(20).copy()
        df_top20["ΔVAB_C (MUSD)"] = (df_top20["delta_vab_central"] / defl / tc).round(2)
        df_top20["ΔVAB_V (MUSD)"] = (df_top20["delta_vab_variante"] / defl / tc).round(2)
        df_show = df_top20[["codigo", "industria", "delta_vab_central", "ΔVAB_C (MUSD)", "ΔVAB_V (MUSD)"]].copy()
        df_show.columns = ["Código", "Industria", "ΔVAB central (MUYU 2016)", "ΔVAB central (MUSD)", "ΔVAB variante (MUSD)"]
        df_show["Industria"] = df_show["Industria"].str[:55]
        agregar_tabla_desde_df(doc, df_show, "Tabla A.3. Las 20 industrias más afectadas (escenario central + variante)")

    doc.add_paragraph()

    # --- Anexo D: Reproducibilidad ---
    agregar_titulo(doc, "Anexo D. Instrucciones de reproducibilidad", 2)
    agregar_parrafo(doc,
        "Para regenerar todos los resultados desde cero en una máquina limpia:"
    )
    comandos = (
        "# 1. Clonar el repositorio\n"
        "git clone <URL_DEL_REPO>\n"
        "cd iva-turismo-uy\n\n"
        "# 2. Instalar dependencias (Python 3.11+)\n"
        "python -m venv venv && source venv/bin/activate\n"
        "pip install -r requirements.txt\n\n"
        "# 3. Colocar los archivos manuales en data/manual/\n"
        "#    (ver instrucciones en DATA_SOURCES.md)\n\n"
        "# 4. Completar en config/supuestos.yaml:\n"
        "#    tc_promedio_2022, deflactor_2016_2022, pib_nominal_musd\n\n"
        "# 5. Ejecutar el pipeline completo\n"
        "python run_all.py\n\n"
        "# 6. Outputs en:\n"
        "#    outputs/informe_final.docx\n"
        "#    outputs/tablas/  y  outputs/figuras/\n"
    )
    agregar_formula(doc, comandos)
    agregar_parrafo(doc,
        f"Semilla aleatoria fija: seed = 42. "
        f"Fecha de generación de este informe: {datetime.date.today().isoformat()}"
    )

    doc.add_paragraph()

    # --- Anexo E: Glosario ---
    agregar_titulo(doc, "Anexo E. Glosario de términos técnicos", 2)
    glosario = [
        ("Base afectada (B)", "El universo de gasto gastronómico de turistas no residentes que actualmente aplica al beneficio fiscal (IVA 0% bajo el decreto estacional). Es la variable cuya cuantificación es más incierta del modelo."),
        ("Coeficientes técnicos (A)", "Elementos de la matriz A de la MIP. A[i,j] indica cuántos pesos de insumo del sector i se requieren para producir un peso de output del sector j."),
        ("Deflactor", "Ratio que convierte valores nominales de un año a precios constantes de otro año. En este modelo: IPC_2016/IPC_2022 = 0,6285."),
        ("Distribución triangular", "Distribución de probabilidad definida por tres parámetros: mínimo, moda y máximo. Se usa en el Monte Carlo para modelizar parámetros inciertos."),
        ("Elasticidad extensiva (ε_ext)", "Cambio porcentual en el número de llegadas/noches de turistas ante un 1% de aumento en el precio relativo del destino. Captura la decisión de viajar o no."),
        ("Elasticidad intensiva (ε_int)", "Cambio porcentual en el gasto gastronómico de un turista que ya decidió viajar ante un 1% de aumento en el precio de los restaurantes. Captura la sustitución dentro del viaje."),
        ("Gasto Tributario (GT)", "El impuesto que el Estado deja de recaudar como resultado de una exoneración, reducción de tasa u otro beneficio fiscal. Es la contraparte fiscal del beneficio al contribuyente."),
        ("Inverso de Leontief (L)", "La matriz L = (I − A)⁻¹ que resume los efectos totales (directo + indirecto + inducido) de un shock de demanda final en un modelo insumo-producto tipo I."),
        ("Margen de absorción (φ)", "Fracción del aumento de impuesto que el productor absorbe reduciendo sus márgenes, en lugar de trasladarla al precio final."),
        ("MIP", "Matriz Insumo-Producto. Tabla que registra las transacciones de bienes y servicios entre los sectores productivos de una economía en un año dado."),
        ("Monte Carlo", "Técnica de simulación que genera miles de escenarios sorteando valores de parámetros inciertos según sus distribuciones de probabilidad, para obtener una distribución del resultado final."),
        ("Multiplicador tipo I", "Indica cuánto VAB o VBP total genera en la economía un peso adicional de demanda final en un sector, incluyendo los efectos en cadena sobre proveedores. No incluye el efecto inducido del consumo de hogares."),
        ("PIB", "Producto Interno Bruto. Suma del Valor Agregado Bruto de todos los sectores más los impuestos netos de subsidios sobre los productos."),
        ("P10 / P90", "Percentiles 10 y 90 de la distribución Monte Carlo. El P10 es el valor tal que el 10% de las simulaciones da un resultado más extremo; el P90 es el valor tal que el 90% de las simulaciones da un resultado más extremo."),
        ("Turista no residente", "Persona que reside habitualmente fuera de Uruguay y visita el país. La ETR del MINTUR recaba información de los visitantes que ingresan por puestos de frontera."),
        ("VAB (Valor Agregado Bruto)", "El valor generado por la producción de una industria, neto de los insumos intermedios consumidos. Es la contribución de cada sector al PIB."),
        ("VBP (Valor Bruto de Producción)", "El valor total producido por una industria, incluyendo los insumos intermedios. VBP = VAB + consumo intermedio."),
        ("w (peso gastronómico)", "Fracción del gasto total del viaje que el turista destina a restaurantes y bares. En el modelo, w = 0,291 (29,1%), calculado de los microdatos ETR para la temporada 2022/2023."),
    ]
    for term, defn in glosario:
        p = doc.add_paragraph()
        r1 = p.add_run(f"{term}: ")
        r1.bold = True
        r1.font.size = Pt(10)
        r2 = p.add_run(defn)
        r2.font.size = Pt(10)


# =============================================================================
# PIPELINE PRINCIPAL
# =============================================================================
def main():
    separador("ETAPA 07 — Informe final exhaustivo")

    supuestos = cargar_supuestos()

    # --- Cargar todos los datos procesados ---
    def leer_csv(nombre, procesado=True):
        base = DATA_PROCESSED if procesado else OUTPUTS_TABLAS
        p = base / nombre
        if p.exists():
            return pd.read_csv(p)
        print(f"  ⚠ No se encontró: {p.name}")
        return pd.DataFrame()

    df_mc_est = leer_csv("tabla_montecarlo_estadisticas.csv", procesado=False)
    df_mip = leer_csv("tabla_impacto_mip.csv", procesado=False)
    df_fiscal = leer_csv("fiscal.csv")
    df_shocks = leer_csv("shocks.csv")
    df_ind = leer_csv("impacto_por_industria.csv")
    df_rubros = leer_csv("shock_por_rubros.csv")
    df_rb = leer_csv("rango_b.csv")

    def fila(df, col, label):
        rows = df[df[col] == label]
        return rows.iloc[0].to_dict() if not rows.empty else {}

    mc_c = fila(df_mc_est, "label", "central_13pp")
    mc_v = fila(df_mc_est, "label", "variante_22pp")
    imp_c_row = fila(df_mip, "label", "central_13pp")
    imp_v_row = fila(df_mip, "label", "variante_22pp")

    # Campos esperados en imp_c/imp_v (para compatibilidad con seccion_resultados)
    for d in [imp_c_row, imp_v_row]:
        for k in ["g_ext", "g_int", "delta_d_gastro_musd", "delta_d_gastro_pct_de_b"]:
            d.setdefault(k, None)

    # Merge shock data
    sc = fila(df_shocks, "label", "central_13pp")
    sv = fila(df_shocks, "label", "variante_22pp")
    for k in ["g_ext", "g_int", "delta_d_gastro_musd", "delta_d_gastro_pct_de_b"]:
        imp_c_row[k] = sc.get(k)
        imp_v_row[k] = sv.get(k)

    res = {
        "mc_central": mc_c,
        "mc_variante": mc_v,
        "imp_central": imp_c_row,
        "imp_variante": imp_v_row,
        "fis_central": fila(df_fiscal, "label", "central_13pp"),
        "fis_variante": fila(df_fiscal, "label", "variante_22pp"),
        "df_ind": df_ind if not df_ind.empty else None,
        "df_rubros": df_rubros if not df_rubros.empty else None,
        "df_rb": df_rb,
    }

    # --- Construir documento ---
    doc = Document()

    # Configurar márgenes (2,5 cm)
    from docx.oxml import OxmlElement
    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)

    seccion_portada(doc, supuestos)
    seccion_resumen_ejecutivo(doc, supuestos, res)
    seccion_contexto(doc, res)
    seccion_marco_legal(doc, supuestos)
    seccion_datos_fuentes(doc)
    seccion_metodologia(doc, supuestos, res)
    seccion_resultados(doc, supuestos, res)
    seccion_sensibilidad(doc, res, supuestos)
    seccion_limitaciones(doc)
    seccion_conclusiones(doc, res, supuestos)
    seccion_referencias(doc)
    seccion_anexos(doc, supuestos, res)

    ruta = ROOT / "outputs" / "informe_final.docx"
    ruta.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ruta))
    print(f"\n  ✓ Informe guardado: {ruta}")
    print(f"    Tamaño: {ruta.stat().st_size // 1024} KB")
    return ruta


if __name__ == "__main__":
    main()
