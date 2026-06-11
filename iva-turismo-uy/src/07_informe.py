"""
Etapa 07 — Generación del informe final en Word (python-docx).

Estructura del informe:
  1. Resumen ejecutivo
  2. Marco legal y definición del escenario
  3. Datos y fuentes
  4. Metodología
  5. Supuestos (tabla autogenerada desde supuestos.yaml)
  6. Resultados
  7. Análisis de sensibilidad
  8. Limitaciones
  9. Anexo de reproducibilidad
"""
import sys
import pathlib
import datetime
import pandas as pd
import yaml

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS, OUTPUTS_FIGURAS,
    cargar_supuestos, obtener_valor, separador
)

try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    DOCX_DISPONIBLE = True
except ImportError:
    DOCX_DISPONIBLE = False
    print("  ⚠ python-docx no instalado — instalar con: pip install python-docx")


# =============================================================================
# HELPERS DE FORMATO
# =============================================================================
def agregar_titulo(doc, texto: str, nivel: int = 1):
    doc.add_heading(texto, level=nivel)


def agregar_parrafo(doc, texto: str, negrita: bool = False):
    p = doc.add_paragraph()
    run = p.add_run(texto)
    run.bold = negrita
    return p


def agregar_tabla_desde_df(doc, df: pd.DataFrame, titulo: str = None):
    """Inserta un DataFrame como tabla Word."""
    if titulo:
        agregar_parrafo(doc, titulo, negrita=True)

    tabla = doc.add_table(rows=1, cols=len(df.columns))
    tabla.style = "Table Grid"

    # Encabezados
    encabezados = tabla.rows[0].cells
    for i, col in enumerate(df.columns):
        encabezados[i].text = str(col)
        for run in encabezados[i].paragraphs[0].runs:
            run.bold = True

    # Filas
    for _, fila in df.iterrows():
        celdas = tabla.add_row().cells
        for i, val in enumerate(fila):
            celdas[i].text = str(val) if val is not None else "N/D"

    doc.add_paragraph()


def agregar_imagen(doc, ruta: pathlib.Path, ancho_pulgadas: float = 5.5, caption: str = None):
    """Inserta una imagen si existe."""
    if ruta.exists():
        doc.add_picture(str(ruta), width=Inches(ancho_pulgadas))
        if caption:
            p = doc.add_paragraph(caption)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.italic = True
    else:
        agregar_parrafo(doc, f"[Imagen no disponible: {ruta.name}]")


# =============================================================================
# SECCIONES DEL INFORME
# =============================================================================
def seccion_resumen_ejecutivo(doc, supuestos, resultados):
    agregar_titulo(doc, "1. Resumen Ejecutivo", 1)

    delta_central = resultados.get("delta_pib_pct_mediana_central", "N/D")
    p10_central = resultados.get("p10_central", "N/D")
    p90_central = resultados.get("p90_central", "N/D")
    delta_variante = resultados.get("delta_pib_pct_mediana_variante", "N/D")
    rec_neta = resultados.get("rec_neta_central", "N/D")

    shock_c = obtener_valor(supuestos, "shock_pp")
    shock_v = obtener_valor(supuestos, "shock_pp_variante")

    texto = (
        f"Este informe estima el impacto sobre el PIB de Uruguay de eliminar la exoneración "
        f"estacional de IVA a servicios gastronómicos pagados por turistas no residentes con "
        f"medios de pago emitidos en el exterior (Decreto 434/023, prorrogado por Decreto 220/2025).\n\n"
        f"Escenario central (+{shock_c} pp de IVA): el efecto estimado sobre el PIB es de "
        f"{delta_central}% (mediana Monte Carlo), con un rango de incertidumbre P10–P90 de "
        f"[{p10_central}%, {p90_central}%]. El impacto es negativo, reflejando la caída en el "
        f"gasto gastronómico de no residentes ante el aumento de precios.\n\n"
        f"Variante (+{shock_v} pp, eliminación total del régimen): impacto estimado de "
        f"{delta_variante}% del PIB.\n\n"
        f"Efecto fiscal neto (escenario central): {rec_neta} millones de USD "
        f"(recaudación ganada menos pérdida por menor actividad). El resultado debe interpretarse "
        f"junto con el Gasto Tributario que el fisco deja de desembolsar.\n\n"
        f"Nota: los resultados se miden en décimas de punto porcentual del PIB. "
        f"La magnitud del shock de IVA es pequeña relativa al tipo de cambio real bilateral "
        f"con Argentina, principal determinante de la demanda turística hacia Uruguay."
    )
    doc.add_paragraph(texto)


def seccion_marco_legal(doc, supuestos):
    agregar_titulo(doc, "2. Marco Legal y Definición del Escenario", 1)
    doc.add_paragraph(
        "El régimen vigente combina dos instrumentos:\n\n"
        "• Ley 17.934 (modificada por Ley 20.212): devolución de 9 puntos de IVA en servicios "
        "gastronómicos a turistas no residentes que abonen con tarjeta emitida en el exterior. "
        "IVA efectivo resultante: 22% − 9% = 13%. Esta es la norma permanente.\n\n"
        "• Decreto 434/023 (prorrogado por Decreto 220/2025): exoneración total de IVA "
        "(tasa 0%) en temporada alta (15 de noviembre al 30 de abril). "
        "El beneficio es adicional al permanente: el turista paga IVA 0% durante la temporada.\n\n"
        "Escenario central: cae únicamente el decreto estacional; queda vigente el piso de 9 puntos "
        "de la Ley 17.934. El precio relativo sube de 100 a 113 (+13 pp).\n\n"
        "Variante: eliminación completa de ambos instrumentos. El precio sube de 100 a 122 (+22 pp).\n\n"
        "El análisis se realiza en equilibrio parcial: no se modela el uso alternativo de la "
        "recaudación ganada ni los efectos de segunda vuelta sobre el tipo de cambio."
    )


def seccion_datos_fuentes(doc):
    agregar_titulo(doc, "3. Datos y Fuentes", 1)
    doc.add_paragraph(
        "La siguiente tabla registra las fuentes de datos del modelo. "
        "Para detalle de fechas de descarga y hashes SHA256, ver DATA_SOURCES.md en el repositorio."
    )
    data_sources = [
        {"ID": "DGI-GT", "Descripción": "Informe Gasto Tributario DGI 2019–2022",
         "URL": "gub.uy/dgi (sección datos y estadísticas)", "Uso": "Base afectada B (vía DGI)"},
        {"ID": "MINTUR-ETR", "Descripción": "Encuesta de Turismo Receptivo (agregados trimestrales y microdatos)",
         "URL": "catalogodatos.gub.uy / gub.uy/ministerio-turismo", "Uso": "Base afectada B (vía ETR); peso w"},
        {"ID": "BCU-MIP", "Descripción": "Matriz Insumo-Producto 2016",
         "URL": "bcu.gub.uy (Estadísticas e Indicadores)", "Uso": "Multiplicadores de Leontief"},
        {"ID": "BCU-CN", "Descripción": "Cuentas Nacionales: VAB y PIB nominal",
         "URL": "bcu.gub.uy (Cuentas Nacionales)", "Uso": "Expresar resultados como % del PIB"},
        {"ID": "BCU-PAGOS", "Descripción": "Reporte Sistema de Pagos Minorista",
         "URL": "bcu.gub.uy (Sistema de Pagos)", "Uso": "Cota superior de consistencia de B"},
    ]
    agregar_tabla_desde_df(doc, pd.DataFrame(data_sources))


def seccion_metodologia(doc, supuestos):
    agregar_titulo(doc, "4. Metodología", 1)
    phi = obtener_valor(supuestos, "phi")
    doc.add_paragraph(
        "El modelo sigue tres etapas: (i) estimación de la base afectada B, "
        "(ii) shock de precio y respuesta de demanda, (iii) propagación input-output.\n"
    )

    agregar_titulo(doc, "4.1 Base afectada (B)", 2)
    doc.add_paragraph(
        "B representa el gasto gastronómico de no residentes pagado con tarjeta del exterior "
        "durante la temporada alta (15-nov al 30-abr). Se estima por triangulación de tres vías:\n\n"
        "1. Vía DGI (preferente): B = GT_estacional / tasa, donde GT_estacional es el Gasto "
        "Tributario reportado por la DGI para el decreto de exoneración estacional.\n\n"
        "2. Vía ETR: gasto en alimentación de no residentes en temporada × share_restaurantes × "
        "share_tarjeta_exterior. La temporada se aproxima como 0.5×T4 + T1 + 0.33×T2.\n\n"
        "3. Vía BCU pagos: total de compras con tarjetas extranjeras como cota superior de consistencia."
    )

    agregar_titulo(doc, "4.2 Shock y respuesta de demanda", 2)
    doc.add_paragraph(
        "Con φ = fracción del shock absorbida por los productores:\n\n"
        "    Δp = (1 − φ) × shock_pp / 100\n"
        "    g_ext = ε_ext × (Δp × w)         [margen extensivo: caída de llegadas]\n"
        "    g_int = ε_int × Δp               [margen intensivo: caída gasto gastronómico condicional]\n"
        "    ΔD_gastro = B × [(1 + g_ext) × (1 + g_int) − 1]\n\n"
        "El margen extensivo (g_ext) afecta todo el gasto turístico. "
        "El margen intensivo (g_int) afecta solo la gastronomía de los turistas que permanecen. "
        "Esto evita el doble conteo."
    )

    agregar_titulo(doc, "4.3 Modelo Insumo-Producto", 2)
    doc.add_paragraph(
        "Se utiliza la Matriz Insumo-Producto 2016 del BCU (modelo de Leontief, tipo I como escenario central):\n\n"
        "    ΔVBP = (I − A)⁻¹ · Δf\n"
        "    ΔVAB = v̂ · ΔVBP\n\n"
        "donde A = matriz de coeficientes técnicos domésticos, Δf = vector de shock de demanda final, "
        "v̂ = diagonal de coeficientes VA/VBP por industria.\n\n"
        "El modelo tipo II (con efecto inducido de hogares) se activa con usar_tipo2: true en "
        "config/supuestos.yaml; se reporta solo como cota superior por tender a sobreestimar."
    )

    agregar_titulo(doc, "4.4 Monte Carlo", 2)
    doc.add_paragraph(
        "Se simula la incertidumbre sobre los parámetros menos conocidos: B, φ, ε_int, ε_ext, w. "
        "Cada parámetro sigue una distribución triangular con los rangos declarados en supuestos.yaml. "
        f"Se generan {obtener_valor(supuestos, 'n_sims'):,} iteraciones con semilla fija = "
        f"{obtener_valor(supuestos, 'seed')}. "
        "Se reportan mediana, media, P10 y P90 del impacto en % del PIB."
    )


def seccion_supuestos(doc, supuestos):
    agregar_titulo(doc, "5. Supuestos del Modelo", 1)
    doc.add_paragraph(
        "Todos los parámetros del modelo se centralizan en config/supuestos.yaml. "
        "La tabla a continuación se genera automáticamente desde ese archivo. "
        "Los parámetros marcados como 'observado' provienen de datos; los marcados como 'supuesto' "
        "son valores asumidos con justificación en la literatura o por diseño del escenario."
    )

    filas = []
    claves_mostrar = [
        "shock_pp", "shock_pp_variante", "phi", "epsilon_intensivo", "epsilon_extensivo",
        "w", "w_fallback", "share_restaurantes", "share_tarjeta_exterior",
        "presion_tributaria_efectiva", "n_sims", "seed", "usar_tipo2",
    ]
    for clave in claves_mostrar:
        if clave not in supuestos:
            continue
        entrada = supuestos[clave]
        if isinstance(entrada, dict):
            rango = entrada.get("rango")
            rango_str = f"[{rango[0]}, {rango[1]}]" if rango else "—"
            filas.append({
                "Parámetro": clave,
                "Valor central": str(entrada.get("valor", "N/D")),
                "Rango MC": rango_str,
                "Fuente": str(entrada.get("fuente", ""))[:60],
                "Justificación": str(entrada.get("justificacion", ""))[:100],
            })
        else:
            filas.append({"Parámetro": clave, "Valor central": str(entrada),
                          "Rango MC": "—", "Fuente": "config", "Justificación": ""})

    agregar_tabla_desde_df(doc, pd.DataFrame(filas))


def seccion_resultados(doc, resultados, supuestos):
    agregar_titulo(doc, "6. Resultados", 1)

    shock_c = obtener_valor(supuestos, "shock_pp")
    shock_v = obtener_valor(supuestos, "shock_pp_variante")

    agregar_titulo(doc, "6.1 Impacto en el PIB", 2)
    doc.add_paragraph(
        "Los resultados se expresan en variación porcentual del PIB. "
        "El impacto es negativo en todos los escenarios: la eliminación de la exoneración "
        "reduce la demanda turística gastronómica, con efectos directos en hostelería y "
        "comida e indirectos en el resto de la cadena de valor.\n"
    )

    if "tabla_mip" in resultados:
        agregar_tabla_desde_df(
            doc, resultados["tabla_mip"],
            titulo=f"Impacto en PIB — Leontief Tipo I (escenario central +{shock_c} pp y variante +{shock_v} pp)"
        )

    agregar_titulo(doc, "6.2 Efecto fiscal neto", 2)
    doc.add_paragraph(
        "Recaudación ganada: turistas que permanecen ahora pagan la tasa de IVA plena. "
        "Recaudación perdida: menor actividad económica implica menor base imponible general.\n"
    )
    if "tabla_fiscal" in resultados:
        agregar_tabla_desde_df(doc, resultados["tabla_fiscal"])

    doc.add_paragraph(
        "Nota metodológica: la recaudación ganada y el impacto en PIB son métricas diferentes "
        "y complementarias del mismo trade-off. No deben sumarse. El PIB mide actividad real; "
        "la recaudación mide ingresos fiscales."
    )


def seccion_sensibilidad(doc):
    agregar_titulo(doc, "7. Análisis de Sensibilidad", 1)
    doc.add_paragraph(
        "El análisis de sensibilidad comprende dos componentes: "
        "(a) la distribución de Monte Carlo del impacto en PIB%, y "
        "(b) el gráfico tornado de sensibilidad one-at-a-time (OAT)."
    )

    agregar_titulo(doc, "7.1 Distribución Monte Carlo", 2)
    agregar_imagen(
        doc,
        OUTPUTS_FIGURAS / "histograma_montecarlo.png",
        caption="Figura 1. Distribución del impacto sobre el PIB (%) — Simulación Monte Carlo."
    )

    agregar_titulo(doc, "7.2 Análisis tornado (sensibilidad OAT)", 2)
    doc.add_paragraph("Escenario central (+13 pp):")
    agregar_imagen(
        doc,
        OUTPUTS_FIGURAS / "tornado_central.png",
        caption="Figura 2. Análisis tornado — Escenario central (+13 pp)."
    )
    doc.add_paragraph("Variante (+22 pp):")
    agregar_imagen(
        doc,
        OUTPUTS_FIGURAS / "tornado_variante.png",
        caption="Figura 3. Análisis tornado — Variante (+22 pp)."
    )


def seccion_limitaciones(doc):
    agregar_titulo(doc, "8. Limitaciones del Modelo", 1)
    limitaciones = [
        "Equilibrio parcial sin modelizar el uso alternativo de la recaudación ganada: "
        "si el fisco reasigna los ingresos fiscales a gasto, el impacto neto sobre el PIB sería menor.",

        "La MIP utilizada corresponde a 2016. La estructura de la economía uruguaya puede haber "
        "cambiado significativamente desde entonces (composición sectorial, integración vertical).",

        "Posible migración al pago en efectivo: la exoneración aplica solo a tarjetas del exterior. "
        "Si los turistas sustituyen hacia efectivo (sin trazabilidad fiscal), la recaudación formal "
        "disminuye sin que la actividad real se vea afectada.",

        "La demanda turística hacia Uruguay está dominada por el tipo de cambio real bilateral con "
        "Argentina. El efecto del IVA es de segundo orden frente a un salto cambiario significativo.",

        "Las elasticidades (ε_int, ε_ext) se toman de la literatura internacional y no han sido "
        "estimadas econométricamente para Uruguay. Este es el supuesto de mayor incertidumbre "
        "del modelo (ver fase 2 — event-study).",

        "El rubro 'alimentación' de la ETR puede incluir compras en supermercados y almacenes, "
        "que no califican para la exoneración. El parámetro share_restaurantes corrige parcialmente "
        "este sesgo pero introduce incertidumbre adicional.",

        "El modelo de Leontief asume coeficientes técnicos fijos y rendimientos constantes a escala. "
        "No captura ajustes de precios relativos, sustitución entre insumos, ni cambios en la frontera "
        "tecnológica.",

        "El modelo tipo I (abierto) excluye el efecto inducido del consumo de los hogares. "
        "El tipo II (cerrado respecto a hogares) tiende a sobreestimar el multiplicador.",
    ]
    for lim in limitaciones:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(lim)


def seccion_reproducibilidad(doc):
    agregar_titulo(doc, "9. Anexo de Reproducibilidad", 1)
    doc.add_paragraph(
        "Para regenerar todos los resultados desde cero en una máquina limpia:\n"
    )
    comandos = (
        "# 1. Clonar el repositorio\n"
        "git clone <URL_DEL_REPO>\n"
        "cd iva-turismo-uy\n\n"
        "# 2. Crear entorno virtual e instalar dependencias\n"
        "python -m venv venv\n"
        "source venv/bin/activate      # Linux/Mac\n"
        "# venv\\Scripts\\activate      # Windows\n"
        "pip install -r requirements.txt\n\n"
        "# 3. Colocar los archivos manuales en data/manual/\n"
        "#    (ver instrucciones en DATA_SOURCES.md y README.md)\n\n"
        "# 4. Ejecutar el pipeline completo\n"
        "python run_all.py\n\n"
        "# 5. El informe se genera en:\n"
        "#    outputs/informe_final.docx\n"
        "#    outputs/tablas/\n"
        "#    outputs/figuras/\n"
    )
    doc.add_paragraph(comandos)
    doc.add_paragraph(
        f"Semilla aleatoria fija: seed = 42. "
        f"Todos los resultados son exactamente reproducibles con este seed y las mismas versiones "
        f"de paquetes listadas en requirements.txt.\n\n"
        f"Fecha de generación de este informe: {datetime.date.today().isoformat()}"
    )


# =============================================================================
# PIPELINE PRINCIPAL
# =============================================================================
def main():
    separador("ETAPA 07 — Generación del informe final")

    if not DOCX_DISPONIBLE:
        print("  ✗ python-docx no disponible. Instalar: pip install python-docx")
        sys.exit(1)

    supuestos = cargar_supuestos()

    # Cargar resultados disponibles
    resultados = {}

    ruta_mc = OUTPUTS_TABLAS / "tabla_montecarlo_estadisticas.csv"
    if ruta_mc.exists():
        df_mc = pd.read_csv(ruta_mc)
        for _, fila in df_mc.iterrows():
            if "central" in str(fila.get("label", "")):
                resultados["delta_pib_pct_mediana_central"] = fila.get("mediana")
                resultados["p10_central"] = fila.get("p10")
                resultados["p90_central"] = fila.get("p90")
            elif "variante" in str(fila.get("label", "")):
                resultados["delta_pib_pct_mediana_variante"] = fila.get("mediana")

    ruta_mip = OUTPUTS_TABLAS / "tabla_impacto_mip.csv"
    if ruta_mip.exists():
        resultados["tabla_mip"] = pd.read_csv(ruta_mip)

    ruta_fiscal = OUTPUTS_TABLAS / "tabla_fiscal.csv"
    if ruta_fiscal.exists():
        df_fis = pd.read_csv(ruta_fiscal)
        resultados["tabla_fiscal"] = df_fis
        fila_c = df_fis[df_fis["label"] == "central_13pp"]
        if not fila_c.empty:
            resultados["rec_neta_central"] = round(fila_c["rec_neta_musd"].iloc[0], 1)

    # Crear documento
    doc = Document()

    # Portada
    doc.add_heading(
        "Impacto sobre el PIB de eliminar la exoneración estacional de IVA "
        "a la gastronomía turística — Uruguay",
        0
    )
    doc.add_paragraph(
        f"Fecha: {datetime.date.today().isoformat()}\n"
        "Marco legal: Ley 17.934 (mod. Ley 20.212) + Decreto 434/023 prorrogado por Decreto 220/2025\n"
        "Metodología: Equilibrio parcial + Modelo Insumo-Producto (MIP BCU 2016) + Monte Carlo\n"
        "Repositorio: iva-turismo-uy (ver Anexo de Reproducibilidad)"
    )
    doc.add_page_break()

    # Secciones
    seccion_resumen_ejecutivo(doc, supuestos, resultados)
    doc.add_page_break()
    seccion_marco_legal(doc, supuestos)
    seccion_datos_fuentes(doc)
    seccion_metodologia(doc, supuestos)
    seccion_supuestos(doc, supuestos)
    seccion_resultados(doc, resultados, supuestos)
    seccion_sensibilidad(doc)
    seccion_limitaciones(doc)
    seccion_reproducibilidad(doc)

    # Guardar
    ruta_salida = ROOT / "outputs" / "informe_final.docx"
    ruta_salida.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(ruta_salida))
    print(f"\n  ✓ Informe guardado: {ruta_salida}")

    return ruta_salida


if __name__ == "__main__":
    main()
