"""
Etapa 02 — Estimación de la base afectada (B).

Triangulación por tres vías:
  1. Vía DGI:   B = GT_estacional / tasa  (preferente)
  2. Vía ETR:   gasto alimentación × share_restaurantes × share_tarjeta_exterior
  3. Vía BCU pagos: cota superior de consistencia (si disponible)

Salida: data/processed/base_afectada.csv con las tres estimaciones y la B central.
Actualiza supuestos.yaml con B_dgi, B_etr, B_bcu.
"""
import sys
import pathlib
import warnings
import pandas as pd
import numpy as np
import yaml

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS,
    cargar_supuestos, obtener_valor, buscar_archivo, separador, checkpoint
)

# Importar pdfplumber con manejo de error si no está instalado
try:
    import pdfplumber
    PDF_DISPONIBLE = True
except ImportError:
    PDF_DISPONIBLE = False
    print("  ⚠ pdfplumber no disponible — no se podrá extraer texto del PDF de DGI.")


# =============================================================================
# VÍA 1 — DGI: Gasto Tributario
# =============================================================================
def extraer_gt_dgi() -> dict:
    """
    Intenta extraer el Gasto Tributario de la exoneración de IVA gastronómico
    del informe PDF de la DGI.

    Devuelve un dict con:
      - gt_estacional_musd: Gasto Tributario estacional en millones de USD
      - gt_total_musd: GT total del régimen si el informe los distingue
      - tasa_usada: tasa IVA efectiva usada para despejar B
      - interpretacion: descripción de qué cifra se tomó del PDF
      - fuente_ok: bool
    """
    resultado = {
        "gt_estacional_musd": None,
        "gt_total_musd": None,
        "tasa_usada": None,
        "interpretacion": "",
        "fuente_ok": False,
    }

    ruta = buscar_archivo("dgi_gasto_tributario_2019_2022.pdf")
    if ruta is None:
        resultado["interpretacion"] = (
            "FALTA ARCHIVO: data/manual/dgi_gasto_tributario_2019_2022.pdf\n"
            "  Descargar desde: https://www.gub.uy/direccion-general-impositiva/datos-y-estadisticas"
        )
        return resultado

    if not PDF_DISPONIBLE:
        resultado["interpretacion"] = "pdfplumber no instalado — instalar con: pip install pdfplumber"
        return resultado

    print(f"  → Leyendo PDF DGI: {ruta.name}")
    texto_completo = []
    with pdfplumber.open(ruta) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                texto_completo.append(texto)

    texto = "\n".join(texto_completo)

    # Buscar secciones relacionadas con Ley 17.934, IVA gastronomía, turismo
    palabras_clave = [
        "17.934", "gastronomía", "gastronomia", "turismo",
        "restaurante", "alimentación", "alimentacion",
        "exoneración estacional", "decreto 434", "220/2025"
    ]
    lineas_relevantes = []
    for linea in texto.split("\n"):
        linea_lower = linea.lower()
        if any(kw.lower() in linea_lower for kw in palabras_clave):
            lineas_relevantes.append(linea.strip())

    if not lineas_relevantes:
        resultado["interpretacion"] = (
            "No se encontraron líneas con palabras clave de gastronomía/turismo en el PDF.\n"
            "  Puede que el PDF esté escaneado (imagen) o que las cifras estén en otra sección.\n"
            "  Revisar manualmente el informe y completar B_dgi en config/supuestos.yaml."
        )
        return resultado

    print(f"  → Líneas relevantes encontradas: {len(lineas_relevantes)}")
    for l in lineas_relevantes[:15]:
        print(f"     {l}")

    # Intentar extraer números de esas líneas
    import re
    numeros_extraidos = []
    for linea in lineas_relevantes:
        # Buscar patrones de números con punto de miles y coma decimal (estilo UY)
        matches = re.findall(r"\b(\d{1,3}(?:\.\d{3})*(?:,\d+)?)\b", linea)
        for m in matches:
            try:
                val = float(m.replace(".", "").replace(",", "."))
                if 1 < val < 500:  # rango razonable para GT en millones USD
                    numeros_extraidos.append((val, linea[:80]))
            except ValueError:
                pass

    if numeros_extraidos:
        resultado["interpretacion"] = (
            f"Extracción automática: {len(numeros_extraidos)} valores en rango [1, 500] "
            f"encontrados en líneas relevantes.\n"
            "  ACCIÓN REQUERIDA: verificar cuál corresponde al GT estacional "
            "y actualizar B_dgi en config/supuestos.yaml con la fuente exacta (página, tabla)."
        )
        print(f"\n  ⚠ REVISIÓN MANUAL NECESARIA — valores encontrados en líneas relevantes:")
        for val, linea in numeros_extraidos[:10]:
            print(f"     {val:>10.1f}  ←  '{linea}'")
    else:
        resultado["interpretacion"] = (
            "No se pudieron extraer valores numéricos en rango razonable. "
            "Revisar el PDF manualmente."
        )

    resultado["fuente_ok"] = False  # siempre requiere confirmación manual
    return resultado


# =============================================================================
# VÍA 2 — ETR: Encuesta de Turismo Receptivo
# =============================================================================
def calcular_b_etr(supuestos: dict) -> dict:
    """
    Calcula B desde los datos de la ETR.

    Aproximación de temporada alta: 0.5×T4 + T1 + 0.33×T2
    (15-nov a 30-abr: medio T4 + T1 completo + un tercio de T2)
    """
    resultado = {
        "gasto_alimentacion_temporada_musd": None,
        "b_etr_musd": None,
        "w_calculado": None,
        "fuente_ok": False,
        "notas": [],
    }

    share_restaurantes = obtener_valor(supuestos, "share_restaurantes")
    share_tarjeta = obtener_valor(supuestos, "share_tarjeta_exterior")

    # Buscar archivo ETR
    ruta = buscar_archivo(
        "mintur_etr_agregados.csv",
        ["mintur_etr_agregados.xlsx", "etr_agregados", "turismo_receptivo"]
    )
    if ruta is None:
        resultado["notas"].append(
            "FALTA ARCHIVO ETR.\n"
            "  Descargar desde: https://catalogodatos.gub.uy/dataset/"
            "ministerio-de-turismo-turismo-receptivo\n"
            "  Guardar como: data/manual/mintur_etr_agregados.csv (o .xlsx)"
        )
        return resultado

    print(f"  → Leyendo ETR: {ruta.name}")
    try:
        if ruta.suffix.lower() in (".xlsx", ".xls"):
            df = pd.read_excel(ruta, sheet_name=None)
            # Intentar encontrar la hoja correcta
            hojas = list(df.keys())
            print(f"     Hojas encontradas: {hojas}")
            # Tomar la primera hoja con más de 10 filas
            for nombre_hoja, hoja in df.items():
                if len(hoja) > 10:
                    df = hoja
                    print(f"     Usando hoja: '{nombre_hoja}'")
                    break
        else:
            # Intentar varias codificaciones comunes en archivos UY
            for enc in ("utf-8", "latin-1", "cp1252"):
                try:
                    df = pd.read_csv(ruta, encoding=enc)
                    break
                except UnicodeDecodeError:
                    continue
    except Exception as e:
        resultado["notas"].append(f"Error leyendo ETR: {e}")
        return resultado

    print(f"     Columnas: {list(df.columns)}")
    print(f"     Filas: {len(df)}")

    # Normalizar nombres de columna
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Identificar columnas de trimestre y rubro
    col_trimestre = next((c for c in df.columns if "trim" in c or "quarter" in c or "t" == c), None)
    col_rubro = next(
        (c for c in df.columns if "rubro" in c or "concepto" in c or "categoria" in c or "item" in c),
        None
    )
    col_gasto = next(
        (c for c in df.columns if "gasto" in c or "monto" in c or "valor" in c or "usd" in c),
        None
    )

    print(f"     Columna trimestre detectada: {col_trimestre}")
    print(f"     Columna rubro detectada:     {col_rubro}")
    print(f"     Columna gasto detectada:     {col_gasto}")

    if not all([col_trimestre, col_rubro, col_gasto]):
        resultado["notas"].append(
            "No se pudo mapear automáticamente la estructura del archivo ETR.\n"
            f"  Columnas disponibles: {list(df.columns)}\n"
            "  ACCIÓN: indicar los nombres de columna exactos en el análisis."
        )
        return resultado

    # Filtrar rubros de alimentación
    df_alim = df[df[col_rubro].str.lower().str.contains(
        "aliment|comida|gastron|restaur", na=False, regex=True
    )].copy()

    if df_alim.empty:
        resultado["notas"].append(
            f"No se encontraron filas con 'alimentación' en la columna '{col_rubro}'.\n"
            f"  Valores únicos de rubro: {df[col_rubro].unique()[:20]}"
        )
        return resultado

    # Extraer gasto por trimestre (T1, T2, T4)
    gasto_por_trimestre = {}
    for trim_nombre, trim_codigo in [("T1", 1), ("T2", 2), ("T4", 4)]:
        mask = df_alim[col_trimestre].astype(str).str.contains(str(trim_codigo), na=False)
        gasto_por_trimestre[trim_nombre] = df_alim.loc[mask, col_gasto].sum()

    print(f"     Gasto alimentación por trimestre (en unidad original):")
    for t, v in gasto_por_trimestre.items():
        print(f"       {t}: {v:,.1f}")

    # Aproximar temporada alta: 0.5×T4 + T1 + 0.33×T2
    gasto_temporada = (
        0.5 * gasto_por_trimestre.get("T4", 0)
        + gasto_por_trimestre.get("T1", 0)
        + 1/3 * gasto_por_trimestre.get("T2", 0)
    )

    resultado["notas"].append(
        "Aproximación temporada alta: 0.5×T4 + T1 + 0.33×T2 "
        "(15-nov a 30-abr como fracción de trimestres calendario)."
    )

    # Calcular w = gasto alimentación / gasto total en la misma aproximación
    df_total = df.groupby(col_trimestre)[col_gasto].sum()
    gasto_total_temporada = (
        0.5 * df_total.get(1, df_total.get("T4", 0))
        + df_total.get(1, df_total.get("T1", 0))
        + 1/3 * df_total.get(2, df_total.get("T2", 0))
    )

    if gasto_total_temporada > 0:
        w_calculado = gasto_temporada / gasto_total_temporada
        resultado["w_calculado"] = float(w_calculado)
        print(f"     w calculado (alimentación/total) = {w_calculado:.3f}")

    # Aplicar shares para llegar a B
    b_etr = gasto_temporada * share_restaurantes * share_tarjeta

    resultado["gasto_alimentacion_temporada_musd"] = float(gasto_temporada)
    resultado["b_etr_musd"] = float(b_etr)
    resultado["fuente_ok"] = True
    resultado["notas"].append(
        f"B_etr = {gasto_temporada:.1f} × {share_restaurantes} (share_restaurantes) "
        f"× {share_tarjeta} (share_tarjeta_exterior) = {b_etr:.1f}"
    )
    return resultado


# =============================================================================
# TRIANGULACIÓN y VALIDACIÓN
# =============================================================================
def triangular_y_validar(b_dgi, b_etr, b_bcu) -> dict:
    """
    Construye la tabla comparativa y elige B central.
    Valida que las fuentes estén dentro del mismo orden de magnitud.
    """
    fuentes = {
        "DGI (preferente)": b_dgi,
        "ETR": b_etr,
        "BCU Pagos (cota sup.)": b_bcu,
    }
    disponibles = {k: v for k, v in fuentes.items() if v is not None}

    print("\n  TABLA COMPARATIVA DE ESTIMACIONES DE B (millones USD):")
    print(f"  {'Vía':<30} {'B (MUSD)':>12}")
    print("  " + "-" * 44)
    for k, v in fuentes.items():
        val_str = f"{v:>10.1f}" if v is not None else "    N/D     "
        print(f"  {k:<30} {val_str}")

    if len(disponibles) == 0:
        return {"b_central": None, "b_min": None, "b_max": None, "alerta": "Sin estimaciones disponibles."}

    # Validación: diferencia mayor a 3× entre fuentes
    valores = list(disponibles.values())
    if len(valores) >= 2:
        razon = max(valores) / min(valores)
        if razon > 3.0:
            msg = (
                f"ALERTA: Las estimaciones de B difieren en un factor de {razon:.1f}x "
                f"(>3× → revisar). Mínimo: {min(valores):.1f}, Máximo: {max(valores):.1f}."
            )
            print(f"\n  ⚠ {msg}")
            return {
                "b_central": None, "b_min": min(valores), "b_max": max(valores),
                "alerta": msg
            }

    # Elegir B central (DGI si está disponible, si no ETR)
    if b_dgi is not None:
        b_central = b_dgi
        fuente_central = "DGI"
    elif b_etr is not None:
        b_central = b_etr
        fuente_central = "ETR"
    else:
        b_central = min(disponibles.values())
        fuente_central = "BCU (cota mínima)"

    b_min = min(disponibles.values())
    b_max = max(disponibles.values())

    print(f"\n  B CENTRAL = {b_central:.1f} MUSD (vía {fuente_central})")
    print(f"  Rango para Monte Carlo: [{b_min:.1f}, {b_max:.1f}] MUSD")

    return {
        "b_central": b_central,
        "b_min": b_min,
        "b_max": b_max,
        "fuente_central": fuente_central,
        "alerta": None,
    }


def actualizar_supuestos_con_b(resultado: dict):
    """Escribe B_dgi, B_etr, B_bcu en supuestos.yaml."""
    from utils import CONFIG_PATH
    with open(CONFIG_PATH, encoding="utf-8") as f:
        supuestos = yaml.safe_load(f)

    if resultado.get("b_central"):
        for clave, val in [
            ("B_dgi", resultado.get("b_dgi")),
            ("B_etr", resultado.get("b_etr")),
            ("B_bcu", resultado.get("b_bcu")),
        ]:
            if clave in supuestos and val is not None:
                supuestos[clave]["valor"] = round(val, 2)

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(supuestos, f, allow_unicode=True, sort_keys=False, default_flow_style=False)

    print("  ✓ supuestos.yaml actualizado con estimaciones de B.")


def main():
    separador("ETAPA 02 — Base afectada (B)")

    supuestos = cargar_supuestos()

    # --- VÍA 1: DGI ---
    print("\n[VÍA 1] DGI — Gasto Tributario")
    res_dgi = extraer_gt_dgi()
    print(f"  → {res_dgi['interpretacion'][:200]}")
    b_dgi = res_dgi.get("gt_estacional_musd")

    # Verificar si hay valor manual en supuestos.yaml
    b_dgi_manual = obtener_valor(supuestos, "B_dgi")
    if b_dgi_manual is not None:
        b_dgi = b_dgi_manual
        print(f"  → Usando valor manual de supuestos.yaml: B_dgi = {b_dgi} MUSD")

    # --- VÍA 2: ETR ---
    print("\n[VÍA 2] MINTUR — ETR")
    res_etr = calcular_b_etr(supuestos)
    for nota in res_etr["notas"]:
        print(f"  → {nota}")
    b_etr = res_etr.get("b_etr_musd")

    b_etr_manual = obtener_valor(supuestos, "B_etr")
    if b_etr_manual is not None and b_etr is None:
        b_etr = b_etr_manual
        print(f"  → Usando valor manual de supuestos.yaml: B_etr = {b_etr} MUSD")

    # Actualizar w en supuestos si se calculó
    if res_etr.get("w_calculado") is not None:
        print(f"  → w calculado de la ETR: {res_etr['w_calculado']:.3f} (actualizar supuestos.yaml)")

    # --- VÍA 3: BCU Pagos ---
    print("\n[VÍA 3] BCU — Sistema de Pagos")
    b_bcu_manual = obtener_valor(supuestos, "B_bcu")
    if b_bcu_manual is not None:
        print(f"  → Valor manual de supuestos.yaml: B_bcu = {b_bcu_manual} MUSD")
    else:
        print("  → No disponible (archivo BCU-PAGOS no descargado).")
    b_bcu = b_bcu_manual

    # --- TRIANGULACIÓN ---
    print()
    resultado = triangular_y_validar(b_dgi, b_etr, b_bcu)
    resultado.update({"b_dgi": b_dgi, "b_etr": b_etr, "b_bcu": b_bcu})

    if resultado.get("alerta") and resultado["b_central"] is None:
        print(f"\n  ✗ PIPELINE DETENIDO: {resultado['alerta']}")
        sys.exit(1)

    # Guardar resultado procesado
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)

    df_b = pd.DataFrame([
        {"via": "DGI (preferente)", "b_musd": b_dgi, "estado": "observado" if b_dgi else "no disponible"},
        {"via": "ETR", "b_musd": b_etr, "estado": "calculado" if b_etr else "no disponible"},
        {"via": "BCU Pagos (cota sup.)", "b_musd": b_bcu, "estado": "observado" if b_bcu else "no disponible"},
        {"via": "CENTRAL (elegida)", "b_musd": resultado.get("b_central"),
         "estado": f"vía {resultado.get('fuente_central', 'N/D')}"},
    ])

    df_b.to_csv(DATA_PROCESSED / "base_afectada.csv", index=False, encoding="utf-8")
    df_b.to_csv(OUTPUTS_TABLAS / "tabla_base_afectada.csv", index=False, encoding="utf-8")
    print(f"\n  ✓ Guardado: data/processed/base_afectada.csv")

    checkpoint(
        "Etapa 02 — Base Afectada",
        [
            f"B vía DGI: {b_dgi if b_dgi else 'NO DISPONIBLE'} MUSD",
            f"B vía ETR: {b_etr if b_etr else 'NO DISPONIBLE'} MUSD",
            f"B vía BCU: {b_bcu if b_bcu else 'NO DISPONIBLE'} MUSD",
            f"B CENTRAL: {resultado.get('b_central', 'NO DEFINIDA')} MUSD",
            "Validar: (1) ¿B central parece razonable? (2) ¿Se puede confirmar la "
            "interpretación del PDF DGI? (3) ¿El archivo ETR tiene la estructura esperada?",
        ]
    )

    return resultado


if __name__ == "__main__":
    main()
