"""
Etapa 04 — Modelo Insumo-Producto de Leontief (MIP 2016 BCU).

Pasos:
  1. Leer la MIP 2016 del BCU (Excel).
  2. Construir la matriz de coeficientes técnicos domésticos A.
  3. Calcular el inverso de Leontief: L = (I − A)⁻¹
  4. Construir el vector de shock Δf desde los resultados de Etapa 03.
  5. ΔVBP = L × Δf
  6. ΔVAB = v̂ × ΔVBP  (v̂ = diagonal de coeficientes VA/VBP)
  7. Descomponer efecto directo vs. indirecto.
  8. Validar multiplicadores y orden de magnitud del impacto.

NOTA SOBRE UNIDADES:
  - La MIP 2016 está en millones de pesos uruguayos (UYU) de 2016.
  - El shock Δf viene en millones de USD corrientes del año_base (2022).
  - Conversión: Δf_uyu2016 = Δf_usd2022 × TCR_promedio_2022 × deflactor_2016→2022
  - El procedimiento de conversión se documenta explícitamente en el código.
"""
import sys
import pathlib
import warnings
import numpy as np
import pandas as pd

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS,
    cargar_supuestos, obtener_valor, buscar_archivo, separador, checkpoint
)


# =============================================================================
# LECTURA Y CONSTRUCCIÓN DE LA MIP
# =============================================================================
def leer_mip_bcu(ruta: pathlib.Path) -> dict:
    """
    Lee el Excel de la MIP 2016 del BCU.

    El BCU publica la MIP en un Excel con varias hojas. Las hojas típicas son:
      - "Doméstica" o "Z_dom": transacciones intermedias domésticas
      - "VBP": Vector de Producción Bruta por industria
      - "VA" o "VAB": Valor Agregado Bruto por industria
      - "Coeficientes" o "A": matriz de coeficientes técnicos (si está precalculada)

    Esta función es flexible: intenta detectar la estructura automáticamente
    y devuelve un dict con las matrices necesarias.
    """
    print(f"  → Leyendo MIP: {ruta.name}")
    xls = pd.ExcelFile(ruta)
    hojas = xls.sheet_names
    print(f"     Hojas disponibles: {hojas}")

    resultado = {
        "Z_dom": None,    # transacciones intermedias domésticas (n×n)
        "vbp": None,      # vector VBP (n,)
        "vab": None,      # vector VAB (n,)
        "A": None,        # coeficientes técnicos (si ya están calculados)
        "industrias": None,  # nombres de las n industrias
        "n": None,
        "hojas_usadas": [],
    }

    # Mapeo heurístico de nombres de hojas
    mapa_hojas = {
        "Z_dom": ["doméstica", "domestica", "z_dom", "intermedia", "transacciones"],
        "vbp": ["vbp", "produccion", "producción", "output"],
        "vab": ["vab", "valor agregado", "va_", "valor_agregado"],
        "A": ["coeficientes", "coef", "matrix a", "a_"],
    }

    hojas_lower = {h.lower(): h for h in hojas}

    for clave, posibles in mapa_hojas.items():
        for posible in posibles:
            for hoja_lower, hoja_real in hojas_lower.items():
                if posible in hoja_lower:
                    try:
                        df = pd.read_excel(ruta, sheet_name=hoja_real, index_col=0)
                        resultado[clave] = df
                        resultado["hojas_usadas"].append(f"{clave} ← '{hoja_real}'")
                        print(f"     {clave}: hoja '{hoja_real}' ({df.shape})")
                        break
                    except Exception as e:
                        print(f"     ⚠ Error leyendo hoja '{hoja_real}': {e}")
            if resultado[clave] is not None:
                break

    # Si no se encontró nada por nombre, intentar la primera hoja cuadrada
    if resultado["Z_dom"] is None:
        print("  ⚠ No se detectó la hoja de transacciones por nombre. Intentando primera hoja cuadrada…")
        for hoja in hojas:
            try:
                df = pd.read_excel(ruta, sheet_name=hoja, index_col=0)
                n = min(df.shape)
                # Una hoja cuadrada con al menos 10 industrias probablemente es Z
                if df.shape[0] == df.shape[1] and n >= 10:
                    resultado["Z_dom"] = df
                    resultado["hojas_usadas"].append(f"Z_dom ← '{hoja}' (detección automática)")
                    print(f"     Z_dom detectada en hoja '{hoja}' ({df.shape})")
                    break
            except Exception:
                continue

    if resultado["Z_dom"] is None:
        raise ValueError(
            "No se pudo identificar la hoja de transacciones intermedias domésticas.\n"
            f"  Hojas disponibles: {hojas}\n"
            "  Indicar manualmente el nombre de la hoja en 04_modelo_mip.py."
        )

    # Inferir industrias del índice de Z_dom
    resultado["industrias"] = list(resultado["Z_dom"].index)
    resultado["n"] = len(resultado["industrias"])
    print(f"     Industrias detectadas: {resultado['n']}")

    return resultado


def construir_A(mip: dict) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Construye la matriz de coeficientes técnicos A, el vector VBP y el vector VA/VBP.

    Devuelve (A, vbp, va_ratio) como arrays numpy.
    """
    Z = mip["Z_dom"].values.astype(float)
    n = Z.shape[0]

    if mip["vbp"] is not None:
        vbp = mip["vbp"].values.flatten().astype(float)
        if len(vbp) != n:
            vbp = vbp[:n]
    else:
        # VBP = suma de columna de Z + demanda final (si no hay hoja separada)
        # Approximación: VBP = suma de toda la fila
        print("  ⚠ Vector VBP no disponible. Estimando como suma de filas de Z.")
        vbp = Z.sum(axis=1)

    # A[i,j] = Z[i,j] / VBP[j]
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        A = Z / vbp[np.newaxis, :]  # broadcast: cada columna j se divide por VBP_j
    A = np.nan_to_num(A, nan=0.0)

    if mip["vab"] is not None:
        vab = mip["vab"].values.flatten().astype(float)
        if len(vab) != n:
            vab = vab[:n]
        va_ratio = np.where(vbp > 0, vab / vbp, 0.0)
    else:
        print("  ⚠ Vector VAB no disponible. Estimando VA/VBP = 1 − sum(columnas de A).")
        va_ratio = np.clip(1 - A.sum(axis=0), 0, 1)

    return A, vbp, va_ratio


def validar_A(A: np.ndarray, industrias: list):
    """
    Validaciones sobre la matriz A:
    - Columnas suman ≤ 1 (ley de conservación: insumos ≤ producción)
    - Sin valores negativos
    """
    sumas_col = A.sum(axis=0)
    if np.any(sumas_col > 1.01):
        industrias_exceden = [industrias[i] for i in np.where(sumas_col > 1.01)[0]]
        print(f"  ⚠ Columnas de A > 1 en: {industrias_exceden[:5]}")
    if np.any(A < -0.001):
        print("  ⚠ Valores negativos en A — revisar lectura del Excel.")


def calcular_leontief(A: np.ndarray) -> np.ndarray:
    """Calcula L = (I − A)^{−1}."""
    n = A.shape[0]
    I = np.eye(n)
    L = np.linalg.inv(I - A)
    return L


def validar_multiplicadores(L: np.ndarray, industrias: list, idx_gastro: int):
    """
    Valida que el multiplicador de producto tipo I de la industria gastronómica
    esté entre 1.2 y 2.0. Fuera de ese rango → revisar construcción de A.
    """
    multiplicadores = L.sum(axis=0)  # suma de columna = multiplicador de producción tipo I
    mult_gastro = multiplicadores[idx_gastro]
    print(f"  Multiplicador tipo I de gastronomía: {mult_gastro:.3f}")
    assert 1.2 <= mult_gastro <= 2.0, (
        f"ALERTA: multiplicador de gastronomía = {mult_gastro:.3f} fuera de [1.2, 2.0]. "
        "Revisar la construcción de la matriz A."
    )
    return multiplicadores


# =============================================================================
# CONVERSIÓN DE UNIDADES Δf: USD corrientes 2022 → UYU constantes 2016
# =============================================================================
def convertir_shock_a_uyu_2016(
    delta_d_musd_2022: float,
    supuestos: dict,
) -> float:
    """
    Convierte el shock en MUSD corrientes del año_base a MUYU constantes 2016.

    Procedimiento documentado:
      1. Multiplicar por el TC promedio 2022 (UYU/USD) → MUYU corrientes 2022
      2. Deflactar por el ratio IPC 2016 / IPC 2022 → MUYU constantes 2016

    Si los parámetros de conversión no están en supuestos.yaml, se detiene.
    """
    tc_2022 = obtener_valor(supuestos, "tc_promedio_2022")
    deflactor = obtener_valor(supuestos, "deflactor_2016_2022")

    if tc_2022 is None or deflactor is None:
        print("  ⚠ Faltan parámetros de conversión de unidades en supuestos.yaml:")
        print("    tc_promedio_2022: tipo de cambio promedio 2022 (UYU/USD)")
        print("    deflactor_2016_2022: IPC_2016 / IPC_2022 (para expresar en pesos 2016)")
        print()
        print("  DATOS REQUERIDOS (agregar a config/supuestos.yaml):")
        print("    tc_promedio_2022:")
        print("      valor: <TC promedio 2022 — fuente BCU>")
        print("      rango: null")
        print("      fuente: 'BCU — Tipo de cambio promedio anual 2022'")
        print("      justificacion: 'Conversión USD→UYU para alinear unidades con la MIP'")
        print()
        print("    deflactor_2016_2022:")
        print("      valor: <IPC_2016/IPC_2022 — fuente INE/BCU>")
        print("      rango: null")
        print("      fuente: 'INE — Índice de Precios al Consumo base 2016'")
        print("      justificacion: 'Deflactar valores nominales 2022 a precios 2016 de la MIP'")
        print()
        print("  Pipeline detenido hasta que se completen estos parámetros.")
        sys.exit(1)

    # Paso 1: USD 2022 → UYU 2022
    delta_muyu_2022 = delta_d_musd_2022 * tc_2022

    # Paso 2: UYU 2022 → UYU 2016 (precios constantes de la MIP)
    delta_muyu_2016 = delta_muyu_2022 * deflactor

    print(f"  Conversión de unidades:")
    print(f"    {delta_d_musd_2022:.2f} MUSD × TC {tc_2022:.2f} = {delta_muyu_2022:.1f} MUYU (corrientes 2022)")
    print(f"    × deflactor {deflactor:.4f} = {delta_muyu_2016:.1f} MUYU (constantes 2016)")

    return delta_muyu_2016


# =============================================================================
# CONSTRUCCIÓN DEL VECTOR Δf
# =============================================================================
def construir_delta_f(
    shock_muyu_2016: float,
    industrias: list,
    mapeo_csv: pathlib.Path,
) -> np.ndarray:
    """
    Construye el vector de shock de demanda final Δf (n,).

    El shock gastronómico se asigna a la(s) industria(s) mapeadas como
    'Servicios de alojamiento y de comida' en el mapeo_rubros_mip.csv.
    """
    n = len(industrias)
    delta_f = np.zeros(n)

    df_mapeo = pd.read_csv(mapeo_csv, encoding="utf-8")
    # Buscar la fila de alimentación/gastronomía
    mask_gastro = df_mapeo["rubro_etr"].str.lower().str.contains("aliment|gastron", regex=True, na=False)
    filas_gastro = df_mapeo[mask_gastro]

    if filas_gastro.empty:
        print("  ⚠ No se encontró mapeo de gastronomía en mapeo_rubros_mip.csv")
        return delta_f

    # Buscar índice en la lista de industrias de la MIP
    industrias_lower = [str(ind).lower() for ind in industrias]
    idx_gastro = None

    for _, fila in filas_gastro.iterrows():
        industria_target = str(fila.get("industria_mip", "")).lower()
        # Intentar coincidencia por nombre o código
        for i, ind in enumerate(industrias_lower):
            if any(kw in ind for kw in ["alojamiento", "comida", "55", "56", "i55"]):
                idx_gastro = i
                break

    if idx_gastro is None:
        # Fallback: buscar por el código MIP en el nombre de la columna
        for i, ind in enumerate(industrias_lower):
            if "restaur" in ind or "gastron" in ind or "aliment" in ind:
                idx_gastro = i
                break

    if idx_gastro is None:
        print("  ⚠ No se pudo mapear la industria gastronómica en las industrias de la MIP.")
        print(f"    Primeras 10 industrias: {industrias[:10]}")
        print("    Ajustar mapeo_rubros_mip.csv o la búsqueda en construir_delta_f().")
        return delta_f

    delta_f[idx_gastro] = shock_muyu_2016
    print(f"  Δf asignado a industria [{idx_gastro}]: '{industrias[idx_gastro]}'")
    print(f"  Valor: {shock_muyu_2016:.1f} MUYU (2016)")

    return delta_f, idx_gastro


# =============================================================================
# CÁLCULO DEL IMPACTO
# =============================================================================
def calcular_impacto(
    L: np.ndarray,
    delta_f: np.ndarray,
    va_ratio: np.ndarray,
    idx_gastro: int,
    industrias: list,
    label: str,
    coef_empleo: np.ndarray | None = None,
) -> dict:
    """
    Calcula ΔVBP, ΔVAB y la descomposición directo/indirecto.
    """
    delta_vbp = L @ delta_f

    va_diag = np.diag(va_ratio)
    delta_vab = va_diag @ delta_vbp

    # Efecto directo: solo la industria que recibe el shock (delta_f distinto de cero)
    delta_vab_directo = delta_vab[idx_gastro]
    delta_vab_indirecto = delta_vab.sum() - delta_vab_directo

    resultado = {
        "label": label,
        "delta_vbp_total_muyu2016": delta_vbp.sum(),
        "delta_vab_total_muyu2016": delta_vab.sum(),
        "delta_vab_directo_muyu2016": delta_vab_directo,
        "delta_vab_indirecto_muyu2016": delta_vab_indirecto,
        "delta_vab_por_industria": dict(zip(industrias, delta_vab.tolist())),
    }

    if coef_empleo is not None:
        delta_empleo = coef_empleo * delta_vbp
        resultado["delta_empleo_total"] = delta_empleo.sum()
        resultado["delta_empleo_directo"] = delta_empleo[idx_gastro]

    return resultado


# =============================================================================
# VALIDACIÓN FINAL: |ΔPIB| < 1% del PIB
# =============================================================================
def validar_magnitud_pib(delta_vab_muyu2016: float, pib_nominal_musd: float, tc_2022: float, deflactor: float):
    """
    Verifica que |ΔPIB| no supere el 1% del PIB. Si supera, es casi seguro un error de unidades.
    """
    # Convertir ΔVAB de MUYU 2016 a MUSD 2022 para comparar con el PIB
    delta_vab_musd2022 = delta_vab_muyu2016 / deflactor / tc_2022
    pct_pib = abs(delta_vab_musd2022) / pib_nominal_musd * 100

    print(f"  |ΔVAB| = {abs(delta_vab_musd2022):.1f} MUSD = {pct_pib:.3f}% del PIB")

    if pct_pib > 1.0:
        raise AssertionError(
            f"ALERTA DE UNIDADES: |ΔVAB| = {pct_pib:.2f}% del PIB > 1%.\n"
            "Casi seguramente hay un error de unidades o escala. Revisar:\n"
            "  1. La unidad de la MIP (¿millones o miles de millones de UYU?)\n"
            "  2. El tipo de cambio y deflactor usados\n"
            "  3. El valor de B (¿MUSD o USD?)"
        )
    return pct_pib


def main():
    separador("ETAPA 04 — Modelo Insumo-Producto (MIP)")

    supuestos = cargar_supuestos()

    # Leer shock procesado
    ruta_shocks = DATA_PROCESSED / "shocks.csv"
    if not ruta_shocks.exists():
        print("  ✗ No se encontró data/processed/shocks.csv — ejecutar Etapa 03 primero.")
        sys.exit(1)

    df_shocks = pd.read_csv(ruta_shocks)
    sc = df_shocks[df_shocks["label"] == "central_13pp"].iloc[0]
    sv = df_shocks[df_shocks["label"] == "variante_22pp"].iloc[0]

    delta_gastro_central = sc["delta_d_gastro_musd"]
    delta_gastro_variante = sv["delta_d_gastro_musd"]

    print(f"\n  ΔD_gastro (escenario central): {delta_gastro_central:.2f} MUSD")
    print(f"  ΔD_gastro (variante):          {delta_gastro_variante:.2f} MUSD")

    # Leer MIP
    ruta_mip = buscar_archivo("bcu_mip_2016.xlsx", ["mip_2016", "mip_bcu"])
    if ruta_mip is None:
        print("\n  ✗ No se encontró la MIP del BCU.")
        print("  INSTRUCCIÓN DE DESCARGA MANUAL:")
        print("    URL:     https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Matriz-Insumo-Producto.aspx")
        print("    Guardar: data/manual/bcu_mip_2016.xlsx")
        print("    Notas:   Descargar el Excel con la MIP 2016.")
        print("             El BCU publica la MIP completa (transacciones domésticas, VBP, VAB).")
        sys.exit(1)

    mip = leer_mip_bcu(ruta_mip)
    A, vbp, va_ratio = construir_A(mip)
    industrias = mip["industrias"]
    validar_A(A, industrias)

    print(f"\n  Construyendo inverso de Leontief ({len(industrias)}×{len(industrias)})…")
    L = calcular_leontief(A)
    print("  ✓ Inverso calculado.")

    # Identificar industria gastronómica en la lista (para validar multiplicador)
    industrias_lower = [str(i).lower() for i in industrias]
    idx_gastro_candidates = [
        i for i, ind in enumerate(industrias_lower)
        if any(kw in ind for kw in ["alojamiento", "comida", "restaur", "55", "56"])
    ]
    if not idx_gastro_candidates:
        print("  ⚠ No se encontró la industria de gastronomía/alojamiento en la MIP.")
        print(f"    Industrias disponibles: {industrias[:20]}")
        sys.exit(1)
    idx_gastro_val = idx_gastro_candidates[0]
    print(f"\n  Industria gastronómica: [{idx_gastro_val}] '{industrias[idx_gastro_val]}'")

    validar_multiplicadores(L, industrias, idx_gastro_val)

    # Conversión de unidades
    print("\n  Convirtiendo unidades: MUSD 2022 → MUYU 2016…")
    shock_uyu_central = convertir_shock_a_uyu_2016(delta_gastro_central, supuestos)
    shock_uyu_variante = convertir_shock_a_uyu_2016(delta_gastro_variante, supuestos)

    # Construir vector Δf
    mapeo_csv = ROOT / "config" / "mapeo_rubros_mip.csv"
    print("\n  Construyendo vector Δf (central)…")
    delta_f_central, idx_gastro = construir_delta_f(shock_uyu_central, industrias, mapeo_csv)
    print("  Construyendo vector Δf (variante)…")
    delta_f_variante, _ = construir_delta_f(shock_uyu_variante, industrias, mapeo_csv)

    # Calcular impactos
    print("\n  Calculando impactos…")
    imp_central = calcular_impacto(L, delta_f_central, va_ratio, idx_gastro, industrias, "central_13pp")
    imp_variante = calcular_impacto(L, delta_f_variante, va_ratio, idx_gastro, industrias, "variante_22pp")

    # Convertir a % del PIB
    pib_nominal_musd = obtener_valor(supuestos, "pib_nominal_musd")
    tc_2022 = obtener_valor(supuestos, "tc_promedio_2022")
    deflactor = obtener_valor(supuestos, "deflactor_2016_2022")

    print("\n  Validando magnitud del impacto…")
    for imp in [imp_central, imp_variante]:
        if all(v is not None for v in [pib_nominal_musd, tc_2022, deflactor]):
            pct = validar_magnitud_pib(
                imp["delta_vab_total_muyu2016"], pib_nominal_musd, tc_2022, deflactor
            )
            imp["delta_vab_pct_pib"] = round(pct, 4)
        else:
            imp["delta_vab_pct_pib"] = None
            print("  ⚠ PIB nominal no disponible — no se puede calcular % del PIB todavía.")

    # Mostrar resumen
    print("\n  RESUMEN DE IMPACTOS (Leontief Tipo I):")
    print(f"  {'Indicador':<45} {'Central':>12} {'Variante':>12}")
    print("  " + "-" * 70)
    for clave in ["delta_vab_total_muyu2016", "delta_vab_directo_muyu2016",
                  "delta_vab_indirecto_muyu2016", "delta_vab_pct_pib"]:
        v1 = imp_central.get(clave)
        v2 = imp_variante.get(clave)
        v1_str = f"{v1:>12.2f}" if v1 is not None else "        N/D "
        v2_str = f"{v2:>12.2f}" if v2 is not None else "        N/D "
        print(f"  {clave:<45} {v1_str} {v2_str}")

    # Guardar resultados
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)

    # Impacto por industria (top 10)
    df_ind = pd.DataFrame([
        {"industria": k, "delta_vab_central": v,
         "delta_vab_variante": imp_variante["delta_vab_por_industria"].get(k)}
        for k, v in imp_central["delta_vab_por_industria"].items()
    ]).sort_values("delta_vab_central")

    df_res = pd.DataFrame([imp_central, imp_variante]).drop(columns=["delta_vab_por_industria"])
    df_res.to_csv(DATA_PROCESSED / "impacto_mip.csv", index=False, encoding="utf-8")
    df_ind.to_csv(DATA_PROCESSED / "impacto_por_industria.csv", index=False, encoding="utf-8")
    df_res.to_csv(OUTPUTS_TABLAS / "tabla_impacto_mip.csv", index=False, encoding="utf-8")
    df_ind.head(20).to_csv(OUTPUTS_TABLAS / "tabla_impacto_industrias.csv", index=False, encoding="utf-8")
    print(f"\n  ✓ Guardados: data/processed/impacto_mip.csv, impacto_por_industria.csv")

    checkpoint(
        "Etapa 04 — Modelo MIP",
        [
            f"ΔVAB central: {imp_central['delta_vab_total_muyu2016']:.1f} MUYU 2016 "
            f"= {imp_central.get('delta_vab_pct_pib', 'N/D')}% del PIB",
            f"ΔVAB variante: {imp_variante['delta_vab_total_muyu2016']:.1f} MUYU 2016 "
            f"= {imp_variante.get('delta_vab_pct_pib', 'N/D')}% del PIB",
            f"Multiplicador tipo I gastronomía: verificar en consola (debe estar entre 1.2–2.0)",
            "Validar: (1) ¿El mapeo rubro→industria MIP es correcto? (2) ¿Las conversiones de unidades son correctas?",
            "Si ΔPIB% parece alto, revisar unidad de la MIP (¿miles de UYU o millones de UYU?).",
        ]
    )

    return imp_central, imp_variante


if __name__ == "__main__":
    main()
