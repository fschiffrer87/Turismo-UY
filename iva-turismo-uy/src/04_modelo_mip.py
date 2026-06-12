"""
Etapa 04 — Modelo Insumo-Producto de Leontief (MIP 2016 BCU, 108 industrias).

Layout VERIFICADO del archivo bcu_mip_2016.xlsx, hoja 'MIP 108 x 108':
  - Filas 9–116:  industrias A.1–A.108 (col 0 = código, col 1 = nombre)
  - Columnas 6–113: demanda intermedia por industria de destino
  - La matriz 108×108 es de ORIGEN NACIONAL (doméstica): verificado que la
    suma de columnas coincide con la fila 117 'Total de usos de origen nacional'.
  - Fila 121: VAB precios básicos | Fila 122: Producción (VBP) | Fila 131: Empleo
  - Unidad: millones de pesos uruguayos (MUYU) corrientes de 2016.

Cálculo:
  A[i,j] = Z[i,j] / VBP[j]      L = (I − A)⁻¹
  ΔVBP = L · Δf                 ΔVAB = v̂ · ΔVBP        ΔL = l̂ · ΔVBP

Conversión de unidades del shock (documentada):
  Δf [MUSD 2022] × TC 41,12 [UYU/USD 2022] × 0,6285 [IPC16/IPC22] = Δf [MUYU 2016]
"""
import sys
import pathlib
import numpy as np
import pandas as pd

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS,
    cargar_supuestos, obtener_valor, buscar_archivo, separador, checkpoint
)

# Layout de la hoja 'MIP 108 x 108' (verificado contra el archivo BCU)
FILA_IND_INICIO, FILA_IND_FIN = 9, 117      # filas 9..116 = A.1..A.108
COL_IND_INICIO, COL_IND_FIN = 6, 114        # cols 6..113
FILA_VAB, FILA_VBP, FILA_EMPLEO = 121, 122, 131


def leer_mip(ruta: pathlib.Path) -> dict:
    """Lee la MIP doméstica, VBP, VAB y empleo del Excel del BCU."""
    print(f"  → Leyendo MIP: {ruta.name} (hoja 'MIP 108 x 108')")
    df = pd.read_excel(ruta, sheet_name="MIP 108 x 108", header=None)

    codigos = df.iloc[FILA_IND_INICIO:FILA_IND_FIN, 0].astype(str).tolist()
    nombres = df.iloc[FILA_IND_INICIO:FILA_IND_FIN, 1].astype(str).tolist()

    Z = df.iloc[FILA_IND_INICIO:FILA_IND_FIN, COL_IND_INICIO:COL_IND_FIN] \
          .apply(pd.to_numeric, errors="coerce").fillna(0).values

    def fila(n):
        return pd.to_numeric(df.iloc[n, COL_IND_INICIO:COL_IND_FIN], errors="coerce") \
                 .fillna(0).values.astype(float)

    vbp, vab, empleo = fila(FILA_VBP), fila(FILA_VAB), fila(FILA_EMPLEO)

    # Verificación de matriz doméstica: suma de columnas de Z = fila 117
    total_nacional = fila(117)
    assert np.allclose(Z.sum(axis=0), total_nacional, rtol=0.01), (
        "La suma de columnas de Z no coincide con 'Total de usos de origen nacional' "
        "— el layout del archivo cambió, revisar leer_mip()."
    )
    print(f"     {len(codigos)} industrias | VBP total: {vbp.sum():,.0f} MUYU 2016 "
          f"| VAB total: {vab.sum():,.0f} MUYU 2016")
    print("     ✓ Verificado: la matriz es de origen nacional (doméstica).")

    return {"Z": Z, "vbp": vbp, "vab": vab, "empleo": empleo,
            "codigos": codigos, "nombres": nombres}


def construir_modelo(mip: dict):
    """Construye A, el inverso de Leontief y los coeficientes VA y empleo."""
    Z, vbp, vab, empleo = mip["Z"], mip["vbp"], mip["vab"], mip["empleo"]
    with np.errstate(divide="ignore", invalid="ignore"):
        A = np.where(vbp > 0, Z / vbp, 0.0)
        va_ratio = np.where(vbp > 0, vab / vbp, 0.0)
        l_ratio = np.where(vbp > 0, empleo / vbp, 0.0)   # empleos por MUYU de VBP

    sumas = A.sum(axis=0)
    assert (A >= -1e-9).all(), "Valores negativos en A — revisar lectura."
    if (sumas > 1.0).any():
        print(f"  ⚠ {int((sumas > 1.0).sum())} columnas de A suman >1 — revisar.")

    L = np.linalg.inv(np.eye(A.shape[0]) - A)
    return A, L, va_ratio, l_ratio


def convertir_a_uyu2016(musd_2022: float, tc: float, deflactor: float) -> float:
    """MUSD corrientes 2022 → MUYU constantes 2016 (procedimiento documentado)."""
    return musd_2022 * tc * deflactor


def construir_delta_f(label: str, df_rubros: pd.DataFrame, df_mapeo: pd.DataFrame,
                      codigos: list, tc: float, deflactor: float) -> np.ndarray:
    """
    Vector Δf (108,): asigna el ΔD de cada rubro ETR a su industria MIP
    según config/mapeo_rubros_mip.csv, convertido a MUYU 2016.
    """
    delta_f = np.zeros(len(codigos))
    d = df_rubros[df_rubros["label"] == label]
    for _, fila in d.iterrows():
        mapeo = df_mapeo[df_mapeo["rubro_etr"] == fila["rubro"]]
        if mapeo.empty:
            print(f"  ⚠ Rubro '{fila['rubro']}' sin mapeo — omitido.")
            continue
        codigo = mapeo["codigo_mip"].iloc[0]
        if codigo not in codigos:
            print(f"  ⚠ Código '{codigo}' no está en la MIP — omitido.")
            continue
        idx = codigos.index(codigo)
        delta_f[idx] += convertir_a_uyu2016(fila["delta_d_musd"], tc, deflactor)
    return delta_f


def calcular_impacto(L, delta_f, va_ratio, l_ratio, codigos, nombres, label) -> dict:
    """ΔVBP, ΔVAB, ΔEmpleo y descomposición directo/indirecto."""
    delta_vbp = L @ delta_f
    delta_vab = va_ratio * delta_vbp
    delta_emp = l_ratio * delta_vbp

    # Directo: VAB generado en las industrias que reciben shock, por su propio Δf
    mask_shock = delta_f != 0
    vab_directo = (va_ratio * delta_f)[mask_shock].sum()
    vab_indirecto = delta_vab.sum() - vab_directo

    return {
        "label": label,
        "delta_f_total_muyu2016": delta_f.sum(),
        "delta_vbp_total_muyu2016": delta_vbp.sum(),
        "delta_vab_total_muyu2016": delta_vab.sum(),
        "delta_vab_directo_muyu2016": vab_directo,
        "delta_vab_indirecto_muyu2016": vab_indirecto,
        "delta_empleo_puestos": delta_emp.sum(),
        "_vab_por_industria": pd.DataFrame({
            "codigo": codigos, "industria": nombres, "delta_vab": delta_vab
        }),
    }


def main():
    separador("ETAPA 04 — Modelo Insumo-Producto (MIP BCU 2016)")
    supuestos = cargar_supuestos()

    tc = obtener_valor(supuestos, "tc_promedio_2022")
    deflactor = obtener_valor(supuestos, "deflactor_2016_2022")
    pib_musd = obtener_valor(supuestos, "pib_nominal_musd")
    if not all([tc, deflactor, pib_musd]):
        print("  ✗ Faltan tc_promedio_2022 / deflactor_2016_2022 / pib_nominal_musd en supuestos.yaml")
        sys.exit(1)

    ruta_rubros = DATA_PROCESSED / "shock_por_rubros.csv"
    if not ruta_rubros.exists():
        print("  ✗ Ejecutar primero la Etapa 03.")
        sys.exit(1)
    df_rubros = pd.read_csv(ruta_rubros)
    df_mapeo = pd.read_csv(ROOT / "config" / "mapeo_rubros_mip.csv")

    ruta_mip = buscar_archivo("bcu_mip_2016.xlsx")
    if ruta_mip is None:
        print("  ✗ Colocar la MIP en data/manual/bcu_mip_2016.xlsx")
        sys.exit(1)

    mip = leer_mip(ruta_mip)
    A, L, va_ratio, l_ratio = construir_modelo(mip)
    codigos, nombres = mip["codigos"], mip["nombres"]

    # Validación 2: multiplicador tipo I de gastronomía en [1.2, 2.0]
    idx_gastro = codigos.index("A.81")
    mult_gastro = L.sum(axis=0)[idx_gastro]
    print(f"\n  Multiplicador tipo I de A.81 (Servicio de alimento y bebida): {mult_gastro:.3f}")
    assert 1.2 <= mult_gastro <= 2.0, (
        f"Multiplicador gastronomía {mult_gastro:.3f} fuera de [1.2, 2.0] — revisar A."
    )

    print(f"\n  Conversión de unidades: MUSD 2022 × TC {tc} × deflactor {deflactor} → MUYU 2016")

    resultados = []
    for label in ["central_13pp", "variante_22pp"]:
        delta_f = construir_delta_f(label, df_rubros, df_mapeo, codigos, tc, deflactor)
        imp = calcular_impacto(L, delta_f, va_ratio, l_ratio, codigos, nombres, label)

        # ΔVAB en % del PIB: reconvertir MUYU 2016 → MUSD 2022
        delta_vab_musd = imp["delta_vab_total_muyu2016"] / deflactor / tc
        pct_pib = delta_vab_musd / pib_musd * 100
        imp["delta_vab_musd2022"] = round(delta_vab_musd, 2)
        imp["delta_vab_pct_pib"] = round(pct_pib, 4)

        # Validación 3: |ΔPIB| < 1%
        assert abs(pct_pib) < 1.0, (
            f"|ΔPIB| = {abs(pct_pib):.2f}% > 1% — casi seguro error de unidades."
        )
        resultados.append(imp)

    imp_c, imp_v = resultados

    print("\n  RESUMEN DE IMPACTOS (Leontief Tipo I):")
    print(f"  {'Indicador':<42} {'Central +13pp':>14} {'Variante +22pp':>15}")
    print("  " + "-" * 73)
    filas_mostrar = [
        ("Δf total (MUYU 2016)", "delta_f_total_muyu2016", 1),
        ("ΔVBP total (MUYU 2016)", "delta_vbp_total_muyu2016", 1),
        ("ΔVAB total (MUYU 2016)", "delta_vab_total_muyu2016", 1),
        ("ΔVAB directo (MUYU 2016)", "delta_vab_directo_muyu2016", 1),
        ("ΔVAB indirecto (MUYU 2016)", "delta_vab_indirecto_muyu2016", 1),
        ("ΔVAB (MUSD 2022)", "delta_vab_musd2022", 2),
        ("ΔVAB (% del PIB)", "delta_vab_pct_pib", 4),
        ("ΔEmpleo (puestos de trabajo)", "delta_empleo_puestos", 0),
    ]
    for nombre, clave, dec in filas_mostrar:
        print(f"  {nombre:<42} {imp_c[clave]:>14.{dec}f} {imp_v[clave]:>15.{dec}f}")

    # Top industrias afectadas
    top = imp_c["_vab_por_industria"].sort_values("delta_vab").head(10)
    print("\n  TOP 10 INDUSTRIAS MÁS AFECTADAS (ΔVAB MUYU 2016, central):")
    for _, f in top.iterrows():
        print(f"    {f['codigo']:<7} {f['industria'][:60]:<62} {f['delta_vab']:>10.1f}")

    # Guardar
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)
    df_res = pd.DataFrame([{k: v for k, v in r.items() if not k.startswith("_")} for r in resultados])
    df_res.to_csv(DATA_PROCESSED / "impacto_mip.csv", index=False, encoding="utf-8")
    df_res.to_csv(OUTPUTS_TABLAS / "tabla_impacto_mip.csv", index=False, encoding="utf-8")

    df_ind = imp_c["_vab_por_industria"].merge(
        imp_v["_vab_por_industria"], on=["codigo", "industria"], suffixes=("_central", "_variante")
    ).sort_values("delta_vab_central")
    df_ind.to_csv(DATA_PROCESSED / "impacto_por_industria.csv", index=False, encoding="utf-8")
    df_ind.head(20).to_csv(OUTPUTS_TABLAS / "tabla_impacto_industrias.csv", index=False, encoding="utf-8")
    print("\n  ✓ Guardado: impacto_mip.csv, impacto_por_industria.csv")

    checkpoint(
        "Etapa 04 — Modelo MIP",
        [
            f"ΔVAB central: {imp_c['delta_vab_musd2022']:.1f} MUSD = {imp_c['delta_vab_pct_pib']:.4f}% del PIB",
            f"ΔVAB variante: {imp_v['delta_vab_musd2022']:.1f} MUSD = {imp_v['delta_vab_pct_pib']:.4f}% del PIB",
            f"Multiplicador tipo I gastronomía: {mult_gastro:.3f} (validado en [1.2, 2.0])",
            f"ΔEmpleo central: {imp_c['delta_empleo_puestos']:.0f} puestos",
            "Validar: mapeo rubros→industrias y conversión de unidades.",
        ]
    )
    return imp_c, imp_v


if __name__ == "__main__":
    main()
