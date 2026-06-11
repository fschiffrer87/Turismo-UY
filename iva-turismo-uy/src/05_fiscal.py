"""
Etapa 05 — Análisis fiscal.

Calcula:
  1. Recaudación GANADA (estática): turistas que siguen viniendo pagan IVA neto.
     = tasa_nueva_efectiva × base_remanente_sin_iva
     donde base_remanente = B × (1 + g_ext) × (1 + g_int)  (turistas que quedan)

  2. Recaudación PERDIDA por menor actividad:
     = presion_tributaria_efectiva × |ΔVAB|  (en unidades comparables)

  3. Efecto NETO = ganado − perdido

  4. Contraste con el Gasto Tributario que el fisco se ahorra (GT_estacional).

ACLARACIÓN EN OUTPUTS: recaudación no es lo mismo que PIB.
  - El PIB mide el impacto en la actividad económica real.
  - La recaudación mide la variación en ingresos del fisco.
  Son dos métricas del mismo trade-off, no se suman.
"""
import sys
import pathlib
import pandas as pd
import numpy as np

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS,
    cargar_supuestos, obtener_valor, separador, checkpoint
)


def calcular_fiscal(
    shock_label: str,
    b_central: float,
    g_ext: float,
    g_int: float,
    shock_pp: float,
    phi: float,
    delta_vab_musd: float,
    presion_tributaria: float,
    gt_estacional_musd: float | None,
) -> dict:
    """
    Calcula los tres componentes fiscales para un escenario dado.

    Unidades: millones de USD (misma base que B).
    """
    # Base remanente: turistas que siguen viniendo tras el shock
    base_remanente = b_central * (1 + g_ext) * (1 + g_int)

    # Tasa IVA efectiva nueva (sobre precio sin IVA)
    # El turista ahora paga IVA 13% (escenario central) o 22% (variante)
    # sobre el precio sin IVA. La tasa efectiva sobre la base sin IVA = shock_pp/100 × (1 − phi)
    tasa_nueva_efectiva = (shock_pp / 100) * (1 - phi)

    # Recaudación ganada (estática): base que permanece × nueva tasa IVA
    rec_ganada_musd = tasa_nueva_efectiva * base_remanente

    # Recaudación perdida por menor actividad económica
    rec_perdida_musd = presion_tributaria * abs(delta_vab_musd)

    rec_neta_musd = rec_ganada_musd - rec_perdida_musd

    resultado = {
        "label": shock_label,
        "b_central_musd": b_central,
        "base_remanente_musd": round(base_remanente, 2),
        "tasa_nueva_efectiva": round(tasa_nueva_efectiva, 4),
        "rec_ganada_musd": round(rec_ganada_musd, 2),
        "rec_perdida_musd": round(rec_perdida_musd, 2),
        "rec_neta_musd": round(rec_neta_musd, 2),
        "gt_estacional_ahorrado_musd": gt_estacional_musd,
        "rec_neta_vs_gt": (
            round(rec_neta_musd / gt_estacional_musd, 2)
            if gt_estacional_musd and gt_estacional_musd > 0
            else None
        ),
        "nota": (
            "ACLARACIÓN: La recaudación ganada es el IVA sobre los turistas que permanecen. "
            "La recaudación perdida es el impacto fiscal indirecto de la menor actividad. "
            "El GT ahorrado es el beneficio fiscal directo de eliminar la exoneración. "
            "Estas cifras complementan (no reemplazan) la estimación del impacto en PIB."
        ),
    }
    return resultado


def main():
    separador("ETAPA 05 — Análisis fiscal")

    supuestos = cargar_supuestos()

    # Leer resultados de etapas anteriores
    ruta_shocks = DATA_PROCESSED / "shocks.csv"
    ruta_impacto = DATA_PROCESSED / "impacto_mip.csv"
    ruta_b = DATA_PROCESSED / "base_afectada.csv"

    for ruta, nombre in [(ruta_shocks, "shocks.csv"),
                         (ruta_impacto, "impacto_mip.csv"),
                         (ruta_b, "base_afectada.csv")]:
        if not ruta.exists():
            print(f"  ✗ No se encontró {ruta} — ejecutar etapas anteriores primero.")
            sys.exit(1)

    df_shocks = pd.read_csv(ruta_shocks)
    df_impacto = pd.read_csv(ruta_impacto)
    df_b = pd.read_csv(ruta_b)

    b_central = float(df_b[df_b["via"] == "CENTRAL (elegida)"]["b_musd"].iloc[0])

    presion = obtener_valor(supuestos, "presion_tributaria_efectiva")
    gt_estacional = obtener_valor(supuestos, "B_dgi")  # B_dgi es el GT si fue despejado

    # Obtener parámetros de cada shock
    sc_row = df_shocks[df_shocks["label"] == "central_13pp"].iloc[0]
    sv_row = df_shocks[df_shocks["label"] == "variante_22pp"].iloc[0]

    imp_c = df_impacto[df_impacto["label"] == "central_13pp"].iloc[0]
    imp_v = df_impacto[df_impacto["label"] == "variante_22pp"].iloc[0]

    # Necesitamos ΔVAB en USD para calcular recaudación perdida
    tc_2022 = obtener_valor(supuestos, "tc_promedio_2022")
    deflactor = obtener_valor(supuestos, "deflactor_2016_2022")

    if tc_2022 is not None and deflactor is not None:
        def uyu2016_a_musd2022(v): return v / deflactor / tc_2022
        delta_vab_c_musd = uyu2016_a_musd2022(imp_c["delta_vab_total_muyu2016"])
        delta_vab_v_musd = uyu2016_a_musd2022(imp_v["delta_vab_total_muyu2016"])
    else:
        print("  ⚠ Sin conversión de unidades — usando ΔVAB en MUYU (comparación fiscal aproximada).")
        delta_vab_c_musd = imp_c["delta_vab_total_muyu2016"]
        delta_vab_v_musd = imp_v["delta_vab_total_muyu2016"]

    # Calcular fiscal
    fiscal_c = calcular_fiscal(
        shock_label="central_13pp",
        b_central=b_central,
        g_ext=sc_row["g_ext"],
        g_int=sc_row["g_int"],
        shock_pp=sc_row["shock_pp"],
        phi=sc_row["phi"],
        delta_vab_musd=delta_vab_c_musd,
        presion_tributaria=presion,
        gt_estacional_musd=gt_estacional,
    )

    fiscal_v = calcular_fiscal(
        shock_label="variante_22pp",
        b_central=b_central,
        g_ext=sv_row["g_ext"],
        g_int=sv_row["g_int"],
        shock_pp=sv_row["shock_pp"],
        phi=sv_row["phi"],
        delta_vab_musd=delta_vab_v_musd,
        presion_tributaria=presion,
        gt_estacional_musd=gt_estacional,
    )

    # Mostrar tabla
    print("\n  ANÁLISIS FISCAL (millones de USD):")
    print(f"  {'Componente':<45} {'Central (+13pp)':>16} {'Variante (+22pp)':>16}")
    print("  " + "-" * 78)
    for clave in ["base_remanente_musd", "tasa_nueva_efectiva", "rec_ganada_musd",
                  "rec_perdida_musd", "rec_neta_musd",
                  "gt_estacional_ahorrado_musd", "rec_neta_vs_gt"]:
        v1 = fiscal_c.get(clave)
        v2 = fiscal_v.get(clave)
        v1_str = f"{v1:>16.3f}" if v1 is not None else "             N/D"
        v2_str = f"{v2:>16.3f}" if v2 is not None else "             N/D"
        print(f"  {clave:<45} {v1_str} {v2_str}")

    print(f"\n  NOTA: {fiscal_c['nota']}")

    # Guardar
    df_fiscal = pd.DataFrame([fiscal_c, fiscal_v]).drop(columns=["nota"])
    df_fiscal.to_csv(DATA_PROCESSED / "fiscal.csv", index=False, encoding="utf-8")
    df_fiscal.to_csv(OUTPUTS_TABLAS / "tabla_fiscal.csv", index=False, encoding="utf-8")
    print(f"\n  ✓ Guardado: data/processed/fiscal.csv")

    checkpoint(
        "Etapa 05 — Fiscal",
        [
            f"Recaudación neta central: {fiscal_c['rec_neta_musd']:.1f} MUSD",
            f"Recaudación neta variante: {fiscal_v['rec_neta_musd']:.1f} MUSD",
            f"GT ahorrado (fisco no desembolsa): {gt_estacional if gt_estacional else 'N/D'} MUSD",
            "Validar: ¿el sign de rec_neta tiene sentido económico?",
            "Recordar: si rec_neta > 0, la medida mejora el balance fiscal neto.",
        ]
    )

    return fiscal_c, fiscal_v


if __name__ == "__main__":
    main()
