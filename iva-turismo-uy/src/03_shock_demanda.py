"""
Etapa 03 — Shock de precio y respuesta de demanda.

Fórmulas (documentadas en README y en el informe final):
  Δp = (1 − φ) × shock_pp / 100
  g_ext = ε_ext × (Δp × w)           ← margen extensivo (todo el gasto turístico)
  g_int = ε_int × Δp                  ← margen intensivo (solo gasto gastronómico)
  ΔD_gastro = B × [(1 + g_ext) × (1 + g_int) − 1]

Resto de rubros (alojamiento, transporte, etc.): solo margen extensivo,
aplicado al gasto de temporada del rubro escalado por la fracción de turistas
efectivamente afectados (afectacion_share = B / gasto_alimentacion_temporada).
Así se evita el doble conteo y no se atribuye el shock a turistas que no
acceden al beneficio (pagan efectivo / no consumen restaurantes).

Unidades: millones de USD corrientes del año base.
"""
import sys
import pathlib
import pandas as pd
import numpy as np

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS,
    cargar_supuestos, obtener_valor, buscar_archivo, separador, checkpoint
)

# Misma ventana de temporada que la Etapa 02
TEMPORADA_INICIO = "2022-11-15"
TEMPORADA_FIN = "2023-04-30"

RUBROS_ETR = [
    "GastoAlimentacion", "GastoAlojamiento", "GastoTransporte",
    "GastoCultural", "GastoTours", "GastoCompras", "GastoOtros",
]


def calcular_shock(shock_pp, phi, b_central, w, epsilon_ext, epsilon_int, label):
    """Calcula la respuesta de demanda gastronómica para un shock dado."""
    assert epsilon_ext <= 0 and epsilon_int <= 0, "elasticidades deben ser negativas"
    assert 0 <= phi < 1, "phi debe estar en [0, 1)"
    assert 0 < w <= 1, "w debe estar en (0, 1]"

    delta_p = (1 - phi) * (shock_pp / 100)
    g_ext = epsilon_ext * (delta_p * w)
    g_int = epsilon_int * delta_p
    delta_d_gastro = b_central * ((1 + g_ext) * (1 + g_int) - 1)

    return {
        "label": label,
        "shock_pp": shock_pp,
        "phi": phi,
        "delta_p": round(delta_p, 4),
        "g_ext": round(g_ext, 4),
        "g_int": round(g_int, 4),
        "b_central_musd": b_central,
        "w": w,
        "epsilon_ext": epsilon_ext,
        "epsilon_int": epsilon_int,
        "delta_d_gastro_musd": round(delta_d_gastro, 2),
        "delta_d_gastro_pct_de_b": round(delta_d_gastro / b_central * 100, 1),
    }


def gasto_temporada_por_rubro() -> pd.Series:
    """Gasto expandido de temporada por rubro de la ETR (MUSD)."""
    ruta = buscar_archivo("mintur_etr_agregados.xlsx")
    if ruta is None:
        raise FileNotFoundError("ETR no encontrada en data/manual/")
    df = pd.read_excel(ruta, sheet_name="Datos")
    df["fi"] = pd.to_datetime(df["FechaIngreso"])
    m = (df["fi"] >= TEMPORADA_INICIO) & (df["fi"] <= TEMPORADA_FIN)
    d = df[m]
    return pd.Series(
        {r: (d[r] * d["Coef"]).sum() / 1e6 for r in RUBROS_ETR}, name="gasto_temporada_musd"
    )


def construir_shock_rubros(sc: dict, gasto_rubros: pd.Series) -> pd.DataFrame:
    """
    Construye el vector de shock por rubro de la ETR.
      - Gastronomía: efecto combinado (extensivo × intensivo) sobre B.
      - Resto: solo extensivo, sobre el gasto de temporada del rubro escalado
        por afectacion_share = B / gasto_alimentacion_temporada.
    """
    b = sc["b_central_musd"]
    alim_temp = gasto_rubros["GastoAlimentacion"]
    afectacion_share = b / alim_temp
    g_ext = sc["g_ext"]

    filas = []
    for rubro in RUBROS_ETR:
        gasto = gasto_rubros[rubro]
        if rubro == "GastoAlimentacion":
            delta = sc["delta_d_gastro_musd"]
            detalle = "extensivo × intensivo sobre B"
        else:
            delta = gasto * afectacion_share * g_ext
            detalle = f"solo extensivo × afectacion_share {afectacion_share:.3f}"
        filas.append({
            "rubro": rubro,
            "gasto_temporada_musd": round(gasto, 1),
            "delta_d_musd": round(delta, 2),
            "detalle": detalle,
            "label": sc["label"],
        })
    return pd.DataFrame(filas)


def main():
    separador("ETAPA 03 — Shock de precio y respuesta de demanda")
    supuestos = cargar_supuestos()

    # B central decidida en Etapa 02 (vía DGI, aprobada por el usuario)
    ruta_b = DATA_PROCESSED / "rango_b.csv"
    if not ruta_b.exists():
        print("  ✗ Ejecutar primero la Etapa 02.")
        sys.exit(1)
    rb = pd.read_csv(ruta_b).iloc[0]
    b_central = float(rb["b_central"])
    print(f"\n  B central: {b_central:.1f} MUSD (rango MC: [{rb['b_min']:.1f}, {rb['b_max']:.1f}])")

    phi = obtener_valor(supuestos, "phi")
    eps_int = obtener_valor(supuestos, "epsilon_intensivo")
    eps_ext = obtener_valor(supuestos, "epsilon_extensivo")
    w = obtener_valor(supuestos, "w") or obtener_valor(supuestos, "w_fallback")
    print(f"  Parámetros: φ={phi}, ε_int={eps_int}, ε_ext={eps_ext}, w={w} "
          f"({'OBSERVADO ETR' if obtener_valor(supuestos, 'w') else 'fallback SUPUESTO'})")

    shock_c = obtener_valor(supuestos, "shock_pp")
    shock_v = obtener_valor(supuestos, "shock_pp_variante")

    sc = calcular_shock(shock_c, phi, b_central, w, eps_ext, eps_int, "central_13pp")
    sv = calcular_shock(shock_v, phi, b_central, w, eps_ext, eps_int, "variante_22pp")

    print("\n  RESULTADOS DEL SHOCK GASTRONÓMICO:")
    print(f"  {'Componente':<32} {'Central (+13pp)':>16} {'Variante (+22pp)':>16}")
    print("  " + "-" * 66)
    for k in ["delta_p", "g_ext", "g_int", "delta_d_gastro_musd", "delta_d_gastro_pct_de_b"]:
        print(f"  {k:<32} {sc[k]:>16.4f} {sv[k]:>16.4f}")

    # Shock por rubro (para el vector Δf de la MIP)
    print("\n  Gasto de temporada por rubro (ETR, MUSD) y shock:")
    gasto_rubros = gasto_temporada_por_rubro()
    df_rub_c = construir_shock_rubros(sc, gasto_rubros)
    df_rub_v = construir_shock_rubros(sv, gasto_rubros)
    df_rubros = pd.concat([df_rub_c, df_rub_v], ignore_index=True)

    print(df_rub_c[["rubro", "gasto_temporada_musd", "delta_d_musd"]].to_string(index=False))
    total_c = df_rub_c["delta_d_musd"].sum()
    total_v = df_rub_v["delta_d_musd"].sum()
    print(f"\n  ΔD total (central): {total_c:.2f} MUSD | (variante): {total_v:.2f} MUSD")

    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)
    pd.DataFrame([sc, sv]).to_csv(DATA_PROCESSED / "shocks.csv", index=False, encoding="utf-8")
    df_rubros.to_csv(DATA_PROCESSED / "shock_por_rubros.csv", index=False, encoding="utf-8")
    pd.DataFrame([sc, sv]).to_csv(OUTPUTS_TABLAS / "tabla_shocks.csv", index=False, encoding="utf-8")
    df_rubros.to_csv(OUTPUTS_TABLAS / "tabla_shock_rubros.csv", index=False, encoding="utf-8")
    print("\n  ✓ Guardado: shocks.csv, shock_por_rubros.csv")

    checkpoint(
        "Etapa 03 — Shock y demanda",
        [
            f"Central (+13 pp): ΔD_gastro = {sc['delta_d_gastro_musd']:.1f} MUSD "
            f"({sc['delta_d_gastro_pct_de_b']:.1f}% de B); ΔD total con arrastre = {total_c:.1f} MUSD",
            f"Variante (+22 pp): ΔD_gastro = {sv['delta_d_gastro_musd']:.1f} MUSD; "
            f"ΔD total = {total_v:.1f} MUSD",
            f"Supuestos: φ={phi}, ε_int={eps_int}, ε_ext={eps_ext}; w={w} (observado)",
            "Revisar mapeo rubros→MIP en config/mapeo_rubros_mip.csv antes de la Etapa 04.",
        ]
    )
    return {"sc": sc, "sv": sv}


if __name__ == "__main__":
    main()
