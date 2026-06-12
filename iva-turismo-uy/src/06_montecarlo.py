"""
Etapa 06 — Simulación Monte Carlo.

Parámetros con distribución triangular:
  B          ~ Triangular(b_min, b_central, b_max)  [MUSD]
  phi        ~ Triangular(0.0, 0.15, 0.30)
  epsilon_int ~ Triangular(-1.5, -1.0, -0.6)  (negativo)
  epsilon_ext ~ Triangular(-1.8, -1.2, -0.8)  (negativo)
  w           ~ Triangular(0.14, 0.18, 0.22)   (o calculado de ETR)

Outputs:
  - Mediana, media, P10, P90 de ΔPIB (% del PIB)
  - Histograma (PNG 300 dpi)
  - Gráfico tornado de sensibilidad (one-at-a-time)
  - CSV con todas las simulaciones

Se corre para escenario central (+13 pp) y variante (+22 pp) con seed=42.
"""
import sys
import pathlib
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")  # sin display
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS, OUTPUTS_FIGURAS,
    cargar_supuestos, obtener_valor, separador, checkpoint
)

# Importar la función de shock para reusar la lógica sin duplicar
sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import cargar_supuestos


def triangular(rng: np.random.Generator, low: float, mode: float, high: float, n: int) -> np.ndarray:
    """Genera n muestras de una distribución triangular."""
    return rng.triangular(low, mode, high, n)


def simular_delta_pib_pct(
    n_sims: int,
    seed: int,
    shock_pp: float,
    b_central: float,
    b_min: float,
    b_max: float,
    phi_vals: tuple,
    eps_int_vals: tuple,
    eps_ext_vals: tuple,
    w_vals: tuple,
    multiplicador_va: float,  # ΔVAB / ΔD (multiplicador promedio del modelo MIP)
    pib_nominal_musd: float,
    label: str,
) -> pd.DataFrame:
    """
    Simula n_sims iteraciones del modelo, devolviendo ΔPIB% para cada una.

    multiplicador_va: ratio ΔVAB / ΔD_gastro extraído del modelo MIP determinístico.
                      Si no está disponible, usar 1.0 (impacto directo sin multiplicador).
    """
    rng = np.random.default_rng(seed)

    B_sim = triangular(rng, b_min, b_central, b_max, n_sims)
    phi_sim = triangular(rng, *phi_vals, n_sims)
    eps_int_sim = triangular(rng, eps_int_vals[0], eps_int_vals[1], eps_int_vals[2], n_sims)
    eps_ext_sim = triangular(rng, eps_ext_vals[0], eps_ext_vals[1], eps_ext_vals[2], n_sims)
    w_sim = triangular(rng, *w_vals, n_sims)

    delta_p = (1 - phi_sim) * (shock_pp / 100)
    g_ext = eps_ext_sim * (delta_p * w_sim)
    g_int = eps_int_sim * delta_p

    delta_d = B_sim * ((1 + g_ext) * (1 + g_int) - 1)

    # ΔVAB ≈ delta_d × multiplicador_va
    delta_vab = delta_d * multiplicador_va

    # Expresar como % del PIB
    delta_pib_pct = delta_vab / pib_nominal_musd * 100

    df = pd.DataFrame({
        "sim": np.arange(n_sims),
        "B_sim": B_sim,
        "phi_sim": phi_sim,
        "eps_int_sim": eps_int_sim,
        "eps_ext_sim": eps_ext_sim,
        "w_sim": w_sim,
        "delta_d_gastro_musd": delta_d,
        "delta_vab_musd": delta_vab,
        "delta_pib_pct": delta_pib_pct,
        "label": label,
    })
    return df


def calcular_estadisticas(df: pd.DataFrame, label: str) -> dict:
    """Calcula estadísticos resumen de ΔPIB%."""
    col = df["delta_pib_pct"]
    return {
        "label": label,
        "mediana": round(col.median(), 4),
        "media": round(col.mean(), 4),
        "p10": round(col.quantile(0.10), 4),
        "p90": round(col.quantile(0.90), 4),
        "min": round(col.min(), 4),
        "max": round(col.max(), 4),
        "n_sims": len(df),
    }


def grafico_histograma(df_c: pd.DataFrame, df_v: pd.DataFrame, ruta_salida: pathlib.Path):
    """Histograma superpuesto de ΔPIB% para los dos escenarios."""
    fig, ax = plt.subplots(figsize=(10, 5))

    bins = np.linspace(
        min(df_c["delta_pib_pct"].min(), df_v["delta_pib_pct"].min()),
        max(df_c["delta_pib_pct"].max(), df_v["delta_pib_pct"].max()),
        60
    )

    ax.hist(df_c["delta_pib_pct"], bins=bins, alpha=0.6,
            color="#2166ac", label="Escenario central (+13 pp)", density=True)
    ax.hist(df_v["delta_pib_pct"], bins=bins, alpha=0.6,
            color="#d6604d", label="Variante (+22 pp)", density=True)

    for df, color, ls in [(df_c, "#2166ac", "--"), (df_v, "#d6604d", ":")]:
        med = df["delta_pib_pct"].median()
        ax.axvline(med, color=color, linestyle=ls, linewidth=1.5,
                   label=f"Mediana {df['label'].iloc[0].split('_')[1]}: {med:.3f}%")

    ax.set_xlabel("Variación del PIB (%)", fontsize=12)
    ax.set_ylabel("Densidad", fontsize=12)
    ax.set_title(
        "Distribución del impacto sobre el PIB — Simulación Monte Carlo\n"
        "(n = {:,} iteraciones, eliminación exoneración estacional IVA gastronomía turística)".format(
            len(df_c)
        ),
        fontsize=11
    )
    ax.legend(fontsize=10)
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f%%"))
    plt.tight_layout()
    fig.savefig(ruta_salida, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"  ✓ Histograma guardado: {ruta_salida.name}")


def grafico_tornado(
    shock_pp: float,
    b_central: float,
    b_min: float,
    b_max: float,
    phi_vals: tuple,
    eps_int_vals: tuple,
    eps_ext_vals: tuple,
    w_vals: tuple,
    multiplicador_va: float,
    pib_nominal_musd: float,
    ruta_salida: pathlib.Path,
    label: str,
):
    """
    Gráfico tornado: sensibilidad one-at-a-time (OAT) sobre los rangos de cada parámetro.
    Para cada parámetro, calcula ΔPIB% con el valor mínimo y máximo (manteniendo el resto central).
    """
    def calcular_delta_pib(b, phi, eps_int, eps_ext, w) -> float:
        delta_p = (1 - phi) * (shock_pp / 100)
        g_ext = eps_ext * (delta_p * w)
        g_int = eps_int * delta_p
        delta_d = b * ((1 + g_ext) * (1 + g_int) - 1)
        delta_vab = delta_d * multiplicador_va
        return delta_vab / pib_nominal_musd * 100

    # Valores centrales
    params_c = {
        "B (MUSD)": b_central,
        "φ (absorción)": phi_vals[1],
        "ε intensivo": eps_int_vals[1],
        "ε extensivo": eps_ext_vals[1],
        "w (peso gastro)": w_vals[1],
    }
    params_rangos = {
        "B (MUSD)": (b_min, b_max),
        "φ (absorción)": (phi_vals[0], phi_vals[2]),
        "ε intensivo": (eps_int_vals[0], eps_int_vals[2]),
        "ε extensivo": (eps_ext_vals[0], eps_ext_vals[2]),
        "w (peso gastro)": (w_vals[0], w_vals[2]),
    }

    delta_c_base = calcular_delta_pib(
        b_central, phi_vals[1], eps_int_vals[1], eps_ext_vals[1], w_vals[1]
    )

    rangos_tornado = []
    for nombre, (vmin, vmax) in params_rangos.items():
        otros = {k: v for k, v in params_c.items() if k != nombre}

        def calc(val):
            args = list(params_c.values())
            idx = list(params_c.keys()).index(nombre)
            args[idx] = val
            return calcular_delta_pib(*args)

        d_min = calc(vmin)
        d_max = calc(vmax)
        rangos_tornado.append({
            "parametro": nombre,
            "delta_pib_min": min(d_min, d_max),
            "delta_pib_max": max(d_min, d_max),
            "amplitud": abs(d_max - d_min),
        })

    df_torn = pd.DataFrame(rangos_tornado).sort_values("amplitud")

    fig, ax = plt.subplots(figsize=(10, 5))
    colores = {"izq": "#92c5de", "der": "#f4a582"}

    for i, row in df_torn.iterrows():
        ax.barh(
            row["parametro"],
            row["delta_pib_max"] - delta_c_base,
            left=delta_c_base,
            color=colores["der"], edgecolor="white"
        )
        ax.barh(
            row["parametro"],
            row["delta_pib_min"] - delta_c_base,
            left=delta_c_base,
            color=colores["izq"], edgecolor="white"
        )

    ax.axvline(delta_c_base, color="black", linewidth=1.0, linestyle="--")
    ax.set_xlabel("Variación del PIB (%)", fontsize=11)
    ax.set_title(
        f"Análisis de sensibilidad (tornado) — {label}\n"
        "Rangos de parámetros one-at-a-time",
        fontsize=11
    )
    ax.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.3f%%"))
    plt.tight_layout()
    fig.savefig(ruta_salida, dpi=300, bbox_inches="tight")
    plt.close(fig)
    print(f"  ✓ Tornado guardado: {ruta_salida.name}")


def main():
    separador("ETAPA 06 — Monte Carlo")

    supuestos = cargar_supuestos()
    n_sims = obtener_valor(supuestos, "n_sims")
    seed = obtener_valor(supuestos, "seed")
    shock_c = obtener_valor(supuestos, "shock_pp")
    shock_v = obtener_valor(supuestos, "shock_pp_variante")

    # Leer B central y rango
    ruta_b = DATA_PROCESSED / "base_afectada.csv"
    if not ruta_b.exists():
        print("  ✗ base_afectada.csv no encontrado.")
        sys.exit(1)
    df_b = pd.read_csv(ruta_b)
    b_central = float(df_b[df_b["via"] == "CENTRAL (elegida)"]["b_musd"].iloc[0])
    b_min_row = df_b[df_b["via"].str.contains("DGI|ETR|BCU")]["b_musd"].dropna()
    b_min = float(b_min_row.min()) if not b_min_row.empty else b_central * 0.7
    b_max = float(b_min_row.max()) if not b_min_row.empty else b_central * 1.3

    # Leer multiplicador del modelo MIP (ratio ΔVAB/ΔD del escenario central)
    ruta_shocks = DATA_PROCESSED / "shocks.csv"
    ruta_mip = DATA_PROCESSED / "impacto_mip.csv"

    if ruta_mip.exists() and ruta_shocks.exists():
        df_mip = pd.read_csv(ruta_mip)
        df_sh = pd.read_csv(ruta_shocks)
        delta_d_c = float(df_sh[df_sh["label"] == "central_13pp"]["delta_d_gastro_musd"].iloc[0])
        delta_vab_c_muyu = float(df_mip[df_mip["label"] == "central_13pp"]["delta_vab_total_muyu2016"].iloc[0])

        tc_2022 = obtener_valor(supuestos, "tc_promedio_2022")
        deflactor = obtener_valor(supuestos, "deflactor_2016_2022")
        if tc_2022 and deflactor:
            delta_vab_c_musd = delta_vab_c_muyu / deflactor / tc_2022
            multiplicador_va = delta_vab_c_musd / delta_d_c if delta_d_c != 0 else 1.0
        else:
            multiplicador_va = 1.0
            print("  ⚠ Sin conversión de unidades — multiplicador MIP = 1.0 (solo efecto directo)")
    else:
        multiplicador_va = 1.0
        print("  ⚠ Resultados MIP no disponibles — usando multiplicador = 1.0")

    print(f"  Multiplicador ΔVAB/ΔD usado en Monte Carlo: {multiplicador_va:.4f}")

    # PIB nominal
    pib_nominal_musd = obtener_valor(supuestos, "pib_nominal_musd")
    if pib_nominal_musd is None:
        print("  ✗ pib_nominal_musd no definido en supuestos.yaml.")
        print("    Completar el campo para poder expresar resultados como % del PIB.")
        sys.exit(1)

    # Distribuciones de parámetros
    phi_rango = supuestos["phi"]["rango"]
    phi_vals = (phi_rango[0], obtener_valor(supuestos, "phi"), phi_rango[1])

    eps_int_rango = supuestos["epsilon_intensivo"]["rango"]
    eps_int_vals = (eps_int_rango[0], obtener_valor(supuestos, "epsilon_intensivo"), eps_int_rango[1])

    eps_ext_rango = supuestos["epsilon_extensivo"]["rango"]
    eps_ext_vals = (eps_ext_rango[0], obtener_valor(supuestos, "epsilon_extensivo"), eps_ext_rango[1])

    # w: usar el rango del parámetro observado si existe; si no, el del fallback
    if obtener_valor(supuestos, "w") is not None and supuestos["w"].get("rango"):
        w_central = obtener_valor(supuestos, "w")
        w_rango = supuestos["w"]["rango"]
    else:
        w_central = obtener_valor(supuestos, "w_fallback")
        w_rango = supuestos["w_fallback"]["rango"]
    w_vals = (w_rango[0], w_central, w_rango[1])

    print(f"\n  Parámetros de la simulación:")
    print(f"    n_sims = {n_sims:,}, seed = {seed}")
    print(f"    B: [{b_min:.1f}, {b_central:.1f}, {b_max:.1f}] MUSD")
    print(f"    φ: {phi_vals}")
    print(f"    ε_int: {eps_int_vals}")
    print(f"    ε_ext: {eps_ext_vals}")
    print(f"    w: {w_vals}")

    # Simular
    print(f"\n  Simulando {n_sims:,} iteraciones (escenario central +{shock_c} pp)…")
    df_mc_c = simular_delta_pib_pct(
        n_sims=n_sims, seed=seed, shock_pp=shock_c,
        b_central=b_central, b_min=b_min, b_max=b_max,
        phi_vals=phi_vals, eps_int_vals=eps_int_vals,
        eps_ext_vals=eps_ext_vals, w_vals=w_vals,
        multiplicador_va=multiplicador_va,
        pib_nominal_musd=pib_nominal_musd,
        label=f"central_{shock_c}pp",
    )

    print(f"  Simulando {n_sims:,} iteraciones (variante +{shock_v} pp)…")
    df_mc_v = simular_delta_pib_pct(
        n_sims=n_sims, seed=seed, shock_pp=shock_v,
        b_central=b_central, b_min=b_min, b_max=b_max,
        phi_vals=phi_vals, eps_int_vals=eps_int_vals,
        eps_ext_vals=eps_ext_vals, w_vals=w_vals,
        multiplicador_va=multiplicador_va,
        pib_nominal_musd=pib_nominal_musd,
        label=f"variante_{shock_v}pp",
    )

    # Estadísticas
    est_c = calcular_estadisticas(df_mc_c, f"central_{shock_c}pp")
    est_v = calcular_estadisticas(df_mc_v, f"variante_{shock_v}pp")

    print("\n  ESTADÍSTICAS DE ΔPIB (%):")
    print(f"  {'Estadístico':<20} {'Central (+13pp)':>16} {'Variante (+22pp)':>16}")
    print("  " + "-" * 54)
    for clave in ["mediana", "media", "p10", "p90", "min", "max"]:
        print(f"  {clave:<20} {est_c[clave]:>16.4f} {est_v[clave]:>16.4f}")

    # Gráficos
    OUTPUTS_FIGURAS.mkdir(parents=True, exist_ok=True)
    grafico_histograma(df_mc_c, df_mc_v, OUTPUTS_FIGURAS / "histograma_montecarlo.png")

    grafico_tornado(
        shock_pp=shock_c, b_central=b_central, b_min=b_min, b_max=b_max,
        phi_vals=phi_vals, eps_int_vals=eps_int_vals,
        eps_ext_vals=eps_ext_vals, w_vals=w_vals,
        multiplicador_va=multiplicador_va, pib_nominal_musd=pib_nominal_musd,
        ruta_salida=OUTPUTS_FIGURAS / "tornado_central.png",
        label=f"Escenario central (+{shock_c} pp)",
    )

    grafico_tornado(
        shock_pp=shock_v, b_central=b_central, b_min=b_min, b_max=b_max,
        phi_vals=phi_vals, eps_int_vals=eps_int_vals,
        eps_ext_vals=eps_ext_vals, w_vals=w_vals,
        multiplicador_va=multiplicador_va, pib_nominal_musd=pib_nominal_musd,
        ruta_salida=OUTPUTS_FIGURAS / "tornado_variante.png",
        label=f"Variante (+{shock_v} pp)",
    )

    # Guardar
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)

    df_mc_c.to_csv(DATA_PROCESSED / "montecarlo_central.csv", index=False, encoding="utf-8")
    df_mc_v.to_csv(DATA_PROCESSED / "montecarlo_variante.csv", index=False, encoding="utf-8")

    df_est = pd.DataFrame([est_c, est_v])
    df_est.to_csv(OUTPUTS_TABLAS / "tabla_montecarlo_estadisticas.csv", index=False, encoding="utf-8")
    print(f"\n  ✓ Guardados: montecarlo_central.csv, montecarlo_variante.csv, estadísticas.")

    checkpoint(
        "Etapa 06 — Monte Carlo",
        [
            f"Mediana ΔPIB central: {est_c['mediana']:.4f}%  [P10={est_c['p10']:.4f}%, P90={est_c['p90']:.4f}%]",
            f"Mediana ΔPIB variante: {est_v['mediana']:.4f}%  [P10={est_v['p10']:.4f}%, P90={est_v['p90']:.4f}%]",
            f"Parámetro con mayor influencia: ver gráfico tornado.",
            "Validar: ¿la distribución es unimodal y razonable? "
            "¿El rango P10–P90 es informativo (no demasiado amplio)?",
        ]
    )

    return df_est


if __name__ == "__main__":
    main()
