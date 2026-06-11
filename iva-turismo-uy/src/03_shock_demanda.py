"""
Etapa 03 — Shock de precio y respuesta de demanda.

Fórmulas (documentadas en README y en el informe final):
  Δp = (1 − φ) × shock_pp / 100
  g_ext = ε_ext × (Δp × w)           ← margen extensivo (todo el gasto turístico)
  g_int = ε_int × Δp                  ← margen intensivo (solo gasto gastronómico)
  ΔD_gastro = B × [(1 + g_ext) × (1 + g_int) − 1]
  ΔD_rubro_j = gasto_j_temporada × g_ext   (para j ≠ gastronomía)

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


def calcular_shock(
    shock_pp: float,
    phi: float,
    b_central: float,
    w: float,
    epsilon_ext: float,
    epsilon_int: float,
    label: str = "escenario_central",
) -> dict:
    """
    Calcula la respuesta de demanda para un shock dado.

    Parámetros
    ----------
    shock_pp    : puntos porcentuales de IVA adicional (ej. 13)
    phi         : fracción del shock absorbida por productores (ej. 0.15)
    b_central   : base afectada en MUSD (gasto gastronómico de no residentes en temporada)
    w           : peso de gastronomía en el gasto total del viaje
    epsilon_ext : elasticidad extensiva (número negativo)
    epsilon_int : elasticidad intensiva (número negativo)
    label       : etiqueta del escenario

    Retorna dict con los componentes del cambio en demanda.
    """
    assert epsilon_ext <= 0, "epsilon_ext debe ser negativo (convención)"
    assert epsilon_int <= 0, "epsilon_int debe ser negativo (convención)"
    assert 0 <= phi < 1, "phi debe estar en [0, 1)"
    assert 0 < w <= 1, "w debe estar en (0, 1]"

    # Fracción del shock que llega al precio al turista (neto de absorción de margen)
    delta_p = (1 - phi) * (shock_pp / 100)

    # Margen extensivo: caída en llegadas/noches (afecta TODO el gasto turístico)
    g_ext = epsilon_ext * (delta_p * w)

    # Margen intensivo: caída del gasto gastronómico condicional al viaje
    g_int = epsilon_int * delta_p

    # Cambio neto en gasto gastronómico afectado (B ya es temporada, no anualizar)
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


def calcular_shock_por_rubros(
    shock_central: dict,
    ruta_etr: pathlib.Path | None,
) -> pd.DataFrame:
    """
    Calcula la caída de demanda (solo margen extensivo) para rubros no gastronómicos.
    Si no hay datos ETR desagregados, devuelve un DataFrame vacío con nota.
    """
    g_ext = shock_central["g_ext"]
    rubros_fallback = {
        "Alimentacion": {"gasto_temporada_musd": shock_central["b_central_musd"], "tipo": "gastronómico"},
        "Alojamiento":  {"gasto_temporada_musd": None, "tipo": "otros"},
        "Transporte":   {"gasto_temporada_musd": None, "tipo": "otros"},
        "Recreacion":   {"gasto_temporada_musd": None, "tipo": "otros"},
        "Compras":      {"gasto_temporada_musd": None, "tipo": "otros"},
    }

    if ruta_etr is None:
        print("  ⚠ ETR no disponible: los rubros no gastronómicos no tienen valores de gasto.")
        print("    El vector de shock Δf solo incluirá la gastronomía.")
        filas = []
        for rubro, datos in rubros_fallback.items():
            filas.append({
                "rubro": rubro,
                "gasto_temporada_musd": datos["gasto_temporada_musd"],
                "delta_d_musd": (
                    datos["gasto_temporada_musd"] * g_ext
                    if datos["gasto_temporada_musd"] is not None and datos["tipo"] != "gastronómico"
                    else None
                ),
                "tipo": datos["tipo"],
                "fuente": "no disponible — ETR faltante",
            })
        return pd.DataFrame(filas)

    # Si la ETR está disponible, leer y calcular (lógica simplificada)
    # La estructura detallada depende del formato real del archivo
    print(f"  → Leyendo rubros de la ETR: {ruta_etr.name}")
    # Implementación completa después de conocer estructura real del archivo
    return pd.DataFrame()


def main():
    separador("ETAPA 03 — Shock de precio y respuesta de demanda")

    supuestos = cargar_supuestos()

    # Leer B central desde archivo procesado
    ruta_b = DATA_PROCESSED / "base_afectada.csv"
    if not ruta_b.exists():
        print("  ✗ No se encontró data/processed/base_afectada.csv")
        print("    Ejecutar primero: python src/02_base_afectada.py")
        sys.exit(1)

    df_b = pd.read_csv(ruta_b)
    fila_central = df_b[df_b["via"] == "CENTRAL (elegida)"].iloc[0]
    b_central = fila_central["b_musd"]

    if pd.isna(b_central) or b_central is None:
        print("  ✗ B central no está definida. Completar la Etapa 02 primero.")
        sys.exit(1)

    b_central = float(b_central)
    print(f"\n  B central cargada: {b_central:.1f} MUSD")

    # Leer parámetros
    phi = obtener_valor(supuestos, "phi")
    eps_int = obtener_valor(supuestos, "epsilon_intensivo")
    eps_ext = obtener_valor(supuestos, "epsilon_extensivo")

    # w: preferir calculado de ETR, luego fallback
    w = obtener_valor(supuestos, "w")
    if w is None:
        w = obtener_valor(supuestos, "w_fallback")
        print(f"  ⚠ Usando w_fallback = {w} (ETR no disponible)")
    else:
        print(f"  w calculado de ETR = {w:.3f}")

    shock_central_pp = obtener_valor(supuestos, "shock_pp")
    shock_variante_pp = obtener_valor(supuestos, "shock_pp_variante")

    # Calcular escenario central (+13 pp)
    print(f"\n  Calculando escenario central: shock = +{shock_central_pp} pp")
    sc = calcular_shock(
        shock_pp=shock_central_pp, phi=phi, b_central=b_central,
        w=w, epsilon_ext=eps_ext, epsilon_int=eps_int,
        label="central_13pp"
    )

    # Calcular variante (+22 pp)
    print(f"  Calculando variante: shock = +{shock_variante_pp} pp")
    sv = calcular_shock(
        shock_pp=shock_variante_pp, phi=phi, b_central=b_central,
        w=w, epsilon_ext=eps_ext, epsilon_int=eps_int,
        label="variante_22pp"
    )

    # Mostrar resultados
    print("\n  RESULTADOS DEL SHOCK:")
    print(f"  {'Componente':<40} {'Central (+13pp)':>16} {'Variante (+22pp)':>16}")
    print("  " + "-" * 74)
    for clave in ["delta_p", "g_ext", "g_int", "delta_d_gastro_musd", "delta_d_gastro_pct_de_b"]:
        print(f"  {clave:<40} {sc[clave]:>16.4f} {sv[clave]:>16.4f}")

    # Rubros ETR para vector Δf
    ruta_etr = buscar_archivo("mintur_etr_agregados.csv", ["mintur_etr_agregados.xlsx"])
    df_rubros = calcular_shock_por_rubros(sc, ruta_etr)

    # Guardar resultados
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)

    df_shocks = pd.DataFrame([sc, sv])
    df_shocks.to_csv(DATA_PROCESSED / "shocks.csv", index=False, encoding="utf-8")
    df_shocks.to_csv(OUTPUTS_TABLAS / "tabla_shocks.csv", index=False, encoding="utf-8")
    print(f"\n  ✓ Guardado: data/processed/shocks.csv")

    if not df_rubros.empty:
        df_rubros.to_csv(DATA_PROCESSED / "shock_por_rubros.csv", index=False, encoding="utf-8")
        print(f"  ✓ Guardado: data/processed/shock_por_rubros.csv")

    checkpoint(
        "Etapa 03 — Shock y demanda",
        [
            f"Escenario central (+13 pp): ΔD_gastro = {sc['delta_d_gastro_musd']:.1f} MUSD "
            f"({sc['delta_d_gastro_pct_de_b']:.1f}% de B)",
            f"Variante (+22 pp): ΔD_gastro = {sv['delta_d_gastro_musd']:.1f} MUSD "
            f"({sv['delta_d_gastro_pct_de_b']:.1f}% de B)",
            f"Supuestos clave: φ={phi}, ε_int={eps_int}, ε_ext={eps_ext}, w={w:.3f}",
            "Validar: ¿las magnitudes son razonables dado B? "
            "¿El signo de ΔD_gastro es negativo (caída de demanda)?",
            "Revisar mapeo de rubros ETR→MIP en config/mapeo_rubros_mip.csv antes de Etapa 04.",
        ]
    )

    return {"sc": sc, "sv": sv}


if __name__ == "__main__":
    main()
