"""
Etapa 02 — Estimación de la base afectada (B).

Triangulación por tres vías:
  1. Vía DGI:   B = GT_A34 / 0,22  (preferente)
     El informe DGI (pág. 35, fila A_34) reporta el GT de la "Reducción de 22
     puntos en la tasa del IVA... gastronómicos... turistas no residentes"
     (Ley 17.934 + decretos). Como el GT corresponde a la exoneración TOTAL
     (22 pts), la base se despeja con tasa 0,22. El GT es ANUAL, por lo que
     se ajusta a temporada con la fracción estacional calculada de la ETR.
  2. Vía ETR:   gasto alimentación temporada (microdatos, ventana exacta
     15-nov a 30-abr) × share_restaurantes × share_tarjeta_exterior
  3. Vía BCU pagos: cota superior de consistencia (si disponible)

Salida: data/processed/base_afectada.csv
"""
import sys
import re
import pathlib
import pandas as pd
import numpy as np
import yaml

sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_PROCESSED, OUTPUTS_TABLAS, CONFIG_PATH,
    cargar_supuestos, obtener_valor, buscar_archivo, separador, checkpoint
)

import pdfplumber

# Año calendario de referencia para alinear GT (DGI) con la ETR
ANIO_REF = 2022
# Temporada de referencia para la vía ETR: 15-nov-2022 a 30-abr-2023
TEMPORADA_INICIO = "2022-11-15"
TEMPORADA_FIN = "2023-04-30"


# =============================================================================
# VÍA 2 (se calcula primero porque la vía DGI necesita la fracción estacional)
# =============================================================================
def calcular_b_etr(supuestos: dict) -> dict:
    """
    Calcula B desde los MICRODATOS de la ETR (cada fila = un viaje encuestado).

    Estructura observada del archivo:
      - Hoja 'Datos': FechaIngreso, Residencia, GastoTotal, GastoAlimentacion,
        ..., Coef (ponderador de grupo), CoefTot (= Coef × Gente).
      - Gasto expandido = Gasto del grupo × Coef.
      - Gastos en USD corrientes.

    Ventaja sobre agregados trimestrales: la temporada 15-nov/30-abr se
    calcula con fechas exactas, sin la aproximación 0,5×T4 + T1 + 0,33×T2.
    """
    resultado = {
        "gasto_alimentacion_temporada_musd": None,
        "b_etr_musd": None,
        "w_calculado": None,
        "share_temporada": None,
        "fuente_ok": False,
        "notas": [],
    }

    ruta = buscar_archivo("mintur_etr_agregados.xlsx", ["mintur_etr_agregados.csv"])
    if ruta is None:
        resultado["notas"].append(
            "FALTA ARCHIVO ETR — colocar en data/manual/mintur_etr_agregados.xlsx"
        )
        return resultado

    print(f"  → Leyendo microdatos ETR: {ruta.name}")
    df = pd.read_excel(ruta, sheet_name="Datos")
    print(f"     Viajes encuestados: {len(df):,}")

    df["fi"] = pd.to_datetime(df["FechaIngreso"])
    # Gasto expandido del grupo: gasto del viaje × ponderador de grupo
    df["gasto_exp"] = df["GastoTotal"] * df["Coef"]
    df["alim_exp"] = df["GastoAlimentacion"] * df["Coef"]

    # --- Temporada alta exacta: 15-nov-2022 a 30-abr-2023 ---
    m_temp = (df["fi"] >= TEMPORADA_INICIO) & (df["fi"] <= TEMPORADA_FIN)
    d_temp = df[m_temp]
    gasto_total_temp = d_temp["gasto_exp"].sum() / 1e6   # MUSD
    alim_temp = d_temp["alim_exp"].sum() / 1e6           # MUSD

    print(f"     Temporada {TEMPORADA_INICIO} a {TEMPORADA_FIN}: "
          f"{len(d_temp):,} obs, gasto total {gasto_total_temp:.1f} MUSD, "
          f"alimentación {alim_temp:.1f} MUSD")

    # --- w: peso de alimentación en el gasto total del viaje (temporada) ---
    w = alim_temp / gasto_total_temp if gasto_total_temp > 0 else None
    print(f"     w (alimentación/total, temporada) = {w:.3f}  [DATO OBSERVADO]")

    # --- Fracción estacional del gasto en alimentación (año calendario de ref) ---
    # Sirve para convertir el GT anual de la DGI a base de temporada.
    m_temp_anio = (
        ((df["fi"] >= f"{ANIO_REF}-01-01") & (df["fi"] <= f"{ANIO_REF}-04-30"))
        | ((df["fi"] >= f"{ANIO_REF}-11-15") & (df["fi"] <= f"{ANIO_REF}-12-31"))
    )
    m_anio = df["fi"].dt.year == ANIO_REF
    alim_temp_anio = df[m_temp_anio]["alim_exp"].sum()
    alim_anio = df[m_anio]["alim_exp"].sum()
    share_temporada = alim_temp_anio / alim_anio if alim_anio > 0 else None
    print(f"     Fracción estacional del gasto alimentación {ANIO_REF}: "
          f"{share_temporada:.3f}  [DATO OBSERVADO]")

    # --- B vía ETR: aplicar shares (SUPUESTOS declarados en config) ---
    share_rest = obtener_valor(supuestos, "share_restaurantes")
    share_tarj = obtener_valor(supuestos, "share_tarjeta_exterior")
    b_etr = alim_temp * share_rest * share_tarj

    resultado.update({
        "gasto_alimentacion_temporada_musd": round(alim_temp, 1),
        "b_etr_musd": round(b_etr, 1),
        "w_calculado": round(float(w), 3),
        "share_temporada": round(float(share_temporada), 3),
        "fuente_ok": True,
    })
    resultado["notas"].append(
        f"B_etr = {alim_temp:.1f} (alimentación temporada, OBSERVADO) "
        f"× {share_rest} (share_restaurantes, SUPUESTO) "
        f"× {share_tarj} (share_tarjeta_exterior, SUPUESTO) = {b_etr:.1f} MUSD"
    )
    return resultado


# =============================================================================
# VÍA 1 — DGI: Gasto Tributario (fila A_34 del informe)
# =============================================================================
def extraer_gt_dgi(supuestos: dict, share_temporada: float | None) -> dict:
    """
    Extrae el GT de la fila A_34 del informe DGI ("Reducción del IVA con
    tarjetas de crédito. Turistas no residentes", Ley 17.934).

    Estructura observada del PDF (pág. 35): la fila A_34 tiene 4 cifras en UYU
    corrientes: GT 2021, GT 2022, GT 2023, GT 2024 (proyección).

    Interpretación documentada: el GT corresponde a la reducción de 22 PUNTOS
    (exoneración total para no residentes con tarjeta del exterior), por lo que
    B_anual = GT / 0,22. Como el GT es anual y el escenario afecta solo la
    temporada, se ajusta: B_temporada = B_anual × share_temporada (de la ETR).

    ADVERTENCIA declarada: la medida A_34 incluye también arrendamiento de
    vehículos sin chofer e intermediación en arrendamiento de inmuebles
    (sesgo al alza en B). Se documenta como limitación.
    """
    resultado = {
        "gt_2022_uyu": None,
        "b_dgi_anual_musd": None,
        "b_dgi_temporada_musd": None,
        "fuente_ok": False,
        "notas": [],
    }

    ruta = buscar_archivo("dgi_gasto_tributario_2019_2022.pdf")
    if ruta is None:
        resultado["notas"].append("FALTA ARCHIVO: data/manual/dgi_gasto_tributario_2019_2022.pdf")
        return resultado

    print(f"  → Leyendo PDF DGI: {ruta.name}")
    linea_a34 = None
    with pdfplumber.open(ruta) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text() or ""
            for linea in texto.split("\n"):
                if "A_34" in linea:
                    linea_a34 = linea
                    break
            if linea_a34:
                break

    if linea_a34 is None:
        resultado["notas"].append(
            "No se encontró la fila A_34 en el PDF. Verificar la edición del informe."
        )
        return resultado

    print(f"     Fila A_34 encontrada: {linea_a34[:120]}…")

    # Extraer las cifras grandes (formato 999.999.999, en UYU)
    numeros = re.findall(r"\b(\d{1,3}(?:\.\d{3}){2,})\b", linea_a34)
    valores = [int(n.replace(".", "")) for n in numeros]
    print(f"     Cifras extraídas (UYU): {valores}")

    if len(valores) != 4:
        resultado["notas"].append(
            f"Se esperaban 4 cifras (GT 2021–2024) y se encontraron {len(valores)}. "
            "Verificar la extracción manualmente."
        )
        return resultado

    # Columnas: GT 2021, GT 2022, GT 2023, GT 2024 (proyección)
    gt_2022_uyu = valores[1]
    print(f"     GT {ANIO_REF} (A_34): {gt_2022_uyu:,} UYU corrientes  [DATO OBSERVADO]")

    # Despeje de B: el GT corresponde a la exoneración total (22 pts)
    tc = obtener_valor(supuestos, "tc_promedio_2022")
    b_anual_musd = gt_2022_uyu / 0.22 / tc / 1e6

    resultado["gt_2022_uyu"] = gt_2022_uyu
    resultado["b_dgi_anual_musd"] = round(b_anual_musd, 1)
    resultado["notas"].append(
        f"B_anual = GT/0,22 = {gt_2022_uyu:,} / 0,22 / TC {tc} = {b_anual_musd:.1f} MUSD (anual)"
    )

    if share_temporada is not None:
        b_temporada = b_anual_musd * share_temporada
        resultado["b_dgi_temporada_musd"] = round(b_temporada, 1)
        resultado["notas"].append(
            f"B_temporada = {b_anual_musd:.1f} × {share_temporada:.3f} "
            f"(fracción estacional, OBSERVADA en ETR) = {b_temporada:.1f} MUSD"
        )
    else:
        resultado["notas"].append(
            "Sin fracción estacional de la ETR — B_dgi queda en términos anuales."
        )

    resultado["notas"].append(
        "ADVERTENCIA: A_34 incluye también arrendamiento de vehículos sin chofer "
        "e intermediación inmobiliaria (sesgo al alza en B_dgi)."
    )
    resultado["fuente_ok"] = True
    return resultado


# =============================================================================
# TRIANGULACIÓN y VALIDACIÓN
# =============================================================================
def triangular_y_validar(b_dgi, b_etr, b_bcu) -> dict:
    """
    Tabla comparativa y elección de B central.
    Si las fuentes difieren en más de 3×, frena y reporta (validación 1).
    """
    fuentes = {
        "DGI temporada (preferente)": b_dgi,
        "ETR temporada": b_etr,
        "BCU Pagos (cota sup.)": b_bcu,
    }
    disponibles = {k: v for k, v in fuentes.items() if v is not None}

    print("\n  TABLA COMPARATIVA DE ESTIMACIONES DE B (millones USD, temporada):")
    print(f"  {'Vía':<32} {'B (MUSD)':>12}")
    print("  " + "-" * 46)
    for k, v in fuentes.items():
        val_str = f"{v:>10.1f}" if v is not None else "    N/D     "
        print(f"  {k:<32} {val_str}")

    if not disponibles:
        return {"b_central": None, "alerta": "Sin estimaciones disponibles."}

    valores = list(disponibles.values())
    alerta = None
    if len(valores) >= 2:
        razon = max(valores) / min(valores)
        print(f"\n  Razón máx/mín entre fuentes: {razon:.2f}×")
        if razon > 3.0:
            alerta = (
                f"Las estimaciones difieren en {razon:.2f}× (>3×). "
                f"Mín: {min(valores):.1f}, Máx: {max(valores):.1f}. "
                "Requiere decisión del usuario antes de continuar."
            )
            print(f"  ⚠ {alerta}")

    if b_dgi is not None:
        b_central, fuente_central = b_dgi, "DGI"
    elif b_etr is not None:
        b_central, fuente_central = b_etr, "ETR"
    else:
        b_central, fuente_central = valores[0], "única disponible"

    return {
        "b_central": b_central,
        "b_min": min(valores),
        "b_max": max(valores),
        "fuente_central": fuente_central,
        "alerta": alerta,
    }


def actualizar_supuestos(res_etr: dict, res_dgi: dict):
    """Persiste en supuestos.yaml los valores observados (w, B por vía)."""
    with open(CONFIG_PATH, encoding="utf-8") as f:
        supuestos = yaml.safe_load(f)

    if res_etr.get("w_calculado") is not None:
        supuestos["w"]["valor"] = res_etr["w_calculado"]
        supuestos["w"]["fuente"] = (
            f"MINTUR ETR microdatos — temporada {TEMPORADA_INICIO}/{TEMPORADA_FIN} (OBSERVADO)"
        )
    if res_dgi.get("b_dgi_temporada_musd") is not None:
        supuestos["B_dgi"]["valor"] = res_dgi["b_dgi_temporada_musd"]
    if res_etr.get("b_etr_musd") is not None:
        supuestos["B_etr"]["valor"] = float(res_etr["b_etr_musd"])

    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(supuestos, f, allow_unicode=True, sort_keys=False, default_flow_style=False)
    print("  ✓ supuestos.yaml actualizado (w observado, B_dgi, B_etr).")


def main():
    separador("ETAPA 02 — Base afectada (B)")
    supuestos = cargar_supuestos()

    # --- VÍA 2 primero (provee la fracción estacional para la vía DGI) ---
    print("\n[VÍA 2] MINTUR — ETR (microdatos)")
    res_etr = calcular_b_etr(supuestos)
    for nota in res_etr["notas"]:
        print(f"  → {nota}")

    # --- VÍA 1: DGI ---
    print("\n[VÍA 1] DGI — Gasto Tributario (fila A_34)")
    res_dgi = extraer_gt_dgi(supuestos, res_etr.get("share_temporada"))
    for nota in res_dgi["notas"]:
        print(f"  → {nota}")

    # --- VÍA 3: BCU Pagos ---
    print("\n[VÍA 3] BCU — Sistema de Pagos")
    b_bcu = obtener_valor(supuestos, "B_bcu")
    print(f"  → {'Valor manual: ' + str(b_bcu) + ' MUSD' if b_bcu else 'No disponible (no bloqueante).'}")

    b_dgi = res_dgi.get("b_dgi_temporada_musd") or res_dgi.get("b_dgi_anual_musd")
    b_etr = res_etr.get("b_etr_musd")

    resultado = triangular_y_validar(b_dgi, b_etr, b_bcu)
    resultado.update({"b_dgi": b_dgi, "b_etr": b_etr, "b_bcu": b_bcu})

    actualizar_supuestos(res_etr, res_dgi)

    # Guardar tabla
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)
    OUTPUTS_TABLAS.mkdir(parents=True, exist_ok=True)
    df_b = pd.DataFrame([
        {"via": "DGI temporada (preferente)", "b_musd": b_dgi,
         "estado": "observado (GT A_34 pdf p.35) × fracción estacional ETR"},
        {"via": "DGI anual (referencia)", "b_musd": res_dgi.get("b_dgi_anual_musd"),
         "estado": "observado (GT A_34 / 0,22)"},
        {"via": "ETR temporada", "b_musd": b_etr,
         "estado": "observado × supuestos (share_restaurantes, share_tarjeta)"},
        {"via": "BCU Pagos (cota sup.)", "b_musd": b_bcu, "estado": "no disponible"},
        {"via": "CENTRAL (elegida)", "b_musd": resultado.get("b_central"),
         "estado": f"vía {resultado.get('fuente_central', 'N/D')}"},
    ])
    df_b.to_csv(DATA_PROCESSED / "base_afectada.csv", index=False, encoding="utf-8")
    df_b.to_csv(OUTPUTS_TABLAS / "tabla_base_afectada.csv", index=False, encoding="utf-8")

    # Guardar también el rango para Monte Carlo
    pd.DataFrame([{
        "b_central": resultado.get("b_central"),
        "b_min": resultado.get("b_min"),
        "b_max": resultado.get("b_max"),
        "alerta": resultado.get("alerta"),
    }]).to_csv(DATA_PROCESSED / "rango_b.csv", index=False, encoding="utf-8")
    print(f"\n  ✓ Guardado: data/processed/base_afectada.csv, rango_b.csv")

    if resultado.get("alerta"):
        if obtener_valor(supuestos, "b_discrepancia_aprobada"):
            print("\n  ⚠ Discrepancia >3× APROBADA por el usuario (b_discrepancia_aprobada=true).")
            print("    Continuando con B central = vía DGI y rango completo en Monte Carlo.")
        else:
            print(f"\n  ⚠ PIPELINE PAUSADO PARA DECISIÓN DEL USUARIO:\n    {resultado['alerta']}")
            sys.exit(1)

    checkpoint(
        "Etapa 02 — Base Afectada",
        [
            f"B vía DGI (temporada): {b_dgi} MUSD",
            f"B vía ETR (temporada): {b_etr} MUSD",
            f"B CENTRAL: {resultado.get('b_central')} MUSD (vía {resultado.get('fuente_central')})",
            f"w observado de la ETR: {res_etr.get('w_calculado')}",
            "Validar interpretación del GT A_34 y los shares supuestos de la vía ETR.",
        ]
    )
    return resultado


if __name__ == "__main__":
    main()
