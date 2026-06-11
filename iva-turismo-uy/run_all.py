"""
run_all.py — Orquestador del pipeline iva-turismo-uy.

Ejecuta todas las etapas en secuencia.
En modo no interactivo (sin terminal TTY), los checkpoints continúan automáticamente.

Uso:
    python run_all.py            # pipeline completo
    python run_all.py --desde 3  # retomar desde la etapa 3
    python run_all.py --solo 6   # correr solo la etapa 6
"""
import sys
import argparse
import pathlib
import importlib
import traceback

# Agregar src/ al path
ROOT = pathlib.Path(__file__).parent
sys.path.insert(0, str(ROOT / "src"))

ETAPAS = [
    ("01", "01_descarga_datos",   "Descarga automática de datos"),
    ("02", "02_base_afectada",    "Estimación de la base afectada (B)"),
    ("03", "03_shock_demanda",    "Shock de precio y respuesta de demanda"),
    ("04", "04_modelo_mip",       "Modelo Insumo-Producto de Leontief"),
    ("05", "05_fiscal",           "Análisis fiscal"),
    ("06", "06_montecarlo",       "Simulación Monte Carlo"),
    ("07", "07_informe",          "Generación del informe final"),
]


def separador_principal(texto):
    print("\n" + "█" * 70)
    print(f"  {texto}")
    print("█" * 70)


def main():
    parser = argparse.ArgumentParser(description="Pipeline iva-turismo-uy")
    parser.add_argument("--desde", type=int, default=1, help="Etapa de inicio (1–7)")
    parser.add_argument("--solo", type=int, default=None, help="Ejecutar solo esta etapa")
    args = parser.parse_args()

    print("\n" + "=" * 70)
    print("  PIPELINE: Impacto en PIB — Exoneración IVA Gastronomía Turística UY")
    print("=" * 70)
    print("  Ejecutar con: python run_all.py")
    print("  Repositorio: iva-turismo-uy/")
    print("  Nota: los checkpoints esperan confirmación en modo interactivo.")

    etapas_a_correr = []
    for num, modulo, descripcion in ETAPAS:
        n = int(num)
        if args.solo is not None:
            if n == args.solo:
                etapas_a_correr.append((num, modulo, descripcion))
        elif n >= args.desde:
            etapas_a_correr.append((num, modulo, descripcion))

    if not etapas_a_correr:
        print("  Sin etapas para ejecutar.")
        return

    print(f"\n  Etapas a ejecutar: {[e[0] for e in etapas_a_correr]}\n")

    for num, modulo, descripcion in etapas_a_correr:
        separador_principal(f"ETAPA {num} — {descripcion}")
        try:
            mod = importlib.import_module(modulo)
            mod.main()
        except SystemExit as e:
            if e.code != 0:
                print(f"\n  Pipeline detenido en etapa {num} (código {e.code}).")
                print("  Resolver los problemas indicados y reiniciar desde esta etapa:")
                print(f"    python run_all.py --desde {num}")
                sys.exit(e.code)
            else:
                print(f"\n  Etapa {num} completada (exit 0).")
                break
        except Exception:
            print(f"\n  ERROR INESPERADO en etapa {num}:")
            traceback.print_exc()
            print(f"\n  Reiniciar desde esta etapa: python run_all.py --desde {num}")
            sys.exit(1)

    print("\n" + "=" * 70)
    print("  Pipeline completado.")
    print("  Outputs en: outputs/tablas/, outputs/figuras/, outputs/informe_final.docx")
    print("=" * 70 + "\n")


if __name__ == "__main__":
    main()
