"""
Etapa 01 — Descarga automática de datos.

Estrategia:
  1. Intentar descargar cada fuente con requests (user-agent de navegador).
  2. Si falla (403, timeout, Cloudflare), imprimir instrucción de descarga manual.
  3. Detectar archivos ya presentes en data/manual/ y usarlos con prioridad.
  4. Actualizar DATA_SOURCES.md con fecha y hash de cada archivo obtenido.
"""
import sys
import time
import pathlib

# Asegurar que src/ esté en el path para importar utils
sys.path.insert(0, str(pathlib.Path(__file__).parent))
from utils import (
    ROOT, DATA_RAW, DATA_MANUAL, DATA_PROCESSED,
    separador, actualizar_data_sources, buscar_archivo, sha256_archivo
)

# Importar requests solo aquí para aislar la dependencia de red
import requests

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
}
TIMEOUT = 30

FUENTES = [
    {
        "id": "DGI-GT",
        "descripcion": "DGI — Informe de Gasto Tributario 2019-2022 (PDF)",
        "url": (
            "https://www.gub.uy/direccion-general-impositiva/sites/"
            "direccion-general-impositiva/files/2023-12/"
            "Informe+de+Gasto+Tributario+2019+-+2022.pdf"
        ),
        "nombre_archivo": "dgi_gasto_tributario_2019_2022.pdf",
        "tipo": "pdf",
        "critico": True,
    },
    {
        "id": "MINTUR-ETR-AGR",
        "descripcion": "MINTUR — ETR Agregados trimestrales (catálogo datos abiertos)",
        "url": "https://catalogodatos.gub.uy/dataset/ministerio-de-turismo-turismo-receptivo",
        "nombre_archivo": "mintur_etr_agregados.csv",
        "tipo": "pagina_catalogo",
        "critico": True,
        "nota": (
            "Este es el catálogo; los CSV individuales están dentro. "
            "El script intentará encontrar el enlace de descarga directa."
        ),
    },
    {
        "id": "BCU-MIP",
        "descripcion": "BCU — Matriz Insumo-Producto 2016 (Excel, SharePoint)",
        "url": (
            "https://www.bcu.gub.uy/Estadisticas-e-Indicadores/"
            "Paginas/Matriz-Insumo-Producto.aspx"
        ),
        "nombre_archivo": "bcu_mip_2016.xlsx",
        "tipo": "sharepoint",
        "critico": True,
    },
    {
        "id": "BCU-CN",
        "descripcion": "BCU — Cuentas Nacionales: VAB y PIB nominal (Excel)",
        "url": (
            "https://www.bcu.gub.uy/Estadisticas-e-Indicadores/"
            "Paginas/Cuentas-Nacionales.aspx"
        ),
        "nombre_archivo": "bcu_cuentas_nacionales.xlsx",
        "tipo": "sharepoint",
        "critico": False,
    },
]

INSTRUCCIONES_MANUALES = {
    "DGI-GT": (
        "  URL:     https://www.gub.uy/direccion-general-impositiva/datos-y-estadisticas\n"
        "           (buscar 'Gasto Tributario' o edición más reciente)\n"
        "  Guardar: data/manual/dgi_gasto_tributario_2019_2022.pdf\n"
        "  Nota:    Si hay una edición más reciente (2023+), usarla en cambio."
    ),
    "MINTUR-ETR-AGR": (
        "  URL:     https://catalogodatos.gub.uy/dataset/ministerio-de-turismo-turismo-receptivo\n"
        "           Descargar el archivo CSV/XLSX con datos trimestrales de gasto por rubro.\n"
        "  Guardar: data/manual/mintur_etr_agregados.csv  (o .xlsx)\n"
        "  También: https://www.gub.uy/ministerio-turismo/datos-y-estadisticas/microdatos\n"
        "           Descargar microdatos del último año completo disponible.\n"
        "  Guardar: data/manual/mintur_etr_microdatos_XXXX.xlsx"
    ),
    "BCU-MIP": (
        "  URL:     https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Matriz-Insumo-Producto.aspx\n"
        "           (el SharePoint del BCU bloquea descargas automáticas)\n"
        "  Guardar: data/manual/bcu_mip_2016.xlsx\n"
        "  Notas:   Descargar el archivo Excel con la MIP 2016 completa.\n"
        "           Verificar que incluya: transacciones intermedias domésticas,\n"
        "           valor de producción bruta (VBP) y valor agregado bruto (VAB) por industria."
    ),
    "BCU-CN": (
        "  URL:     https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Cuentas-Nacionales.aspx\n"
        "  Guardar: data/manual/bcu_cuentas_nacionales.xlsx\n"
        "  Necesitamos: VAB de 'Alojamiento y servicios de comida' y PIB nominal anual."
    ),
}


def intentar_descarga(fuente: dict) -> pathlib.Path | None:
    """
    Intenta descargar un archivo. Devuelve el Path si tiene éxito, None si falla.
    Hace un solo reintento tras 3 segundos.
    """
    # Si ya está en data/manual/, usarlo directamente
    archivo_manual = buscar_archivo(fuente["nombre_archivo"])
    if archivo_manual and DATA_MANUAL in archivo_manual.parents:
        print(f"  ✓ Encontrado en data/manual/: {archivo_manual.name}")
        return archivo_manual

    # Tipos que sabemos que fallarán (SharePoint, páginas de catálogo)
    if fuente["tipo"] in ("sharepoint", "pagina_catalogo"):
        print(f"  ⚠ Tipo '{fuente['tipo']}' — descarga automática no aplica.")
        return None

    destino = DATA_RAW / fuente["nombre_archivo"]
    if destino.exists():
        print(f"  ✓ Ya existe en data/raw/: {destino.name}")
        return destino

    print(f"  → Intentando descargar: {fuente['url'][:70]}…")
    for intento in range(1, 3):
        try:
            resp = requests.get(fuente["url"], headers=HEADERS, timeout=TIMEOUT, stream=True)
            if resp.status_code == 200:
                with open(destino, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=65536):
                        f.write(chunk)
                print(f"  ✓ Descargado ({destino.stat().st_size // 1024} KB)")
                return destino
            else:
                print(f"  ✗ HTTP {resp.status_code} en intento {intento}")
        except Exception as e:
            print(f"  ✗ Error en intento {intento}: {e}")
        if intento == 1:
            time.sleep(3)

    return None


def main():
    separador("ETAPA 01 — Descarga de datos")

    DATA_RAW.mkdir(parents=True, exist_ok=True)
    DATA_MANUAL.mkdir(parents=True, exist_ok=True)
    DATA_PROCESSED.mkdir(parents=True, exist_ok=True)

    pendientes = []

    for fuente in FUENTES:
        print(f"\n[{fuente['id']}] {fuente['descripcion']}")
        ruta = intentar_descarga(fuente)

        if ruta:
            actualizar_data_sources(fuente["id"], fuente["url"], ruta, "automático")
            print(f"  SHA256: {sha256_archivo(ruta)[:32]}…")
        else:
            pendientes.append(fuente["id"])
            nivel = "CRÍTICO" if fuente["critico"] else "NO CRÍTICO"
            print(f"\n  {'!' * 60}")
            print(f"  [{nivel}] No se pudo descargar '{fuente['id']}'.")
            print(f"  INSTRUCCIONES DE DESCARGA MANUAL:")
            print(INSTRUCCIONES_MANUALES.get(fuente["id"], "  Sin instrucciones disponibles."))
            print(f"  {'!' * 60}")

    # Resumen final
    separador("RESUMEN — Etapa 01")
    criticos_pendientes = [
        fid for fid in pendientes
        if any(f["id"] == fid and f["critico"] for f in FUENTES)
    ]
    opcionales_pendientes = [fid for fid in pendientes if fid not in criticos_pendientes]

    if not pendientes:
        print("  ✓ Todas las fuentes descargadas correctamente.")
    else:
        if criticos_pendientes:
            print(f"  ✗ CRÍTICOS pendientes de descarga manual: {', '.join(criticos_pendientes)}")
        if opcionales_pendientes:
            print(f"  ⚠ Opcionales pendientes: {', '.join(opcionales_pendientes)}")

    print()
    print("  PRÓXIMOS PASOS:")
    print("  1. Descargar manualmente los archivos indicados arriba.")
    print("  2. Colocarlos en data/manual/ con los nombres indicados.")
    print("  3. Ejecutar nuevamente este script para verificar (o pasar a la etapa 02).")

    return criticos_pendientes


if __name__ == "__main__":
    pendientes = main()
    if pendientes:
        print(f"\n  Pipeline detenido: faltan archivos críticos → {pendientes}")
        sys.exit(1)
