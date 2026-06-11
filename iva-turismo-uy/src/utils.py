"""
Utilidades compartidas por todos los módulos del pipeline.
"""
import hashlib
import os
import pathlib
import re
import sys
import datetime
import yaml

ROOT = pathlib.Path(__file__).parent.parent
CONFIG_PATH = ROOT / "config" / "supuestos.yaml"
DATA_RAW = ROOT / "data" / "raw"
DATA_MANUAL = ROOT / "data" / "manual"
DATA_PROCESSED = ROOT / "data" / "processed"
OUTPUTS_TABLAS = ROOT / "outputs" / "tablas"
OUTPUTS_FIGURAS = ROOT / "outputs" / "figuras"


def cargar_supuestos() -> dict:
    """Carga el archivo YAML de supuestos y devuelve un dict plano de valores."""
    with open(CONFIG_PATH, encoding="utf-8") as f:
        raw = yaml.safe_load(f)
    return raw


def obtener_valor(supuestos: dict, clave: str):
    """
    Devuelve el 'valor' de un parámetro o el parámetro directo si no es un dict.
    Levanta KeyError si la clave no existe.
    """
    entrada = supuestos[clave]
    if isinstance(entrada, dict):
        return entrada.get("valor")
    return entrada


def sha256_archivo(ruta: pathlib.Path) -> str:
    """Calcula el hash SHA256 de un archivo."""
    h = hashlib.sha256()
    with open(ruta, "rb") as f:
        for bloque in iter(lambda: f.read(65536), b""):
            h.update(bloque)
    return h.hexdigest()


def buscar_archivo(nombre_preferente: str, nombres_alternativos: list = None) -> pathlib.Path | None:
    """
    Busca un archivo primero en data/manual/ (prioridad) y luego en data/raw/.
    Acepta coincidencia exacta o parcial (sin extensión).
    Devuelve el Path si lo encuentra, None si no.
    """
    candidatos = [nombre_preferente] + (nombres_alternativos or [])
    for directorio in [DATA_MANUAL, DATA_RAW]:
        for nombre in candidatos:
            # coincidencia exacta
            ruta = directorio / nombre
            if ruta.exists():
                return ruta
            # coincidencia parcial (el archivo contiene el nombre sin extensión)
            stem = pathlib.Path(nombre).stem
            for archivo in directorio.iterdir():
                if stem.lower() in archivo.name.lower():
                    return archivo
    return None


def separador(titulo: str):
    """Imprime un separador visual en consola."""
    print("\n" + "=" * 70)
    print(f"  {titulo}")
    print("=" * 70)


def checkpoint(etapa: str, resumen: list[str]):
    """
    Imprime el resumen de una etapa y espera confirmación del usuario.
    En modo no interactivo (CI/test), continúa automáticamente.
    """
    separador(f"CHECKPOINT — {etapa}")
    for linea in resumen:
        print(f"  • {linea}")
    print()
    if sys.stdin.isatty():
        respuesta = input("  ¿Continuar con la siguiente etapa? [s/N]: ").strip().lower()
        if respuesta not in ("s", "si", "sí", "y", "yes"):
            print("  Pipeline pausado por el usuario.")
            sys.exit(0)
    else:
        print("  (modo no interactivo — continuando automáticamente)")


def actualizar_data_sources(id_fuente: str, url: str, ruta: pathlib.Path, metodo: str):
    """Actualiza DATA_SOURCES.md con la fecha real y el hash del archivo descargado."""
    ds_path = ROOT / "DATA_SOURCES.md"
    fecha = datetime.date.today().isoformat()
    hash_val = sha256_archivo(ruta)
    texto = ds_path.read_text(encoding="utf-8")
    patron = rf"(\| {re.escape(id_fuente)} \|[^|]+\|[^|]+\|) pendiente (\|) pendiente (\|)"
    reemplazo = rf"\g<1> {fecha} \g<2> {hash_val[:16]}… \g<3>"
    texto_nuevo = re.sub(patron, reemplazo, texto)
    ds_path.write_text(texto_nuevo, encoding="utf-8")
