# iva-turismo-uy

**Modelo para estimar el impacto sobre el PIB de Uruguay de eliminar la exoneración estacional de IVA a la gastronomía turística.**

Marco legal: Ley 17.934 (mod. Ley 20.212) + Decreto 434/023 prorrogado por Decreto 220/2025.

---

## Cómo reproducir los resultados desde cero

### 1. Requisitos previos
- Python 3.11 o superior
- Git

### 2. Clonar e instalar
```bash
git clone <URL_DEL_REPO>
cd iva-turismo-uy
python -m venv venv
source venv/bin/activate        # Linux/Mac
# venv\Scripts\activate         # Windows
pip install -r requirements.txt
```

### 3. Colocar los archivos de datos manuales

El pipeline intentará descargar automáticamente algunas fuentes. Para las que requieren descarga manual (especialmente el BCU), colocar los archivos en `data/manual/` con los nombres exactos indicados:

| Archivo en `data/manual/` | Fuente | URL |
|---|---|---|
| `dgi_gasto_tributario_2019_2022.pdf` | DGI — Gasto Tributario | https://www.gub.uy/direccion-general-impositiva/datos-y-estadisticas |
| `mintur_etr_agregados.csv` (o `.xlsx`) | MINTUR — ETR Agregados | https://catalogodatos.gub.uy/dataset/ministerio-de-turismo-turismo-receptivo |
| `mintur_etr_microdatos_XXXX.xlsx` | MINTUR — ETR Microdatos | https://www.gub.uy/ministerio-turismo/datos-y-estadisticas/microdatos |
| `bcu_mip_2016.xlsx` | BCU — MIP 2016 | https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Matriz-Insumo-Producto.aspx |
| `bcu_cuentas_nacionales.xlsx` | BCU — Cuentas Nacionales | https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Cuentas-Nacionales.aspx |

> **Nota sobre el BCU:** el SharePoint del BCU bloquea descargas automáticas. Descargar manualmente los Excel y colocarlos en `data/manual/`.

### 4. Completar parámetros de conversión de unidades

Antes de correr la Etapa 04, agregar los siguientes parámetros en `config/supuestos.yaml` con datos del BCU:

```yaml
tc_promedio_2022:
  valor: <TC promedio 2022 UYU/USD — fuente BCU>
  rango: null
  fuente: "BCU — Tipo de cambio promedio anual 2022"
  justificacion: "Conversión USD→UYU para alinear unidades con la MIP 2016"

deflactor_2016_2022:
  valor: <IPC_2016/IPC_2022 — fuente INE>
  rango: null
  fuente: "INE — Índice de Precios al Consumo"
  justificacion: "Deflactar valores nominales 2022 a precios constantes 2016 de la MIP"

pib_nominal_musd:
  valor: <PIB nominal 2022 en MUSD — fuente BCU>
  rango: null
  fuente: "BCU — Cuentas Nacionales"
  justificacion: "Denominador para expresar el impacto como % del PIB"
```

### 5. Ejecutar el pipeline

```bash
# Pipeline completo (con checkpoints interactivos entre etapas)
python run_all.py

# Retomar desde una etapa específica
python run_all.py --desde 3

# Ejecutar solo una etapa
python run_all.py --solo 4
```

El pipeline se detiene en cada checkpoint y pide confirmación antes de continuar. En modo no interactivo (CI/scripts), continúa automáticamente.

### 6. Outputs

```
outputs/
├── informe_final.docx          # Informe completo en Word
├── tablas/
│   ├── tabla_base_afectada.csv
│   ├── tabla_shocks.csv
│   ├── tabla_impacto_mip.csv
│   ├── tabla_impacto_industrias.csv
│   ├── tabla_fiscal.csv
│   └── tabla_montecarlo_estadisticas.csv
└── figuras/
    ├── histograma_montecarlo.png
    ├── tornado_central.png
    └── tornado_variante.png
```

---

## Estructura del repositorio

```
iva-turismo-uy/
├── README.md
├── requirements.txt
├── run_all.py                   # orquestador del pipeline
├── DATA_SOURCES.md              # registro de fuentes con hash SHA256
├── config/
│   ├── supuestos.yaml           # TODOS los parámetros del modelo
│   └── mapeo_rubros_mip.csv     # mapeo rubros ETR → industrias MIP
├── data/
│   ├── raw/                     # descargas automáticas
│   ├── manual/                  # archivos colocados manualmente
│   └── processed/               # resultados intermedios
├── src/
│   ├── utils.py                 # funciones compartidas
│   ├── 01_descarga_datos.py     # descargas automáticas + instrucciones manuales
│   ├── 02_base_afectada.py      # triangulación de B (DGI + ETR + BCU)
│   ├── 03_shock_demanda.py      # shock de precio y respuesta de demanda
│   ├── 04_modelo_mip.py         # modelo Leontief con MIP 2016
│   ├── 05_fiscal.py             # análisis fiscal
│   ├── 06_montecarlo.py         # Monte Carlo + gráficos
│   └── 07_informe.py            # generación del informe Word
├── outputs/
│   ├── tablas/
│   ├── figuras/
│   └── informe_final.docx
└── fase2_econometria/
    └── NOTES.md                 # diseño del event-study (no implementado)
```

---

## Escenarios del modelo

| Escenario | Shock | Descripción |
|---|---|---|
| Central | +13 pp | Cae el decreto estacional; queda el piso de 9 pts de la Ley 17.934. Precio: 100 → 113. |
| Variante | +22 pp | Eliminación total del régimen. Precio: 100 → 122. |

---

## Fórmulas clave

```
Δp = (1 − φ) × shock_pp / 100
g_ext = ε_ext × (Δp × w)
g_int = ε_int × Δp
ΔD_gastro = B × [(1 + g_ext) × (1 + g_int) − 1]
ΔVBP = (I − A)⁻¹ · Δf
ΔVAB = v̂ · ΔVBP
```

Donde: φ = absorción de margen, w = peso de gastronomía en el gasto total del viaje, B = base afectada (MUSD), A = coeficientes técnicos domésticos de la MIP, Δf = vector de shock de demanda final, v̂ = diagonal de coeficientes VA/VBP.

---

## Principios de transparencia

- **Datos vs. supuestos:** todo output indica si un valor proviene de datos observados o de un parámetro supuesto.
- **Parámetros centralizados:** ningún número hardcodeado en los scripts; todos los parámetros viven en `config/supuestos.yaml`.
- **Replicabilidad:** semilla aleatoria fija (seed=42), versiones de paquetes pinneadas en `requirements.txt`.
- **Trazabilidad:** hashes SHA256 de todos los archivos de datos en `DATA_SOURCES.md`.
