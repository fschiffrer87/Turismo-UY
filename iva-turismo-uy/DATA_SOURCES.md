# DATA SOURCES — iva-turismo-uy

Este archivo se actualiza automáticamente por `src/01_descarga_datos.py`.
Cada fila registra URL, fecha de descarga (o recepción manual) y hash SHA256.

| ID | Descripción | URL | Fecha | SHA256 | Método |
|----|-------------|-----|-------|--------|--------|
| DGI-GT | Informe Gasto Tributario DGI 2019–2022 | https://www.gub.uy/direccion-general-impositiva/sites/direccion-general-impositiva/files/2023-12/Informe+de+Gasto+Tributario+2019+-+2022.pdf | pendiente | pendiente | automático / manual |
| MINTUR-ETR-AGR | ETR Agregados trimestrales MINTUR | https://catalogodatos.gub.uy/dataset/ministerio-de-turismo-turismo-receptivo | pendiente | pendiente | automático / manual |
| MINTUR-ETR-MIC | ETR Microdatos anuales MINTUR | https://www.gub.uy/ministerio-turismo/datos-y-estadisticas/microdatos | pendiente | pendiente | manual |
| BCU-MIP | Matriz Insumo-Producto 2016 BCU | https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Matriz-Insumo-Producto.aspx | pendiente | pendiente | manual (SharePoint) |
| BCU-CN | Cuentas Nacionales BCU (VAB, PIB nominal) | https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Cuentas-Nacionales.aspx | pendiente | pendiente | automático / manual |
| BCU-PAGOS | Reporte Sistema de Pagos Minorista BCU | https://www.bcu.gub.uy/Estadisticas-e-Indicadores/Paginas/Sistema-de-Pagos.aspx | pendiente | pendiente | manual |

## Notas
- **automático**: descargado por `01_descarga_datos.py` con `requests`.
- **manual**: requiere descarga manual por restricciones de acceso (Cloudflare, SharePoint, autenticación). Ver instrucciones en `README.md`.
- El script actualiza este archivo con fecha real y hash al completar cada descarga exitosa.
