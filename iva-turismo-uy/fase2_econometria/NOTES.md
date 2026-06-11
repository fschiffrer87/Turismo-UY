# Fase 2 — Diseño del Event-Study (NO IMPLEMENTAR AÚN)

## Objetivo
Estimar econométricamente las elasticidades ε_int y ε_ext para Uruguay, reemplazando los supuestos de la literatura internacional utilizados en la Fase 1.

## Estrategia de identificación
Explotar la variación temporal y cruzada generada por los cambios del régimen de devolución de IVA a turistas entre 2012 y 2025, como experimento cuasi-natural.

### Eventos explotables (cambios de política)
| Año | Evento | Variación en el incentivo |
|-----|--------|--------------------------|
| 2012 | Introducción de la devolución permanente (Ley 17.934) | IVA 22% → IVA 13% para tarjetas del exterior |
| 2018–2019 | Modificaciones al régimen de devolución (monto mínimo, categorías) | Cambio marginal en la tasa efectiva |
| 2023 | Decreto 434/023: exoneración total en temporada alta | IVA 13% → IVA 0% en nov–abr |
| 2025 | Prórroga Decreto 220/2025 | Continuidad del régimen estacional |

### Hipótesis de identificación
Los cambios de política son exógenos a la demanda turística corriente (no responden a shocks de turismo del período). La variación en el incentivo fiscal es ortogonal a los determinantes habituales del turismo (TCR, actividad económica de origen).

## Datos
- **Variable dependiente:** llegadas de turistas no residentes (ETR, desagregada por nacionalidad y destino) y gasto total por rubro y nacionalidad (ETR trimestral, serie 2010–2024).
- **Variable de tratamiento:** tasa de IVA efectiva para no residentes con tarjeta del exterior, según el régimen vigente en cada trimestre.
- **Controles:**
  - Tipo de cambio real bilateral Uruguay–Argentina (fuente BCU/INDEC).
  - Nivel de actividad de Argentina (EMAE mensual, fuente INDEC).
  - Estacionalidad (dummies trimestrales y efectos fijos de mes/año).
  - Efectos fijos por nacionalidad (para el panel de llegadas por origen).

## Especificación base
```
ln(Y_{it}) = α_i + γ_t + β × IVA_t + δ × TCR_t + ζ × EMAE_t + ε_{it}
```
- i: nacionalidad del turista (Argentina, Brasil, Chile, resto)
- t: trimestre
- Y_{it}: llegadas o gasto en gastronomía de la nacionalidad i en trimestre t
- IVA_t: tasa de IVA efectiva (0 en temporada alta post-2023, 13 en permanente, 22 antes de 2012)
- β: elasticidad-precio de la demanda turística (parámetro de interés)

## Diseño event-study
- Ventana: ±8 trimestres alrededor de cada cambio de política.
- Prueba de anticipación: coeficientes pre-evento deben ser ≈ 0.
- Heterogeneidad: estimar β por nacionalidad (Argentina vs. resto) — se espera mayor sensibilidad de argentinos al IVA (TCR ya capta la competitividad precio).

## Amenazas a la validez interna
1. **Confusión con el TCR:** los cambios de política coinciden en parte con movimientos cambiarios; controlar por TCR bilateral es crítico.
2. **Efectos espurios de pandemia (2020–2021):** excluir esos trimestres o tratarlos como observaciones faltantes.
3. **Turismo de compras vs. turismo de ocio:** la ETR no distingue claramente; puede ser necesario usar microdatos.

## Outputs esperados
- Estimación de β (elasticidad precio del turismo) con intervalos de confianza.
- Test de pre-tendencias (gráfico de coeficientes evento a evento).
- Comparación con los supuestos de la Fase 1: ¿las elasticidades de la literatura son razonables para Uruguay?

## Nota sobre magnitud esperada
La literatura de demanda turística latinoamericana sugiere elasticidades precio en torno a −1.0 a −1.5 (Crouch, 1995; Peng et al., 2015). Sin embargo, la dominancia del TCR Argentina–Uruguay en la función de demanda puede reducir el poder estadístico para identificar el efecto del IVA (segundo orden relativo).
