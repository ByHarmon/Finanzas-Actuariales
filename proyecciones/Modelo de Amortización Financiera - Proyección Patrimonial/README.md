# 📊 Modelo de Amortización Financiera — Proyección Patrimonial 65–95 años

Modelo determinístico en Excel que proyecta, **mes a mes**, la situación patrimonial de una persona durante su retiro, para responder una pregunta central:

> **¿El capital inicial, la pensión y los rendimientos de inversión alcanzan para cubrir los gastos durante todo el horizonte de vida proyectado? Si no, ¿a qué edad se agota el patrimonio?**

Construido en Python (`openpyxl`) y entregado como un archivo `.xlsx` 100% funcional, editable y sin macros ni dependencias externas.

---

## 🎯 ¿Qué hace este modelo?

- Proyecta el patrimonio **mes a mes** (no año a año) durante todo el horizonte de retiro.
- Calcula por separado el crecimiento de la **pensión**, el **gasto general** y el **gasto médico**, cada uno con su propia lógica de actualización.
- Incorpora un **aging factor** actuarial: el gasto médico no solo crece con inflación, también aumenta conforme la persona envejece.
- Modela el **rendimiento de inversión** (CETES 28 días) y su **retención fiscal**, reinvirtiendo el saldo mes a mes.
- Determina automáticamente si el patrimonio **alcanza** hasta la edad final o **en qué mes/edad se agota**.
- Es **100% paramétrico**: cambiar la edad inicial, la edad final, el capital, la pensión o cualquier supuesto recalcula todo el modelo sin tocar una sola fórmula.

---

## 🧮 Supuestos y lógica financiera

| Variable | Frecuencia de actualización | Fórmula |
|---|---|---|
| Pensión | **Anual** (constante 12 meses, luego salta) | `Pensión_nueva = Pensión_anterior × (1 + INPC anual)` |
| Gasto general | **Mensual** | `Gasto_nuevo = Gasto_anterior × (1 + inflación subyacente mensual)` |
| Gasto médico | **Mensual**, con doble factor | `Gasto_nuevo = Gasto_anterior × (1 + inflación médica mensual) × (1 + aging mensual)` |
| Rendimiento de inversión | **Mensual**, sobre saldo inicial del mes | `Rendimiento = Saldo_inicial × tasa CETES mensual` |

Toda tasa anual efectiva se convierte a mensual con:

```
tasa_mensual = (1 + tasa_anual)^(1/12) - 1
```

### Aging factor (gasto médico)

El aging representa el aumento en la **utilización** de servicios médicos por edad, independiente de la inflación de precios:

| Rango de edad | Aging anual (supuesto, editable) |
|---|---|
| 65–69 | 3.0% |
| 70–74 | 2.5% |
| 75–79 | 2.0% |
| 80–84 | 1.0% |
| 85–89 | 0.5% |
| 90+ | 0.0% |

> ⚠️ Todos los supuestos (INPC, inflación subyacente, inflación médica, CETES, retención fiscal y aging) son **aproximaciones para fines de modelación**, no predicciones oficiales ni tasas garantizadas a 30 años. Son 100% editables.

---

## 📁 Estructura del archivo Excel

El libro tiene 3 hojas:

### 1. `PARAMETROS`
Todas las celdas editables (fondo amarillo) del modelo: edades, capital, pensión, gasto, distribución médico/general, tasas anuales y curva de aging. También calcula automáticamente las tasas mensuales equivalentes.

### 2. `AMORTIZACION`
Una fila por cada mes del horizonte (hasta 720 filas / 60 años de holgura), con 17 columnas: mes, año de proyección, edad, saldo inicial, rendimiento, retención fiscal, pensión, gasto general, gasto médico, aging aplicado, gasto total, flujo neto, saldo final y estado (OK / AGOTADO).

El número de meses activos se calcula automáticamente como `(Edad final − Edad inicial) × 12`; las filas fuera de ese rango quedan en blanco sin necesidad de editar fórmulas.

### 3. `RESUMEN`
Panel ejecutivo con los resultados clave:
- Si el patrimonio alcanza o no hasta la edad final.
- Edad y mes exactos de agotamiento (si aplica).
- Pensión, gasto y gasto médico proyectados al final del horizonte.
- Dos gráficas: evolución del patrimonio, y pensión vs. gasto mensual.

---

## 🚀 Cómo usarlo

1. Descarga el archivo `.xlsx` de este repositorio.
2. Ábrelo en Excel (recomendado) o Google Sheets/LibreOffice.
3. Ve a la hoja `PARAMETROS` y modifica cualquier celda amarilla (edad, capital, pensión, inflaciones, etc.).
4. Ve a la hoja `RESUMEN` para ver el resultado actualizado al instante.
5. Consulta la hoja `AMORTIZACION` si quieres revisar el detalle mes a mes.

No requiere macros, complementos ni configuración adicional.

---

## 🔍 Preguntas que responde el modelo

1. ¿El capital inicial alcanza hasta la edad final?
2. ¿A qué edad/mes se agota el patrimonio, si no alcanza?
3. ¿Cuál es el saldo proyectado al final del horizonte?
4. ¿Cómo evolucionan la pensión y el gasto mes a mes?
5. ¿Qué porcentaje del gasto total representa el gasto médico conforme aumenta la edad?
6. (Vía Buscar Objetivo de Excel) ¿Cuánto capital inicial se necesitaría para llegar a la edad final con saldo ≥ 0?
7. (Vía Buscar Objetivo de Excel) ¿Qué pensión inicial se necesitaría para que el capital alcance?

---

## 🛠️ Stack técnico

- **Python 3** + [`openpyxl`](https://openpyxl.readthedocs.io/) para la generación del archivo.
- Fórmulas nativas de Excel (sin VBA/macros) para que el modelo sea 100% transparente y editable por cualquier persona.
- Formato condicional, paneles congelados, filtros y gráficas nativas incluidas en el archivo.

---

## 🗺️ Roadmap — próxima versión (modelo probabilístico)

Este repositorio contiene la versión **determinística**. Las siguientes mejoras están planeadas:

- [ ] Histórico real mensual de CETES 28 días e INPC.
- [ ] Escenarios pesimista / base / optimista.
- [ ] Simulación Monte Carlo de rendimientos e inflación.
- [ ] Percentil de longevidad (ej. 95%) en vez de edad fija.
- [ ] Análisis de sensibilidad al gasto médico y al aging factor.
- [ ] Cálculo de ISR real (no solo retención provisional).
- [ ] Cálculo automático (sin Buscar Objetivo manual) del capital/pensión necesarios.

---

## ⚠️ Disclaimer

Este modelo es una **herramienta educativa y de planeación personal**, no constituye asesoría financiera, actuarial ni fiscal. Los supuestos de inflación, rendimientos y aging son aproximaciones y deben ajustarse con datos reales y, de preferencia, revisión de un asesor financiero certificado antes de tomar decisiones de retiro.

---

## 📄 Licencia

MIT — úsalo, modifícalo y compártelo libremente citando la fuente.
