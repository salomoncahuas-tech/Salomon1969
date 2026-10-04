# Datos base del Proyecto IN Piura

## Identificación
- **Nombre:** Recuperación del servicio de regulación de riesgos naturales y recuperación de ecosistemas degradados en la Cuenca Alta del Río Piura
- **CUI:** 2669244 · **Sistema:** Invierte.pe (DGPMI-MEF)
- **Entidad:** Autoridad Nacional de Infraestructura (ANIN) · DIME · SESDI
- **Fase actual:** Preinversión (Estudio de Perfil)
- **Financiamiento y modalidad de ejecución:** únicamente Administración Directa (confirmado por el usuario)

## Presupuesto
- **Inversión total:** S/ 372,442,809 (tipo de cambio referencial 3.40 S/ por USD) ≈ **USD 109.5 M**
- Incluye: estudio de preinversión, estudio definitivo o expediente técnico, ejecución (5 años) y operación y mantenimiento (15 años).

## Ámbito
- **3 provincias:** Morropón, Huancabamba, Ayabaca.
- **15 distritos:** Frías, Canchaque, Huarmaca, Huancabamba, Lalaquiz, San Miguel de El Faique, Buenos Aires, Chalaco, Chulucanas, Morropón, Salitral, San Juan de Bigote, Santa Catalina de Mossa, Santo Domingo, Yamango.
- **Área de intervención directa:** entre 10,000 y 15,000 ha.
- **Bloques preliminares de intervención:** 117, que suman **12,270.235 ha**.
- **Codificación de bloques:** `M[microcuenca]-B[bloque]` (ej. M5-B1) en la nomenclatura del proyecto; el catálogo V6 los registra mezclando códigos numéricos (`27`) y `M17B10` sin guion.
- **Sistema de coordenadas:** UTM WGS84 Zona 17S (EPSG:32717).

## Fuentes de verdad en el repositorio
- `datos/unidades_liberacion_areas.csv`: catálogo V6 vigente (117 bloques y lotes SUS, con provincia, distrito, área y asistente).
- `README_LIBERACION_AREAS.md`: bloques retirados en V6 y cómo se enlazan con la base de datos.
- Excel de centros poblados, centroides, MSAVI y plantillas FDT en la raíz del repositorio. Los de nombre "V5" o "v6" distinguen versiones del catálogo: usa la más reciente salvo que el usuario indique otra.

## Líneas de intervención (resumen)
1. **Infraestructura natural (verde):** revegetación, reforestación con especies nativas; viveros permanentes y temporales.
2. **Infraestructura marrón:** zanjas de infiltración, terrazas, diques para cárcavas, clausura y manejo de pastizales, barreras y cercos vivos; obras complementarias (estabilización de taludes, muros, gaviones, diques en quebradas, drenaje).
3. **Gobernanza y gestión comunitaria:** planes de manejo, capacidades en GdR, alerta temprana comunitaria, MERESE.

Detalle técnico y costos en `gdr-ccc.md`.

## Indicador de brecha
Porcentaje de superficie de ecosistemas degradados que brindan servicios ecosistémicos que requieren recuperación (RM N° 00213-2024-MINAM).
