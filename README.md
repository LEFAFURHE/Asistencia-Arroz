# 🌾 Asistencia Técnica Integral en Cultivo de Arroz (Programa AMTEC - Fedearroz)

Repositorio técnico y aplicativo interactivo para la asistencia técnica integral, seguimiento fenológico día a día, monitoreo fitosanitario con umbrales económicos, manejo de riego con manguera de 2 pulgadas y teledetección satelital (SIG / NDVI / Radar SAR) en el cultivo de arroz bajo la metodología **AMTEC (Adopción Masiva de Tecnología - Fedearroz)**.

---

## 🚀 Aplicativo Interactivo (Cronograma y SIG)

El archivo index.html contiene el panel interactivo completo con 8 módulos:
1. **🌱 Fenología:** Los 9 estadíos fenológicos con Días Después de Germinación (DDG) y Días Calendario.
2. **📋 Cronograma General:** Flujo cronológico de labores desde pre-siembra (DDS -30) hasta cosecha.
3. **🏗️ Preparación de Suelos:** Labores manuales de adecuación, nivelación (<5 cm) y caballones.
4. **🔍 Monitoreo Fitosanitario:** Umbrales económicos AMTEC para Sogata, Diatraea, Spodoptera, Oebalus, Pyricularia y malezas.
5. **💧 Sistema de Riego:** Dinámica de láminas de agua con motor de 16 L y manguera de 2 pulgadas.
6. **👷 Jornales:** Balance detallado de los ~150 jornales/ha requeridos para labores 100% manuales con bombas de 20 L.
7. **🧪 Insumos y Fertilización:** Programa nutricional fraccionado y manejo fitosanitario condicional (MIP).
8. **🛰️ Teledetección y SIG:** Catálogo de índices satelitales (NDVI, NDRE, LSWI, MNDWI, SAR Sentinel-1, PSRI) e integración en QGIS.

> 🌐 **Para visualizarlo online de forma gratuita:** Habilita **GitHub Pages** en la configuración de este repositorio (Settings > Pages > Deploy from a branch: main / root).

---

## 📖 Bitácora Día a Día de Manejo

El documento [Bitacora_Dia_a_Dia_Arroz_AMTEC.md](./Bitacora_Dia_a_Dia_Arroz_AMTEC.md) desglosa el seguimiento agronómico día a día para un lote sembrado el **12 de Septiembre**, incluyendo:
* **07 de Octubre (DDS 25 / DDG 20):** 1ª Abonada (Urea 50kg + SAM 50kg + KCl 75kg/Ha).
* **14 de Octubre (DDS 32 / DDG 27):** Aplicación de Propanil (6 L/Ha) + Ozlo [Clorantraniliprol + Thiametoxam] (200 cc/Ha).
* **17 de Octubre (DDS 35 / DDG 30):** 2ª Abonada (Urea 50kg + SAM 50kg + KCl 75kg/Ha).
* **DDS 50-55 (Noviembre):** 3ª Abonada en Iniciación Panicular (Urea 50kg + SAM 25kg + KCl 50kg/Ha).
* **DDS 65 (Embuchamiento):** Protección fungicida obligatoria para cuello de panícula.
* **DDS 100:** Drenaje final del lote.
* **DDS 110-115:** Cosecha manual con humedad de grano entre 20% y 22%.

---

## 🛰️ Estudios con SIG y Sensores Remotos

Se integran datos multiespectrales y radar en el proyecto QGIS [Monitoreo_Arroz.qgz](./Monitoreo_Arroz.qgz) junto a las capas móviles de KoboCollect [Datos_Monitoreo_Arroz.gpkg](./Datos_Monitoreo_Arroz.gpkg):
* **NDVI (Sentinel-2 B8/B4):** Vigor y fallas de emergencia en etapas tempranas.
* **NDRE (Sentinel-2 B8/B5):** Contenido de clorofila y nitrógeno foliar sin saturación en dosel cerrado.
* **NDWI / LSWI (B8A/B11):** Diagnóstico anticipado de estrés hídrico 3 a 5 días antes de síntomas visuales.
* **MNDWI (B3/B11):** Mapeo de la lámina de agua y detección de fugas en caballones.
* **Radar SAR (Sentinel-1 Banda C):** Monitoreo continuo de biomasa y fenología **100% inmune a la nubosidad**.
* **PSRI:** Detección precoz de focos de Pyricularia y Añublo de la vaina.

---

## 📂 Estructura del Repositorio

`	ext
├── index.html                           # Aplicativo web interactivo (GitHub Pages)
├── Cronograma_Arroz_AMTEC.html          # Copia del cronograma para apertura local
├── Bitacora_Dia_a_Dia_Arroz_AMTEC.md    # Bitácora detallada día por día
├── README.md                            # Presentación y documentación general
├── Monitoreo_Arroz.qgz                  # Proyecto SIG en QGIS 3.x
├── Datos_Monitoreo_Arroz.gpkg           # Base de datos geoespacial de monitoreos
├── Formulario_Arroz.qml                 # Estilos QML para visualización en QGIS
├── Formularios_Kobo/                    # Formularios XLSForm para captura en campo
├── .gitignore                           # Exclusión de archivos pesados (>100MB) y temporales
└── Cartillas_AMTEC_PDF/                 # Documentación técnica de referencia AMTEC
`

---

## 📜 Referencias
Desarrollado bajo los lineamientos técnicos de **FEDEARROZ - Fondo Nacional del Arroz (FNA)** dentro del programa **AMTEC** (*Adopción Masiva de Tecnología*).
