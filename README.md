# SpringLog-IA — paquete final de proyecto

## Resumen
Prototipo reproducible de clasificación asistida de severidad para logs sintéticos del microservicio `payment-service`. Incluye análisis exploratorio, sanitización mediante reglas, comparación de clasificadores, modelo serializado, informes y presentación. No es una integración productiva ni una prueba sobre telemetría real.

## Entregables
- `01_FICHA_DEL_PROBLEMA.docx` y `01_FICHA_DEL_PROBLEMA.pdf`: definición del problema, alcance y criterios de éxito.
- `02_IA_GENERATIVA.docx`: prompts, ejemplos de entrada/salida, capacidades, límites y pasos siguientes.
- `03_ANALISIS_DATOS.docx`, `03_ANALISIS_DATOS.pdf` y `03_ANALISIS_DATOS.xlsx`: EDA y dataset preparado; el libro contiene resumen, severidades, instancias, excepciones, calidad, modelos y datos limpios.
- `04_MODELO_O_PROTOTIPO.docx`, `04_MODELO_O_PROTOTIPO.py` y `04_MODELO_O_PROTOTIPO.joblib`: arquitectura, evaluación, script y modelo serializado.
- `05_EVALUACION_RESULTADOS.docx`: métricas reproducibles, matriz de confusión, casos borde y muestra estratificada de 25 predicciones.
- `06_ANALISIS_ETICO_NEGOCIO.docx`: privacidad, sesgos, supervisión, riesgos y escenario económico explícitamente hipotético.
- `07_PRESENTACION_EJECUTIVA.pptx` y `07_PRESENTACION_EJECUTIVA.pdf`: presentación de 15 diapositivas.
- `07_VIDEO_RESUMEN_EJECUTIVO.mp4`: video de aproximadamente 4 minutos y 48 segundos con narración sintética en español, basado en la presentación. Puede reemplazarse por una grabación personal.
- `07_GUION_RESUMEN_EJECUTIVO.docx`: guion de narración para adaptar o volver a grabar.

## Resultados de la última ejecución registrada
- Dataset: 500 registros, 8 columnas de origen y 11 columnas en la versión preparada; 0 duplicados exactos.
- Etiquetas: INFO 300 (60 %), WARN 100 (20 %), ERROR 75 (15 %) y CRITICAL 25 (5 %). La distribución fue fijada al generar datos sintéticos.
- Random Forest: accuracy 93,0 %, F1 macro 0,933 y F1 ponderado 0,925 en una partición estratificada de prueba de 100 filas (`random_state=42`).
- Matriz de confusión: 5/5 CRITICAL y 15/15 ERROR correctos; 13/20 WARN correctos, con 7 WARN clasificados como INFO.
- Latencia local: mediana 21,5 ms y P95 23,6 ms por registro en la ejecución registrada; depende del hardware y no equivale al MTTR.

**Advertencia de validez:** la muestra es pequeña y sintética. Regresión logística y Naive Bayes obtuvieron 100 % en la misma partición, lo que sugiere que algunas plantillas son demasiado separables. Las métricas no demuestran generalización a producción; la severidad crítica tiene solo cinco casos de prueba. La muestra de 25 predicciones es una hoja de trabajo para revisión humana, no una auditoría independiente.

## Requisitos y ejecución
Python 3.10+ recomendado. Instala dependencias con:

```bash
pip install -r requirements.txt
```

Desde la raíz del paquete, ejecuta en orden:

```bash
python 03_ANALISIS_DATOS.py
python 04_MODELO_O_PROTOTIPO.py
```

El primer script valida el CSV de origen, normaliza los 389 marcadores literales `None` como `NO_APLICA`, aplica reglas de sanitización, crea `data/payment_service_logs_500_clean.csv` y `outputs/payment_service_logs_500_clean.csv`, y genera las gráficas y resúmenes del EDA.

El segundo script toma el dataset preparado, construye características desde `logger_class`, `exception_type` y `log_message`, enmascara tokens de severidad dentro del texto, compara cuatro clasificadores, calcula métricas en la partición reservada, genera las gráficas y serializa el pipeline. La evaluación usa el split de prueba; el `.joblib` final se reentrena sobre todos los registros para la demostración posterior.

Para reconstruir informes, libro y presentación desde los últimos resultados guardados, ejecuta opcionalmente:

```bash
python build_deliverables.py
```

## Contrato de inferencia
El `.joblib` recibe una lista de cadenas ya compuestas y sanitizadas con el formato `logger_class + exception_type + log_message`. Consulta las funciones `sanitize` y `build_features` en `04_MODELO_O_PROTOTIPO.py`. No envíes logs sensibles a servicios externos sin sanitización local y revisión de seguridad.

## Limitaciones y siguientes pasos
1. Validar etiquetas con especialistas y evaluar mensajes de plantillas nuevas o logs reales anonimizados.
2. Usar una prueba temporal o separada por plantillas para reducir el riesgo de fuga de información.
3. Añadir características estructuradas como duración, estado HTTP, reintentos y métricas del pool.
4. Medir el pipeline completo y el impacto en diagnóstico mediante un piloto en modo sombra.
5. Las estimaciones de ROI son un escenario de negocio basado en supuestos, no ahorros observados.
