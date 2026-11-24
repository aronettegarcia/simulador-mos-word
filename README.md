# Simulador MOS Word (MO-100 / MO-110)

## Instalación
1. Instale Python 3.10 o posterior.
2. En esta carpeta ejecute `python -m venv .venv` y active el entorno.
3. Ejecute `pip install -r requirements.txt`.
4. Ejecute `python app.py` y abra `http://127.0.0.1:5000`.

Incluye ocho plantillas DOCX en `templates/`, interfaz Flask y evaluación inicial de 40 criterios XML.

## Importante: alcance de la evaluación
La evaluación es una aproximación automatizada basada en OpenXML. Word puede almacenar ciertos ajustes en estilos, relaciones, partes XML distintas o estructuras equivalentes; por ello, una comprobación negativa no siempre significa que la tarea esté mal. Los conjuntos de estilos, TOC actualizada, columnas limitadas a un rango exacto, compatibilidad, controles de cambios y algunos formatos visuales requieren validación más sofisticada o revisión en Microsoft Word. Este paquete no es un producto oficial de Certiport/Microsoft y no reproduce el motor propietario de examen.

## Seguridad antes de publicar
Cambie `SECRET_KEY`, desactive `debug=True`, configure HTTPS, límites de tamaño de archivo, autenticación, almacenamiento aislado y limpieza programada. Para una evaluación real, la fecha límite debe persistirse en servidor y asociarse a un intento autenticado; el reloj del navegador es solo presentación. La sesión Flask aquí usa cookie firmada y no sustituye un sistema de supervisión de examen.

## Puntuación
El código devuelve tareas aprobadas por proyecto. La escala total del examen es 1000: 40 tareas x 25 puntos. Aprobación: 700 puntos. El servidor acumula las tareas aprobadas de los proyectos entregados: 40 tareas x 25 puntos = 1000 puntos. Los proyectos no enviados cuentan como tareas no aprobadas. El resultado global es provisional hasta completar los ocho proyectos o agotar el tiempo.
