# Reportes_incidencias

Sistema web serverless desarrollado para automatizar la captura de datos en campo, orquestar flujos de auditoría de calidad y visualizar métricas en tiempo real. Este proyecto elimina el uso de reportes manuales, conectando la operación técnica directamente con los tableros de toma de decisiones gerenciales mediante un flujo de datos continuo.

## Arquitectura y Tecnologías

* Frontend: HTML5, CSS3, Vanilla JavaScript.
* Backend: Google Apps Script (V8 Engine).
* Base de Datos: Google Sheets.
* Almacenamiento: Google Drive API (Procesamiento de imágenes y generación de PDF).
* Integración de Datos: Zoho Desk API (vía Webhooks).
* Business Intelligence (BI): Looker Studio.

## Funcionamiento del Sistema

El proyecto se divide en tres módulos principales, enrutados dinámicamente desde un único controlador backend:

### 1. Captura de Datos en Campo
Interfaz web responsiva diseñada para uso en dispositivos móviles por parte de los técnicos.
* Autocompletado de datos: Se conecta mediante un Webhook a la API de Zoho Desk para validar el número de ticket y cruzar la información del cliente y equipo de forma automática.
* Optimización de rendimiento: Las evidencias fotográficas son procesadas y comprimidas localmente mediante la API Canvas de HTML5 en el navegador antes de su conversión a Base64, reduciendo el consumo de datos móviles y tiempos de carga.
* Ruteo inteligente: El sistema evalúa las respuestas del técnico (ej. solicitud de refacciones o apoyo en sitio) para disparar correos electrónicos con plantillas HTML a los departamentos específicos correspondientes.

### 2. Flujo de Aprobaciones y Soporte
Vistas de interfaz controladas por parámetros en la URL para el personal administrativo y de ingeniería.
* Panel de Calidad: Permite a los departamentos auditar incidencias y documentar causas raíz y acciones preventivas.
* Panel de Soporte: Interfaz exclusiva para que Ingeniería documente la resolución de problemas técnicos en campo.
* Control de concurrencia: El backend verifica el estado del registro en la base de datos para bloquear formularios que ya fueron dictaminados, garantizando la integridad de la información.

### 3. Generación Documental y Analítica
* Reportes automatizados: Utilizando DocumentApp, el backend compila los datos y las imágenes almacenadas en Drive para generar un reporte técnico formal en formato PDF, disparando su descarga automática al usuario.
* Business Intelligence: La base de datos centralizada alimenta un tablero interactivo en Looker Studio. Mediante campos calculados, modelado de datos y sentencias SQL, se transforman registros en bruto en KPIs dinámicos (tasas de resolución, incidencias principales, rendimiento por sucursal).

## Despliegue

1. Configurar un proyecto de Google Apps Script vinculado a un archivo de Google Sheets.
2. Añadir los archivos fuente (.js y .html) al entorno de desarrollo.
3. Declarar las variables de entorno (IDs de carpetas de Drive y Spreadsheet) en las constantes globales.
4. Desplegar como Aplicación Web con permisos de ejecución.
5. Configurar el endpoint generado en el Webhook de Zoho Desk para la recepción de cargas JSON.

## Nota de Privacidad

Este repositorio contiene una versión pública del código fuente. Por motivos de seguridad corporativa, las credenciales, identificadores de carpetas, tokens de integración y direcciones de correo electrónico han sido anonimizados.
