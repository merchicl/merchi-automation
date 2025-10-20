# Merchi Automation (Apps Script + Web)
Código de Cotizador, Órdenes y Bocetos/Aprobaciones.

## Estructura
- `*/appsscript/*.gs` lógica de Google Apps Script
- `*/web/*.html` vistas servidas por HtmlService
- `shared/utils.gs` funciones reutilizables

## Flujo de ramas
- `main`: estable/producción
- `dev`: integración y pruebas

## Despliegue
Sincronizar manualmente los archivos del repo a cada proyecto de Apps Script correspondiente.
Mantener IDs/secretos en Properties Service (no en el repo).
