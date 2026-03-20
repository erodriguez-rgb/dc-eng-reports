# DC Eng Reports - Servidor de Generación de Informes

Servidor Flask para generar informes diarios en formato DOCX.

## Variables de entorno requeridas
- `API_TOKEN` - Token secreto para autenticar las peticiones

## Endpoints
- `GET /health` - Verificar que el servidor está activo
- `POST /generate` - Generar informe DOCX (requiere header `X-API-Token`)

## Deploy en Railway
1. Conectar este repo en Railway
2. Configurar variable de entorno `API_TOKEN`
3. Railway detecta el Procfile automáticamente
