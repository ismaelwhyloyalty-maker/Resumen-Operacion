# Email Monitoring Agent - Resumen de Operaciones

Este agente monitorea tu Outlook de Microsoft 365 cada 5 minutos y genera un reporte CSV con:

- **Fecha Recepción**: Fecha y hora del correo
- **Tipo Operación**: Detecta automáticamente (Exportación, Importación, Cruces, Otra)
- **Asunto**: Asunto del correo
- **Caja/Trailer**: Extrae códigos (R + dígitos o 5 dígitos) del cuerpo del correo
- **Observaciones**: Primeros 500 caracteres del cuerpo del correo
- **Remitente**: Email completo del remitente

## Funcionamiento

### Detección automática de tipos:

1. **Exportación**: Remitente es Cecilia Ramón + asunto contiene "Loyalty Relacion CCP"
2. **Importación**: Asunto contiene "Carta Porte" + remitente es de Radar
3. **Cruces**: Remitente es de Radar (especialmente Abraham Alvarez)
4. **Otra**: Cualquier otro correo

### Extracción de Caja/Trailer:

Busca en el cuerpo del correo:
- Códigos que empiezan con "R" seguido de dígitos (ej: R344, R389)
- Números de 5 dígitos (ej: 50772, 51753)

## Configuración

Los secrets ya están configurados:
- `EMAIL_USER`: ismael@whyloyalty.com
- `EMAIL_PASSWORD`: Tu contraseña de Outlook

## Reporte

El archivo `reporte_operaciones.csv` se actualiza automáticamente cada 5 minutos con los últimos correos monitoreados.

### Para visualizar el reporte:
1. Ve a la rama `main` de tu repositorio
2. Haz clic en `reporte_operaciones.csv`
3. GitHub lo muestra como tabla

### Para descargar:
1. Haz clic en el botón "Download"
2. Abre en Excel o Google Sheets
