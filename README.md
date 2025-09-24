# Conciliapp QA

Conciliapp QA es una aplicación de gestión y conciliación de cobranzas, diseñada para facilitar el registro, seguimiento y análisis de pagos en organizaciones que manejan múltiples vendedores, clientes y bancos. El sistema está construido sobre Google Apps Script y utiliza hojas de cálculo de Google para el almacenamiento y procesamiento de datos.

## Características principales

- **Autenticación de usuarios**: Control de acceso por roles (Analista, Administrador, Vendedor).
- **Gestión de cobranzas**: Registro de pagos, asignación de registros a analistas, actualización de estados (Pendiente, Procesado, Rechazado).
- **Filtros avanzados**: Búsqueda y filtrado por estado, sucursal, vendedor, cliente, banco y fecha.
- **Exportación de reportes**: Generación de reportes en PDF de registros filtrados y análisis para analistas y administradores.
- **Sincronización con API**: Obtención de datos de vendedores, clientes y facturas desde fuentes externas.
- **Particionamiento de datos**: Organización de registros por mes, vendedor o banco para optimizar el rendimiento.
- **Gestión de overrides**: Asignación directa de registros a analistas específicos por parte de administradores.
- **Notificaciones**: Avisos automáticos de pagos confirmados y actualizaciones de registros.

## Estructura del repositorio

- `Codigo.js`: Lógica principal del servidor y funciones Apps Script.
- `Index.html`: Interfaz de usuario para registro de cobranzas.
- `AnalystView.html`: Panel de analista para gestión y revisión de registros.
- `Auth.html`, `Auth.js.js`: Pantalla y lógica de autenticación.
- `dashboard.html`, `Report.html`: Vistas adicionales para reportes y dashboard.
- `styles.html`: Estilos compartidos para las vistas.
- `appsscript.json`: Configuración del proyecto Apps Script.
- `routersheets.txt`, `routersheets/`: Archivos auxiliares y de configuración.
- `Cobranza_tinito_QA - analista.csv`: Ejemplo de datos de analistas.

## Instalación y despliegue

1. Clona el repositorio y abre el proyecto en Google Apps Script.
2. Configura las hojas de cálculo y los triggers necesarios desde el menú de Apps Script.
3. Personaliza los parámetros de conexión a la API y las propiedades del script según tu entorno.
4. Publica el WebApp y comparte la URL con los usuarios autorizados.

## Uso

- Accede a la URL del WebApp y autentícate con tu usuario.
- Registra cobranzas, consulta registros recientes y descarga reportes en PDF.
- Los analistas pueden revisar, procesar o rechazar registros asignados.
- Los administradores pueden gestionar asignaciones directas y sincronizar datos desde la API.

## Contribución

Las contribuciones son bienvenidas. Por favor, abre un issue o pull request para sugerencias, mejoras o correcciones.

## Licencia

Este proyecto está bajo licencia MIT.
