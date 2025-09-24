# Plan de Mejora: Sistema de Notificaciones

Este documento describe el plan para refactorizar y completar el sistema de notificaciones de pagos confirmados para vendedores.

## 1. Resumen del Problema

La implementación actual es funcional pero incompleta. Sus principales debilidades son:

- **Efímera:** Las notificaciones solo se muestran si el pago se confirmó en las últimas 24 horas. No hay historial.
- **Sin Estado:** El sistema no registra si una notificación ha sido leída, por lo que se muestra repetidamente.
- **Ineficiente:** El mecanismo de búsqueda escanea todas las hojas del documento, lo que no es escalable y degrada el rendimiento.
- **UX Limitada:** No existe un centro de notificaciones para consultarlas ni son interactivas.

## 2. Objetivos de la Mejora

- **Persistencia:** Todas las notificaciones deben guardarse de forma permanente.
- **Gestión de Estado:** Implementar la funcionalidad de "leído" y "no leído".
- **Eficiencia:** Optimizar la creación y consulta de notificaciones para no impactar el rendimiento de la aplicación.
- **Mejora de UX:** Crear una interfaz de usuario intuitiva con un centro de notificaciones y un indicador de avisos nuevos.

## 3. Arquitectura Propuesta

### 3.1. Nuevo Almacén de Datos: Hoja "Notificaciones"

Se creará una nueva hoja de cálculo llamada `Notificaciones` que servirá como almacén persistente. Tendrá la siguiente estructura:

| Columna | Nombre Técnico | Descripción | Ejemplo |
|---|---|---|---|
| A | ID_Notificacion | Identificador único de la notificación. | `NOTIF_kjs93hsd` |
| B | ID_Registro_Pago | ID del registro de pago confirmado (para enlaces). | `REG_banco_xyz` |
| C | CodVendedor | Código del vendedor al que pertenece la notificación. | `085` |
| D | Mensaje | Texto que se mostrará en la notificación. | `El pago de la factura F-123 ha sido confirmado.` |
| E | FechaCreacion | Timestamp de cuándo se generó la notificación. | `2025-09-25T10:30:00Z` |
| F | Estado | Estado de la notificación. | `NO_LEIDO` / `LEIDO` |
| G | FechaLectura | Timestamp de cuándo se marcó como leída. | `2025-09-25T11:00:00Z` |

### 3.2. Modificación del Flujo de Sincronización

La función `sincronizarRegistrosBanco` en `Codigo.js` será modificada. Además de guardar los pagos en las hojas `BANCO-*`, deberá:

1.  Identificar un pago recién sincronizado.
2.  Generar un `Mensaje` descriptivo.
3.  Crear una nueva fila en la hoja `Notificaciones` con el `ID_Registro_Pago`, `CodVendedor`, `Mensaje`, `FechaCreacion` y el `Estado` inicial en `NO_LEIDO`.

Esto desacopla la *creación* de la *consulta* de notificaciones, mejorando la eficiencia.

### 3.3. Nuevas Funciones de Backend (Codigo.js)

Se crearán o modificarán las siguientes funciones:

- **`getNotificacionesUsuario(token)`:**
    - Reemplazará a `getNotificacionesPagosConfirmados`.
    - Consultará **únicamente** la hoja `Notificaciones`.
    - Devolverá un objeto con:
        - Una lista de las últimas 5 notificaciones no leídas.
        - El número total de notificaciones no leídas (para un contador en la UI).
- **`marcarNotificacionesComoLeidas(token, idsNotificaciones)`:**
    - Recibirá un array de `ID_Notificacion`.
    - Actualizará el `Estado` a `LEIDO` y registrará la `FechaLectura` en la hoja `Notificaciones` para las filas correspondientes.
- **`getHistorialNotificaciones(token)`:**
    - Permitirá obtener todas las notificaciones (leídas y no leídas) de un usuario, posiblemente con paginación, para un historial completo.

### 3.4. Cambios en la Interfaz de Usuario (Frontend - Index.html)

1.  **Icono de Campana (`<i class="bell-icon">`):**
    - Se añadirá en la cabecera, junto al nombre de usuario.
    - Mostrará un **contador** con el número de notificaciones no leídas.
2.  **Panel de Notificaciones:**
    - Al hacer clic en la campana, se desplegará un panel con la lista de notificaciones no leídas.
    - Al abrir el panel, se llamará a `marcarNotificacionesComoLeidas` para actualizar su estado en el backend.
    - El panel incluirá un enlace "Ver todas", que podría dirigir a una nueva página o modal con el historial completo.

## 4. Plan de Implementación por Pasos

### Paso 1: Preparar el Backend (Hoja de Cálculo)

1.  Crear la nueva hoja `Notificaciones` en el Spreadsheet con las columnas definidas en la sección 3.1.
2.  Actualizar `SheetManager.SHEET_CONFIG` en `Codigo.js` para incluir la configuración de la nueva hoja.

### Paso 2: Modificar el Proceso de Creación de Notificaciones

1.  Editar la función `sincronizarRegistrosBanco` en `Codigo.js`.
2.  Después de que un nuevo registro de pago es añadido a una hoja `BANCO-*`, añadir la lógica para insertar una nueva fila en la hoja `Notificaciones`.

### Paso 3: Implementar la Gestión de Notificaciones (API)

1.  Eliminar o descontinuar la función `getNotificacionesPagosConfirmados`.
2.  Implementar las nuevas funciones en `Codigo.js`: `getNotificacionesUsuario`, `marcarNotificacionesComoLeidas` y `getHistorialNotificaciones`.
3.  Exponer estas funciones para que puedan ser llamadas desde el frontend.

### Paso 4: Desarrollar la Interfaz de Usuario (UI)

1.  En `Index.html` y `styles.html`, añadir el icono de campana y el contador.
2.  Diseñar y maquetar el panel desplegable de notificaciones.
3.  En el JavaScript de `Index.html`:
    - Al cargar la página, llamar a `getNotificacionesUsuario` para obtener el número de no leídas y actualizar el contador.
    - Implementar el evento `onclick` en la campana para mostrar el panel y llamar a la API para obtener el detalle de las notificaciones.
    - Al mostrar el panel, invocar `marcarNotificacionesComoLeidas` para limpiar el contador.

### Paso 5: Pruebas y Despliegue

1.  **Prueba unitaria:** Verificar que `sincronizarRegistrosBanco` cree correctamente las notificaciones.
2.  **Prueba de integración:** Simular un inicio de sesión de vendedor y comprobar:
    - Que el contador de la campana muestre el número correcto.
    - Que el panel muestre las notificaciones.
    - Que el estado cambie a "leído" después de verlas.
    - Que el contador se reinicie a cero.
3.  Desplegar la nueva versión del Web App.

## 5. Consideraciones Adicionales

- **Limpieza de datos:** Se podría diseñar una estrategia futura para archivar o eliminar notificaciones muy antiguas (ej: más de 1 año) para mantener el rendimiento de la hoja `Notificaciones`.
- **Notificaciones por Email:** Una vez que el sistema persistente esté en su lugar, se podría extender para enviar resúmenes diarios por correo electrónico a los vendedores con sus pagos confirmados.
