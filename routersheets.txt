# Propuesta: Minimizar lecturas de Sheets y Particionamiento lógico de datos (GAS)

Objetivo
- Reducir latencia y consumo de cuotas de Google Sheets.
- Evitar leer “todo” y mover el peso de cómputo al momento de escritura y a caches rápidos.
- Mantener consultas rápidas para “recientes” y escalables para históricos.

---

## 1) Minimizar lecturas de Sheets

### Principios
- Leer solo lo necesario: columnas y filas estrictamente requeridas.
- Preferir “últimos N” y rangos fijos en vez de recorrer toda la hoja.
- Precalcular e indexar en la escritura (materializar vistas de consulta frecuente).
- Cachear resultados y metadatos (última fila, índices) en CacheService.

### Técnicas concretas

1) Proyección de columnas (solo las necesarias)
- Si solo vas a construir la tabla de registros, evita traer columnas no mostradas.
```javascript
function readProjectedRange_(sheet, startRow, numRows) {
  // Ejemplo: columnas A:I (9 col), ajusta a tus cabeceras reales
  const lastCol = 9;
  return sheet.getRange(startRow, 1, numRows, lastCol).getValues();
}
```

2) Ventana móvil de “últimos N”
- Para “recientes” (por vendedor o global), traer únicamente los últimos N registros.
```javascript
function readLastN_(sheet, N) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const headerRow = 1;
  const startRow = Math.max(headerRow + 1, lastRow - N + 1);
  const numRows = lastRow - startRow + 1;
  return readProjectedRange_(sheet, startRow, numRows);
}
```

3) Rango por fecha (si el log está ordenado por inserción)
- Si guardas la fecha/hora en una columna, puedes usar una búsqueda binaria para delimitar filas por rango (requiere datos ordenados crecientemente por fecha).
```javascript
function findRowBoundsByDate_(values, dateColIndex, from, to) {
  // values: matriz sin header; dateColIndex base 0; from/to: Date
  // Simplificado: lineal o binaria si garantizas orden. Aquí, lineal con corte temprano.
  const res = [];
  for (let i = values.length - 1; i >= 0; i--) { // desde el final (más reciente)
    const d = new Date(values[i][dateColIndex]);
    if (d >= from && d <= to) res.push(values[i]);
    if (d < from) break; // corte temprano si está ordenado
  }
  return res.reverse();
}
```

4) Índices en memoria/caché (para búsquedas por vendedor/cliente)
- Mantén un índice ligero `{ vendedor -> [rowPointers] }` con:
  - hoja (partition), fila, fecha.
  - Actualiza el índice al insertar un registro.
  - Guárdalo en CacheService (y sombra en PropertiesService si quieres resiliencia).
- Consultas rápidas: obtienes los punteros y luego haces lecturas directas solo a esas filas (ideal con `Sheets Advanced Service` batchGet si están en distintas hojas).

5) Vistas materializadas “Últimos N por vendedor”
- Mantener por vendedor una hoja “LN_{VEND}” con los últimos N registros, actualizada al escribir.
- Consultar “recientes” solo leyendo desde esa hoja, que es pequeña y acotada.

6) Lotes en una sola llamada
- Cuando requieras múltiples rangos del mismo spreadsheet, usa una sola operación:
  - Apps Script nativo: `getRangeList(['A2:I101', 'A205:I305']).getRanges().map(r => r.getValues())`
  - Sheets Advanced Service: `Sheets.Spreadsheets.values.batchGet(...)` para varios rangos.

7) Evita `getDisplayValues()` salvo para UI
- Usa `getValues()` (tipos nativos) y da formato en cliente/plantilla. Menos bytes y conversión.

8) Metadata en caché
- Cachea `lastRow`, mapeo de cabeceras (nombre -> índice), y nombres de pestañas/particiones.
- Invalida/actualiza estos metadatos solo al escribir, no en cada lectura.

---

## 2) Particionamiento lógico de datos

### Estrategias de partición

- Por mes/año (recomendada para logs de cobranzas)
  - Hojas: `REG_2025_09`, `REG_2025_10`, ...
  - Ventajas: tamaño acotado por hoja; consultas a “recientes” quedan en la partición activa.
  - Historias cruzadas (rango > 1 mes) se resuelven leyendo 1–3 hojas en la mayoría de casos.

- Por vendedor (útil si hay alto tráfico por vendedor)
  - Hojas: `V_001_2025_09`, `V_002_2025_09`, ...
  - Ventajas: paraleliza lectura/escritura por vendedor y reduce contención.
  - Complejidad: mayor cantidad de hojas; necesitas un router robusto.

- Híbrido (mes + vendedor)
  - Para cargas muy altas. Evaluar solo si los volúmenes lo ameritan.

### Router de particiones

- Función para resolver a qué hoja escribir/leer según fecha (y opcionalmente vendedor).
- Crea la hoja si no existe (con cabecera estándar) y congela la fila de títulos.

```javascript
function partitionSheetName_(date, vendedorCode) {
  const yyyy = date.getFullYear();
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  // Variante por mes
  return `REG_${yyyy}_${mm}`;
  // Variante híbrida:
  // return `V_${vendedorCode}_${yyyy}_${mm}`;
}

function ensurePartitionSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);
  }
  return sh;
}
```

### Escrituras con actualización de vistas/índices

- Al insertar un registro:
  1) Determina partición (por fecha y, si aplica, vendedor).
  2) Inserta en la hoja de partición.
  3) Actualiza:
     - Hoja “LN_{VEND}” (últimos N del vendedor).
     - Hoja “LN_ALL” (últimos N global).
     - Índice `IDX` (fila, hoja, vendedor, fecha).
  4) Invalida o actualiza entradas de caché relacionadas.

```javascript
function appendRegistro_(ss, record, header) {
  const date = new Date(record.fechaEnvio || new Date());
  const vend = record.vendedorCodigo || record.vendedor;
  const sheetName = partitionSheetName_(date /*, vend */);
  const sh = ensurePartitionSheet_(ss, sheetName, header);

  const row = header.map(h => record[h] ?? '');
  sh.appendRow(row);

  // Actualizar vistas materializadas
  updateUltimosNPorVendedor_(ss, vend, row, header, 100); // N = 100
  updateUltimosNGlobal_(ss, row, header, 100);

  // Actualizar índice (hoja IDX)
  updateIndex_(ss, { hoja: sheetName, fila: sh.getLastRow(), vendedor: vend, fecha: date });

  // Invalidar/actualizar cachés clave
  // CacheService: script:registros:user:{email}:page=1, script:registros:all:page=1, etc.
}
```

### Consulta de históricos

- Por rango de fechas:
  1) Determina las particiones afectadas (p. ej., mes actual y anterior).
  2) Lee solo esas hojas y aplica filtro por fecha/vendedor en memoria (los datasets ya son más acotados).
- Por vendedor:
  - Primero intenta “LN_{VEND}” (si el rango solicitado cabe).
  - Si no, cae a las particiones relevantes.

---

## 3) Índices y Vistas materializadas

### Índice (hoja `IDX`)
- Columnas sugeridas: `hoja`, `fila`, `fechaISO`, `vendedor`, `cliente`, `factura`, `monto`
- Al escribir, append en `IDX`.
- Búsquedas rápidas:
  - Filtra `IDX` por vendedor y rango de fechas para obtener punteros (hoja/fila).
  - Lee filas exactas con `getRangeList` para minimizar lecturas.

```javascript
function updateIndex_(ss, idxRow) {
  const sh = ensurePartitionSheet_(ss, 'IDX', ['hoja','fila','fechaISO','vendedor','cliente','factura','monto']);
  sh.appendRow([
    idxRow.hoja,
    idxRow.fila,
    new Date(idxRow.fecha).toISOString().slice(0,10),
    idxRow.vendedor || '',
    idxRow.cliente || '',
    idxRow.factura || '',
    idxRow.monto || ''
  ]);
}
```

### Vistas materializadas “LN” (últimos N)
- `LN_ALL`: últimos N globales
- `LN_{VEND}`: últimos N por vendedor
- Mantener tamaño acotado (sobrescribir cuando excede N).
- Consultas de “recientes” usan estas vistas, no la tabla grande.

```javascript
function updateUltimosNPorVendedor_(ss, vend, row, header, N) {
  const name = `LN_${vend}`;
  const sh = ensurePartitionSheet_(ss, name, header);
  sh.appendRow(row);
  // Recortar si excede
  const lastRow = sh.getLastRow();
  const headerRow = 1;
  const dataRows = lastRow - headerRow;
  if (dataRows > N) {
    const toDelete = dataRows - N;
    sh.deleteRows(headerRow + 1, toDelete);
  }
}
```

---

## 4) Triggers de mantenimiento

- Rotación mensual
  - A inicios de mes, crea la partición del mes con cabecera.
  - Opcional: mover las últimas X filas del mes anterior a la vista `LN_ALL` si aplica.

- Compactación de vistas LN
  - Asegurar que `LN_*` no excedan N.

- Precalentamiento de cachés (+ índices)
  - Recalcular cachés “recientes” (all y top vendedores) a primera hora.

```javascript
function rotacionMensual_() {
  const ss = SpreadsheetApp.getActive();
  const now = new Date();
  const name = partitionSheetName_(now /*, null */);
  ensurePartitionSheet_(ss, name, getHeader_());
}
```

---

## 5) Plan de migración

1) Definir cabecera estándar (orden y nombres de columnas).
2) Implementar router de particiones (por mes).
3) Redirigir la escritura de nuevos registros a la partición activa.
4) Crear y poblar `LN_ALL` y `LN_{VEND}` al vuelo con cada inserción.
5) Adaptar lectura de “recientes” para usar `LN_*`.
6) Implementar consultas por rango: resolver particiones afectadas y leer solo esas hojas.
7) Opcional: retro-migrar históricos a hojas por mes (script de una sola vez).
8) Añadir triggers de rotación mensual y precalentamiento de cachés.
9) Métricas de tiempo de respuesta y lecturas/llamadas (ajustar N y TTL según uso).

---

## 6) Riesgos y mitigaciones

- Complejidad de consultas跨-partición:
  - Mitigar con índice `IDX` y/o limitar rangos de fechas en UI.
- Consistencia de vistas LN:
  - Actualizar atómicamente tras escribir (LockService si hay alta concurrencia).
- Límite de hojas:
  - Años de operación → muchas hojas; evaluar particionar por archivo a partir de cierto umbral (p.ej. un spreadsheet por año).

---

## 7) Métricas y validación

- Medir tiempo de:
  - Escritura + actualización de LN/IDX.
  - Lectura de “recientes” y de rango típico.
- Contar lecturas a Sheets por endpoint (antes/después).
- Ajustar N (últimos) y columnas proyectadas según patrones de uso.

---

## 8) Checklist de adopción

- [ ] Router de particiones (por mes).
- [ ] Escritura en partición + actualización de LN_ALL y LN_{VEND}.
- [ ] Consulta “recientes” desde LN_*.
- [ ] Consulta por rango leyendo solo particiones afectadas.
- [ ] Índice `IDX` (opcional, recomendado si se requieren búsquedas combinadas).
- [ ] Triggers: rotación mensual, compactación LN, precalentamiento de cachés.
- [ ] Métricas básicas de rendimiento.

---

Con estas prácticas, las lecturas se vuelven acotadas y predecibles, evitando escanear hojas completas. El particionamiento por mes y las vistas materializadas permiten mantener una experiencia fluida, incluso con crecimiento sostenido de datos.
