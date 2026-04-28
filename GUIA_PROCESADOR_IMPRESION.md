# Guía: Procesador de Archivos de Impresión Carabineros

## ¿Qué es?
Un procesador que lee archivos de impresión de notificaciones de Carabineros (formato XLS/XLSX), extrae los IDs de notificación y genera un CSV listo para automatización.

## Archivo de Entrada
- **Formato**: XLS o XLSX
- **Contenido**: Reportes de impresión con columnas que incluyen ID de notificación
- **Ejemplos de nombres de columna soportados**:
  - `ID`, `ID_NOTIFICACION`, `Notificación`, `numero`, etc.
  - El sistema detecta automáticamente cuál es

## Flujo de Uso

### Paso 1: Seleccionar tipo "Carabineros"
```
Selector: [Terreno | Carabineros]
                      ↓ (selecciona)
```
→ Aparecerá la sección "PROCESADOR DE IMPRESIÓN"

### Paso 2: Seleccionar archivo
```
[Seleccionar archivo] → Elige tu archivo XLS/XLSX
```
→ Se mostrará la ruta del archivo

### Paso 3: (Opcional) Previsualizar
```
[Previsualizar] → Verá los primeros 5 IDs encontrados
```
→ Confirma que el archivo se leyó correctamente

### Paso 4: Procesar
```
[Procesar] → Se abre diálogo pidiendo:
  ├─ Hora (ej: 1205)
  └─ Código (ej: D2)
```
→ Se generará el CSV

### Paso 5: Usar en Automatización
```
¿Deseas usar este archivo en el procesador de Carabineros?
[Sí] → Se carga automáticamente
      → Se abre el tab de procesador normal
      → Puedes ejecutar la automatización
```

## Archivo de Salida

### Nombre
```
[nombre_original]_procesado.csv
```

### Estructura
```
rit,anio,id_notificacion,codigo,hora,observacion
,,,12345,D2,1205,
,,,12346,D2,1205,
,,,12347,D2,1205,
```

**Notas**:
- `rit` y `anio` están vacíos (se busca solo por id_notificacion)
- `codigo` y `hora` se repiten para todos los registros
- `observacion` vacío para que lo llenes manualmente si necesitas

## Ejemplo Completo

### 1. Tienes archivo: `DetalleImpresion_13-04-2026.xls`
```
ID     | Nombre         | Estado
-------|----------------|-------
12345  | Juan Pérez     | Activo
12346  | María García   | Activo
12347  | Carlos López   | Activo
```

### 2. Seleccionas el archivo en la interfaz
### 3. Ejecutas "Procesar" con:
- Hora: `1205`
- Código: `D2`

### 4. Se genera: `DetalleImpresion_13-04-2026_procesado.csv`
```
rit,anio,id_notificacion,codigo,hora,observacion
,,,12345,D2,1205,
,,,12346,D2,1205,
,,,12347,D2,1205,
```

### 5. La interfaz ofrece cargar el CSV
- Si aceptas: se carga automáticamente en el procesador de Carabineros
- Puedes ejecutar la automatización directamente

## Archivos Soportados

| Formato | ¿Funciona? | Notas |
|---------|-----------|-------|
| .xlsx   | ✅ Sí     | Excel moderno (mejor compatibilidad) |
| .xls    | ✅ Sí     | Excel antiguo |
| .csv    | ✅ Sí     | Si tiene columna de ID |

## Detección Automática de Columnas

El sistema intenta detectar la columna de ID de notificación en este orden:
1. `ID`
2. `id`
3. `ID_NOTIFICACION`
4. `id_notificacion`
5. `ID Notificación`
6. `Número`
7. `numero`
8. `Notificación`
9. `notificacion`
10. `NOTIFICACION`
11. `ID_NOTI`
12. `id_noti`

Si no encuentra ninguno, usa la primera columna como fallback.

## Validaciones

| Regla | Descripción |
|-------|------------|
| Hora | Debe tener 4 dígitos (HHMM), ej: 1205 |
| Código | No puede estar vacío, ej: D2 |
| Archivo | Debe existir y ser un Excel válido |
| IDs | Se ignoran filas vacías y headers |

## Solución de Problemas

### "No se encontraron IDs de notificación en el archivo"
- Verifica que el archivo tenga una columna con IDs
- Verifica que el nombre de la columna sea reconocible (ver lista arriba)
- Intenta usar "Previsualizar" primero

### "Hora debe tener formato HHMM"
- La hora debe tener EXACTAMENTE 4 dígitos
- Ej: ✅ `1205`, ❌ `205`, ❌ `12:05`

### "Código no puede estar vacío"
- Ingresa un código válido, ej: `D2`

### El archivo se ve raro en Excel
- Asegúrate de que sea un Excel válido
- Intenta guardar como XLSX desde Excel
- Algunos archivos .xls pueden estar dañados

## Integración con Procesador Normal

Una vez generado el CSV desde impresión:

1. **Opción A**: Cargarlo automáticamente (botón en diálogo)
   - Se abre el tab de Carabineros con el CSV cargado
   - Puedes ejecutar la automatización

2. **Opción B**: Usarlo manualmente
   - Guarda el CSV en una carpeta
   - Cárgalo manualmente en el procesador normal

## Mejores Prácticas

1. **Previsualizar primero** para confirmar que se leyó bien
2. **Usar un código y hora consistentes** para el lote
3. **Revisar el CSV generado** antes de ejecutar automatización
4. **Mantener una copia** del archivo de impresión original
5. **Agregar observaciones** si necesitas notas personalizadas

## FAQ

**P: ¿Puedo cambiar código y hora después?**
A: Sí, edita el CSV generado en Excel y vuelve a procesarlo

**P: ¿Qué pasa con registros duplicados?**
A: Se incluyen todos. El procesador normal elimina duplicados si es necesario

**P: ¿Puedo procesar múltiples archivos?**
A: Sí, repite el proceso para cada archivo

**P: ¿Los CSVs generados se guardan automáticamente?**
A: Sí, en la misma carpeta que el archivo de impresión original

---

**Última actualización**: 27-04-2026
