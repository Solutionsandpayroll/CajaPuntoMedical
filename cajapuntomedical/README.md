# Sistema de Caja Menor - Punto Medical

Sistema web para el registro y control de la caja y caja menor, integrado con Google Sheets mediante Google Apps Script.

## 📁 Estructura del Proyecto

```
cajapuntomedical/
├── index.html              # Aplicación web (HTML + CSS)
├── app.js                  # Lógica JavaScript de la aplicación
├── Code.gs                 # Apps Script principal (Acople de Caja + Consulta)
├── CajaMenor.gs            # Apps Script para Caja Menor (Spreadsheet separado)
└── README.md               # Este archivo
```

## 🚀 Guía de Instalación

### Paso 1: Configurar Google Sheets

Tienes **dos Spreadsheets** independientes:

| Spreadsheet | Uso |
|---|---|
| `Code.gs` | Acople de Caja mensual + Consultas + hoja PLANTILLA |
| `CajaMenor.gs` | Registros de Caja Menor |

Para cada uno, copia el **ID del Spreadsheet** desde la URL:
```
https://docs.google.com/spreadsheets/d/[ESTE_ES_TU_ID]/edit
```

#### Hoja PLANTILLA (requerida en el Spreadsheet principal)
El Spreadsheet principal debe tener una hoja llamada exactamente `PLANTILLA`. Esta hoja se usa como plantilla para generar el bloque de cada acople diario y para exportar el comprobante PDF. Su estructura es:

| Celda | Contenido |
|---|---|
| `B2` | Fecha del acople (se llena automáticamente) |
| `A4:A18` | Conceptos de entrada (se llenan automáticamente) |
| `B4:B18` | Montos de entrada (se llenan automáticamente) |
| `C4:C18` | Conceptos de salida (se llenan automáticamente) |
| `D4:D18` | Montos de salida (se llenan automáticamente) |
| `B19` | Total entradas (se llena automáticamente) |
| `D19` | Total salidas (se llena automáticamente) |
| `B20` | Saldo del día (se llena automáticamente) |

### Paso 2: Configurar Google Apps Script

#### Script principal (`Code.gs`)
1. En el Spreadsheet principal, ve a **Extensiones** > **Apps Script**
2. Borra el contenido por defecto y pega el contenido de `Code.gs`
3. Actualiza el ID en la línea de configuración:
   ```javascript
   const SPREADSHEET_ID = 'TU_ID_DEL_SPREADSHEET_PRINCIPAL';
   ```
4. Guarda el proyecto

#### Script de Caja Menor (`CajaMenor.gs`)
1. En el Spreadsheet de Caja Menor, ve a **Extensiones** > **Apps Script**
2. Borra el contenido por defecto y pega el contenido de `CajaMenor.gs`
3. Actualiza el ID:
   ```javascript
   const SPREADSHEET_ID = 'TU_ID_DEL_SPREADSHEET_CAJA_MENOR';
   ```
4. Guarda el proyecto

### Paso 3: Desplegar como Web App (ambos scripts)

Repite este proceso para **cada uno** de los dos scripts:

1. Haz clic en **Implementar** > **Nueva implementación**
2. Configura:
   - **Tipo**: Aplicación web
   - **Ejecutar como**: Yo (tu cuenta)
   - **Quién tiene acceso**: Cualquier persona *(no "con cuenta Google")*
3. Haz clic en **Implementar** y autoriza cuando se solicite
4. Copia la URL generada

> ⚠️ **Importante:** Cada vez que modifiques el código del script debes crear una **nueva implementación**. Editar una implementación existente no actualiza el endpoint.

### Paso 4: Configurar la Aplicación Web

En `app.js`, actualiza las dos URLs al inicio del archivo:

```javascript
const SCRIPT_URL = 'URL_DEL_SCRIPT_PRINCIPAL';           // Code.gs
const CAJA_MENOR_SCRIPT_URL = 'URL_DEL_SCRIPT_CAJA_MENOR'; // CajaMenor.gs
const PLANTILLA_GID = 'GID_DE_TU_HOJA_PLANTILLA';         // ID numérico de la hoja PLANTILLA
```

Para obtener el `PLANTILLA_GID`, abre la hoja PLANTILLA en Sheets y revisa la URL:
```
...spreadsheets/d/.../edit#gid=ESTE_ES_EL_GID
```

### Paso 5: Publicar la Aplicación

#### Opción A: GitHub Pages (Gratuito)
1. Sube el proyecto a un repositorio de GitHub
2. Ve a **Settings** > **Pages**, selecciona la rama
3. Disponible en `https://tuusuario.github.io/nombrerepo/`

#### Opción B: Vercel (Gratuito)
1. Crea cuenta en [vercel.com](https://vercel.com)
2. Conecta tu repositorio de GitHub y despliega

#### Opción C: Netlify (Gratuito)
1. Crea cuenta en [netlify.com](https://netlify.com)
2. Arrastra la carpeta del proyecto y obtén tu URL

#### Opción D: Uso Local
Abre `index.html` directamente en el navegador.
> Nota: Las peticiones GET funcionan normalmente. Las POST usan `Content-Type: text/plain` para evitar el bloqueo CORS de Apps Script.

---

## 📋 Estructura de las Hojas

### Hojas de mes (Spreadsheet principal)
Cada mes crea automáticamente una hoja con nombre `MES AÑO` (ej: `MARZO 2026`):

| Col A | Col B | Col C | Col D |
|---|---|---|---|
| CAJA DE MES AÑO (fusionado A1:D1) | | | |
| CONCEPTO DE ENTRADA | ENTRADAS DE DINERO | CONCEPTO DE SALIDA | SALIDAS DE DINERO |
| *concepto* | *monto* | *concepto* | *monto* |
| TOTAL | `=SUM(B3:B18)` | TOTAL | `=SUM(D3:D18)` |

Cada acople diario ocupa un bloque de 19 filas copiado desde la hoja PLANTILLA.

#### Celdas especiales para Consulta
| Celda | Uso |
|---|---|
| `I2` | Valor de caja física contada. **Si está vacía**, la Caja Final = Caja Total (sin aplicar Sobra/Falta) |
| `L2` | Fórmula de SOBRA/FALTA. Solo se aplica si `I2` tiene valor |

### Hojas de Caja Menor (Spreadsheet separado)
Cada mes crea una hoja `MES AÑO` con encabezados en fila 9 y datos desde fila 10:

| Col C | Col D | Col E | Col F | Col G | Col H | Col I |
|---|---|---|---|---|---|---|
| FECHA | DETALLE | NIT | PROVEEDOR | N° FAC | VALOR | OBSERVACIONES |

---

## ✨ Funcionalidades

### Pestaña: Acople de Caja

**Agregar Movimientos**
- Selecciona el tipo (Entrada / Salida), ingresa el monto y el concepto
- Los movimientos se acumulan en una lista antes de confirmar
- Puedes eliminar cualquier movimiento antes del cierre

**Cierre de Caja**
- Envía todos los movimientos del día al script principal (`Code.gs`)
- Llena la hoja PLANTILLA con los datos del día y la copia al mes correspondiente
- Muestra un modal con enlace para **descargar el comprobante en PDF** (rango `A2:D20` de la hoja PLANTILLA)
- Soporta hasta **15 entradas y 15 salidas** por día

### Pestaña: Consultar Caja

**Resumen mensual**
Selecciona mes y año para ver:
- **Total Entradas** del mes
- **Total Salidas** del mes
- **Saldo** (diferencia entradas - salidas acumulada en el mes)
- **Caja Total del Mes** — saldo acumulado histórico desde AGOSTO 2025 más el saldo del mes actual
- **SOBRA/FALTA** — leído desde celda `L2` de la hoja (solo si `I2` tiene valor)
- **Caja Final**:
  - Si `I2` tiene valor → `Caja Total + Sobra/Falta`
  - Si `I2` está vacía → `Caja Final = Caja Total`

**Lógica de estructura histórica**
El sistema detecta automáticamente dos estructuras de hoja:
- **Antigua** (AGOSTO 2025 – FEBRERO 2026): lee totales en columnas/filas del formato anterior
- **Nueva** (MARZO 2026 en adelante): lee totales en el formato actual

La acumulación histórica arranca desde un valor base de `$3.326.232` (agosto 2025) y suma recursivamente mes a mes hasta el mes consultado.

### Pestaña: Caja Menor

**Registrar gasto**
Campos disponibles:
- Fecha (requerida)
- Detalle: CAFETERÍA / TRANSPORTE / ASEO / PARQUEADERO / ENVÍOS / OTROS (requerido)
- Proveedor
- NIT
- N° Factura
- Valor en COP (requerido)
- Observaciones

Los datos se guardan en el Spreadsheet de Caja Menor en la hoja del mes correspondiente. Solo se permiten registros desde **NOVIEMBRE 2025** en adelante.

---

## 🔧 Solución de Problemas

### Error de CORS en POST
Las peticiones POST usan `Content-Type: text/plain` en lugar de `application/json`. Esto evita el preflight `OPTIONS` que bloquea Apps Script. El script sigue leyendo `e.postData.contents` normalmente.

Si persiste el error, verifica:
1. El script esté desplegado con acceso **"Cualquier persona"** (no "con cuenta Google")
2. Haber creado una **nueva implementación** después del último cambio de código

### Los totales difieren del Excel
El script aplica redondeo a 2 decimales (`Math.round(v * 100) / 100`) en todos los valores antes de devolverlos, para evitar errores de acumulación de punto flotante en JavaScript (ej: `0.1 + 0.2 = 0.30000000000000004`).

Si la diferencia persiste, puede ser que tu hoja tenga fórmulas con `ROUND()` adicional — revisar las celdas de total del mes.

### La hoja PLANTILLA no existe
El script `Code.gs` lanza el error `No existe la hoja PLANTILLA` si no encuentra esa hoja. Créala manualmente en el Spreadsheet principal con el nombre exacto `PLANTILLA`.

### No se encuentra la hoja del mes
Si el mes no existe al hacer un acople, el script la crea automáticamente. Si no se crea, verifica que el `SPREADSHEET_ID` en el script sea correcto y que la cuenta tenga permisos de edición.

### Error de autorización
1. Vuelve a crear una nueva implementación en Apps Script
2. Autoriza **todas** las solicitudes de permisos (especialmente acceso a Sheets)

---

## 📞 Soporte

Para asistencia técnica, contactar al administrador del sistema.

---

**Punto Medical** - Sistema de Gestión de Caja © 2026