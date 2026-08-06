# Configurar Google de nuevo (hoja + Drive)

## Destinos (ya puestos en el código)

| Qué | Enlace / ID |
|-----|-------------|
| **Hoja Excel (Google Sheets)** | https://docs.google.com/spreadsheets/d/1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek/edit |
| ID hoja | `1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek` |
| **Carpeta PDF + Excel** | https://drive.google.com/drive/folders/1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv |
| ID carpeta | `1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv` |

---

## Paso 1 — Abrir la hoja e instalar el script

1. Abra la **hoja nueva**:  
   https://docs.google.com/spreadsheets/d/1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek/edit  
2. Menú: **Extensiones → Apps Script**.  
3. Si hay código viejo, **borre todo**.  
4. Abra en el PC el archivo:  
   `apps-script\Codigo.gs`  
   (o `PEGAR-EN-APPS-SCRIPT-v9-word.gs`)  
5. **Seleccione todo → Copiar → Pegar** en el editor de Apps Script.  
6. Compruebe que arriba diga:  
   - `SPREADSHEET_ID = '1obx8kVIXXxk2P65LVXxHju3I1iUf3SsPgozlxKmZ5ek'`  
   - `PDF_FOLDER_ID = '1dqy2HuWf6IYVGasuyrtff2id7Jf3aRXv'`  
7. **Ctrl + S** (guardar). Nombre del proyecto, por ejemplo: `Ficha Social UNCP`.

---

## Paso 2 — Publicar como aplicación web (CRÍTICO)

1. En Apps Script: **Implementar → Nueva implementación**.  
2. Tipo: **Aplicación web**.  
3. Configuración:  
   - **Descripción:** ficha uncp 2026  
   - **Ejecutar como:** Yo (su cuenta de Google)  
   - **Quién tiene acceso:** **Cualquier persona** ← obligatorio  
     Si pone “Solo yo” o “Usuarios de Google”, la entrevista muestra  
     **Error de red / Failed to fetch / 401**.  
4. **Implementar**.  
5. La primera vez pida **permisos**:  
   - Avanzado → Ir a (proyecto) (no seguro) → **Permitir**.  
6. **Copie la URL** que termina en `/exec`  
   Ejemplo: `https://script.google.com/macros/s/XXXX/exec`

### Si ya sale “Error de red al guardar”

1. Apps Script → **Implementar → Administrar implementaciones**  
2. **Lápiz** de la implementación activa  
3. **Quién tiene acceso = Cualquier persona**  
4. **Versión → Nueva versión**  
5. **Implementar**  
6. Pruebe la URL `/exec` en **ventana de incógnito**: debe verse **JSON** (`"ok":true`), **no** login de Google.  
7. En la entrevista: Ctrl+F5 y vuelva a guardar.

> Si ya existía una implementación antigua, use  
> **Implementar → Administrar implementaciones → lápiz → Nueva versión → Implementar**  
> y copie de nuevo la URL `/exec`.

---

## Paso 3 — Probar que el script responde

1. Pegue la URL `/exec` en el navegador (barra de direcciones) y Enter.  
2. Debe verse algo como JSON:

```json
{
  "ok": true,
  "version": "2026-03-pdf-v10-excel",
  "pdfReady": true,
  "sheetOk": true,
  "folderAccessible": true
}
```

3. Si `folderAccessible` es **false**:  
   - Abra la carpeta Drive.  
   - Comparta con la **misma cuenta** de Google que usa el script (como **Editor**).  
4. Si `sheetOk` es **false**:  
   - La cuenta del script debe poder editar la hoja.

---

## Paso 4 — Conectar la entrevista (pantalla local)

1. Abra la entrevista con **`Abrir entrevista.bat`** (no con doble clic al HTML).  
2. En la pantalla de resumen, entre a **admin** (5 clics en “UNCP · Ficha social” o la tecla que usen).  
3. Pulse **⚙ URL / config**.  
4. Pegue la URL `/exec` del Paso 2.  
5. **💾 Guardar URL**.  
6. Pulse **🔎 Diagnosticar**.  
   - Debe decir script listo / versión v10 y carpeta accesible.

**Borrar URL vieja del navegador (si falla):**  
En la consola del navegador (F12):

```js
localStorage.removeItem('fichaUNCP_sheetsUrl');
location.reload();
```

Luego vuelva a pegar la URL nueva.

---

## Paso 5 — Prueba completa

1. Complete una ficha de prueba (o cargue borrador).  
2. **☁ Guardar ficha en la nube**.  
3. Compruebe:  
   - **Hoja:** pestaña `Fichas` con una fila nueva + enlaces.  
   - **Carpeta Drive:** PDF + Excel (y Word si aplica).

---

## Resumen rápido

```
Hoja  →  Extensiones → Apps Script  →  pegar Codigo.gs  →  guardar
     →  Implementar → Aplicación web → Cualquier persona → copiar /exec

Entrevista → Admin → pegar /exec → Guardar → Diagnosticar → Guardar ficha
```

---

## Archivos del proyecto ya actualizados

- `apps-script/Codigo.gs` — IDs nuevos  
- `PEGAR-EN-APPS-SCRIPT-v9-word.gs` — igual  
- `ficha-entrevista.js` — enlaces hoja/carpeta  
- `index.html` — botones admin ↗ Hoja y ↗ Carpeta  

Solo falta **implementar el script en Google** y **pegar la URL `/exec`** en la entrevista.
