# Calculadora Índice FACPCE — AXI

Herramienta para consultar el índice FACPCE (Res. JG 539/18) usado en ajuste por inflación contable (AXI).

---

## Estructura del proyecto

```
facpce-axi/
├── index.html              ← App web (abrir en browser)
├── update_indices.py       ← Script de actualización de datos
├── indices.json            ← Datos generados (se sobreescribe con cada update)
└── .github/
    └── workflows/
        └── update.yml      ← Automatización en GitHub Actions
```

---

## Cómo usar

### Opción A — Uso local simple (sin GitHub)

1. Instalar dependencias:
   ```bash
   pip install requests beautifulsoup4 openpyxl
   ```

2. Correr el script cuando necesites actualizar:
   ```bash
   python update_indices.py
   ```
   Esto genera/actualiza `indices.json`.

3. Abrir `index.html` en el browser.
   - La app carga automáticamente `indices.json` si está en la misma carpeta.
   - Si no encuentra el JSON, intenta descargar el Excel directamente via proxy.
   - El botón **Actualizar** en la barra superior refresca los datos.

---

### Opción B — GitHub Pages (recomendado)

Con esta opción la app queda online y accesible desde cualquier dispositivo.

**Setup inicial:**

1. Crear un repositorio en GitHub y subir todos los archivos.

2. Activar GitHub Pages:
   - Ir a Settings → Pages
   - Source: **Deploy from a branch**
   - Branch: `main`, folder: `/ (root)`
   - Guardar

3. Tu app va a estar disponible en:
   `https://TU-USUARIO.github.io/NOMBRE-REPO/`

**Para actualizar los datos:**

1. Ir a tu repositorio en GitHub
2. Click en la pestaña **Actions**
3. Click en **"Actualizar Índice FACPCE"** en el panel izquierdo
4. Click en **"Run workflow"** → **"Run workflow"**
5. En ~30 segundos los datos se actualizan y GitHub Pages los publica

El botón **Actualizar** en la app llama primero a `indices.json` (datos del servidor) y, si no lo encuentra, intenta el proxy CORS.

---

## Frecuencia de actualización

FACPCE publica el Excel con el índice del mes anterior, generalmente entre el día 3 y 10 del mes siguiente.

El workflow tiene un respaldo automático el **día 5 de cada mes** (configurable en `update.yml`), pero la idea es que lo corrás manualmente cuando veas que publicaron el dato nuevo.

---

## Funcionalidades de la app

- **Últimos 12 meses**: tabla con índice, variación mensual y acumulado indexado
- **Gráfico**: variación mensual en barras
- **Calcular acumulado indexado** entre dos períodos cualesquiera (mm/aaaa)
- **Actualizar valor nominal** a otra fecha usando el coeficiente AXI exacto

### Fórmulas utilizadas

**Variación mensual:**
```
var_m = (índice_m / índice_m-1 − 1) × 100
```

**Acumulado indexado entre período A y período B:**
```
acum = (índice_B / índice_A − 1) × 100
```
> ⚠️ No es suma de variaciones mensuales. Es el cociente real de índices (capitalización compuesta).

**Coeficiente AXI para actualizar un valor:**
```
valor_actualizado = valor_nominal × (índice_destino / índice_base)
```

---

## Actualizar la URL del Excel

Cuando FACPCE publica un nuevo archivo, la URL cambia. El script `update_indices.py` intenta detectarla automáticamente scrapando el sitio. Si falla, usa la URL de respaldo definida en la variable `FALLBACK_URL` al inicio del script.

Actualizar esa variable con la URL más reciente si la detección automática no funciona.
