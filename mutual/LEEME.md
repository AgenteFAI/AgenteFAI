# Buscador de información — Mutual de Seguridad CChC

Este programa es como un **asistente sentado a tu lado**: tú abres la puerta
(entras al portal con tu RUT y clave) y él recorre las carpetas del portal.
Saca "fotocopias" (capturas de pantalla), copia las tablas a un Excel y guarda los PDF.

Busca 3 tipos de información:

| Categoría              | Palabras que busca (ejemplos)                       |
|------------------------|-----------------------------------------------------|
| Certificados           | certificado, afiliación, adhesión, siniestralidad, tasa |
| Accidentes y licencias | accidente, denuncia, DIAT, DIEP, licencia, reposo   |
| Facturas y pagos       | factura, pago, cotización, deuda, estado de cuenta  |

> **Tu clave nunca pasa por el programa.** La escribes tú en el navegador, como siempre.

---

## 1. Instalar (solo la primera vez)

1. Instala **Python 3** desde https://www.python.org/downloads/
   - En Windows, marca la casilla **"Add Python to PATH"** al instalar.
2. Abre una terminal **en esta carpeta** (`mutual`):
   - Windows: abre la carpeta, escribe `cmd` en la barra de dirección y presiona Enter.
   - Mac: clic derecho sobre la carpeta → "Nueva terminal en la carpeta".
3. Copia y pega estos dos comandos (uno a la vez):

```
pip install -r requisitos.txt
python -m playwright install chromium
```

En Mac usa `pip3` y `python3` en vez de `pip` y `python`.

## 2. Usar

```
python buscar_mutual.py
```

El programa tiene 3 pasos:

1. **Ingresar**: se abre un navegador. Entra a la Sucursal Virtual de Empresas
   con tu RUT y clave. Cuando ya estés dentro, vuelve a la terminal y presiona **Enter**.
2. **Búsqueda automática**: el programa recorre solo los menús que coinciden con
   las palabras de la tabla de arriba. **No toques el navegador** mientras tanto.
3. **Modo guiado**: ahora navegas tú. Entra a cualquier página que quieras guardar
   y presiona **Enter** en la terminal. Si haces clic en un botón "Descargar", el
   archivo también se guarda. Escribe `salir` para terminar.

## 3. Resultados

Todo queda en `resultados_mutual/<fecha y hora>/`:

| Archivo o carpeta | Qué contiene                                              |
|-------------------|-----------------------------------------------------------|
| `resumen.xlsx`    | **Abre este primero.** Índice, archivos, mapa del portal y una hoja por tabla |
| `descargas/`      | Los PDF y otros archivos                                  |
| `capturas/`       | Una imagen de cada página visitada                        |
| `textos/`         | El texto de cada página (para buscar con Ctrl+F)          |

## Seguridad

- El programa **nunca** abre enlaces como "Cerrar sesión", "Pagar", "Eliminar" o "Anular".
- Solo **mira y copia**: no llena formularios ni envía nada.
- La carpeta `resultados_mutual` tiene datos de tu empresa y de tus trabajadores.
  Por eso está bloqueada para que **no se suba a GitHub** (archivo `.gitignore`).

## Si algo no funciona

- **"No encontré el navegador"** → ejecuta `python -m playwright install chromium`.
- **La búsqueda automática no encuentra nada** → algunos menús son botones, no enlaces.
  Usa el modo guiado (paso 3), que funciona con cualquier página.
- **Quieres buscar otras palabras** → cambia la lista `PALABRAS_CLAVE` al comienzo
  de `buscar_mutual.py`.
- **Para mejorar el programa**, puedes compartir la hoja **"Mapa del portal"** del Excel
  (son los nombres y direcciones del menú). Revísala antes, por si aparece algún dato personal.
