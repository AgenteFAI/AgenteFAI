# ===========================================
# Buscador de información - Mutual de Seguridad CChC
# ===========================================
#
# ¿Qué hace este programa?
#   Es como un asistente sentado a tu lado: tú abres la puerta
#   (ingresas con tu clave) y él recorre las carpetas del portal,
#   sacando fotocopias (capturas de pantalla), copiando las tablas
#   a un Excel y guardando los PDF que encuentre.
#
# Tu clave NUNCA pasa por este programa: la escribes tú en la
# ventana del navegador, igual que siempre.
#
# Busca 3 tipos de información:
#   - Certificados
#   - Accidentes y licencias (denuncias DIAT/DIEP, reposos)
#   - Facturas y pagos (cotizaciones, deudas, estados de cuenta)
#
# Cómo instalarlo y usarlo: lee LEEME.md (en esta misma carpeta).
# ===========================================

import os
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright


# -------------------------------------------
# CONFIGURACIÓN (puedes cambiar estos valores)
# -------------------------------------------

URL_INICIO = "https://www.mutual.cl/"

MAX_PAGINAS = 30        # tope de páginas que visita la búsqueda automática
PROFUNDIDAD = 2         # 1 = solo el menú principal, 2 = también sus sub-páginas
PAUSA_SEGUNDOS = 1.5    # espera entre páginas (para no saturar el portal)
MAX_ARCHIVOS_POR_PAGINA = 20

# Palabras que buscamos (en minúscula y SIN tildes).
# Basta con el comienzo de la palabra: "certificad" sirve para
# "certificado", "certificados", "certificación"...
PALABRAS_CLAVE = {
    "Certificados": [
        "certificad", "afiliacion", "adhesion", "siniestralidad", "tasa",
    ],
    "Accidentes y licencias": [
        "accidente", "denuncia", "diat", "diep", "licencia", "reposo",
        "siniestro", "enfermedad profesional", "subsidio",
    ],
    "Facturas y pagos": [
        "factura", "pago", "cotizacion", "deuda", "estado de cuenta",
        "cartola", "boleta", "nota de credito",
    ],
}

# Enlaces que el programa NUNCA abre por su cuenta (por seguridad).
PALABRAS_PROHIBIDAS = [
    "salir", "cerrar sesion", "logout", "log out", "logoff", "signout",
    "sign out", "desconectar", "pagar", "eliminar", "borrar", "anular",
]

# Si un enlace termina así, es un archivo para descargar (no una página).
EXTENSIONES_ARCHIVO = (".pdf", ".xls", ".xlsx", ".csv", ".zip", ".doc", ".docx")


# -------------------------------------------
# "CUADERNO DE NOTAS": aquí se va anotando lo encontrado
# -------------------------------------------

carpeta = None              # carpeta de resultados (se crea al iniciar)
paginas_guardadas = []      # una entrada por cada página guardada
tablas_guardadas = []       # cada tabla encontrada
archivos_guardados = []     # PDF y otros archivos descargados
urls_descargadas = set()    # para no descargar dos veces lo mismo
descargas_pendientes = []   # descargas que el navegador inició (al hacer clic)
mapa_del_portal = []        # todos los enlaces del menú principal
SITIO_INICIAL = ""          # sitio web donde ingresaste (se completa al ingresar)


# -------------------------------------------
# Pequeños códigos en JavaScript que se ejecutan DENTRO de la página.
# (El navegador entiende JavaScript, así que le "preguntamos" en su idioma.)
# -------------------------------------------

JS_ENLACES = """
() => Array.from(document.querySelectorAll('a[href]')).map(a => ({
    texto: (a.innerText || a.title || a.getAttribute('aria-label') || '')
               .replace(/\\s+/g, ' ').trim(),
    url: a.href
}))
"""

JS_TEXTO = "() => [document.body ? document.body.innerText : '']"

JS_TITULOS = """
() => [Array.from(document.querySelectorAll('h1, h2'))
           .map(h => h.innerText).join(' ')]
"""

# Copia las tablas: las normales (<table>) y las "de diseño moderno" (role=grid).
JS_TABLAS = """
() => {
    const limpiar = (s) => (s || '').replace(/\\s+/g, ' ').trim();
    const tablas = [];
    const sirve = (filas) => filas.length >= 2 && filas.some(f => f.length >= 2);

    document.querySelectorAll('table').forEach(t => {
        if (t.querySelector('table')) return;   // tabla que solo ordena el diseño
        const filas = Array.from(t.rows).map(r =>
            Array.from(r.cells).map(c => limpiar(c.innerText)));
        if (sirve(filas)) tablas.push(filas);
    });

    document.querySelectorAll('[role=grid], [role=table]').forEach(g => {
        if (g.tagName === 'TABLE') return;      // ya se copió arriba
        const filas = Array.from(g.querySelectorAll('[role=row]')).map(r =>
            Array.from(r.querySelectorAll(
                '[role=cell], [role=gridcell], [role=columnheader], [role=rowheader]'
            )).map(c => limpiar(c.innerText)));
        if (sirve(filas)) tablas.push(filas);
    });
    return tablas;
}
"""


# -------------------------------------------
# FUNCIONES DE AYUDA (herramientas pequeñas)
# -------------------------------------------

def normalizar(texto):
    """Pasa a minúsculas y quita tildes: 'Afiliación' -> 'afiliacion'."""
    texto = unicodedata.normalize("NFD", texto or "")
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return texto.lower()


def categorias_de(texto):
    """Dice a qué categorías pertenece un texto, según PALABRAS_CLAVE."""
    t = normalizar(texto)
    return [cat for cat, palabras in PALABRAS_CLAVE.items()
            if any(p in t for p in palabras)]


def es_prohibido(texto):
    t = normalizar(texto)
    return any(p in t for p in PALABRAS_PROHIBIDAS)


def es_del_portal(url):
    """True si el enlace es de la Mutual (o del mismo sitio donde ingresaste)."""
    partes = urlparse(url)
    if partes.scheme not in ("http", "https"):
        return False   # descarta "javascript:", "mailto:", etc.
    sitio = partes.hostname or ""
    return (sitio == SITIO_INICIAL or sitio == "mutual.cl"
            or sitio.endswith(".mutual.cl"))


def es_archivo(url):
    ruta = urlparse(url).path.lower()
    return ruta.endswith(EXTENSIONES_ARCHIVO)


def limpiar_nombre(texto, largo=50):
    """Convierte un texto en un nombre de archivo seguro."""
    t = re.sub(r"[^a-z0-9]+", "_", normalizar(texto)).strip("_")
    return t[:largo] or "pagina"


def nombre_libre(subcarpeta, base, extension):
    """Devuelve un nombre de archivo que todavía no existe (no sobrescribe)."""
    destino = carpeta / subcarpeta / f"{base}{extension}"
    numero = 2
    while destino.exists():
        destino = carpeta / subcarpeta / f"{base}_{numero}{extension}"
        numero += 1
    return destino


def extension_segun_contenido(contenido):
    """Si no sabemos qué archivo es, miramos sus primeras letras:
    todos los PDF empiezan con '%PDF' (como una firma)."""
    return ".pdf" if contenido[:4] == b"%PDF" else ".bin"


def para_excel(valor):
    """Quita caracteres raros que Excel no acepta (los números quedan igual)."""
    if not isinstance(valor, str):
        return valor
    return ILLEGAL_CHARACTERS_RE.sub("", valor)[:32000]


def en_todos_los_marcos(pagina, codigo_js):
    """Ejecuta el código en la página y en sus 'marcos' (iframes).
    Un iframe es una página metida dentro de otra, como una foto enmarcada
    colgada en una pared: hay que mirarla aparte."""
    resultado = []
    for marco in pagina.frames:
        try:
            resultado.extend(marco.evaluate(codigo_js))
        except PlaywrightError:
            pass   # algunos marcos no se dejan leer; los saltamos
    return resultado


def esperar_carga(pagina):
    try:
        pagina.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightError:
        pass   # hay portales que nunca quedan 100% "quietos"; seguimos igual
    pagina.wait_for_timeout(PAUSA_SEGUNDOS * 1000)


def titulo_de(pagina):
    try:
        return pagina.title() or pagina.url
    except PlaywrightError:
        return pagina.url


def pide_clave(pagina):
    """True si la página muestra un campo de contraseña (= pantalla de ingreso)."""
    try:
        return pagina.locator("input[type=password]").count() > 0
    except PlaywrightError:
        return False


def preguntar(mensaje):
    """Como input(), pero si no hay teclado disponible responde 'salir'."""
    try:
        return input(mensaje).strip().lower()
    except EOFError:
        return "salir"


# -------------------------------------------
# GUARDAR ARCHIVOS Y PÁGINAS
# -------------------------------------------

def descargar_archivo(contexto, url, texto_enlace, categoria):
    """Descarga un archivo usando la misma sesión del navegador.
    Devuelve True si guardó un archivo."""
    if url in urls_descargadas:
        return False
    urls_descargadas.add(url)
    try:
        respuesta = contexto.request.get(url, timeout=30000)
    except PlaywrightError:
        print("     (no se pudo descargar)", url)
        return False
    tipo = respuesta.headers.get("content-type", "")
    if not respuesta.ok or "text/html" in tipo:
        return False   # no era un archivo, era una página web

    # Nombre del archivo: el que manda el portal, o uno armado por nosotros
    disposicion = respuesta.headers.get("content-disposition", "")
    encontrado = re.search(r'filename="?([^";]+)"?', disposicion)
    nombre_original = encontrado.group(1) if encontrado else Path(urlparse(url).path).name
    contenido = respuesta.body()
    extension = Path(nombre_original).suffix.lower()
    if not extension:
        extension = ".pdf" if "pdf" in tipo else extension_segun_contenido(contenido)
    base = limpiar_nombre(Path(nombre_original).stem or texto_enlace)

    destino = nombre_libre("descargas", base, extension)
    destino.write_bytes(contenido)
    archivos_guardados.append({"categoria": categoria, "archivo": destino.name,
                               "enlace": texto_enlace, "url": url})
    print("     Archivo guardado:", destino.name)
    return True


def anotar_descarga(descarga):
    descargas_pendientes.append(descarga)


def vigilar_descargas(pestana):
    """Le pide a una pestaña que nos avise cada vez que descargue algo."""
    pestana.on("download", anotar_descarga)


def guardar_descargas_pendientes(categoria="Descarga manual"):
    """Guarda lo que el navegador descargó (por ejemplo, cuando TÚ haces
    clic en un botón 'Descargar' durante el modo guiado)."""
    while descargas_pendientes:
        descarga = descargas_pendientes.pop(0)
        nombre = Path(descarga.suggested_filename)
        destino = nombre_libre("descargas", limpiar_nombre(nombre.stem),
                               nombre.suffix.lower() or ".bin")
        try:
            descarga.save_as(destino)
        except PlaywrightError:
            print("     (una descarga no se pudo completar)")
            continue
        if destino.suffix == ".bin":   # sin extensión: revisamos si es un PDF
            nueva = extension_segun_contenido(destino.read_bytes())
            if nueva != ".bin":
                corregido = nombre_libre("descargas", destino.stem, nueva)
                destino = destino.rename(corregido)
        archivos_guardados.append({"categoria": categoria, "archivo": destino.name,
                                   "enlace": "(descarga del navegador)",
                                   "url": descarga.url})
        print("     Archivo guardado:", destino.name)


def guardar_pagina(contexto, pagina, categoria):
    """Guarda TODO lo útil de la página que se ve en pantalla:
    texto, captura, tablas y archivos enlazados."""
    numero = len(paginas_guardadas) + 1
    titulo = titulo_de(pagina)
    base = f"{numero:02d}_{limpiar_nombre(titulo)}"

    # 1) Texto completo (sirve para buscar con Ctrl+F después)
    texto = "\n\n".join(en_todos_los_marcos(pagina, JS_TEXTO))
    archivo_texto = carpeta / "textos" / f"{base}.txt"
    archivo_texto.write_text(f"{titulo}\n{pagina.url}\n\n{texto}", encoding="utf-8")

    # 2) Captura de pantalla (la "fotocopia")
    captura = carpeta / "capturas" / f"{base}.png"
    try:
        pagina.screenshot(path=str(captura), full_page=True)
    except PlaywrightError:
        captura = None

    # 3) Tablas -> irán al Excel
    tablas = en_todos_los_marcos(pagina, JS_TABLAS)
    for filas in tablas:
        tablas_guardadas.append({"pagina": numero, "titulo": titulo,
                                 "url": pagina.url, "filas": filas})

    # 4) Archivos enlazados (PDF, Excel...) o botones "Descargar" que sean enlaces
    archivos = 0
    for enlace in en_todos_los_marcos(pagina, JS_ENLACES):
        if archivos >= MAX_ARCHIVOS_POR_PAGINA:
            break
        url, texto_enlace = enlace["url"], enlace["texto"]
        parece_archivo = (es_archivo(url) or "descarg" in normalizar(texto_enlace)
                          or "download" in url.lower() or "descarga" in url.lower())
        if parece_archivo and es_del_portal(url) and not es_prohibido(texto_enlace + " " + url):
            if descargar_archivo(contexto, url, texto_enlace, categoria):
                archivos += 1

    paginas_guardadas.append({
        "numero": numero, "categoria": categoria, "titulo": titulo,
        "url": pagina.url, "tablas": len(tablas), "archivos": archivos,
        "captura": captura.name if captura else "",
    })
    print(f"     Guardada: {len(tablas)} tabla(s), {archivos} archivo(s)")


# -------------------------------------------
# PASO 2: BÚSQUEDA AUTOMÁTICA
# -------------------------------------------

def enlaces_interesantes(pagina):
    """Del menú de la página, elige los enlaces que hablan de lo que buscamos."""
    elegidos = {}
    for enlace in en_todos_los_marcos(pagina, JS_ENLACES):
        url, texto = enlace["url"], enlace["texto"]
        if not es_del_portal(url) or es_prohibido(texto + " " + url):
            continue
        # Buscamos las palabras en el texto del enlace y en su dirección
        categorias = categorias_de(texto + " " + urlparse(url).path)
        if categorias and url not in elegidos:
            elegidos[url] = {"texto": texto or url, "url": url,
                             "categoria": ", ".join(categorias)}
    return list(elegidos.values())


def busqueda_automatica(contexto, pagina):
    url_inicio = pagina.url

    # Anotamos el "mapa" del portal: todos los enlaces de la página principal
    for enlace in en_todos_los_marcos(pagina, JS_ENLACES):
        if es_del_portal(enlace["url"]):
            mapa_del_portal.append({
                "texto": enlace["texto"], "url": enlace["url"],
                "categoria": ", ".join(categorias_de(enlace["texto"] + " " + enlace["url"])),
            })

    print("\n  -> Página principal del portal")
    guardar_pagina(contexto, pagina, "Inicio del portal")

    # La "cola" es la lista de páginas por visitar, como una fila en el banco:
    # se atiende primero al que llegó primero.
    cola = [(enlace, 1) for enlace in enlaces_interesantes(pagina)]
    visitadas = {url_inicio}
    print(f"\n  Encontré {len(cola)} enlace(s) relacionados en el menú.")

    while cola and len(visitadas) <= MAX_PAGINAS:
        enlace, nivel = cola.pop(0)
        if enlace["url"] in visitadas or enlace["url"] in urls_descargadas:
            continue
        visitadas.add(enlace["url"])
        print(f"\n  -> [{enlace['categoria']}] {enlace['texto'][:70]}")

        if es_archivo(enlace["url"]):
            descargar_archivo(contexto, enlace["url"], enlace["texto"], enlace["categoria"])
            continue

        # Abrimos el enlace en la MISMA pestaña, para no perder la sesión
        try:
            pagina.goto(enlace["url"], timeout=30000)
            esperar_carga(pagina)
        except PlaywrightError as error:
            guardar_descargas_pendientes(enlace["categoria"])   # quizás era un archivo
            if "Download is starting" not in str(error):
                print("     (no se pudo abrir esta página)")
            continue

        if pide_clave(pagina):
            print("\n  ATENCIÓN: apareció una pantalla de ingreso. Parece que la sesión")
            print("  se cerró, así que detengo la búsqueda automática aquí.")
            print("  Vuelve a ingresar en el navegador y usa el modo guiado.")
            break

        guardar_pagina(contexto, pagina, enlace["categoria"])
        guardar_descargas_pendientes(enlace["categoria"])

        if nivel < PROFUNDIDAD:
            for nuevo in enlaces_interesantes(pagina):
                if nuevo["url"] not in visitadas:
                    cola.append((nuevo, nivel + 1))

    # Volvemos a la página principal para que sigas desde ahí
    try:
        pagina.goto(url_inicio, timeout=30000)
        esperar_carga(pagina)
    except PlaywrightError:
        pass


# -------------------------------------------
# PASO 3: MODO GUIADO (tú navegas, el programa guarda)
# -------------------------------------------

def pestana_visible(contexto):
    """Devuelve la pestaña que estás mirando en el navegador."""
    for pagina in reversed(contexto.pages):
        try:
            if pagina.evaluate("document.visibilityState") == "visible":
                return pagina
        except PlaywrightError:
            pass
    return contexto.pages[-1] if contexto.pages else None


def modo_guiado(contexto):
    print("\n" + "=" * 60)
    print("PASO 3: MODO GUIADO")
    print("  En el navegador, entra a cualquier página que quieras guardar")
    print("  (o haz clic en sus botones 'Descargar'). Luego vuelve aquí:")
    print("    - Enter      = guardar la página que estás viendo")
    print("    - 'salir'    = terminar y crear el Excel")
    print("=" * 60)
    while True:
        respuesta = preguntar("\nEnter para guardar / 'salir' para terminar: ")
        if respuesta in ("salir", "s", "q", "exit"):
            break
        pagina = pestana_visible(contexto)
        if pagina is None:
            print("  Se cerró el navegador.")
            break
        esperar_carga(pagina)
        guardar_descargas_pendientes()
        textos = " ".join(en_todos_los_marcos(pagina, JS_TITULOS))
        categorias = categorias_de(titulo_de(pagina) + " " + pagina.url + " " + textos)
        print(f"  -> {titulo_de(pagina)[:70]}")
        guardar_pagina(contexto, pagina, ", ".join(categorias) or "Guardada a mano")


# -------------------------------------------
# PASO 4: EXCEL CON EL RESUMEN
# -------------------------------------------

def nombre_hoja(texto):
    """Excel exige nombres de hoja cortos (máx. 31) y sin []:*?/\\ ."""
    return re.sub(r"[\[\]:*?/\\]", "", texto)[:31]


def agregar_hoja(libro, titulo, encabezados, filas, anchos):
    hoja = libro.create_sheet(nombre_hoja(titulo))
    hoja.append(encabezados)
    for fila in filas:
        hoja.append([para_excel(v) for v in fila])
    for columna, ancho in zip("ABCDEFGH", anchos):
        hoja.column_dimensions[columna].width = ancho
    return hoja


def crear_excel():
    libro = Workbook()
    libro.remove(libro.active)   # quitamos la hoja vacía que viene por defecto

    agregar_hoja(libro, "Indice",
                 ["N°", "Categoría", "Título de la página", "Dirección (URL)",
                  "Tablas", "Archivos", "Captura"],
                 [[p["numero"], p["categoria"], p["titulo"], p["url"],
                   p["tablas"], p["archivos"], p["captura"]] for p in paginas_guardadas],
                 [5, 25, 40, 60, 8, 9, 35])

    agregar_hoja(libro, "Archivos descargados",
                 ["Categoría", "Archivo (en carpeta 'descargas')", "Texto del enlace", "URL"],
                 [[a["categoria"], a["archivo"], a["enlace"], a["url"]]
                  for a in archivos_guardados],
                 [25, 40, 40, 60])

    agregar_hoja(libro, "Mapa del portal",
                 ["Texto del enlace", "Categoría detectada", "URL"],
                 [[m["texto"], m["categoria"], m["url"]] for m in mapa_del_portal],
                 [40, 25, 80])

    # Una hoja por cada tabla encontrada
    for i, tabla in enumerate(tablas_guardadas, start=1):
        hoja = libro.create_sheet(nombre_hoja(f"T{i:02d} {tabla['titulo']}"))
        hoja.append([para_excel(f"Página {tabla['pagina']}: {tabla['titulo']}")])
        hoja.append([para_excel(tabla["url"])])
        hoja.append([])
        for fila in tabla["filas"]:
            hoja.append([para_excel(celda) for celda in fila])

    destino = carpeta / "resumen.xlsx"
    libro.save(destino)
    return destino


# -------------------------------------------
# PROGRAMA PRINCIPAL
# -------------------------------------------

def main():
    global carpeta, SITIO_INICIAL

    # Opcional: python buscar_mutual.py https://otra-direccion...
    url = sys.argv[1] if len(sys.argv) > 1 else URL_INICIO
    # Solo para pruebas automáticas: MUTUAL_SIN_VENTANA=1 oculta el navegador
    sin_ventana = os.environ.get("MUTUAL_SIN_VENTANA") == "1"

    fecha = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    carpeta = Path(__file__).resolve().parent / "resultados_mutual" / fecha
    for sub in ("capturas", "descargas", "textos"):
        (carpeta / sub).mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        try:
            navegador = p.chromium.launch(headless=sin_ventana)
        except PlaywrightError:
            print("No encontré el navegador. Ejecuta primero:")
            print("    python -m playwright install chromium")
            return

        contexto = navegador.new_context(accept_downloads=True,
                                         viewport={"width": 1366, "height": 900})
        # Cada pestaña nueva nos avisará cuando empiece una descarga
        contexto.on("page", vigilar_descargas)

        try:
            pagina = contexto.new_page()
            pagina.goto(url, timeout=60000)

            print("=" * 60)
            print("PASO 1: INGRESA AL PORTAL")
            print("  Se abrió un navegador. Entra a la Sucursal Virtual de")
            print("  Empresas con tu RUT y clave, como lo haces siempre.")
            print("  (Este programa no ve ni guarda tu clave.)")
            print("=" * 60)
            preguntar("\nCuando ya estés DENTRO del portal, presiona Enter aquí... ")

            pagina = pestana_visible(contexto)
            if pagina is None:
                print("Se cerró el navegador. Fin.")
                return
            esperar_carga(pagina)
            SITIO_INICIAL = urlparse(pagina.url).hostname or ""

            print("\n" + "=" * 60)
            print("PASO 2: BÚSQUEDA AUTOMÁTICA (no toques el navegador)")
            print("=" * 60)
            busqueda_automatica(contexto, pagina)

            modo_guiado(contexto)
            guardar_descargas_pendientes()

        except KeyboardInterrupt:
            print("\nDetenido por ti (Ctrl+C). Guardo lo que alcancé a reunir...")
        except PlaywrightError as error:
            print("\nEl navegador se cerró o falló:", str(error).splitlines()[0])
            print("Guardo lo que alcancé a reunir...")
        finally:
            excel = crear_excel()
            try:
                navegador.close()
            except PlaywrightError:
                pass

    print("\n" + "=" * 60)
    print("LISTO")
    print(f"  Páginas guardadas:  {len(paginas_guardadas)}")
    print(f"  Tablas encontradas: {len(tablas_guardadas)}")
    print(f"  Archivos (PDF...):  {len(archivos_guardados)}")
    print(f"\n  Todo quedó en: {carpeta}")
    print(f"  Abre primero:  {excel.name}")
    print("=" * 60)


if __name__ == "__main__":
    main()
