#!/usr/bin/env python3
"""
Avisos a padres de familia
--------------------------
Programa con ventana para enviar correos de incidencias (tarea, conducta,
reprobación, etc.) a los padres desde tu cuenta de Gmail.

Archivos (en la misma carpeta que este programa):
  alumnos.csv          una fila por alumno: alumno, grupo, correo_padre
  materias.csv         lista de materias (una por fila, con el encabezado "materia")
  plantillas.json      textos de cada tipo de aviso (los puedes editar)
  config.json          tu correo, nombre y firma
  secrets.txt          (opcional) tus claves: contraseña de aplicación y clave de la IA
  registro_avisos.csv  se crea solo: historial de todo lo enviado

Claves: el programa las busca primero en las variables de entorno y luego en
secrets.txt (líneas NOMBRE=valor). Si no las encuentra, te las pide al enviar el
primer correo (o al redactar con IA) de cada sesión y NO las guarda.
Nombres: AVISOS_CLAVE (contraseña de aplicación de Gmail), GEMINI_API_KEY,
ANTHROPIC_API_KEY y AVISOS_IA_CLAVE (otros servicios de IA).

Redacción con IA (tipo «Otra situación»): describes lo ocurrido y una IA redacta el
borrador. El nombre del alumno NO se envía a la IA: se cambia por el marcador {alumno}
y se vuelve a poner aquí.
En config.json eliges el proveedor con "proveedor": "anthropic" (por defecto),
"gemini" u "openai" (cualquier servicio con formato chat/completions: OpenRouter,
DeepSeek, OpenAI, Ollama local...). Opcionalmente "modelo_ia" y "url_ia".
"""

import csv
import io
import json
import os
import re
import shutil
import smtplib
import ssl
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

CARPETA = Path(__file__).resolve().parent
ARCHIVO_ALUMNOS = CARPETA / "alumnos.csv"
ARCHIVO_MATERIAS = CARPETA / "materias.csv"
ARCHIVO_PLANTILLAS = CARPETA / "plantillas.json"
ARCHIVO_CONFIG = CARPETA / "config.json"
ARCHIVO_SECRETOS = CARPETA / "secrets.txt"
ARCHIVO_REGISTRO = CARPETA / "registro_avisos.csv"

CAMPOS_REGISTRO = [
    "fecha_envio", "fecha_incidencia", "alumno", "grupo", "materia",
    "tipo", "correo_padre", "estado", "detalle", "error",
]

MESES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
    "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]

PATRON_CORREO = re.compile(r"[^@\s]+@[^@\s]+\.[^@\s]+")
DOMINIOS_EJEMPLO = {"example.com", "example.org", "example.net"}

# --- Redacción con IA ---
OPCION_IA = "Otra situación (redactar con IA)"
TIPO_REGISTRO_IA = "Redactado con IA"
MODELO_IA = "claude-haiku-5-5"          # modelo predeterminado de Anthropic
URL_IA = "https://api.anthropic.com/v1/messages"
VERSION_API = "2023-06-01"
URL_GEMINI = "https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent"
MODELO_GEMINI = "gemini-3.8-flash"
MODELO_GEMINI_RESPALDO = "gemini-3.5-flash-lite"   # se prueba si el principal está saturado
# Segundos de espera antes de cada reintento cuando el servidor de IA está saturado.
ESPERA_REINTENTOS = (2, 5)
# Proveedores de IA que se pueden elegir con "proveedor" en config.json
PROVEEDORES = {
    "anthropic": {"nombre": "Anthropic (Claude)", "url": URL_IA, "modelo": MODELO_IA,
                  "variable": "ANTHROPIC_API_KEY",
                  "ayuda": "Pega tu clave de API de Anthropic (empieza con sk-ant-)."},
    "gemini": {"nombre": "Google (Gemini)", "url": URL_GEMINI, "modelo": MODELO_GEMINI,
               "variable": "GEMINI_API_KEY", "respaldo": MODELO_GEMINI_RESPALDO,
               "ayuda": "Pega tu clave de API de Google AI Studio."},
    # Cualquier servicio con formato «chat/completions» de OpenAI: OpenRouter, DeepSeek,
    # OpenAI, Ollama (local), etc. Requiere "url_ia" y "modelo_ia" en config.json.
    "openai": {"nombre": "el servicio de IA configurado", "url": "", "modelo": "",
               "variable": "AVISOS_IA_CLAVE",
               "ayuda": "Pega la clave de API de tu servicio de IA."},
}
CAMPOS_IA = ("alumno", "grupo", "materia", "fecha")
PARTICULAS = {"del", "los", "las", "van", "von", "san"}
TONOS = {
    "Cordial y firme": "cordial pero firme: deja claro que la situación debe corregirse, "
                       "sin regañar",
    "Más suave": "muy amable y conciliador, centrado en apoyar al alumno",
    "Más serio y formal": "formal y serio: es un asunto que requiere la atención de la "
                          "familia, sin ser agresivo",
}


# ----------------------------------------------------------------------
# Lógica (sin ventana)
# ----------------------------------------------------------------------

def fecha_larga(d: date) -> str:
    return f"{d.day} de {MESES[d.month - 1]} de {d.year}"


def parsear_fecha(texto: str) -> date:
    return datetime.strptime(texto.strip(), "%d/%m/%Y").date()


def leer_texto(ruta: Path) -> str:
    """
    Lee un archivo de texto sin importar cómo lo guardó el programa que lo creó.
    Excel en español suele guardar los CSV en Windows-1252 (ANSI) en lugar de UTF-8,
    y el Bloc de notas a veces agrega una marca BOM al inicio.
    """
    datos = Path(ruta).read_bytes()
    for codificacion in ("utf-8-sig", "cp1252"):
        try:
            return datos.decode(codificacion)
        except UnicodeDecodeError:
            continue
    return datos.decode("latin-1")


def cargar_json(ruta: Path):
    try:
        return json.loads(leer_texto(ruta))
    except json.JSONDecodeError as e:
        raise ValueError(
            f"{Path(ruta).name} tiene un error de formato en la línea {e.lineno}, "
            f"columna {e.colno}: {e.msg}. Revisa comillas y comas.")


# Otros nombres que también se aceptan para cada clave (por si en secrets.txt la
# escribiste con un nombre parecido). No distingue mayúsculas de minúsculas.
ALIAS_SECRETOS = {
    "AVISOS_CLAVE": ("GMAIL_APP_PASSWORD", "GMAIL_PASSWORD", "GMAIL_CLAVE", "CLAVE_GMAIL",
                     "CONTRASENA_GMAIL", "CONTRASENA_APLICACION", "CORREO_CLAVE",
                     "EMAIL_PASSWORD", "APP_PASSWORD"),
    "GEMINI_API_KEY": ("GOOGLE_API_KEY", "GEMINI_KEY", "CLAVE_GEMINI"),
    "ANTHROPIC_API_KEY": ("CLAUDE_API_KEY", "ANTHROPIC_KEY", "CLAVE_ANTHROPIC",
                          "CLAVE_CLAUDE"),
}
COMILLAS = "\"'\u201c\u201d\u2018\u2019"


def cargar_secretos(ruta: Path = ARCHIVO_SECRETOS) -> dict:
    """
    Lee secrets.txt: una línea NOMBRE=valor por clave. Se ignoran las líneas vacías,
    las que empiezan con # y los valores en blanco; las comillas alrededor son opcionales.
    Los nombres se guardan en mayúsculas.
    """
    secretos = {}
    if not Path(ruta).exists():
        return secretos
    for linea in leer_texto(ruta).splitlines():
        linea = linea.strip()
        if not linea or linea.startswith("#"):
            continue
        if linea.lower().startswith("export "):
            linea = linea[7:].strip()
        if "=" in linea:
            nombre, valor = linea.split("=", 1)
        elif ":" in linea:
            nombre, valor = linea.split(":", 1)
        else:
            continue
        nombre, valor = nombre.strip().upper(), valor.strip().strip(COMILLAS).strip()
        if nombre and valor:
            secretos[nombre] = valor
    return secretos


def obtener_secreto(nombre: str, ruta: Path = ARCHIVO_SECRETOS) -> str:
    """Busca una clave: primero en las variables de entorno y luego en secrets.txt."""
    nombres = (nombre, *ALIAS_SECRETOS.get(nombre, ()))
    for n in nombres:
        valor = (os.environ.get(n) or "").strip()
        if valor:
            return valor
    archivo = cargar_secretos(ruta)
    for n in nombres:
        valor = (archivo.get(n.upper()) or "").strip()
        if valor:
            return valor
    return ""


def diagnostico_secretos(nombre: str, ruta: Path = ARCHIVO_SECRETOS) -> str:
    """Explica por qué no se encontró una clave (para mostrarlo cuando el programa la pide)."""
    ruta = Path(ruta)
    if not ruta.exists():
        extra = ""
        if ruta.with_name(ruta.name + ".txt").exists():
            extra = f" Existe «{ruta.name}.txt»: quita el «.txt» de más al nombre."
        return (f"No encontré {nombre} en una variable de entorno ni el archivo "
                f"{ruta.name} en:\n{ruta.parent}{extra}")
    try:
        presentes = sorted(cargar_secretos(ruta))
    except OSError as e:
        return f"No pude leer {ruta.name} ({e})."
    if presentes:
        lista = ", ".join(presentes)
        return (f"En {ruta.name} (carpeta {ruta.parent}) no hay una línea «{nombre}=...» "
                f"con valor. Nombres con valor que sí vi: {lista}.")
    return (f"El archivo {ruta.name} (carpeta {ruta.parent}) no tiene ninguna clave con "
            f"valor. Escribe una línea como {nombre}=tu_clave y guarda el archivo.")


def separar_correos(texto: str):
    """'a@x.com; b@y.com' -> ['a@x.com', 'b@y.com']"""
    return [c.strip() for c in re.split(r"[;,]", texto) if c.strip()]


def cargar_alumnos(ruta: Path = ARCHIVO_ALUMNOS):
    """
    Devuelve (alumnos, avisos). Una entrada por alumno (mismo nombre y grupo).
    Las filas con problemas se omiten y se reportan.
    Compatibilidad: si el archivo trae una columna 'materia' (formato anterior),
    se ignora y las filas repetidas del mismo alumno se unen en una sola.
    """
    por_clave, avisos = {}, []
    with io.StringIO(leer_texto(ruta), newline="") as f:
        for n, fila in enumerate(csv.DictReader(f), start=2):
            fila = {(k or "").strip().lower(): (v or "").strip() for k, v in fila.items()}
            nombre = fila.get("alumno", "")
            if not nombre:
                continue
            correos = separar_correos(fila.get("correo_padre", ""))
            malos = [c for c in correos if not PATRON_CORREO.fullmatch(c)]
            if not correos or malos:
                avisos.append(f"Línea {n} ({nombre}): correo no válido, se omitió.")
                continue
            grupo = fila.get("grupo", "")
            clave = (nombre.lower(), grupo.lower())
            if clave in por_clave:
                for c in correos:
                    if c not in por_clave[clave]["correos"]:
                        por_clave[clave]["correos"].append(c)
                continue
            por_clave[clave] = {"alumno": nombre, "grupo": grupo, "correos": correos}
    alumnos = sorted(por_clave.values(), key=lambda a: (a["grupo"], a["alumno"]))
    return alumnos, avisos


def cargar_materias(ruta: Path = ARCHIVO_MATERIAS):
    """Lee materias.csv: una materia por fila (primera columna; el encabezado es opcional)."""
    materias = []
    with io.StringIO(leer_texto(ruta), newline="") as f:
        for i, fila in enumerate(csv.reader(f)):
            if not fila:
                continue
            valor = fila[0].strip()
            if not valor or (i == 0 and valor.lower() == "materia"):
                continue
            if valor not in materias:
                materias.append(valor)
    return materias


def cargar_plantillas(ruta: Path = ARCHIVO_PLANTILLAS):
    plantillas = cargar_json(ruta)
    for nombre, p in plantillas.items():
        if "asunto" not in p or "cuerpo" not in p:
            raise ValueError(f"La plantilla '{nombre}' necesita 'asunto' y 'cuerpo'.")
        if isinstance(p["cuerpo"], list):
            p["cuerpo"] = "\n".join(p["cuerpo"])
    return plantillas


def armar_mensaje(plantilla: dict, alumno: dict, materia: str, fecha: date,
                  detalle: str, cfg: dict):
    """Devuelve (asunto, cuerpo) con los campos ya sustituidos."""
    datos = {
        "alumno": alumno["alumno"],
        "grupo": alumno["grupo"],
        "materia": materia,
        "fecha": fecha_larga(fecha),
        "profesor": cfg.get("profesor", ""),
        "detalle": detalle.strip(),
    }
    try:
        asunto = plantilla["asunto"].format(**datos)
        cuerpo = plantilla["cuerpo"].format(**datos)
    except KeyError as e:
        raise ValueError(
            f"La plantilla usa un campo desconocido: {e}. "
            f"Campos válidos: {', '.join('{' + k + '}' for k in datos)}"
        )
    cuerpo = re.sub(r"\n{3,}", "\n\n", cuerpo).strip()
    firma = cfg.get("firma", "").strip()
    if firma:
        cuerpo += "\n" + firma
    return asunto.strip(), cuerpo


def construir_email(cfg: dict, destinatarios, asunto: str, cuerpo: str) -> EmailMessage:
    msg = EmailMessage()
    msg["From"] = formataddr((cfg.get("profesor", ""), cfg["correo"]))
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = asunto
    msg.set_content(cuerpo)
    return msg


def es_correo_de_ejemplo(destinatarios) -> bool:
    return any(d.split("@")[-1].lower() in DOMINIOS_EJEMPLO for d in destinatarios)


def enviar_lote(cfg: dict, clave: str, mensajes):
    """
    mensajes: lista de dicts con la clave 'email' (EmailMessage).
    Devuelve [(mensaje, estado, error)]. Un fallo en un correo no detiene los demás.
    Si falla el inicio de sesión, lanza la excepción.
    """
    servidor = cfg.get("servidor_smtp", "smtp.gmail.com")
    puerto = int(cfg.get("puerto_smtp", 465))
    contexto = ssl.create_default_context()

    if puerto == 465:
        conexion = smtplib.SMTP_SSL(servidor, puerto, context=contexto, timeout=30)
    else:
        conexion = smtplib.SMTP(servidor, puerto, timeout=30)

    resultados = []
    with conexion as s:
        if puerto != 465:
            s.starttls(context=contexto)
        s.login(cfg["correo"], clave)
        for m in mensajes:
            try:
                s.send_message(m["email"])
                resultados.append((m, "enviado", ""))
            except Exception as e:  # noqa: BLE001
                resultados.append((m, "error", str(e)))
    return resultados


def leer_enviados_previos():
    """Conjunto de (alumno, materia, tipo, fecha_incidencia) ya enviados."""
    previos = set()
    if not ARCHIVO_REGISTRO.exists():
        return previos
    with io.StringIO(leer_texto(ARCHIVO_REGISTRO), newline="") as f:
        for fila in csv.DictReader(f):
            if fila.get("estado") == "enviado":
                previos.add((fila.get("alumno"), fila.get("materia"),
                             fila.get("tipo"), fila.get("fecha_incidencia")))
    return previos


def registrar(resultados, ruta: Path = ARCHIVO_REGISTRO):
    nuevo = not ruta.exists()
    ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(ruta, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if nuevo:
            w.writerow(CAMPOS_REGISTRO)
        for m, estado, error in resultados:
            w.writerow([
                ahora, m["fecha_txt"], m["alumno"]["alumno"], m["alumno"]["grupo"],
                m["materia"], m["tipo"], "; ".join(m["destinatarios"]),
                estado, m["detalle"], error,
            ])


def validar_config(cfg: dict):
    if not cfg.get("correo") or "tucorreo@" in cfg["correo"]:
        raise ValueError("Abre config.json y escribe tu correo real en el campo \"correo\".")


# ----------------------------------------------------------------------
# Redacción con IA
# ----------------------------------------------------------------------

class ErrorIA(Exception):
    """Fallo al consultar la IA. El mensaje está pensado para mostrarse al usuario."""

    def __init__(self, mensaje, autenticacion=False, saturado=False):
        super().__init__(mensaje)
        self.autenticacion = autenticacion
        self.saturado = saturado      # el servidor de IA está saturado (errores 5xx)


INSTRUCCIONES_IA = """Eres un asistente que ayuda a un profesor de secundaria en México a redactar correos breves dirigidos a las familias de sus alumnos.

El profesor te describe qué ocurrió. Redacta el correo con estas reglas:
- Español de México, tratando de "ustedes" a la familia. Empieza con "Estimada familia de {alumno}:".
- Usa EXACTAMENTE estos marcadores, tal cual, cuando corresponda: {alumno} (nombre del alumno), {grupo}, {materia} y {fecha} (fecha del hecho, ya escrita en palabras). No uses ningún otro marcador ni llaves fuera de estos cuatro.
- Nunca escribas nombres reales ni inventes nombres, datos, fechas, calificaciones, sanciones o compromisos que el profesor no mencionó.
- Describe los hechos de forma objetiva y respetuosa, sin etiquetas ni juicios sobre el alumno o la familia, sin diagnósticos y sin tono acusatorio.
- Si el profesor no mencionó una medida o sanción, no la inventes: invita a la familia a platicar con el alumno y ofrece comunicarse para cualquier duda.
- Sé breve: entre 80 y 150 palabras. Texto plano, sin Markdown.
- Termina con "Atentamente," y nada más: no escribas nombre ni firma (el programa los agrega).

Responde únicamente con este formato exacto:
ASUNTO: <asunto breve; puede usar {alumno} o {materia}>
CUERPO:
<texto del correo>"""


def _variantes_nombre(alumno: dict):
    nombre = " ".join(alumno["alumno"].split())
    variantes = {nombre} if nombre else set()
    for parte in nombre.split():
        if len(parte) >= 3 and parte.lower() not in PARTICULAS:
            variantes.add(parte)
    return variantes


def a_marcadores(texto: str, alumnos, materia: str = "", fecha_txt: str = "",
                 escapar: bool = False) -> str:
    """
    Cambia los datos reales por marcadores: nombre (completo o por partes) -> {alumno},
    grupo -> {grupo}, materia -> {materia}, fecha en palabras -> {fecha}.
    Con escapar=True duplica antes las llaves del texto, para poder usarlo como plantilla.
    """
    if escapar:
        texto = texto.replace("{", "{{").replace("}", "}}")
    reemplazos = set()
    for a in alumnos:
        for v in _variantes_nombre(a):
            reemplazos.add((v, "{alumno}"))
        if len(a.get("grupo", "")) >= 2:
            reemplazos.add((a["grupo"], "{grupo}"))
    if materia:
        reemplazos.add((materia, "{materia}"))
    if fecha_txt:
        reemplazos.add((fecha_txt, "{fecha}"))
    # Los más largos primero: «Ana Pérez López» se cambia completo antes que «Ana».
    for original, marcador in sorted(reemplazos, key=lambda r: len(r[0]), reverse=True):
        texto = re.sub(rf"(?<!\w){re.escape(original)}(?!\w)", marcador, texto,
                       flags=re.IGNORECASE)
    return texto


def plantillizar(texto: str, alumno: dict, materia: str, fecha: date, cfg: dict) -> str:
    """Convierte un mensaje ya redactado (y quizá editado a mano) en texto de plantilla."""
    texto = texto.strip()
    firma = cfg.get("firma", "").strip()
    if firma:
        # La firma la agrega el programa al enviar; no debe quedar dentro de la plantilla,
        # aunque la hayas editado y ya no esté al final.
        pos = texto.rfind(firma)
        if pos != -1:
            texto = texto[:pos] + texto[pos + len(firma):]
            texto = re.sub(r"\n{3,}", "\n\n", texto).strip()
    return a_marcadores(texto, [alumno], materia, fecha_larga(fecha), escapar=True)


def sanear_campos(texto: str, campos=CAMPOS_IA) -> str:
    """Deja como marcadores solo los campos permitidos; cualquier otra llave queda literal."""
    t = texto.replace("{", "{{").replace("}", "}}")
    for c in campos:
        t = t.replace("{{" + c + "}}", "{" + c + "}")
    return t


def construir_peticion_ia(brief_marcado: str, tono: str):
    descripcion_tono = TONOS.get(tono) or next(iter(TONOS.values()))
    usuario = ("Tono deseado: " + descripcion_tono +
               "\n\nLo que ocurrió, según el profesor (materia: {materia}):\n" + brief_marcado)
    return INSTRUCCIONES_IA, usuario


_RUIDO = r"[\s*_#>`\-]*"          # adornos de Markdown que la IA a veces pone
_ETIQ_ASUNTO = re.compile(r"^" + _RUIDO + r"asunto[\s*_]*[:\uff1a][\s*_]*(.*)$",
                          re.IGNORECASE | re.MULTILINE)
_ETIQ_CUERPO = re.compile(r"^" + _RUIDO + r"(?:cuerpo|mensaje|texto|correo)"
                          r"(?:\s+del\s+(?:correo|mensaje))?[\s*_]*[:\uff1a][\s*_]*(.*)$",
                          re.IGNORECASE | re.MULTILINE)


def _limpiar_md(texto: str) -> str:
    return texto.replace("**", "").replace("__", "").strip()


def parsear_borrador(texto: str) -> dict:
    """
    Extrae asunto y cuerpo de la respuesta de la IA. Tolera negritas, encabezados,
    texto antes del borrador y respuestas sin la etiqueta CUERPO o sin ASUNTO.
    """
    original = texto
    texto = re.sub(r"^\s*```\w*\s*|\s*```\s*$", "", texto.strip())
    asunto, cuerpo = "", ""
    ma = _ETIQ_ASUNTO.search(texto)
    mc = _ETIQ_CUERPO.search(texto, ma.end() if ma else 0)
    if ma:
        asunto = ma.group(1)
        resto_asunto_fin = mc.start() if mc else len(texto)
        if not asunto.strip():      # el asunto quedó en la línea siguiente
            asunto = texto[ma.end():resto_asunto_fin].strip().split("\n")[0]
    if mc:
        cuerpo = (mc.group(1) + "\n" + texto[mc.end():]).strip()
    elif ma:
        # Sin etiqueta CUERPO: el cuerpo es lo que sigue a la línea del asunto.
        cuerpo = texto[ma.end():].strip()
        if not ma.group(1).strip():
            cuerpo = cuerpo.split("\n", 1)[1].strip() if "\n" in cuerpo else ""
    else:
        # Sin etiquetas: si hay un saludo reconocible, lo usamos como inicio del cuerpo.
        m = re.search(r"^[\s*_#>]*(estimad[oa]s?\b.*)", texto, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if m:
            cuerpo = m.group(1)
            asunto = "{materia}: aviso sobre {alumno}"
    asunto = " ".join(_limpiar_md(asunto).split()).strip("\"\u201c\u201d").strip()
    cuerpo = re.sub(r"\n{3,}", "\n\n", _limpiar_md(cuerpo))
    if not asunto and not cuerpo:
        vista = " ".join(original.split())
        if len(vista) > 200:
            vista = vista[:200] + "…"
        raise ErrorIA("La IA respondió en un formato inesperado. Pulsa «Redactar con IA» "
                      "otra vez." + (f"\n\nLo que respondió: «{vista}»" if vista else
                                      " La respuesta llegó vacía."))
    if not asunto:
        asunto = "{materia}: aviso sobre {alumno}"
    if not cuerpo:
        raise ErrorIA("La IA devolvió un borrador sin texto. Pulsa «Redactar con IA» otra vez.")
    return {"asunto": sanear_campos(asunto), "cuerpo": sanear_campos(cuerpo)}


def es_local(url: str) -> bool:
    return (urllib.parse.urlparse(url).hostname or "") in ("localhost", "127.0.0.1", "::1")


def datos_proveedor(cfg: dict) -> dict:
    """Resuelve proveedor, URL y modelo a partir de config.json (con valores por defecto)."""
    pid = str(cfg.get("proveedor") or "anthropic").strip().lower()
    if pid not in PROVEEDORES:
        raise ErrorIA(f"En config.json, el proveedor «{pid}» no existe. "
                      f"Usa: {', '.join(PROVEEDORES)}.")
    base = PROVEEDORES[pid]
    url = str(cfg.get("url_ia") or base["url"]).strip()
    modelo = str(cfg.get("modelo_ia") or base["modelo"]).strip()
    if not url or not modelo:
        raise ErrorIA(f"Con el proveedor «{pid}» debes escribir \"url_ia\" y \"modelo_ia\" "
                      "en config.json (mira LEEME.txt).")
    local = es_local(url)
    return {"id": pid, "nombre": "el servidor de IA local" if local else base["nombre"],
            "url": url.replace("{modelo}", modelo), "modelo": modelo,
            "variable": base["variable"], "ayuda": base["ayuda"], "local": local,
            "respaldo": str(cfg.get("modelo_respaldo") or base.get("respaldo") or "").strip()}


def clave_ia_de_entorno(cfg: dict):
    try:
        variable = datos_proveedor(cfg)["variable"]
    except ErrorIA:
        return None
    return obtener_secreto(variable) or obtener_secreto("AVISOS_IA_CLAVE") or None


def _mayus(texto: str) -> str:
    return texto[:1].upper() + texto[1:]


def _error_http(e: urllib.error.HTTPError, nombre: str = "la IA") -> ErrorIA:
    try:
        info = json.loads(e.read().decode("utf-8", "replace")).get("error", "")
        detalle = info if isinstance(info, str) else str(info.get("message") or "")
    except Exception:  # noqa: BLE001
        detalle = ""
    bajo = detalle.lower()
    if e.code in (401, 403) or (e.code == 400 and "api key" in bajo and "not valid" in bajo):
        return ErrorIA(f"La clave de API no es válida o no tiene permiso para {nombre}. "
                       "Revísala en el sitio del proveedor.", autenticacion=True)
    if e.code == 402 or (e.code == 400 and "credit" in bajo):
        return ErrorIA(f"La cuenta no tiene saldo para {nombre}. Agrega créditos en el sitio "
                       "del proveedor.")
    if e.code == 429:
        return ErrorIA("Se alcanzó el límite de uso (por minuto o por día en los planes "
                       "gratuitos). Espera un poco e intenta de nuevo; si el límite es "
                       "diario, hasta mañana.")
    if e.code in (500, 502, 503, 504, 529):
        return ErrorIA(f"{_mayus(nombre)} está saturado o no disponible. Intenta de nuevo en unos "
                       "minutos.", saturado=True)
    return ErrorIA(f"{_mayus(nombre)} respondió con el error {e.code}: "
                   f"{detalle or e.reason}")


def _enviar_json(url: str, cabeceras: dict, cuerpo: dict, p: dict, timeout: int) -> dict:
    """Envía la petición; si el servidor está saturado (5xx) reintenta un par de veces."""
    esperas = list(ESPERA_REINTENTOS)
    while True:
        try:
            return _enviar_json_una_vez(url, cabeceras, cuerpo, p, timeout)
        except ErrorIA as e:
            if not e.saturado or not esperas:
                raise
            time.sleep(esperas.pop(0))


def _enviar_json_una_vez(url: str, cabeceras: dict, cuerpo: dict, p: dict, timeout: int) -> dict:
    req = urllib.request.Request(url, data=json.dumps(cuerpo).encode("utf-8"),
                                 method="POST", headers=cabeceras)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        raise _error_http(e, p["nombre"])
    except urllib.error.URLError as e:
        pista = " ¿Está iniciado el servidor local?" if p["local"] else ""
        raise ErrorIA(f"No hay conexión con {p['nombre']} ({e.reason}).{pista}")
    except ValueError:
        raise ErrorIA(f"La respuesta de {p['nombre']} no se pudo leer. Intenta de nuevo.")
    except OSError as e:
        raise ErrorIA(f"{_mayus(p['nombre'])} tardó demasiado o se cortó la "
                      f"conexión ({e}).")


def llamar_ia(cfg: dict, clave: str, sistema: str, usuario: str, timeout: int = 90) -> str:
    """Consulta el proveedor de IA configurado y devuelve el texto de la respuesta.

    Si el modelo está saturado incluso tras los reintentos y hay un modelo de respaldo
    (Gemini lo trae por defecto; en otros proveedores, "modelo_respaldo" en config.json),
    prueba con ese antes de rendirse."""
    p = datos_proveedor(cfg)
    try:
        return _llamar_ia(p, clave, sistema, usuario, timeout)
    except ErrorIA as e:
        if not (e.saturado and p["respaldo"] and p["respaldo"] != p["modelo"]):
            raise
    p2 = datos_proveedor({**cfg, "modelo_ia": p["respaldo"], "url_ia": cfg.get("url_ia")})
    return _llamar_ia(p2, clave, sistema, usuario, timeout)


def _llamar_ia(p: dict, clave: str, sistema: str, usuario: str, timeout: int) -> str:
    cabeceras = {"content-type": "application/json"}
    if p["id"] == "anthropic":
        cabeceras.update({"x-api-key": clave, "anthropic-version": VERSION_API})
        # Margen amplio: los modelos recientes de Claude «razonan» antes de responder y esos
        # tokens cuentan dentro de max_tokens; con un límite bajo podían agotarlo sin escribir.
        cuerpo = {"model": p["modelo"], "max_tokens": 4096, "system": sistema,
                  "messages": [{"role": "user", "content": usuario}]}
    elif p["id"] == "gemini":
        cabeceras["x-goog-api-key"] = clave
        cuerpo = {"systemInstruction": {"parts": [{"text": sistema}]},
                  "contents": [{"role": "user", "parts": [{"text": usuario}]}],
                  # Margen amplio: en los modelos con «razonamiento» este límite también
                  # cuenta los pasos internos, no solo el texto final.
                  "generationConfig": {"maxOutputTokens": 4096}}
    else:  # "openai": formato chat/completions (OpenRouter, DeepSeek, OpenAI, Ollama...)
        if clave:
            cabeceras["Authorization"] = "Bearer " + clave
        cuerpo = {"model": p["modelo"],
                  "messages": [{"role": "system", "content": sistema},
                               {"role": "user", "content": usuario}]}

    datos = _enviar_json(p["url"], cabeceras, cuerpo, p, timeout)

    texto = ""
    if p["id"] == "anthropic":
        texto = "".join(b.get("text", "") for b in datos.get("content", [])
                        if b.get("type") == "text")
    elif p["id"] == "gemini":
        candidatos = datos.get("candidates") or []
        if not candidatos:
            motivo = (datos.get("promptFeedback") or {}).get("blockReason")
            raise ErrorIA("Gemini no devolvió respuesta"
                          + (f" (bloqueada: {motivo})" if motivo else "")
                          + ". Reformula el texto e intenta de nuevo.")
        partes = (candidatos[0].get("content") or {}).get("parts") or []
        texto = "".join(x.get("text", "") for x in partes if not x.get("thought"))
    else:
        opciones = datos.get("choices") or []
        mensaje = (opciones[0].get("message") or {}) if opciones else {}
        texto = mensaje.get("content") or ""
        if isinstance(texto, list):
            texto = "".join(b.get("text", "") for b in texto if isinstance(b, dict))
    if not texto.strip():
        motivo = datos.get("stop_reason") if p["id"] == "anthropic" else None
        if motivo == "refusal":
            raise ErrorIA("Claude rechazó redactar este aviso. Reformula lo ocurrido con "
                          "palabras más neutrales y pulsa «Redactar con IA» otra vez.")
        if motivo == "max_tokens":
            raise ErrorIA("Claude se quedó sin espacio para responder. Pulsa «Redactar con IA» "
                          "otra vez; si se repite, acorta lo que escribiste.")
        extra = f" (motivo: {motivo})" if motivo else ""
        raise ErrorIA(f"{_mayus(p['nombre'])} no devolvió texto{extra}. Intenta de nuevo.")
    return texto


def guardar_plantilla(nombre: str, asunto: str, cuerpo: str,
                      ruta: Path = ARCHIVO_PLANTILLAS, sobrescribir: bool = False):
    """Agrega (o reemplaza) una plantilla en plantillas.json, dejando copia en .bak."""
    nombre = nombre.strip()
    if not nombre:
        raise ValueError("Escribe un nombre para la plantilla.")
    if nombre == OPCION_IA:
        raise ValueError("Ese nombre está reservado. Elige otro.")
    armar_mensaje({"asunto": asunto, "cuerpo": cuerpo}, {"alumno": "Nombre", "grupo": "1A"},
                  "Materia", date.today(), "", {})  # lanza ValueError si hay llaves mal puestas
    actuales = cargar_json(ruta)
    if nombre in actuales and not sobrescribir:
        raise FileExistsError(nombre)
    actuales[nombre] = {"asunto": asunto, "cuerpo": cuerpo.split("\n")}
    shutil.copy2(ruta, ruta.with_name(ruta.name + ".bak"))
    temporal = ruta.with_name(ruta.name + ".tmp")
    temporal.write_text(json.dumps(actuales, ensure_ascii=False, indent=2) + "\n",
                        encoding="utf-8")
    os.replace(temporal, ruta)


# ----------------------------------------------------------------------
# Ventana
# ----------------------------------------------------------------------

def iniciar_ventana():
    import queue
    import threading
    import tkinter as tk
    from tkinter import messagebox, simpledialog, ttk

    class App(tk.Tk):
        def __init__(self, cfg, plantillas, alumnos, materias, avisos_carga):
            super().__init__()
            self.title("Avisos a padres de familia")
            self.geometry("1100x780")
            self.minsize(950, 660)

            self.cfg = cfg
            self.plantillas = plantillas
            self.alumnos = alumnos
            self.materias = materias
            self.visibles = []
            self.clave = obtener_secreto("AVISOS_CLAVE").replace(" ", "") or None

            # Redacción con IA
            self.borrador = None            # {"asunto": ..., "cuerpo": ...} con marcadores
            # Para conservar lo que editas a mano en la vista previa:
            self._ctx = None                # (alumno, tipo, borrador) de la última vista
            self._auto = ("", "")           # asunto y cuerpo generados en la última vista
            self._vals = None               # (materia, fecha) de la última vista
            self.clave_ia = clave_ia_de_entorno(cfg)
            self.consentimiento_ia = False
            self.cola_ia = queue.Queue()
            self._brief_enviado = ""

            self.var_grupo = tk.StringVar(value="Todos")
            self.var_busqueda = tk.StringVar()
            self.var_tipo = tk.StringVar(value=list(plantillas)[0])
            self.var_materia = tk.StringVar(value=materias[0])
            self.var_fecha = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
            self.var_detalle = tk.StringVar()
            self.var_asunto = tk.StringVar()
            self.var_tono = tk.StringVar(value=list(TONOS)[0])

            self._construir()

            for var in (self.var_busqueda,):
                var.trace_add("write", lambda *_: self.refrescar_lista())
            for var in (self.var_fecha, self.var_detalle):
                var.trace_add("write", lambda *_: self.actualizar_vista())

            self.refrescar_lista()
            if avisos_carga:
                messagebox.showwarning(
                    "Revisa alumnos.csv", "\n".join(avisos_carga[:15]), parent=self)

        # ---------- construcción de la interfaz ----------
        def _construir(self):
            self.columnconfigure(1, weight=1)
            self.rowconfigure(0, weight=1)

            # Panel izquierdo: alumnos
            izq = ttk.LabelFrame(self, text="1. Alumnos (Ctrl o Shift para varios)", padding=8)
            izq.grid(row=0, column=0, sticky="ns", padx=(10, 5), pady=10)
            izq.rowconfigure(3, weight=1)

            grupos = ["Todos"] + sorted({a["grupo"] for a in self.alumnos if a["grupo"]})
            ttk.Label(izq, text="Grupo:").grid(row=0, column=0, sticky="w")
            cb = ttk.Combobox(izq, textvariable=self.var_grupo, values=grupos,
                              state="readonly", width=12)
            cb.grid(row=0, column=1, sticky="w", padx=5)
            cb.bind("<<ComboboxSelected>>", lambda _e: self.refrescar_lista())

            ttk.Label(izq, text="Buscar:").grid(row=1, column=0, sticky="w", pady=(6, 0))
            ttk.Entry(izq, textvariable=self.var_busqueda, width=24).grid(
                row=1, column=1, sticky="w", padx=5, pady=(6, 0))

            botones = ttk.Frame(izq)
            botones.grid(row=2, column=0, columnspan=2, sticky="w", pady=6)
            ttk.Button(botones, text="Seleccionar todos",
                       command=self.seleccionar_todos).pack(side="left")
            ttk.Button(botones, text="Limpiar",
                       command=self.limpiar_seleccion).pack(side="left", padx=5)

            marco = ttk.Frame(izq)
            marco.grid(row=3, column=0, columnspan=2, sticky="nsew")
            marco.rowconfigure(0, weight=1)
            self.listbox = tk.Listbox(marco, selectmode="extended", width=42,
                                      exportselection=False, activestyle="none")
            sb = ttk.Scrollbar(marco, command=self.listbox.yview)
            self.listbox.config(yscrollcommand=sb.set)
            self.listbox.grid(row=0, column=0, sticky="nsew")
            sb.grid(row=0, column=1, sticky="ns")
            self.listbox.bind("<<ListboxSelect>>", lambda _e: self.actualizar_vista())

            self.lbl_conteo = ttk.Label(izq, text="")
            self.lbl_conteo.grid(row=4, column=0, columnspan=2, sticky="w", pady=(6, 0))

            # Panel derecho: aviso
            der = ttk.LabelFrame(self, text="2. Aviso", padding=8)
            der.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
            der.columnconfigure(1, weight=1)
            der.rowconfigure(7, weight=1)

            ttk.Label(der, text="Tipo:").grid(row=0, column=0, sticky="w")
            self.cb_tipo = ttk.Combobox(der, textvariable=self.var_tipo,
                                        values=list(self.plantillas) + [OPCION_IA],
                                        state="readonly")
            self.cb_tipo.grid(row=0, column=1, sticky="ew", padx=5, pady=2)
            self.cb_tipo.bind("<<ComboboxSelected>>", lambda _e: self._cambio_tipo())

            ttk.Label(der, text="Materia:").grid(row=1, column=0, sticky="w")
            cbm = ttk.Combobox(der, textvariable=self.var_materia,
                               values=self.materias, state="readonly")
            cbm.grid(row=1, column=1, sticky="ew", padx=5, pady=2)
            cbm.bind("<<ComboboxSelected>>", lambda _e: self.actualizar_vista())

            ttk.Label(der, text="Fecha (dd/mm/aaaa):").grid(row=2, column=0, sticky="w")
            ttk.Entry(der, textvariable=self.var_fecha, width=14).grid(
                row=2, column=1, sticky="w", padx=5, pady=2)

            ttk.Label(der, text="Detalle (opcional):").grid(row=3, column=0, sticky="w")
            self.ent_detalle = ttk.Entry(der, textvariable=self.var_detalle)
            self.ent_detalle.grid(row=3, column=1, sticky="ew", padx=5, pady=2)

            # Zona de redacción con IA (solo se ve con «Otra situación»)
            self.marco_ia = ttk.LabelFrame(
                der, text="Describe lo que pasó y la IA redacta el correo", padding=6)
            self.marco_ia.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(8, 0))
            self.marco_ia.columnconfigure(0, weight=1)
            self.txt_brief = tk.Text(self.marco_ia, height=4, wrap="word",
                                     font=("Segoe UI", 10))
            self.txt_brief.grid(row=0, column=0, sticky="ew")
            ttk.Label(
                self.marco_ia, wraplength=650, justify="left",
                text="Escribe breve lo ocurrido. El nombre del alumno elegido NO se envía a la "
                     "IA: se cambia por un marcador y se vuelve a poner aquí. Evita escribir "
                     "nombres de otros alumnos."
            ).grid(row=1, column=0, sticky="w", pady=(4, 4))
            fila_ia = ttk.Frame(self.marco_ia)
            fila_ia.grid(row=2, column=0, sticky="ew")
            ttk.Label(fila_ia, text="Tono:").pack(side="left")
            ttk.Combobox(fila_ia, textvariable=self.var_tono, values=list(TONOS),
                         state="readonly", width=20).pack(side="left", padx=5)
            self.btn_ia = ttk.Button(fila_ia, text="Redactar con IA", command=self.redactar_ia)
            self.btn_ia.pack(side="left", padx=5)
            self.btn_guardar = ttk.Button(fila_ia, text="Guardar como plantilla…",
                                          command=self.guardar_como_plantilla,
                                          state="disabled")
            self.btn_guardar.pack(side="left", padx=5)
            self.marco_ia.grid_remove()

            self.lbl_para = ttk.Label(der, text="", wraplength=650, justify="left")
            self.lbl_para.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 2))

            ttk.Label(der, text="Asunto:").grid(row=6, column=0, sticky="w")
            self.ent_asunto = ttk.Entry(der, textvariable=self.var_asunto)
            self.ent_asunto.grid(row=6, column=1, sticky="ew", padx=5, pady=2)

            marco_txt = ttk.Frame(der)
            marco_txt.grid(row=7, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
            marco_txt.rowconfigure(0, weight=1)
            marco_txt.columnconfigure(0, weight=1)
            self.txt_cuerpo = tk.Text(marco_txt, wrap="word", font=("Segoe UI", 10))
            sb2 = ttk.Scrollbar(marco_txt, command=self.txt_cuerpo.yview)
            self.txt_cuerpo.config(yscrollcommand=sb2.set)
            self.txt_cuerpo.grid(row=0, column=0, sticky="nsew")
            sb2.grid(row=0, column=1, sticky="ns")

            self.lbl_estado = ttk.Label(der, text="", wraplength=650, justify="left")
            self.lbl_estado.grid(row=8, column=0, columnspan=2, sticky="w", pady=(6, 0))

            ttk.Button(der, text="Enviar aviso(s)", command=self.enviar).grid(
                row=9, column=1, sticky="e", pady=(8, 0))

        # ---------- selección ----------
        def seleccionar_todos(self):
            self.listbox.select_set(0, "end")
            self.actualizar_vista()

        def limpiar_seleccion(self):
            self.listbox.selection_clear(0, "end")
            self.actualizar_vista()

        def refrescar_lista(self):
            grupo = self.var_grupo.get()
            texto = self.var_busqueda.get().strip().lower()
            self.visibles = [
                a for a in self.alumnos
                if (grupo == "Todos" or a["grupo"] == grupo) and texto in a["alumno"].lower()
            ]
            self.listbox.delete(0, "end")
            for a in self.visibles:
                etiqueta = a["alumno"]
                if grupo == "Todos" and a["grupo"]:
                    etiqueta += f"  ({a['grupo']})"
                self.listbox.insert("end", etiqueta)
            self.lbl_conteo.config(text=f"{len(self.visibles)} mostrados")
            self.actualizar_vista()

        # ---------- vista previa ----------
        def _fecha_valida(self):
            try:
                return parsear_fecha(self.var_fecha.get())
            except ValueError:
                return None

        def _poner_cuerpo(self, texto, editable):
            self.txt_cuerpo.config(state="normal")
            self.txt_cuerpo.delete("1.0", "end")
            self.txt_cuerpo.insert("1.0", texto)
            self.txt_cuerpo.config(state="normal" if editable else "disabled")

        def actualizar_vista(self):
            sel = self.listbox.curselection()
            if not sel:
                self.lbl_para.config(text="Selecciona uno o más alumnos.")
                self._ctx = None
                self.var_asunto.set("")
                self.ent_asunto.config(state="normal")
                self._poner_cuerpo("", editable=False)
                self.lbl_estado.config(text="")
                return

            fecha = self._fecha_valida()
            if fecha is None:
                self.lbl_estado.config(text="Fecha no válida. Usa el formato dd/mm/aaaa.",
                                       foreground="red")
                return
            self.lbl_estado.config(text="", foreground="")

            al = self.visibles[sel[0]]
            tipo = self.var_tipo.get()
            if tipo == OPCION_IA:
                if self.borrador is None:
                    self.lbl_para.config(
                        text="Escribe lo que pasó y pulsa «Redactar con IA».")
                    self._ctx = None
                    self.var_asunto.set("")
                    self.ent_asunto.config(state="normal")
                    self._poner_cuerpo("", editable=False)
                    return
                plantilla, detalle = self.borrador, ""
            else:
                plantilla, detalle = self.plantillas[tipo], self.var_detalle.get()
            uno = len(sel) == 1
            materia = self.var_materia.get()
            ctx = (tuple(sel), tipo, al["alumno"], al["grupo"],
                   (self.borrador["asunto"], self.borrador["cuerpo"]) if self.borrador else None)
            try:
                asunto, cuerpo = armar_mensaje(plantilla, al, materia, fecha, detalle, self.cfg)
                auto = (asunto, cuerpo)
                conservado = False
                if uno and ctx == self._ctx and self._vals:
                    # Mismo alumno y mismo tipo: si editaste el asunto o el texto, se
                    # conservan tus cambios (ajustando alumno, materia y fecha).
                    materia0, fecha0 = self._vals
                    asunto_act = self.var_asunto.get().strip()
                    cuerpo_act = self.txt_cuerpo.get("1.0", "end").strip()
                    if asunto_act and asunto_act != self._auto[0].strip():
                        asunto = self._rearmar("asunto", asunto_act, al, materia0, fecha0,
                                               materia, fecha, detalle)
                        conservado = True
                    if cuerpo_act and cuerpo_act != self._auto[1].strip():
                        cuerpo = self._rearmar("cuerpo", cuerpo_act, al, materia0, fecha0,
                                               materia, fecha, detalle)
                        conservado = True
            except ValueError as e:
                self.lbl_estado.config(text=str(e), foreground="red")
                return
            self._ctx, self._auto, self._vals = ctx, auto, (materia, fecha)
            if uno:
                self.lbl_para.config(text="Para: " + ", ".join(al["correos"]))
            else:
                self.lbl_para.config(
                    text=f"{len(sel)} alumnos seleccionados. Cada familia recibe su propio "
                         f"correo con su nombre. Vista previa del primero: {al['alumno']}.")
            self.ent_asunto.config(state="normal")
            self.var_asunto.set(asunto)
            self.ent_asunto.config(state="normal" if uno else "disabled")
            self._poner_cuerpo(cuerpo, editable=uno)
            if conservado:
                self.lbl_estado.config(
                    text="Se conservaron los cambios que hiciste al asunto o al texto. "
                         "Para volver al original, cambia de tipo o de alumno.")

        def _rearmar(self, campo, texto, al, materia0, fecha0, materia, fecha, detalle):
            """Vuelve a armar un texto editado a mano con la materia y fecha actuales."""
            plantilla = {"asunto": "x", "cuerpo": "x"}
            plantilla[campo] = plantillizar(texto, al, materia0, fecha0, self.cfg)
            asunto, cuerpo = armar_mensaje(plantilla, al, materia, fecha, detalle, self.cfg)
            return asunto if campo == "asunto" else cuerpo

        # ---------- redacción con IA ----------
        def _cambio_tipo(self):
            if self.var_tipo.get() == OPCION_IA:
                self.marco_ia.grid()
                self.ent_detalle.config(state="disabled")
            else:
                self.marco_ia.grid_remove()
                self.ent_detalle.config(state="normal")
            self.actualizar_vista()

        def _pedir_clave_ia(self, p):
            if p["local"] or self.clave_ia:
                return True
            clave = simpledialog.askstring(
                "Clave de API",
                p["ayuda"] + "\n\n" + diagnostico_secretos(p["variable"])
                + "\n\nLo que escribas aquí no se guarda; se pide una vez por sesión.",
                show="*", parent=self)
            if not clave or not clave.strip():
                return False
            self.clave_ia = clave.strip()
            return True

        def redactar_ia(self):
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showinfo(
                    "Falta elegir",
                    "Primero selecciona al menos un alumno: su nombre se protege antes "
                    "de consultar a la IA.", parent=self)
                return
            brief = self.txt_brief.get("1.0", "end").strip()
            if len(brief) < 15:
                messagebox.showinfo("Describe lo ocurrido",
                                    "Escribe al menos una frase sobre lo que pasó.", parent=self)
                return
            fecha = self._fecha_valida()
            if fecha is None:
                messagebox.showerror("Fecha no válida", "Usa el formato dd/mm/aaaa.",
                                     parent=self)
                return
            try:
                p = datos_proveedor(self.cfg)
            except ErrorIA as e:
                messagebox.showerror("Revisa config.json", str(e), parent=self)
                return
            if not p["local"] and not self.consentimiento_ia:
                if not messagebox.askyesno(
                        "Enviar el texto a la IA",
                        "Para redactar el borrador, lo que escribiste se enviará por internet "
                        f"a {p['nombre']}.\n\n"
                        "El nombre del alumno se reemplaza por un marcador y NO se envía. "
                        "Tampoco se envían correos ni calificaciones.\n\n"
                        "Revisa que el texto no incluya nombres de otros alumnos ni datos "
                        "que tu escuela no permita compartir.\n\n¿Continuar?", parent=self):
                    return
                self.consentimiento_ia = True
            if not self._pedir_clave_ia(p):
                return

            alumnos_sel = [self.visibles[i] for i in sel]
            brief_marcado = a_marcadores(brief, alumnos_sel, self.var_materia.get(),
                                         fecha_larga(fecha))
            sistema, usuario = construir_peticion_ia(brief_marcado, self.var_tono.get())
            self._brief_enviado = brief_marcado
            self.btn_ia.config(state="disabled", text="Redactando…")
            self.lbl_estado.config(text="Consultando a la IA… puede tardar unos segundos.",
                                   foreground="")
            threading.Thread(target=self._trabajo_ia, args=(sistema, usuario),
                             daemon=True).start()
            self.after(150, self._revisar_cola_ia)

        def _trabajo_ia(self, sistema, usuario):
            try:
                texto = llamar_ia(self.cfg, self.clave_ia or "", sistema, usuario)
                self.cola_ia.put(("ok", parsear_borrador(texto)))
            except ErrorIA as e:
                self.cola_ia.put(("error", e))
            except Exception as e:  # noqa: BLE001
                self.cola_ia.put(("error", ErrorIA(f"Error inesperado: {e}")))

        def _revisar_cola_ia(self):
            try:
                estado, valor = self.cola_ia.get_nowait()
            except queue.Empty:
                self.after(150, self._revisar_cola_ia)
                return
            self.btn_ia.config(state="normal", text="Redactar con IA")
            if estado == "error":
                if valor.autenticacion:
                    self.clave_ia = None
                self.lbl_estado.config(text="")
                messagebox.showerror("No se pudo redactar", str(valor), parent=self)
                return
            self.borrador = valor
            self._ctx = None            # borrador nuevo: no conservar ediciones anteriores
            self.btn_guardar.config(state="normal")
            self.actualizar_vista()
            resumen = " ".join(self._brief_enviado.split())
            if len(resumen) > 220:
                resumen = resumen[:220] + "…"
            self.lbl_estado.config(
                text=f"Revisa el borrador antes de enviar. Esto fue lo que se envió a la IA: "
                     f"«{resumen}»", foreground="")

        def guardar_como_plantilla(self):
            if self.borrador is None:
                return
            asunto, cuerpo = self.borrador["asunto"], self.borrador["cuerpo"]
            sel = self.listbox.curselection()
            fecha = self._fecha_valida()
            if len(sel) == 1 and fecha is not None:
                # Si editaste el texto a mano, se guarda tu versión (con marcadores).
                al, materia = self.visibles[sel[0]], self.var_materia.get()
                esp_asunto, esp_cuerpo = armar_mensaje(
                    self.borrador, al, materia, fecha, "", self.cfg)
                act_asunto = self.var_asunto.get().strip()
                act_cuerpo = self.txt_cuerpo.get("1.0", "end").strip()
                if act_asunto and act_asunto != esp_asunto.strip():
                    asunto = plantillizar(act_asunto, al, materia, fecha, self.cfg)
                if act_cuerpo and act_cuerpo != esp_cuerpo.strip():
                    cuerpo = plantillizar(act_cuerpo, al, materia, fecha, self.cfg)

            nombre = simpledialog.askstring(
                "Guardar como plantilla",
                "Nombre de la plantilla (aparecerá en la lista «Tipo»):", parent=self)
            if not nombre or not nombre.strip():
                return
            nombre = nombre.strip()
            try:
                try:
                    guardar_plantilla(nombre, asunto, cuerpo)
                except FileExistsError:
                    if not messagebox.askyesno(
                            "Ya existe",
                            f"Ya hay una plantilla llamada «{nombre}». ¿Reemplazarla?\n"
                            "(Se guarda una copia en plantillas.json.bak)", parent=self):
                        return
                    guardar_plantilla(nombre, asunto, cuerpo, sobrescribir=True)
            except (ValueError, OSError) as e:
                messagebox.showerror("No se pudo guardar la plantilla", str(e), parent=self)
                return

            self.plantillas = cargar_plantillas()
            self.cb_tipo.config(values=list(self.plantillas) + [OPCION_IA])
            self.borrador = None
            self.btn_guardar.config(state="disabled")
            self.txt_brief.delete("1.0", "end")
            self.var_tipo.set(nombre)
            self._cambio_tipo()
            messagebox.showinfo(
                "Plantilla guardada",
                f"Guardé «{nombre}». Ya aparece en la lista «Tipo» y la puedes retocar "
                "en plantillas.json.", parent=self)

        # ---------- envío ----------
        def _pedir_clave(self):
            if self.clave:
                return True
            clave = simpledialog.askstring(
                "Contraseña de aplicación",
                "Escribe tu contraseña de aplicación de Google (16 letras).\n\n"
                + diagnostico_secretos("AVISOS_CLAVE")
                + "\n\nLo que escribas aquí no se guarda; se pide una vez por sesión.",
                show="*", parent=self)
            if not clave:
                return False
            self.clave = clave.replace(" ", "")
            return True

        def enviar(self):
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showinfo("Falta elegir", "Selecciona al menos un alumno.", parent=self)
                return
            fecha = self._fecha_valida()
            if fecha is None:
                messagebox.showerror("Fecha no válida", "Usa el formato dd/mm/aaaa.", parent=self)
                return

            tipo = self.var_tipo.get()
            ia = tipo == OPCION_IA
            if ia and self.borrador is None:
                messagebox.showinfo(
                    "Falta el borrador",
                    "Describe lo que pasó y pulsa «Redactar con IA» antes de enviar.",
                    parent=self)
                return
            materia = self.var_materia.get()
            fecha_txt = fecha.strftime("%d/%m/%Y")
            if ia:
                plantilla = self.borrador
                detalle = self.txt_brief.get("1.0", "end").strip()   # queda en el historial
                detalle_plantilla = ""
                tipo = TIPO_REGISTRO_IA
            else:
                plantilla = self.plantillas[tipo]
                detalle = detalle_plantilla = self.var_detalle.get().strip()
            uno = len(sel) == 1

            mensajes = []
            try:
                for i in sel:
                    al = self.visibles[i]
                    if uno:
                        asunto = self.var_asunto.get().strip()
                        cuerpo = self.txt_cuerpo.get("1.0", "end").strip()
                    else:
                        asunto, cuerpo = armar_mensaje(
                            plantilla, al, materia, fecha, detalle_plantilla, self.cfg)
                    if not asunto or not cuerpo:
                        messagebox.showerror("Mensaje vacío",
                                             "El asunto y el cuerpo no pueden estar vacíos.",
                                             parent=self)
                        return
                    mensajes.append({
                        "alumno": al, "materia": materia, "tipo": tipo, "fecha_txt": fecha_txt,
                        "detalle": detalle, "destinatarios": al["correos"],
                        "email": construir_email(self.cfg, al["correos"], asunto, cuerpo),
                    })
            except ValueError as e:
                messagebox.showerror("Error en la plantilla", str(e), parent=self)
                return

            ejemplo = [m["alumno"]["alumno"] for m in mensajes
                       if es_correo_de_ejemplo(m["destinatarios"])]
            if ejemplo:
                messagebox.showerror(
                    "Correos de ejemplo",
                    "Estos alumnos todavía tienen correos de ejemplo (example.com). "
                    "Cámbialos en alumnos.csv:\n\n" + "\n".join(ejemplo[:10]), parent=self)
                return

            repetidos = []
            if not ia:   # los textos redactados con IA son siempre situaciones distintas
                previos = leer_enviados_previos()
                repetidos = [m["alumno"]["alumno"] for m in mensajes
                             if (m["alumno"]["alumno"], materia, tipo, fecha_txt) in previos]

            pregunta = f"¿Enviar {len(mensajes)} correo(s) de tipo «{tipo}» en {materia}?"
            if repetidos:
                pregunta += ("\n\nOJO: ya se envió este mismo aviso, con la misma fecha, a:\n"
                             + "\n".join(repetidos[:10]))
            if not messagebox.askyesno("Confirmar envío", pregunta, parent=self):
                return

            if not self._pedir_clave():
                return

            self.lbl_estado.config(text="Enviando…", foreground="")
            self.update_idletasks()
            try:
                resultados = enviar_lote(self.cfg, self.clave, mensajes)
            except smtplib.SMTPAuthenticationError:
                self.clave = None
                self.lbl_estado.config(text="")
                messagebox.showerror(
                    "No se pudo iniciar sesión",
                    "Google rechazó la contraseña. Revisa que sea la contraseña de "
                    "aplicación (16 letras, sin espacios) y que el correo de config.json "
                    "sea el correcto.", parent=self)
                return
            except Exception as e:  # noqa: BLE001
                self.lbl_estado.config(text="")
                messagebox.showerror("No se pudo enviar",
                                     f"Error de conexión o del servidor:\n{e}", parent=self)
                return

            try:
                registrar(resultados)
                aviso_registro = ""
            except OSError as e:
                aviso_registro = (f"\n\nNo se pudo escribir registro_avisos.csv ({e}). "
                                  "¿Lo tienes abierto en Excel?")

            ok = [r for r in resultados if r[1] == "enviado"]
            mal = [r for r in resultados if r[1] != "enviado"]
            texto = f"Enviados: {len(ok)}."
            if mal:
                texto += f"\nCon error: {len(mal)}\n" + "\n".join(
                    f"- {m['alumno']['alumno']}: {err}" for m, _s, err in mal[:10])
            texto += aviso_registro
            self.lbl_estado.config(text=f"Enviados: {len(ok)}. Con error: {len(mal)}.")
            (messagebox.showwarning if (mal or aviso_registro) else messagebox.showinfo)(
                "Resultado", texto, parent=self)
            if ok:
                if ia:
                    self.borrador = None
                    self.btn_guardar.config(state="disabled")
                    self.txt_brief.delete("1.0", "end")
                self.limpiar_seleccion()
                self.var_detalle.set("")

    # --- arranque con manejo de errores amigables ---
    raiz = tk.Tk()
    raiz.withdraw()
    try:
        cfg = cargar_json(ARCHIVO_CONFIG)
        validar_config(cfg)
        plantillas = cargar_plantillas()
        alumnos, avisos = cargar_alumnos()
        if not alumnos:
            raise ValueError("alumnos.csv no tiene alumnos válidos.")
        materias = cargar_materias()
        if not materias:
            raise ValueError("materias.csv está vacío. Escribe una materia por fila.")
    except FileNotFoundError as e:
        messagebox.showerror(
            "Falta un archivo",
            f"No encuentro: {Path(e.filename).name}\n"
            "Debe estar en la misma carpeta que avisos_padres.py.")
        raiz.destroy()
        return
    except (ValueError, KeyError, json.JSONDecodeError) as e:
        messagebox.showerror("Revisa tus archivos de configuración", str(e))
        raiz.destroy()
        return
    raiz.destroy()

    App(cfg, plantillas, alumnos, materias, avisos).mainloop()


if __name__ == "__main__":
    iniciar_ventana()
