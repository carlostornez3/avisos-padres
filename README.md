# Avisos a padres de familia

Programa de escritorio (Python + Tkinter) para que un docente envíe avisos de incidencias
(tarea, conducta, reprobación, reconocimiento…) a las familias desde su propio Gmail,
con vista previa editable, historial de envíos y, opcionalmente, redacción asistida por IA.

- Sin dependencias externas: solo la biblioteca estándar de Python (incluye Tkinter).
- Envío por SMTP de Gmail con **contraseña de aplicación** (requiere verificación en dos pasos).
- Alumnos y materias en tablas separadas (`alumnos.csv`, `materias.csv`).
- Plantillas editables (`plantillas.json`) con los campos `{alumno}`, `{grupo}`, `{materia}`, `{fecha}`, `{profesor}`, `{detalle}`.
- Tipo «Otra situación (redactar con IA)»: describes lo ocurrido, la IA redacta, tú revisas, editas y envías; opcionalmente lo guardas como plantilla.
- Privacidad: el nombre del alumno, grupo, materia y fecha se cambian por marcadores **antes** de enviar el texto a la IA y se restituyen localmente.
- IA compatible: Anthropic (predeterminado), Google Gemini y cualquier servicio con formato `chat/completions` (OpenRouter, OpenAI, Ollama local…).

## Requisitos

Python 3.9 o superior (en Windows, instálalo desde python.org marcando «Add Python to PATH»).

## Primera vez

1. Clona o descarga el repositorio.
2. Crea tus archivos de trabajo copiando los de ejemplo (en PowerShell):

   ```powershell
   copy alumnos.example.csv alumnos.csv
   copy config.example.json config.json
   copy secrets.example.txt secrets.txt
   ```

3. Edita `config.json` (tu correo, nombre y firma), `alumnos.csv` (alumnos y correos de las familias; el programa se niega a enviar a direcciones `example.com`) y `secrets.txt` (contraseña de aplicación de Gmail y, si usas IA, su clave).
4. Ejecuta:

   ```powershell
   python .\avisos_padres.py
   ```

La guía de uso completa (incluye cómo crear la contraseña de aplicación, elegir proveedor de IA y variables de entorno) está en [`LEEME.txt`](LEEME.txt).

## Datos personales: qué NO se sube

`alumnos.csv`, `config.json`, `secrets.txt` y `registro_avisos.csv` contienen datos de menores, correos y claves, y están en `.gitignore`. En el repositorio solo hay archivos de ejemplo. Antes de cada commit revisa con `git status` que ninguno de esos cuatro aparezca.

## Uso responsable

El programa envía correos reales a familias reales. Revisa siempre la vista previa. Si usas IA, lo que escribas se envía al proveedor elegido (el nombre del alumno se protege, pero no el resto del texto): revisa las políticas de tu escuela y de tu proveedor, y no incluyas datos sensibles.

## Licencia

MIT. Ver [`LICENSE`](LICENSE).
