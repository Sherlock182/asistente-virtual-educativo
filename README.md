# 🎓 Asistente Virtual Educativo

Aplicación de escritorio para Windows pensada para apoyar el trabajo diario de un docente: chat con IA (Groq), recordatorios con notificaciones, control de asistencia con Excel, calendario, juegos educativos (sopa de letras, laberinto, ahorcado, crucigrama) y un editor de PDF.

## 📥 Descargar y usar la aplicación (recomendado para la mayoría de usuarios)

No necesitas instalar Python ni nada más:

1. Ve a la sección **[Releases](../../releases)** de este repositorio.
2. Descarga el archivo `Asistente Virtual Educativo.exe` de la versión más reciente.
3. Haz doble clic para abrirlo. La primera vez, Windows puede mostrar un aviso de "Editor desconocido" — dale a **Más información → Ejecutar de todas formas** (es normal para aplicaciones que no están firmadas digitalmente).
4. La app crea automáticamente una carpeta **"Asistente Virtual"** en tu Escritorio, donde guarda todos tus datos (alumnos, asistencia, recordatorios, calendario, etc.). No la borres ni la muevas si quieres conservar tu información.

## 🧑‍💻 Ejecutar desde el código fuente (para desarrolladores)

Requisitos: Python 3.10 o superior, en Windows.

```bash
git clone <URL-de-este-repositorio>
cd <carpeta-del-repositorio>
pip install -r requirements.txt
python asistente_virtual_x_RECUPERADO.py
```

## 🔑 Clave de API de Groq

El chat con IA usa la API de Groq (https://groq.com). Cada persona usa **su propia clave gratuita**:

1. Créala en https://console.groq.com/keys.
2. En la app, ve a **Configuración → Chat con IA**, pégala en "Tu clave de API de Groq" y presiona **Guardar**. La app comprueba que funcione.

La clave se guarda solo en tu computadora (carpeta "Asistente Virtual"); no va en el código ni en el `.exe`.

## 🔄 Actualizaciones automáticas

La app revisa este repositorio al abrirse y cada 6 horas. Cuando hay una versión nueva en **Releases**, aparece **"Nueva"** en el menú lateral → **Actualizaciones**. Con un clic descarga la versión nueva, guarda todo, se cierra y se vuelve a abrir actualizada. No hace falta reinstalar y los datos de la carpeta "Asistente Virtual" no se tocan.

**Para publicar una versión nueva** (después de hacer tus commits):

```bash
git tag v1.1.0
git push origin v1.1.0
```

GitHub arma el `.exe` solo (ver `.github/workflows/publicar-version.yml`), escribe el número de versión dentro de la app y usa los mensajes de tus commits como lista de novedades. El número siempre debe ser mayor que el anterior (v1.1.0 → v1.2.0 → v2.0.0).

## 🛠️ Generar tu propio ejecutable (.exe)

Si modificas el código y quieres generar un nuevo `.exe`:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=logo.ico --name "Asistente Virtual Educativo" asistente_virtual_x_RECUPERADO.py
```

El resultado queda en la carpeta `dist/`.

## ✨ Funcionalidades principales

- **Chat con IA (Groq):** respuestas de texto, con opción de lectura en voz alta (voz neuronal en español).
- **Recordatorios:** con notificaciones (incluyendo aviso por voz, repetido cada minuto si no se atiende).
- **Asistencia:** registro mensual por alumno, cálculo de notas, exportación y sincronización automática con Excel.
- **Calendario:** carga y conserva un PDF de calendario entre sesiones.
- **Juegos educativos:** Sopa de Letras, Laberinto, Ahorcado y Crucigrama, todos con temas educativos configurables y generación de hojas para imprimir en Word.
- **Editor de PDF:** edición de texto existente, inserción de texto nuevo, y reemplazo/movimiento/redimensión de imágenes dentro del PDF.
- **Preferencias recordadas:** tamaño y posición de ventana, última pestaña abierta, y configuración de voz/notificaciones se recuerdan entre sesiones.
- **Protección de datos:** todos los datos se guardan en una carpeta fija en el Escritorio, detectando si la app se ejecuta en la misma computadora para conservar los datos, o en una diferente para iniciar de cero.

## 👤 Autor

**Desarrollado por:** Ing. Osea Nahun Montalvo Rogel
**Correo electrónico:** 4603642023@mail.utec.edu.sv
**GitHub:** [github.com/Sherlock182](https://github.com/Sherlock182)
