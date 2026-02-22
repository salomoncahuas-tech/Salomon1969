# Guia para Compartir tu Aplicacion Web mediante un Enlace

Tu aplicacion **English Assessment Platform** (carpeta `english_assessment/`) es una
app web Flask que puede desplegarse en internet para que cualquier persona acceda
con un simple enlace.

A continuacion tienes **3 opciones gratuitas** ordenadas de mas facil a mas avanzada.

---

## Opcion 1: Render.com (Recomendada - Gratis)

Render es la plataforma mas sencilla para desplegar apps Flask gratis.

### Pasos:

1. **Crear cuenta en Render**
   - Ve a [https://render.com](https://render.com)
   - Registrate con tu cuenta de GitHub

2. **Conectar tu repositorio**
   - Haz clic en **"New +"** > **"Web Service"**
   - Selecciona tu repositorio **Salomon1969** de GitHub
   - Si no aparece, haz clic en "Configure account" para dar permisos

3. **Configurar el servicio**
   - **Name:** `english-assessment` (o el nombre que prefieras)
   - **Region:** Oregon (US West) o el mas cercano
   - **Branch:** `master` (o la rama que prefieras)
   - **Root Directory:** `english_assessment`
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app --bind 0.0.0.0:$PORT`
   - **Instance Type:** Free

4. **Variables de entorno** (en la seccion Environment)
   - `SECRET_KEY` = (genera una clave segura, ejemplo: `mi-clave-secreta-2024`)
   - `PYTHON_VERSION` = `3.11.0`

5. **Hacer clic en "Create Web Service"**
   - Render construira y desplegara tu app automaticamente
   - En unos minutos tendras tu enlace: `https://english-assessment.onrender.com`

6. **Compartir el enlace**
   - Copia la URL y compartela con tus alumnos
   - Ejemplo: `https://english-assessment.onrender.com`

> **Nota:** En el plan gratuito de Render, la app se "duerme" tras 15 min de
> inactividad. La primera visita despues de estar dormida tarda ~30 segundos
> en cargar.

---

## Opcion 2: PythonAnywhere (Gratis)

Ideal si prefieres un panel de control sencillo sin necesidad de Git.

### Pasos:

1. **Crear cuenta**
   - Ve a [https://www.pythonanywhere.com](https://www.pythonanywhere.com)
   - Registrate con una cuenta gratuita (Beginner)

2. **Subir los archivos**
   - Ve a la pestana **"Files"**
   - Crea una carpeta `english_assessment`
   - Sube todos los archivos de la carpeta `english_assessment/` de tu proyecto

3. **Crear un entorno virtual**
   - Ve a la pestana **"Consoles"** > "Bash"
   - Ejecuta:
     ```bash
     mkvirtualenv --python=/usr/bin/python3.10 myenv
     cd english_assessment
     pip install -r requirements.txt
     ```

4. **Configurar la Web App**
   - Ve a la pestana **"Web"**
   - Haz clic en **"Add a new web app"**
   - Selecciona **"Manual configuration"** > Python 3.10
   - En **"Source code"**: `/home/TU_USUARIO/english_assessment`
   - En **"Virtualenv"**: `/home/TU_USUARIO/.virtualenvs/myenv`

5. **Editar el archivo WSGI**
   - Haz clic en el enlace del archivo WSGI
   - Reemplaza todo el contenido con:
     ```python
     import sys
     path = '/home/TU_USUARIO/english_assessment'
     if path not in sys.path:
         sys.path.append(path)
     from app import app as application
     ```

6. **Recargar la app**
   - Haz clic en el boton verde **"Reload"**
   - Tu app estara en: `https://TU_USUARIO.pythonanywhere.com`

---

## Opcion 3: Railway.app (Gratis con limites)

Railway ofrece un plan gratuito con 500 horas/mes.

### Pasos:

1. **Crear cuenta**
   - Ve a [https://railway.app](https://railway.app)
   - Inicia sesion con GitHub

2. **Nuevo proyecto**
   - Haz clic en **"New Project"**
   - Selecciona **"Deploy from GitHub Repo"**
   - Escoge tu repositorio **Salomon1969**

3. **Configurar**
   - En **Settings** > **Root Directory**: `english_assessment`
   - En **Variables**, agrega:
     - `SECRET_KEY` = `tu-clave-secreta`
   - Railway detectara automaticamente que es una app Python

4. **Generar dominio**
   - Ve a **Settings** > **Networking** > **Generate Domain**
   - Obtendras un enlace como: `https://english-assessment-production.up.railway.app`

---

## Cargar datos de ejemplo

Una vez desplegada, para cargar las evaluaciones de ejemplo necesitas ejecutar
el seeder. Dependiendo de la plataforma:

### En Render:
- Ve a tu servicio > **"Shell"**
- Ejecuta: `cd /opt/render/project/src && python -c "from seed_data import seed_assessments; seed_assessments()"`

### En PythonAnywhere:
- Abre una consola Bash
- Ejecuta:
  ```bash
  cd ~/english_assessment
  workon myenv
  python -c "from seed_data import seed_assessments; seed_assessments()"
  ```

### En Railway:
- Cambia temporalmente el Start Command a:
  `python -c "from seed_data import seed_assessments; seed_assessments()" && gunicorn app:app`
- Haz redeploy, luego vuelve al comando original

---

## Credenciales de Administrador

Una vez desplegada la aplicacion, accede al panel de administracion:

- **URL:** `https://TU-DOMINIO/admin/login`
- **Usuario:** `admin`
- **Contrasena:** `admin123`

> **IMPORTANTE:** Cambia la contrasena del administrador despues del primer
> inicio de sesion en un entorno de produccion.

---

## Compartir con tus Alumnos

Una vez desplegada, tienes varias formas de compartir:

1. **Enlace directo:** Comparte la URL principal
   - Ejemplo: `https://english-assessment.onrender.com`

2. **Codigo de acceso:** Cada evaluacion tiene un codigo unico
   - Los alumnos van a `/access` e ingresan el codigo
   - Puedes compartir solo el codigo por WhatsApp o en clase

3. **Codigo QR:** Genera un QR de tu enlace en [https://www.qr-code-generator.com](https://www.qr-code-generator.com)
   - Imprimelo y pegalo en el salon de clases

4. **Google Classroom / WhatsApp:** Pega el enlace directamente

---

## Resumen Rapido

| Plataforma     | Dificultad | Costo    | Enlace personalizado        | Velocidad |
|----------------|------------|----------|-----------------------------|-----------|
| Render.com     | Facil      | Gratis   | nombre.onrender.com         | Media     |
| PythonAnywhere | Media      | Gratis   | usuario.pythonanywhere.com  | Buena     |
| Railway.app    | Facil      | Gratis*  | nombre.up.railway.app       | Rapida    |

*Railway: 500 horas gratis/mes, suficiente para uso escolar.
