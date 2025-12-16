# 🧰 OfficeToolbox Suite

<div align="center">

![Python](https://img.shields.io/badge/Python-3.11%20|%203.12-3776AB?style=flat&logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?style=flat&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

**OfficeToolbox** es una suite de escritorio Open Source *Todo en Uno* para la gestión documental y multimedia. Permite convertir, unir, extraer y optimizar archivos de forma **100% local**, sin subir datos a la nube, garantizando seguridad y privacidad.

</div>

---

## 🚀 Características

| Categoría | Herramientas |
| :--- | :--- |
| **📄 Documentos** | Office a PDF, PDF a Word, PDF a PowerPoint, Unir PDFs |
| **🎥 Multimedia** | Imágenes a PDF, Video a Fotos, Extraer Imágenes de PDF |
| **🚀 Optimización** | Optimizar PDF (Compresión), Optimizar Imágenes (JPG/PNG) |

---

## 🛠️ Instalación y Desarrollo (Importante)

> [!IMPORTANT]
> **REQUISITO CRÍTICO DE VERSIÓN**
> Este proyecto **requiere estrictamente Python 3.11 o 3.12**.
> ❌ **NO utilices Python 3.14 (Beta) ni 3.13**, ya que algunas dependencias aún no son compatibles.

### 1️⃣ Clonar el repositorio y configurar el entorno

Abre tu terminal (PowerShell o CMD) y ejecuta:

```powershell
# Clonar el repositorio
git clone [https://github.com/GrAlmazan/office-toolbox-suite.git](https://github.com/GrAlmazan/office-toolbox-suite.git)
cd office-toolbox-suite

# Crear entorno virtual forzando la versión 3.11 (o 3.12)
py -3.11 -m venv .venv

# Activar el entorno virtual
.\.venv\Scripts\Activate
```

### 2️⃣ Instalar dependencias

Con el entorno virtual activado:

```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

---

## 📦 Compilación (Crear Ejecutable .exe)

Para generar el archivo `OfficeToolbox.exe` portable y standalone, asegúrate de estar en el entorno virtual (`.venv`) y ejecuta el siguiente comando en una sola línea:

```powershell
pyinstaller --noconfirm --onedir --windowed --name "OfficeToolbox" --icon="icon.ico" --add-data "icon.ico;." --hidden-import "docx2pdf" --hidden-import "comtypes.stream" --hidden-import "pythoncom" --hidden-import "PIL" --hidden-import "pdf2docx" --hidden-import "cv2" --hidden-import "numpy" --hidden-import "pptx" --hidden-import "fitz" main.py
```

### 📂 Resultado
Una vez finalizado el proceso, encontrarás el ejecutable en:
`dist/OfficeToolbox/OfficeToolbox.exe`

---

## 📄 Licencia

Este proyecto está bajo la **Licencia MIT**.
Desarrollado por **GrAlmazan**.