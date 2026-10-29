# 🎓 VerbFlow — Plataforma Institucional de Aprendizaje de Inglés
**Institución Educativa Promoción Social — Palermo**

VerbFlow es una plataforma web interactiva y modular diseñada para el fortalecimiento de competencias en el idioma inglés (tiempos verbales, práctica oral con reconocimiento de voz, roleplays interactivos y suites para estudiantes, docentes y directivos).

---

## 📁 Estructura del Proyecto

```text
PromocionSocialEnglish/
├── .stitch/                    # Diseños y prototipos visuales (Stitch UI)
├── Docs/                       # Documentación formal de ingeniería de software
│   ├── DocumentacionSoftware/  # Especificaciones oficiales en formato Word (.docx):
│   │   ├── SAD (Documento de Arquitectura y API)
│   │   ├── SRS (Especificación de Requisitos de Software)
│   │   ├── Modelo de Datos (Diagramas y Diccionario de Datos)
│   │   ├── STP (Plan de Pruebas de Software)
│   │   ├── Manual de Usuario y Operación
│   │   └── Política de Tratamiento de Datos Personales
│   ├── build_documents.py      # Generador automatizado de documentación
│   └── requirements.txt        # Dependencias de Python para compilar documentos
├── Software/                   # Código fuente de la aplicación web
│   ├── backend/                # Scripts de base de datos (schema.sql Supabase)
│   ├── css/                    # Hojas de estilo institucionales (Campus Blue & responsivo)
│   ├── data/                   # Bancos de preguntas, tiempos verbales y escenarios
│   ├── html/                   # Páginas web por rol:
│   │   ├── index.html          # Portal Estudiantil
│   │   ├── teacher.html        # Suite Docente
│   │   ├── owner.html          # Panel Directivo / Administrador
│   │   └── roleplay.html       # Módulo de Práctica Oral & Simulación
│   ├── Img/                    # Recursos multimedia y gráficos
│   └── js/                     # Lógica frontend (Speech Engine, Auth, Supabase, Theme)
├── .gitignore                  # Exclusiones de Git (entornos virtuales, temporales)
└── README.md                   # Guía general de desarrollo y colaboración
```

---

## 🔄 Protocolo de Documentación y Sincronización

> **REGLA FUNDAMENTAL DEL EQUIPO:**  
> Cada vez que se agregue una nueva funcionalidad, se modifique una pantalla o se altere el modelo de datos en `Software/`, **debe actualizarse en sincronía la documentación correspondiente**:
> 1. Si cambia la base de datos o endpoints: actualizar `backend/schema.sql`, `generate_model_doc.py` y `generate_sad_doc.py`.
> 2. Si cambia un flujo de usuario o requerimiento: actualizar `generate_srs_doc.py` y `generate_manual_doc.py`.
> 3. Ejecutar los scripts de `Docs/` para regenerar los documentos oficiales en `.docx`.
> 4. Registrar los cambios en los commits de Git.

---

## 👥 Flujo de Trabajo en Equipo (Colaboración en GitHub)

Para que ambos puedan trabajar al mismo tiempo sin sobrescribir el trabajo del otro y manteniendo la página funcionando en línea:

### 1. Clonar el repositorio
```bash
git clone https://github.com/Kevin81A/PromocionSocialEnglish.git
cd PromocionSocialEnglish
```

### 2. Trabajar siempre en ramas (`branches`)
Nunca trabajen directamente sobre `main`. Creen una rama para cada tarea:
```bash
# Crear y cambiarse a una nueva rama
git checkout -b feature/nombre-de-la-mejora
```

### 3. Guardar cambios y subir la rama
```bash
git add .
git commit -m "feat: descripción clara del cambio realizado"
git push origin feature/nombre-de-la-mejora
```

### 4. Crear un Pull Request (PR) en GitHub
* Entren a GitHub y abran un **Pull Request** hacia `main`.
* El compañero revisa los cambios y aprueba el **Merge**.
* Una vez unido a `main`, la página en vivo se actualiza automáticamente.

### 5. Mantener la copia local actualizada
Antes de empezar a trabajar cada día:
```bash
git checkout main
git pull origin main
```

---

## 🚀 Cómo Ejecutar la Aplicación Localmente

1. **Abrir directamente:**  
   Puedes abrir `Software/html/index.html` en cualquier navegador web moderno (Edge, Chrome, Firefox).

2. **Con servidor local (Recomendado para APIs y Speech Engine):**
   ```bash
   # Opción con Python:
   python -m http.server 8000
   # Luego abrir en el navegador: http://localhost:8000/Software/html/
   ```
