# -*- coding: utf-8 -*-
import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_styling(table, col_widths=None, header_bg="1F3A60", alt_bg="F1F5F9"):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if col_widths and idx < len(col_widths):
                cell.width = Inches(col_widths[idx])
                
    # Header row
    for cell in table.rows[0].cells:
        set_cell_background(cell, header_bg)
        for p in cell.paragraphs:
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(9.5)
                run.font.name = "Segoe UI"
                
    # Data rows
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg = alt_bg if r_idx % 2 == 1 else "FFFFFF"
        for cell in row.cells:
            set_cell_background(cell, bg)
            for p in cell.paragraphs:
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                for run in p.runs:
                    run.font.size = Pt(9)
                    run.font.name = "Segoe UI"

def add_heading_with_spacing(doc, text, level):
    p = doc.add_heading(text, level=level)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    for run in p.runs:
        run.font.name = "Segoe UI"
    return p

def add_p(doc, text, bold_prefix=None, style='Normal'):
    p = doc.add_paragraph(style=style)
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        r1.font.bold = True
        r1.font.name = "Segoe UI"
        r1.font.size = Pt(10)
    r2 = p.add_run(text)
    r2.font.name = "Segoe UI"
    r2.font.size = Pt(10)
    return p

def add_bullet(doc, title, text):
    p = doc.add_paragraph(style='List Paragraph')
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    r1 = p.add_run(title + ": ")
    r1.font.bold = True
    r1.font.name = "Segoe UI"
    r1.font.size = Pt(10)
    r2 = p.add_run(text)
    r2.font.name = "Segoe UI"
    r2.font.size = Pt(10)
    return p

def add_table_data(doc, headers, data, col_widths=None):
    table = doc.add_table(rows=len(data) + 1, cols=len(headers))
    table.style = 'Table Grid'
    for c_idx, h in enumerate(headers):
        table.cell(0, c_idx).text = h
    for r_idx, row in enumerate(data, start=1):
        for c_idx, val in enumerate(row):
            table.cell(r_idx, c_idx).text = str(val)
    set_table_styling(table, col_widths=col_widths)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    return table

def generate_sad_doc(filepath):
    doc = docx.Document()
    
    # Title
    p_title = doc.add_heading("VerbFlow — Documento de Arquitectura de Software y Especificación de la API (SAD)", level=1)
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(12)
    r_meta = p_meta.add_run("Versión 1.0 (Definitiva) · Autor: Kevin · Basado en IEEE 42010 / C4 Model · Octubre de 2026")
    r_meta.font.italic = True
    r_meta.font.size = Pt(9.5)
    r_meta.font.color.rgb = RGBColor(100, 116, 139)
    
    # 1. Introducción y Objetivos de Arquitectura
    add_heading_with_spacing(doc, "1. Introducción y Objetivos de Arquitectura", level=2)
    add_p(doc, "Este documento define la arquitectura técnica global, los patrones de diseño, las restricciones de seguridad y el catálogo detallado de endpoints de la API REST para VerbFlow. Está adaptado y optimizado para la operación privada en la Institución Educativa Promoción Social (Palermo, Huila, Colombia), satisfaciendo los requerimientos del SRS v0.3 y las 19 entidades del Modelo de Datos v0.3.")
    
    add_heading_with_spacing(doc, "1.1 Metas de Calidad Arquitectónica", level=3)
    add_bullet(doc, "Seguridad y Privacidad Estricta (Ley 1581 / Ley 1098)", "Control de acceso basado en roles (RBAC) y Row Level Security (RLS). Cero almacenamiento de audios o biométricos de menores en servidores. Respuestas de exámenes oficiales ocultas del cliente hasta el momento de calificación.")
    add_bullet(doc, "Tolerancia a Conectividad Escolar Inestable (RNF-14)", "Mecanismo de autoguardado progresivo e idempotente para respuestas de examen, asegurando que un corte de fluido eléctrico o de red wifi no signifique pérdida de calificaciones.")
    add_bullet(doc, "Rendimiento y Ligereza (RNF-05 / RNF-06)", "Cargas de página inferiores a 2.5 segundos en redes móviles 4G. Respuestas de API con latencias inferiores a 250 ms en el 95% de las peticiones para soportar 300 estudiantes concurrentes rindiendo pruebas.")
    add_bullet(doc, "Sostenibilidad de Costos", "Arquitectura 'serverless-friendly' / PaaS de costo cero o mínimo: reconocimiento de voz ejecutado en el cliente (Web Speech API) y backend relacional PostgreSQL gestionado.")

    # 2. Vista General de la Arquitectura (Modelo C4)
    add_heading_with_spacing(doc, "2. Vista General de la Arquitectura (Modelo C4)", level=2)
    add_p(doc, "La arquitectura se describe mediante el estándar C4 (Contexto, Contenedores y Componentes) para garantizar claridad conceptual.")
    
    add_heading_with_spacing(doc, "2.1 Nivel 1: Contexto del Sistema", level=3)
    add_p(doc, "VerbFlow interactúa con tres actores humanos (Estudiante, Docente y Owner) y un subsistema de hardware/navegador (Web Speech API):")
    add_bullet(doc, "Estudiante (10° y 11°)", "Accede desde computador o smartphone a la interfaz web para estudiar, interactuar por voz en diálogos contextuales y rendir exámenes oficiales mediante código VF-XXXX.")
    add_bullet(doc, "Docente de Inglés", "Gestiona sus grupos escolares, programa evaluaciones con ventanas temporales, supervisa notas en tiempo real y descarga planillas.")
    add_bullet(doc, "Owner (Administrador TheKeas)", "Supervisa la plataforma, valida credenciales docentes para autorizar su ingreso y audita operaciones críticas.")
    add_bullet(doc, "Web Speech Engine (Navegador)", "Subsistema nativo del cliente que ejecuta reconocimiento de voz (SpeechRecognition) y síntesis de audio (SpeechSynthesis) sin llamadas a servidores externos.")

    add_heading_with_spacing(doc, "2.2 Nivel 2: Contenedores de la Solución", level=3)
    h_cont = ["Contenedor", "Tecnología", "Responsabilidad Principal"]
    d_cont = [
        ["Cliente Web (SPA)", "HTML5 semántico, CSS3 (Electric Academy), JavaScript ES6+ / React", "Renderizado de interfaces táctiles, visualizador de ondas de audio, captura de voz y modo examen seguro."],
        ["Capa de API REST", "PostgreSQL Functions / Supabase REST Gateway / Django REST Framework", "Enrutamiento seguro, validación de esquemas JSON, control de sesiones JWT y cálculo de notas oficiales."],
        ["Base de Datos Relacional", "PostgreSQL 15+ con RLS activado", "Persistencia de 19 tablas relacionales, integridad referencial y auditoría inmutable de calificaciones."],
        ["Almacenamiento Estático", "GitHub Pages / Vercel CDN", "Distribución de archivos HTML, CSS, JS e imágenes optimizadas con caché HTTP."]
    ]
    add_table_data(doc, h_cont, d_cont, col_widths=[1.8, 2.4, 3.0])

    # 3. Mecanismos de Seguridad y Patrones Críticos
    add_heading_with_spacing(doc, "3. Mecanismos de Seguridad y Patrones Críticos", level=2)
    
    add_heading_with_spacing(doc, "3.1 Autenticación, JWT y Control de Roles (RBAC)", level=3)
    add_p(doc, "El acceso se gestiona mediante tokens de acceso (Access Tokens de 60 min de vigencia) y tokens de actualización (Refresh Tokens de 7 días). El payload del JWT incluye el ID de usuario, rol ('student', 'teacher', 'owner') y estado de la cuenta ('approved', 'pending').")
    add_p(doc, "Regla de Aprobación Docente: Un usuario registrado como 'teacher' no puede emitir peticiones a endpoints de creación de clases o exámenes hasta que su estado cambie a 'approved' en la tabla 'profiles' por autorización explícita del Owner.")

    add_heading_with_spacing(doc, "3.2 Blindaje de Exámenes Oficiales (RNF-04)", level=3)
    add_bullet(doc, "Ocultamiento de Respuestas", "Al iniciar un examen con código VF-XXXX, el endpoint '/api/examen/iniciar/' devuelve únicamente los enunciados y opciones de respuesta con su ID. El atributo booleano 'es_correcta' y las frases de 'respuesta_aceptada' NUNCA viajan en el payload.")
    add_bullet(doc, "Cálculo Atómico en Servidor", "Al invocar '/api/examen/intentos/{id}/finalizar/', una función transaccional en el backend realiza el cruce entre las respuestas enviadas y las respuestas correctas congeladas, asignando la nota final en escala 1.0 a 5.0.")
    add_bullet(doc, "Fórmula Institucional Oficial", "Puntaje = 1.00 + (aciertos / total_preguntas) * 4.00. La nota se redondea a 2 decimales para la base de datos y a 1 decimal para boletines escolares.")

    add_heading_with_spacing(doc, "3.3 Arquitectura de Voz y Protección de Datos Biométricos", level=3)
    add_p(doc, "Para blindar legalmente a la I.E. Promoción Social bajo la Ley 1581 de 2012 y el Código de Infancia y Adolescencia (Ley 1098 de 2006):")
    add_bullet(doc, "Ejecución 100% en el Edge", "La captura del micrófono es procesada en memoria RAM por la Web Speech API del navegador. Ningún paquete de audio PCM o MP3 se transmite al servidor ni a proveedores terceros.")
    add_bullet(doc, "Algoritmo Word-Diff Matcher", "El cliente recibe las palabras clave y fonemas esperados ('palabras_clave' de 'linea_dialogo') y realiza la comparación lingüística local, mostrando de inmediato en pantalla las palabras correctas (verde) y las alertas fonéticas (rojo con tips en español).")

    # 4. Catálogo Completo de Endpoints de la API REST
    add_heading_with_spacing(doc, "4. Catálogo Completo de Endpoints de la API REST", level=2)
    add_p(doc, "Convenciones generales: Todas las rutas inician con el prefijo /api/v1/. Toda petición autenticada requiere el encabezado 'Authorization: Bearer <access_token>'. Las respuestas exitosas devuelven HTTP 200 (OK) o 201 (Created). Las peticiones con errores devuelven un objeto uniforme con campos 'error', 'codigo' y 'detalle'.")

    # 4.1 Módulo Auth
    add_heading_with_spacing(doc, "4.1 Módulo de Autenticación y Cuentas (/api/v1/auth/)", level=3)
    h_ep = ["Método", "Endpoint", "Acceso", "Descripción y Payload"]
    d_ep_auth = [
        ["POST", "/auth/registro/estudiante/", "Público", "Registra un alumno. Body: { nickname, contrasena, nombre_completo, grado: 10|11, consentimiento_1581: true }."],
        ["POST", "/auth/registro/docente/", "Público", "Registra un profesor. Nace en estado 'pendiente'. Body: { nickname, contrasena, nombre_completo, correo }."],
        ["POST", "/auth/login/", "Público", "Inicia sesión. Body: { nickname, contrasena }. Retorna: { access_token, refresh_token, usuario: { id, nick, rol, estado, clase_id } }."],
        ["POST", "/auth/refresh/", "Público", "Renueva token. Body: { refresh_token }. Retorna nuevo access_token."],
        ["GET", "/auth/me/", "Autenticado", "Devuelve información del perfil en sesión, su clase asignada y roles activos."],
        ["GET", "/auth/docentes/pendientes/", "Solo Owner", "Lista solicitudes de profesores en estado 'pendiente' para revisión."],
        ["POST", "/auth/docentes/{id}/aprobar/", "Solo Owner", "Aprueba al docente ({id}), cambiando su estado a 'activo' y permitiendo creación de clases."],
        ["POST", "/auth/docentes/{id}/rechazar/", "Solo Owner", "Rechaza la solicitud del docente con motivo opcional: { motivo }."]
    ]
    add_table_data(doc, h_ep, d_ep_auth, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 4.2 Módulo Pedagógico y Diálogos
    add_heading_with_spacing(doc, "4.2 Módulo Pedagógico y Diálogos Contextuales (/api/v1/pedagogico/)", level=3)
    d_ep_ped = [
        ["GET", "/pedagogico/temas/", "Público", "Lista los 7 tiempos verbales con metadatos, slugs, colores y orden curricular."],
        ["GET", "/pedagogico/temas/{slug}/", "Público", "Detalle de un tiempo verbal: contexto de uso, fórmulas, línea de tiempo y audios de ejemplo."],
        ["GET", "/pedagogico/escenarios/", "Estudiante", "Lista escenarios de role-play disponibles por tema con sus tonos emocionales."],
        ["GET", "/pedagogico/escenarios/{id}/", "Estudiante", "Carga el diálogo completo por turnos, objetivos comunicativos, avatar y pautas de feedback fonético."],
        ["GET", "/pedagogico/practica/lote/", "Estudiante", "Genera un paquete aleatorio de 5 a 25 preguntas para práctica casual en memoria. Incluye respuestas correctas para calificación en cliente."]
    ]
    add_table_data(doc, h_ep, d_ep_ped, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 4.3 Módulo Banco de Preguntas
    add_heading_with_spacing(doc, "4.3 Módulo del Banco de Preguntas (/api/v1/banco/)", level=3)
    d_ep_banco = [
        ["GET", "/banco/preguntas/", "Docente, Owner", "Filtra preguntas por tema, dificultad y tipo. El docente ve preguntas globales y sus preguntas privadas."],
        ["POST", "/banco/preguntas/", "Docente, Owner", "Crea una nueva pregunta (opción múltiple o completar) con sus opciones o respuestas aceptadas."],
        ["GET", "/banco/preguntas/{id}/", "Docente, Owner", "Detalle exhaustivo de la pregunta y su justificación pedagógica."],
        ["PUT", "/banco/preguntas/{id}/", "Docente, Owner", "Edición segura: si la pregunta ya fue usada en exámenes, se archiva y se crea versión n+1 para no alterar notas históricas."],
        ["DELETE", "/banco/preguntas/{id}/", "Docente, Owner", "Borrado lógico (estado = 'archivada')."]
    ]
    add_table_data(doc, h_ep, d_ep_banco, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 4.4 Módulo Clases
    add_heading_with_spacing(doc, "4.4 Módulo de Clases (/api/v1/clases/)", level=3)
    d_ep_clase = [
        ["GET", "/clases/", "Docente, Owner", "Lista las clases del docente en sesión (o todas si es Owner)."],
        ["POST", "/clases/", "Docente", "Crea un nuevo grupo escolar. Body: { nombre: 'Inglés 10-A', grado: 10, anio_lectivo: 2026 }."],
        ["GET", "/clases/{id}/", "Docente, Owner", "Detalle de la clase y listado de estudiantes matriculados."],
        ["POST", "/clases/{id}/estudiantes/", "Docente", "Asigna manualmente a un estudiante al curso (o se autoasigna al presentar examen con código)."],
        ["DELETE", "/clases/{id}/estudiantes/{est_id}/", "Docente", "Desvincula a un estudiante del grupo (libera su clase_id)."]
    ]
    add_table_data(doc, h_ep, d_ep_clase, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 4.5 Módulo Evaluaciones (Docente)
    add_heading_with_spacing(doc, "4.5 Módulo de Evaluaciones (/api/v1/evaluaciones/)", level=3)
    d_ep_eval = [
        ["GET", "/evaluaciones/", "Docente, Owner", "Lista las evaluaciones creadas por el profesor con estado (borrador, publicada, cerrada)."],
        ["POST", "/evaluaciones/", "Docente", "Configura una prueba: temas, cantidad de preguntas, duración, fechas y genera código único VF-XXXX."],
        ["GET", "/evaluaciones/{id}/", "Docente, Owner", "Configuración completa y preguntas asignadas a la evaluación."],
        ["POST", "/evaluaciones/{id}/publicar/", "Docente", "Congela las preguntas en 'evaluacion_pregunta' y activa la prueba para los estudiantes."],
        ["POST", "/evaluaciones/{id}/cerrar/", "Docente", "Cierra anticipadamente la recepción de intentos de examen."]
    ]
    add_table_data(doc, h_ep, d_ep_eval, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 4.6 Módulo Presentación de Examen (Estudiante)
    add_heading_with_spacing(doc, "4.6 Módulo de Presentación de Exámenes (/api/v1/examen/)", level=3)
    d_ep_exam = [
        ["POST", "/examen/validar-codigo/", "Estudiante", "Verifica código VF-XXXX. Valida vigencia temporal, clase del estudiante y tope de intentos."],
        ["POST", "/examen/iniciar/", "Estudiante", "Crea la fila en 'intento', fija 'limite_en' y retorna preguntas SIN respuestas correctas."],
        ["POST", "/examen/intentos/{id}/responder/", "Estudiante", "Autoguardado progresivo: guarda opción elegida o texto escrito ante cada interacción."],
        ["POST", "/examen/intentos/{id}/evento/", "Estudiante", "Registra telemetría de foco (salida de pestaña, salida de pantalla completa)."],
        ["POST", "/examen/intentos/{id}/finalizar/", "Estudiante", "Cierra el intento, calcula la nota oficial en escala 1.0 a 5.0 y retorna feedback si es inmediato."]
    ]
    add_table_data(doc, h_ep, d_ep_exam, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 4.7 Módulo Reportes y Calificaciones
    add_heading_with_spacing(doc, "4.7 Módulo de Calificaciones y Reportes (/api/v1/reportes/)", level=3)
    d_ep_rep = [
        ["GET", "/reportes/evaluaciones/{id}/planilla/", "Docente, Owner", "Sábana de notas en vivo: lista de alumnos, aciertos, puntaje (1.0 - 5.0) e incidencias de foco."],
        ["GET", "/reportes/evaluaciones/{id}/exportar-csv/", "Docente", "Descarga archivo CSV/Excel listo para ingresar al sistema de notas del colegio."],
        ["GET", "/reportes/estudiante/mis-notas/", "Estudiante", "Historial de pruebas oficiales rendidas por el estudiante autenticado."],
        ["GET", "/reportes/auditoria/", "Solo Owner", "Bitácora completa de operaciones administrativas y cambios de estado."]
    ]
    add_table_data(doc, h_ep, d_ep_rep, col_widths=[1.0, 2.3, 1.4, 2.5])

    # 5. Formato Estándar de Cargas y Errores HTTP
    add_heading_with_spacing(doc, "5. Formato Estándar de Cargas y Errores HTTP", level=2)
    add_p(doc, "Para garantizar predictibilidad y facilidad de depuración en el cliente, la API maneja respuestas estandarizadas:")
    add_bullet(doc, "Estructura de Éxito (HTTP 200 / 201)", "{ 'ok': true, 'datos': { ... }, 'mensaje': 'Operación exitosa' }")
    add_bullet(doc, "Estructura de Error (HTTP 4xx / 5xx)", "{ 'ok': false, 'error': 'CODIGO_ERROR', 'detalle': 'Explicación clara en español', 'campo': 'nickname' }")
    
    h_err = ["Código HTTP", "Código de Error", "Motivo de Disparo"]
    d_err = [
        ["400 Bad Request", "VALIDACION_FALLIDA", "Campos obligatorios ausentes o formato inválido (ej. nickname con caracteres raros)."],
        ["401 Unauthorized", "TOKEN_EXPIRADO / NO_AUTORIZADO", "Falta encabezado Authorization o el token JWT ha caducado."],
        ["403 Forbidden", "PERMISO_DENEGADO / DOCENTE_PENDIENTE", "Rol insuficiente o profesor aún no aprobado por el Owner."],
        ["404 Not Found", "RECURSO_NO_ENCONTRADO", "Código VF-XXXX inexistente o ID de clase no localizado."],
        ["409 Conflict", "INTENTOS_AGOTADOS / NICK_DUPLICADO", "El estudiante ya presentó el número máximo de intentos permitidos en el examen."],
        ["423 Locked", "EXAMEN_FUERA_DE_HORARIO", "Se intenta acceder antes de la hora de apertura o después del cierre."]
    ]
    add_table_data(doc, h_err, d_err, col_widths=[1.5, 2.3, 3.4])

    # 6. Diagramas de Flujo y Secuencia Críticos
    add_heading_with_spacing(doc, "6. Diagramas de Flujo y Secuencia Críticos", level=2)
    
    add_heading_with_spacing(doc, "6.1 Flujo de Presentación Segura de Examen Oficial", level=3)
    add_p(doc, "1. El Estudiante digita 'VF-4821' en pantalla -> POST /api/v1/examen/validar-codigo/.")
    add_p(doc, "2. El Backend valida: ¿Existe? ¿Está dentro de la ventana de tiempo? ¿El alumno es de esa clase? ¿Le quedan intentos?")
    add_p(doc, "3. El Estudiante pulsa 'Comenzar' -> POST /api/v1/examen/iniciar/ -> Se crea 'intento' (estado='en_curso') y se calcula 'limite_en'.")
    add_p(doc, "4. El Backend responde con preguntas barajadas SIN la columna 'es_correcta' ni 'respuesta_aceptada'. El frontend activa 'Modo Examen' (pantalla completa y bloqueo de chuleta).")
    add_p(doc, "5. Cada vez que el alumno marca o escribe -> POST /api/v1/examen/intentos/{id}/responder/ (autoguardado progresivo).")
    add_p(doc, "6. Al terminar o agotarse el tiempo -> POST /api/v1/examen/intentos/{id}/finalizar/ -> El servidor cruza respuestas, aplica fórmula 1.0 a 5.0, fija estado='finalizado' y bloquea reingresos.")

    add_heading_with_spacing(doc, "6.2 Flujo de Práctica Oral con Micrófono y Feedback Fonético", level=3)
    add_p(doc, "1. El Estudiante selecciona un escenario de role-play (ej. 'El Alibi del Detective') -> GET /api/v1/pedagogico/escenarios/1/.")
    add_p(doc, "2. El Interlocutor virtual emite su turno con síntesis de voz (Web Speech API a 1.0x o 0.75x).")
    add_p(doc, "3. Es el turno del alumno: la web muestra objetivo en español y estructura clave. El alumno presiona [Espacio] o el botón de Micrófono 3D.")
    add_p(doc, "4. Web Speech API transcribe la voz en memoria RAM del dispositivo en tiempo real.")
    add_p(doc, "5. El algoritmo Word-Diff compara contra 'palabras_clave': pinta verde aciertos y rojo desvíos.")
    add_p(doc, "6. Si hay error, el Coach Fonético despliega el consejo correctivo en español y permite comparar con el audio nativo antes de reintentar.")

    # 7. Consideraciones de Despliegue e Infraestructura
    add_heading_with_spacing(doc, "7. Consideraciones de Despliegue e Infraestructura", level=2)
    add_bullet(doc, "Frontend", "Hospedado en GitHub Pages, Vercel o Netlify como sitio web estático optimizado.")
    add_bullet(doc, "Backend y Base de Datos", "PostgreSQL gestionado en Supabase (plan gratuito/comunitario) con funciones RPC y políticas RLS activas, o contenedor Docker en Render/Railway.")
    add_bullet(doc, "Variables de Entorno Críticas", "DATABASE_URL, JWT_SECRET, INSTITUCION_NOMBRE='I.E. Promoción Social', INSTITUCION_CIUDAD='Palermo, Huila', OWNER_NICK='thekeas'.")
    add_bullet(doc, "Respaldos Automáticos", "Dump diario automático de la base de datos PostgreSQL cifrado en almacenamiento secundario.")

    doc.save(filepath)
    print(f"Documento SAD guardado exitosamente en {filepath}")

generate_sad_doc("Docs/VerbFlow — Documento de Arquitectura y Especificación de la API (SAD).docx")
