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

def generate_srs(filepath):
    doc = docx.Document()
    
    # Title
    p_title = doc.add_heading("VerbFlow — Especificación de Requisitos de Software (SRS)", level=1)
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(12)
    r_meta = p_meta.add_run("Versión 0.3 (Enfoque Comunicativo y Audiovisual) · Autor: Kevin · Basado en IEEE 29148 · Octubre de 2026")
    r_meta.font.italic = True
    r_meta.font.size = Pt(9.5)
    r_meta.font.color.rgb = RGBColor(100, 116, 139)
    
    # 1. Introducción
    add_heading_with_spacing(doc, "1. Introducción", level=2)
    
    add_heading_with_spacing(doc, "1.1 Propósito", level=3)
    add_p(doc, "Este documento define los requisitos funcionales y no funcionales de VerbFlow, una plataforma web privada diseñada para la Institución Educativa Promoción Social (Palermo, Huila, Colombia). Su objetivo es proporcionar un entorno especializado e interactivo para que los estudiantes de grados décimo (10°) y once (11°) dominen los tiempos verbales en inglés mediante un enfoque comunicativo, audiovisual y práctico centrado en la pronunciación en contexto, diálogos situacionales y evaluaciones de clase.")
    
    add_heading_with_spacing(doc, "1.2 Alcance y Paradigma Pedagógico", level=3)
    add_p(doc, "VerbFlow evoluciona el aprendizaje tradicional hacia el Enfoque Comunicativo (Communicative Language Teaching - CLT), permitiendo:")
    add_bullet(doc, "Módulo de Diálogos Contextuales y Emocionales (Role-Play)", "Conversaciones situacionales interactivas donde los tiempos verbales se aplican según la emoción y el contexto real (sorpresa, urgencia, formalidad, alegría, relatos de anécdotas).")
    add_bullet(doc, "Evaluación de Pronunciación en Tiempo Real (Speaking Feedback)", "El estudiante habla al micrófono; la web escucha mediante Web Speech API, compara palabra por palabra y brinda retroalimentación fonética inmediata en español indicando errores exactos (omisión de -ed, confusión de sonidos /b/ vs /v/, /s/ inicial, vocales cortas/largas).")
    add_bullet(doc, "Entorno Audiovisual Inmersivo", "Visualizador interactivo de ondas de voz (waveform), avatares con estados de ánimo contextuales, líneas de tiempo visuales y reproducción de audio a doble velocidad (1.0x natural y 0.75x de entrenamiento).")
    add_bullet(doc, "Estudio sintáctico y Chuleta de Fórmulas", "Teoría clara y accesible de fórmulas afirmativas, negativas e interrogativas para consulta rápida.")
    add_bullet(doc, "Práctica libre casual y privada", "Ejercicios interactivos evaluados en el navegador del estudiante sin persistencia en servidor para evitar saturación y preservar privacidad.")
    add_bullet(doc, "Evaluaciones Oficiales de Clase (Modo Examen)", "Pruebas asignadas por docentes con código VF-XXXX, autoguardado de respuestas en servidor, telemetría de foco y calificación formal en escala colombiana de 1.0 a 5.0.")
    add_bullet(doc, "Suite Docente y Panel Owner", "Gestión de cursos, banco de preguntas con opción múltiple y respuesta escrita, control de acceso de profesores y exportación de calificaciones a Excel/CSV.")
    add_p(doc, "Exclusiones explícitas (Fuera de alcance v1.0): Gamificación basada en XP acumulado, rachas diarias obligatorias y rankings públicos (suprimida para evitar ansiedad académica y cumplir con minimización de datos personales de menores de edad); pagos; y almacenamiento de grabaciones de voz en servidores externos.")

    add_heading_with_spacing(doc, "1.3 Definiciones y siglas", level=3)
    add_bullet(doc, "Owner", "Administrador general y custodio institucional de la plataforma.")
    add_bullet(doc, "Docente (Teacher)", "Profesor de inglés facultado para gestionar clases y publicar evaluaciones.")
    add_bullet(doc, "Estudiante", "Alumno de 10° u 11° matriculado en un grupo académico específico.")
    add_bullet(doc, "Web Speech API", "Estándar nativo del navegador para reconocimiento (SpeechRecognition) y síntesis de voz (SpeechSynthesis) en el cliente.")
    add_bullet(doc, "Word-Diff Matcher", "Algoritmo de alineación fonética que compara el texto transcrito contra el texto objetivo para resaltar aciertos en verde y desvíos en rojo.")
    add_bullet(doc, "Modo Examen", "Estado seguro en el cliente que inhabilita teoría, chuleta y pistas durante una prueba oficial de clase.")
    add_bullet(doc, "Código VF-XXXX", "Identificador alfanumérico único para ingresar a una evaluación de clase programada.")
    add_bullet(doc, "Escala MEN", "Calificación académica oficial de Colombia de 1.00 a 5.00 (mínima aprobatoria: 3.00).")

    add_heading_with_spacing(doc, "1.4 Referencias y marco normativo", level=3)
    add_bullet(doc, "Prototipo de Interfaz (Mockup)", "Proyecto interactivo 'Ingles Web' (https://kevin81a.github.io/VerbFlow/).")
    add_bullet(doc, "Ley 1581 de 2012 y Decreto 1377 de 2013", "Protección de datos personales en Colombia.")
    add_bullet(doc, "Ley 1098 de 2006", "Código de la Infancia y la Adolescencia (protección y prevalencia de derechos de menores).")
    add_bullet(doc, "W3C Web Speech API Specification", "Estándar de interfaz de voz en navegadores.")
    add_bullet(doc, "Estándar IEEE 29148:2018", "Ingeniería de Requisitos de Software.")

    # 2. Descripción general
    add_heading_with_spacing(doc, "2. Descripción general", level=2)
    add_heading_with_spacing(doc, "2.1 Perspectiva del producto", level=3)
    add_p(doc, "VerbFlow es una plataforma web SPA (Single Page Application) moderna, reactiva y liviana que combina interfaces táctiles enriquecidas con reconocimiento de voz en el navegador y un backend transaccional en PostgreSQL con seguridad por roles.")
    
    add_heading_with_spacing(doc, "2.2 Roles de usuario y permisos", level=3)
    h_roles = ["Rol", "Descripción institucional", "Permisos principales"]
    d_roles = [
        ["Owner", "Administrador institucional (TheKeas).", "Aprobar o rechazar docentes, gestionar banco global, auditar accesos y supervisar métricas globales."],
        ["Docente", "Profesor de inglés de la I.E. Promoción Social.", "Crear clases, generar códigos VF-XXXX, diseñar evaluaciones, aportar preguntas propias (privadas), monitorear exámenes en vivo y exportar notas."],
        ["Estudiante", "Alumno regular de grado 10° u 11°.", "Estudiar teoría, practicar diálogos contextuales con micrófono, practicar gramática libremente, presentar exámenes de su clase y consultar notas."]
    ]
    add_table_data(doc, h_roles, d_roles, col_widths=[1.5, 2.3, 3.2])

    add_heading_with_spacing(doc, "2.3 Entorno operativo y compatibilidad de voz", level=3)
    add_p(doc, "Optimizado para Chrome, Microsoft Edge, Safari y navegadores basados en Chromium en computadores y teléfonos móviles (Android/iOS). La funcionalidad de micrófono requiere permisos de audio del navegador y opera con auriculares o micrófonos integrados de gama estándar escolar.")

    add_heading_with_spacing(doc, "2.4 Restricciones del proyecto", level=3)
    add_bullet(doc, "Privacidad absoluta de voz de menores (Ley 1581)", "El procesamiento de voz ocurre exclusivamente en la memoria del navegador. No se envía, graba ni almacena ningún archivo de audio del estudiante en servidores o bases de datos.")
    add_bullet(doc, "Cero costo de infraestructura de voz", "Al emplear Web Speech API nativa, la institución no incurre en costos por consumo de tokens ni APIs de terceros.")
    add_bullet(doc, "Tolerancia a fallos de conectividad", "Las evaluaciones oficiales guardan progresivamente cada respuesta en el servidor; los módulos de práctica y diálogo operan fluidamente aun con conectividad intermitente.")

    # 3. Requisitos funcionales
    add_heading_with_spacing(doc, "3. Requisitos funcionales", level=2)
    h_rf = ["ID", "Requisito funcional", "Prioridad"]

    # 3.1 Autenticación
    add_heading_with_spacing(doc, "3.1 Autenticación y gestión de usuarios", level=3)
    d_rf_auth = [
        ["RF-01", "El sistema permitirá el registro de usuarios con nickname único (case-insensitive), contraseña encriptada, nombre completo, grado (10° u 11°) y tipo de cuenta (Estudiante o Docente). La institución queda predeterminada como I.E. Promoción Social.", "M"],
        ["RF-02", "El sistema exigirá aceptar la declaración de política de tratamiento de datos personales y autorización de acudiente antes de completar el registro (Ley 1581 / Ley 1098).", "M"],
        ["RF-03", "El sistema autenticará sesiones de forma segura mediante tokens JWT.", "M"],
        ["RF-04", "Las cuentas registradas como Docente nacerán en estado 'pendiente' hasta que el Owner las apruebe.", "M"],
        ["RF-05", "El sistema validará que cada estudiante solo pertenezca a exactamente una clase a la vez (asociación directa clase_id).", "M"],
        ["RF-06", "El usuario podrá solicitar la anonimización de su cuenta bajo el derecho de hábeas data, preservando únicamente registros académicos exigidos por ley.", "M"]
    ]
    add_table_data(doc, h_rf, d_rf_auth, col_widths=[1.2, 5.0, 1.0])

    # 3.2 Diálogos Contextuales y Práctica Oral (NUEVO ENFOQUE)
    add_heading_with_spacing(doc, "3.2 Módulo de Diálogos Contextuales y Práctica Oral (Role-Play)", level=3)
    d_rf_dialogo = [
        ["RF-15", "El sistema ofrecerá escenarios de diálogo basados en situaciones de la vida real (anécdotas, disculpas, planes, debates) asociados a cada tiempo verbal.", "M"],
        ["RF-16", "Cada escenario de diálogo definirá un contexto y una emoción explícita (alegría, sorpresa, urgencia, formalidad, duda) reflejada en el diseño y en la voz del interlocutor.", "M"],
        ["RF-17", "El sistema presentará la conversación por turnos audiovisuales: el interlocutor virtual habla primero con entonación contextual y luego se le da el turno de habla al estudiante con un objetivo comunicativo claro en español.", "M"],
        ["RF-18", "El estudiante podrá activar el micrófono para emitir su respuesta hablada; la interfaz mostrará un visualizador de ondas de voz en tiempo real.", "M"],
        ["RF-19", "El motor de reconocimiento cotejará la transcripción con la frase esperada mediante el algoritmo Word-Diff: resaltará en verde las palabras correctas y en rojo las palabras erróneas o no reconocidas.", "M"],
        ["RF-19a", "El sistema generará retroalimentación fonética correctiva en español explicando la causa del error (ej. confusión entre 'washing' y 'watching', omisión de la terminación '-ed', vocal corta vs larga).", "M"],
        ["RF-19b", "El estudiante podrá escuchar la pronunciación modelo nativa a velocidad estándar (1.0x) y a velocidad lenta (0.75x) para afinar su oído.", "M"]
    ]
    add_table_data(doc, h_rf, d_rf_dialogo, col_widths=[1.2, 5.0, 1.0])

    # 3.3 Modo Estudio Gramatical
    add_heading_with_spacing(doc, "3.3 Modo Estudio Gramatical", level=3)
    d_rf_estudio = [
        ["RF-10", "El sistema exhibirá el contenido curricular de cada tiempo verbal: contexto de uso, estructuras sintácticas (afirmativa, negativa, interrogativa), reglas ortográficas y ejemplos.", "M"],
        ["RF-11", "El sistema mostrará una línea de tiempo visual e interactiva que posicione la acción verbal (pasado, presente, futuro) respecto al momento del habla.", "D"],
        ["RF-12", "El sistema reproducirá el audio de las oraciones de ejemplo mediante síntesis de voz Web Speech API.", "M"],
        ["RF-13", "El sistema ofrecerá una 'Chuleta de Fórmulas' accesible rápidamente desde cualquier vista de estudio y práctica.", "M"]
    ]
    add_table_data(doc, h_rf, d_rf_estudio, col_widths=[1.2, 5.0, 1.0])

    # 3.4 Modo Práctica Libre
    add_heading_with_spacing(doc, "3.4 Modo Práctica Libre (casual y en navegador)", level=3)
    d_rf_practica = [
        ["RF-20", "El estudiante podrá seleccionar temas verbales, nivel de dificultad (Fácil, Media, Difícil, Ponderada) y cantidad de preguntas (5 a 25).", "M"],
        ["RF-21", "El banco generará preguntas combinando opciones múltiples y respuestas escritas / completar.", "M"],
        ["RF-22", "Al finalizar la práctica libre, el sistema mostrará aciertos, errores y explicaciones pedagógicas de inmediato en el navegador.", "M"],
        ["RF-23", "La práctica libre se ejecutará enteramente en memoria local del cliente sin persistir intentos en la base de datos para no saturar el sistema.", "M"]
    ]
    add_table_data(doc, h_rf, d_rf_practica, col_widths=[1.2, 5.0, 1.0])

    # 3.5 Pruebas Oficiales de Clase
    add_heading_with_spacing(doc, "3.5 Pruebas Oficiales de Clase (Modo Examen)", level=3)
    d_rf_examen = [
        ["RF-30", "El estudiante accederá a la evaluación oficial ingresando el código único VF-XXXX asignado por su docente.", "M"],
        ["RF-31", "El sistema validará que la prueba esté activa, en horario vigente y que el estudiante pertenezca a la clase destinataria.", "M"],
        ["RF-32", "Al iniciar, se activará el Modo Examen: bloqueo de teoría, línea de tiempo y chuleta de fórmulas.", "M"],
        ["RF-33", "El sistema controlará el número de intentos permitidos (por defecto 1) y aplicará temporizador con cierre automático.", "M"],
        ["RF-34", "Guardado progresivo de respuestas en el servidor para evitar pérdida de datos ante desconexiones.", "M"],
        ["RF-35", "Registro de señales de supervisión académica (pérdida de foco, cambio de pestaña, salida de pantalla completa).", "D"],
        ["RF-36", "Calificación automática en escala colombiana de 1.00 a 5.00 calculada en el backend.", "M"]
    ]
    add_table_data(doc, h_rf, d_rf_examen, col_widths=[1.2, 5.0, 1.0])

    # 3.6 Suite Docente y Owner
    add_heading_with_spacing(doc, "3.6 Suite del Docente y Panel del Owner", level=3)
    d_rf_admin = [
        ["RF-40", "El docente podrá crear grupos (10° y 11°), publicar evaluaciones automáticas o manuales y generar códigos VF-XXXX.", "M"],
        ["RF-41", "El docente monitoreará entregas y notas en tiempo real y podrá exportar planillas a Excel / CSV.", "M"],
        ["RF-42", "El docente podrá agregar preguntas y diálogos propios al banco (nacen con visibilidad privada para sus clases).", "M"],
        ["RF-50", "El Owner visualizará las solicitudes de registro docente pendientes y podrá aprobarlas o rechazarlas con un clic.", "M"],
        ["RF-51", "El Owner supervisará el banco global, publicará contenido y fiscalizará la auditoría del sistema.", "M"]
    ]
    add_table_data(doc, h_rf, d_rf_admin, col_widths=[1.2, 5.0, 1.0])

    # 3.7 Resolución de Gamificación
    add_heading_with_spacing(doc, "3.7 Resolución sobre Gamificación", level=3)
    add_p(doc, "Se reitera la exclusión formal de mecánicas de puntos XP, niveles acumulados, rachas de días consecutivos y rankings públicos para priorizar el enfoque pedagógico comunicativo y garantizar la protección de datos personales de menores de edad.")

    # 4. Requisitos no funcionales
    add_heading_with_spacing(doc, "4. Requisitos no funcionales", level=2)
    h_rnf = ["ID", "Categoría", "Especificación técnica"]
    d_rnf = [
        ["RNF-01", "Seguridad", "Contraseñas cifradas con algoritmos robustos (Argon2 / PBKDF2) y sesiones autenticadas por JWT."],
        ["RNF-02", "Seguridad", "Comunicación cliente-servidor cifrada obligatoriamente sobre HTTPS."],
        ["RNF-03", "Seguridad", "Las claves de respuestas de exámenes oficiales se validan exclusivamente en el servidor (RNF-04)."],
        ["RNF-04", "Privacidad de Voz", "Cero almacenamiento de grabaciones de voz en servidores (procesamiento 100% en memoria del navegador según Ley 1581)."],
        ["RNF-05", "Rendimiento", "Carga inicial inferior a 3 segundos en 4G. Respuesta de reconocimiento de voz menor a 800 ms."],
        ["RNF-06", "Concurrencia", "Soporte para al menos 300 estudiantes simultáneos rindiendo pruebas o practicando sin degradación."],
        ["RNF-07", "Usabilidad y A11y", "Interfaz 'Electric Academy' de alto contraste, compatible con lectores de pantalla y operable desde 360 px de ancho."],
        ["RNF-08", "Integridad", "Transacciones atómicas en base de datos para calificaciones y tolerancia a fallos de red en el colegio."]
    ]
    add_table_data(doc, h_rnf, d_rnf, col_widths=[1.2, 1.6, 4.4])

    # 5. Reglas de negocio
    add_heading_with_spacing(doc, "5. Reglas de negocio", level=2)
    add_bullet(doc, "RN-01", "Un docente requiere aprobación del Owner antes de poder crear cursos o publicar evaluaciones.")
    add_bullet(doc, "RN-02", "Un estudiante solo puede pertenecer a exactamente una clase a la vez.")
    add_bullet(doc, "RN-03", "El código VF-XXXX solo habilita la prueba dentro de la ventana de tiempo fijada por el docente.")
    add_bullet(doc, "RN-04", "Durante el Modo Examen queda inhabilitado el acceso a teoría, fórmulas y diálogos.")
    add_bullet(doc, "RN-05", "Tratamiento de datos de menores condicionado a consentimiento informado del acudiente (Ley 1581 / Ley 1098).")
    add_bullet(doc, "RN-06", "Aislamiento por docente: Un docente solo accede a las calificaciones y alumnos de sus propias clases.")
    add_bullet(doc, "RN-07", "Escala oficial de notas: Puntaje = 1.0 + (aciertos / total_preguntas) * 4.0. Escala de 1.00 a 5.00 con aprobación en 3.00.")

    # 6. Casos de uso principales
    add_heading_with_spacing(doc, "6. Casos de uso principales", level=2)
    h_cu = ["ID", "Caso de uso", "Actor", "Resultado esperado"]
    d_cu = [
        ["CU-01", "Registrarse e iniciar sesión", "Estudiante, Docente", "Cuenta registrada (Docente queda en 'pendiente')."],
        ["CU-02", "Aprobar docente", "Owner", "Cuenta docente activada con permisos plenos."],
        ["CU-03", "Practicar diálogo contextual con micrófono", "Estudiante", "Interacción hablada por turnos con evaluación fonética y retroalimentación en español."],
        ["CU-04", "Practicar gramática libremente", "Estudiante", "Ejercicios mixtos resueltos y calificados inmediatamente en el cliente."],
        ["CU-05", "Crear clase y publicar examen", "Docente", "Generación de código VF-XXXX y congelación de preguntas."],
        ["CU-06", "Presentar examen oficial", "Estudiante", "Respuestas guardadas progresivamente y nota en escala 1.0 a 5.0 registrada."],
        ["CU-07", "Monitorear en vivo y exportar notas", "Docente", "Visualización en tiempo real y descarga de planilla en Excel/CSV."]
    ]
    add_table_data(doc, h_cu, d_cu, col_widths=[1.1, 2.3, 1.6, 2.2])

    doc.save(filepath)
    print(f"SRS v0.3 guardado exitosamente en {filepath}")

generate_srs("Docs/VerbFlow — Especificación de Requisitos de Software (SRS).docx")
