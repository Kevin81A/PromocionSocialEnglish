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

def generate_modelo_datos(filepath):
    doc = docx.Document()
    
    # Title
    p_title = doc.add_heading("VerbFlow — Modelo de Datos", level=1)
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(12)
    r_meta = p_meta.add_run("Versión 0.3 (Enfoque Comunicativo y Audiovisual) · Autor: Kevin · Complementa el SRS v0.3 · Motor: PostgreSQL · Octubre de 2026")
    r_meta.font.italic = True
    r_meta.font.size = Pt(9.5)
    r_meta.font.color.rgb = RGBColor(100, 116, 139)
    
    # 1. Alcance y decisiones confirmadas
    add_heading_with_spacing(doc, "1. Alcance y decisiones del modelo v0.3", level=2)
    add_p(doc, "Este documento define la arquitectura relacional de VerbFlow para PostgreSQL, adaptada al enfoque comunicativo y audiovisual para la Institución Educativa Promoción Social (Palermo, Huila).")
    
    add_bullet(doc, "1. Módulo de Diálogos Contextuales (Role-Play)", "Se incorporan las entidades 'escenario_dialogo' y 'linea_dialogo' para soportar conversaciones situacionales por turnos, con emociones, avatares, transcripciones esperadas y metadatos de retroalimentación fonética.")
    add_bullet(doc, "2. Reconocimiento de voz en el cliente", "El reconocimiento y evaluación de pronunciación se procesan en memoria del navegador mediante Web Speech API. La base de datos almacena el patrón fonético esperado y los tips correctivos en español, pero NO almacena archivos de audio del estudiante, protegiendo a los menores (Ley 1581 / Ley 1098).")
    add_bullet(doc, "3. Preguntas de opción múltiple y respuesta escrita", "Banco curricular en 'pregunta', 'opcion_pregunta' y 'respuesta_aceptada' completamente activo para la v1.0.")
    add_bullet(doc, "4. Institución única y privada", "Configuración global fija para la I.E. Promoción Social sin tabla 'institucion'.")
    add_bullet(doc, "5. Aislamiento docente con supervisión Owner", "Varios profesores independientes; toda cuenta docente requiere visto bueno obligatorio del Owner.")
    add_bullet(doc, "6. Relación 1 Estudiante = 1 Clase", "Asociación directa 'clase_id' en la entidad 'usuario' sin tabla intermedia.")
    add_bullet(doc, "7. Exclusión total de gamificación", "Sin tablas de XP, niveles, rachas ni rankings. Práctica casual inmediata.")
    add_bullet(doc, "8. Calificación oficial 1.0 a 5.0", "Escala estándar del Ministerio de Educación Nacional de Colombia para exámenes oficiales.")

    # 2. Resumen de entidades (19 tablas)
    add_heading_with_spacing(doc, "2. Resumen de entidades (19 tablas)", level=2)
    headers_ent = ["Dominio funcional", "Tablas", "Propósito"]
    data_ent = [
        ["Usuarios y cumplimiento legal", "usuario, consentimiento, solicitud_docente, auditoria", "Accesos, roles (owner, docente, estudiante) y cumplimiento de la Ley 1581."],
        ["Módulo Comunicativo y Oral (Nuevo)", "escenario_dialogo, linea_dialogo", "Situaciones de la vida real con contexto emocional, turnos y pautas fonéticas."],
        ["Contenido pedagógico y banco", "tema, formula, audio_ejemplo, pregunta, opcion_pregunta, respuesta_aceptada", "Temarios de tiempos verbales, chuleta de fórmulas y preguntas (opción múltiple y completar)."],
        ["Gestión de clases y exámenes", "clase, evaluacion, evaluacion_tema, evaluacion_pregunta", "Administración de cursos 10°/11°, congelación de pruebas con código VF-XXXX."],
        ["Intentos y auditoría de examen", "intento, intento_respuesta, evento_examen", "Registro exclusivo de exámenes oficiales de clase y métricas de integridad."]
    ]
    add_table_data(doc, headers_ent, data_ent, col_widths=[2.0, 2.5, 2.5])

    # 3. Diccionario de datos
    add_heading_with_spacing(doc, "3. Diccionario de datos", level=2)
    h_cols = ["Columna", "Tipo", "Restricciones", "Descripción"]

    # 3.1 usuario
    add_heading_with_spacing(doc, "3.1 usuario", level=3)
    d_usuario = [
        ["id", "UUID", "PK", "Identificador único."],
        ["nickname", "varchar(30)", "NOT NULL, UNIQUE (lower)", "Nombre de usuario único case-insensitive."],
        ["contrasena_hash", "varchar(255)", "NOT NULL", "Hash Argon2 / PBKDF2."],
        ["nombre_completo", "varchar(150)", "NOT NULL", "Nombre según documento escolar."],
        ["rol", "enum", "NOT NULL: owner, docente, estudiante", "Rol en la plataforma."],
        ["estado", "enum", "NOT NULL: pendiente, activo, suspendido, rechazado, eliminado", "Docentes nacen en 'pendiente'."],
        ["clase_id", "UUID", "FK → clase, NULL", "Clase asignada al estudiante."],
        ["grado", "smallint", "CHECK (grado IN (10, 11)), NULL", "Grado escolar cursado."],
        ["correo", "varchar(254)", "NULL, UNIQUE si existe", "Correo opcional."],
        ["ultimo_acceso_en", "timestamptz", "NULL", "Último inicio de sesión."],
        ["creado_en", "timestamptz", "NOT NULL, default now()", "Fecha de registro."],
        ["actualizado_en", "timestamptz", "NOT NULL, default now()", "Última modificación."]
    ]
    add_table_data(doc, h_cols, d_usuario, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.2 consentimiento
    add_heading_with_spacing(doc, "3.2 consentimiento (Ley 1581)", level=3)
    d_consent = [
        ["id", "UUID", "PK", "Identificador."],
        ["usuario_id", "UUID", "FK → usuario, NOT NULL", "Usuario que acepta."],
        ["tipo", "enum", "NOT NULL: mayor_edad, acudiente", "Condición declarada."],
        ["version_politica", "varchar(20)", "NOT NULL", "Versión de la política."],
        ["aceptado_en", "timestamptz", "NOT NULL, default now()", "Marca legal."],
        ["ip", "inet", "NULL", "Dirección IP para auditoría."]
    ]
    add_table_data(doc, h_cols, d_consent, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.3 solicitud_docente
    add_heading_with_spacing(doc, "3.3 solicitud_docente", level=3)
    d_solic = [
        ["id", "UUID", "PK", "Identificador."],
        ["usuario_id", "UUID", "FK → usuario, NOT NULL, UNIQUE", "Docente solicitante."],
        ["estado", "enum", "NOT NULL: pendiente, aprobada, rechazada", "Estado del trámite."],
        ["revisado_por", "UUID", "FK → usuario, NULL", "Owner que aprobó/rechazó."],
        ["revisado_en", "timestamptz", "NULL", "Fecha de resolución."],
        ["motivo_rechazo", "text", "NULL", "Comentarios en caso de rechazo."]
    ]
    add_table_data(doc, h_cols, d_solic, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.4 auditoria
    add_heading_with_spacing(doc, "3.4 auditoria", level=3)
    d_aud = [
        ["id", "bigint", "PK autoincremental", "Identificador."],
        ["usuario_id", "UUID", "FK → usuario, NULL", "Operador."],
        ["accion", "varchar(60)", "NOT NULL", "Acción sensible ejecutada."],
        ["entidad", "varchar(60)", "NOT NULL", "Tabla impactada."],
        ["entidad_id", "varchar(64)", "NOT NULL", "ID de registro."],
        ["ip", "inet", "NULL", "Dirección IP."],
        ["creado_en", "timestamptz", "NOT NULL, default now()", "Fecha y hora."]
    ]
    add_table_data(doc, h_cols, d_aud, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.5 escenario_dialogo (NUEVA TABLA COMUNICATIVA)
    add_heading_with_spacing(doc, "3.5 escenario_dialogo (Módulo Comunicativo y Oral)", level=3)
    add_p(doc, "Define situaciones contextuales de la vida cotidiana para practicar los tiempos verbales con carga emocional y conversacional.")
    d_escenario = [
        ["id", "smallint", "PK", "Identificador del escenario."],
        ["tema_id", "smallint", "FK → tema, NOT NULL", "Tiempo verbal principal en práctica."],
        ["titulo", "varchar(120)", "NOT NULL", "Ej.: 'El Alibi del Detective', 'El Café Inesperado'."],
        ["contexto_situacional", "text", "NOT NULL", "Descripción del contexto y objetivo en español."],
        ["emocion", "enum", "NOT NULL: neutral, alegre, sorprendido, urgente, formal, dudoso", "Tono emocional predominante."],
        ["avatar_personaje", "varchar(100)", "NOT NULL", "Identificador del personaje interlocutor."],
        ["orden", "smallint", "NOT NULL", "Secuencia sugerida."],
        ["activo", "boolean", "NOT NULL, default true", "Disponibilidad pedagógica."]
    ]
    add_table_data(doc, h_cols, d_escenario, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.6 linea_dialogo (NUEVA TABLA COMUNICATIVA)
    add_heading_with_spacing(doc, "3.6 linea_dialogo (Turnos y Guía Fonética)", level=3)
    add_p(doc, "Almacena los turnos de la conversación, el texto objetivo en inglés, las metas comunicativas y las pautas fonéticas de retroalimentación.")
    d_linea = [
        ["id", "integer", "PK autoincremental", "Identificador del turno."],
        ["escenario_id", "smallint", "FK → escenario_dialogo, NOT NULL", "Escenario asociado."],
        ["orden_turno", "smallint", "NOT NULL", "Número secuencial de turno en el diálogo."],
        ["hablante", "enum", "NOT NULL: personaje, estudiante", "Quién emite la frase."],
        ["texto_ingles", "varchar(250)", "NOT NULL", "Frase dicha por el personaje o frase meta que el alumno debe pronunciar."],
        ["traduccion_es", "varchar(250)", "NOT NULL", "Traducción contextual en español."],
        ["objetivo_comunicativo", "varchar(250)", "NULL", "Instrucción para el estudiante (ej.: 'Dile que estabas viendo tele')."],
        ["audio_url", "varchar(300)", "NULL", "Audio nativo pregrabado con emoción (o usa Web Speech TTS)."],
        ["palabras_clave", "jsonb", "NULL", "Lista de palabras fonéticamente críticas a evaluar."],
        ["tips_fonetica", "jsonb", "NULL", "Mensajes correctivos en español si falla palabras específicas."]
    ]
    add_table_data(doc, h_cols, d_linea, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.7 tema
    add_heading_with_spacing(doc, "3.7 tema", level=3)
    d_tema = [
        ["id", "smallint", "PK", "Identificador."],
        ["slug", "varchar(60)", "NOT NULL, UNIQUE", "Slug URL (ej.: past-continuous)."],
        ["nombre", "varchar(100)", "NOT NULL", "Nombre del tiempo verbal."],
        ["orden", "smallint", "NOT NULL", "Secuencia pedagógica."],
        ["uso", "text", "NOT NULL", "Explicación contextual en español."],
        ["reglas_ortografia", "text", "NULL", "Reglas (-ing, -ed, irregulares)."],
        ["linea_tiempo", "jsonb", "NULL", "Datos de coordenadas para renderizar la línea de tiempo."],
        ["activo", "boolean", "NOT NULL, default true", "Disponibilidad."]
    ]
    add_table_data(doc, h_cols, d_tema, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.8 formula
    add_heading_with_spacing(doc, "3.8 formula", level=3)
    d_formula = [
        ["id", "integer", "PK autoincremental", "Identificador."],
        ["tema_id", "smallint", "FK → tema, NOT NULL", "Tema verbal."],
        ["tipo", "enum", "NOT NULL: afirmativa, negativa, interrogativa", "Estructura oracional."],
        ["estructura", "varchar(200)", "NOT NULL", "Sintaxis de la Chuleta de Fórmulas."],
        ["ejemplo", "varchar(250)", "NOT NULL", "Oración de ejemplo."]
    ]
    add_table_data(doc, h_cols, d_formula, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.9 audio_ejemplo
    add_heading_with_spacing(doc, "3.9 audio_ejemplo", level=3)
    d_audio = [
        ["id", "integer", "PK autoincremental", "Identificador."],
        ["tema_id", "smallint", "FK → tema, NOT NULL", "Tema asociado."],
        ["texto", "varchar(250)", "NOT NULL", "Texto a reproducir."],
        ["archivo_url", "varchar(300)", "NULL", "Audio MP3 (o Web Speech API si NULL)."]
    ]
    add_table_data(doc, h_cols, d_audio, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.10 pregunta
    add_heading_with_spacing(doc, "3.10 pregunta", level=3)
    d_preg = [
        ["id", "UUID", "PK", "Identificador."],
        ["tema_id", "smallint", "FK → tema, NOT NULL", "Tiempo verbal."],
        ["tipo", "enum", "NOT NULL: opcion_multiple, completar", "Tipo de ejercicio."],
        ["enunciado", "text", "NOT NULL", "Texto de la pregunta."],
        ["dificultad", "enum", "NOT NULL: facil, media, dificil", "Complejidad."],
        ["explicacion", "text", "NOT NULL", "Justificación pedagógica mostrada al corregir."],
        ["autor_id", "UUID", "FK → usuario, NOT NULL", "Docente u Owner creador."],
        ["visibilidad", "enum", "NOT NULL: global, privada", "Privada para el profesor; global si es oficial."],
        ["estado", "enum", "NOT NULL: activa, archivada", "Borrado lógico."],
        ["version", "integer", "NOT NULL, default 1", "Versión incremental para blindar notas históricas."],
        ["pregunta_origen_id", "UUID", "FK → pregunta, NULL", "Versión previa en caso de edición."]
    ]
    add_table_data(doc, h_cols, d_preg, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.11 opcion_pregunta
    add_heading_with_spacing(doc, "3.11 opcion_pregunta", level=3)
    d_opc = [
        ["id", "UUID", "PK", "Identificador."],
        ["pregunta_id", "UUID", "FK → pregunta, NOT NULL", "Pregunta de opción múltiple."],
        ["texto", "varchar(250)", "NOT NULL", "Alternativa de respuesta."],
        ["es_correcta", "boolean", "NOT NULL", "Secreta en exámenes oficiales."],
        ["orden", "smallint", "NOT NULL", "Orden canónico."]
    ]
    add_table_data(doc, h_cols, d_opc, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.12 respuesta_aceptada
    add_heading_with_spacing(doc, "3.12 respuesta_aceptada", level=3)
    d_acep = [
        ["id", "UUID", "PK", "Identificador."],
        ["pregunta_id", "UUID", "FK → pregunta, NOT NULL", "Pregunta de completar."],
        ["texto", "varchar(200)", "NOT NULL", "Respuesta escrita válida normalizada."]
    ]
    add_table_data(doc, h_cols, d_acep, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.13 clase
    add_heading_with_spacing(doc, "3.13 clase", level=3)
    d_clase = [
        ["id", "UUID", "PK", "Identificador."],
        ["docente_id", "UUID", "FK → usuario, NOT NULL", "Profesor a cargo."],
        ["nombre", "varchar(100)", "NOT NULL", "Ej.: Inglés 10-A, Inglés 11-B."],
        ["grado", "smallint", "NOT NULL, CHECK (grado IN (10, 11))", "Nivel escolar."],
        ["anio_lectivo", "smallint", "NOT NULL", "Año académico."],
        ["estado", "enum", "NOT NULL: activa, archivada", "Estado del curso."]
    ]
    add_table_data(doc, h_cols, d_clase, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.14 evaluacion
    add_heading_with_spacing(doc, "3.14 evaluacion", level=3)
    d_eval = [
        ["id", "UUID", "PK", "Identificador."],
        ["clase_id", "UUID", "FK → clase, NOT NULL", "Clase asignada."],
        ["titulo", "varchar(150)", "NOT NULL", "Nombre del examen."],
        ["codigo", "varchar(10)", "NOT NULL, UNIQUE", "Código único VF-XXXX."],
        ["modo_seleccion", "enum", "NOT NULL: automatica, manual", "Generada o elegida."],
        ["dificultad", "enum", "NULL: ponderada, facil, media, dificil", "Filtro en automática."],
        ["cantidad_preguntas", "smallint", "NOT NULL, CHECK (cantidad BETWEEN 5 AND 50)", "Total de preguntas."],
        ["duracion_min", "smallint", "NULL", "Minutos límite (NULL = sin límite)."],
        ["intentos_max", "smallint", "NOT NULL, default 1", "Máximo de intentos."],
        ["abre_en", "timestamptz", "NOT NULL", "Inicio de ventana."],
        ["cierra_en", "timestamptz", "NOT NULL, CHECK (cierra_en > abre_en)", "Cierre de ventana."],
        ["mostrar_resultado", "enum", "NOT NULL: inmediato, diferido", "Modo de entrega."],
        ["barajar", "boolean", "NOT NULL, default true", "Aleatoriedad."],
        ["estado", "enum", "NOT NULL: borrador, publicada, cerrada, archivada", "Ciclo de vida."],
        ["creada_por", "UUID", "FK → usuario, NOT NULL", "Docente autor."]
    ]
    add_table_data(doc, h_cols, d_eval, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.15 evaluacion_tema y evaluacion_pregunta
    add_heading_with_spacing(doc, "3.15 evaluacion_tema y evaluacion_pregunta", level=3)
    d_eval_comp = [
        ["evaluacion_tema", "evaluacion_id (FK), tema_id (FK)", "PK compuesta", "Temas incluidos en pruebas automáticas."],
        ["evaluacion_pregunta", "evaluacion_id (FK), pregunta_id (FK), orden, puntos", "PK compuesta", "Preguntas congeladas al publicar la prueba."]
    ]
    add_table_data(doc, ["Tabla", "Columnas", "Restricciones", "Descripción"], d_eval_comp, col_widths=[1.5, 2.2, 1.3, 2.0])

    # 3.16 intento
    add_heading_with_spacing(doc, "3.16 intento", level=3)
    d_int = [
        ["id", "UUID", "PK", "Identificador."],
        ["estudiante_id", "UUID", "FK → usuario, NOT NULL", "Estudiante evaluado."],
        ["evaluacion_id", "UUID", "FK → evaluacion, NOT NULL", "Examen oficial presentado."],
        ["numero_intento", "smallint", "NOT NULL, default 1", "Ordinal de intento."],
        ["total_preguntas", "smallint", "NOT NULL", "Preguntas totales."],
        ["estado", "enum", "NOT NULL: en_curso, finalizado, expirado, anulado", "Estado del intento."],
        ["inicio_en", "timestamptz", "NOT NULL, default now()", "Inicio cronometrado."],
        ["limite_en", "timestamptz", "NULL", "Hora máxima calculada."],
        ["fin_en", "timestamptz", "NULL", "Entrega efectiva."],
        ["aciertos", "smallint", "NULL", "Aciertos logrados."],
        ["puntaje", "numeric(3,2)", "NULL, CHECK (puntaje BETWEEN 1.00 AND 5.00)", "Nota definitiva (1.00 a 5.00)."]
    ]
    add_table_data(doc, h_cols, d_int, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.17 intento_respuesta
    add_heading_with_spacing(doc, "3.17 intento_respuesta", level=3)
    d_iresp = [
        ["id", "UUID", "PK", "Identificador."],
        ["intento_id", "UUID", "FK → intento, NOT NULL", "Intento asociado."],
        ["pregunta_id", "UUID", "FK → pregunta, NOT NULL", "Pregunta respondida."],
        ["orden", "smallint", "NOT NULL", "Orden mostrado."],
        ["opcion_elegida_id", "UUID", "FK → opcion_pregunta, NULL", "Opción en selección múltiple."],
        ["texto_respuesta", "varchar(200)", "NULL", "Texto digitado en completar."],
        ["es_correcta", "boolean", "NULL", "Validación en servidor."],
        ["respondida_en", "timestamptz", "NOT NULL, default now()", "Guardado progresivo."]
    ]
    add_table_data(doc, h_cols, d_iresp, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 3.18 evento_examen
    add_heading_with_spacing(doc, "3.18 evento_examen", level=3)
    d_iev = [
        ["id", "bigint", "PK autoincremental", "Identificador."],
        ["intento_id", "UUID", "FK → intento, NOT NULL", "Intento en curso."],
        ["tipo", "enum", "NOT NULL: perdida_foco, salida_pestana, pantalla_completa_salida, reconexion", "Señal académica."],
        ["ocurrido_en", "timestamptz", "NOT NULL, default now()", "Marca temporal."],
        ["detalle", "jsonb", "NULL", "Telemetría complementaria."]
    ]
    add_table_data(doc, h_cols, d_iev, col_widths=[1.5, 1.2, 1.8, 2.5])

    # 4. Reglas de cálculo y escala oficial
    add_heading_with_spacing(doc, "4. Reglas de cálculo y escala oficial", level=2)
    add_p(doc, "Puntaje = 1.0 + (aciertos / total_preguntas) * 4.0", bold_prefix="Fórmula institucional (MEN): ")
    add_bullet(doc, "Rango de notas", "1.00 (0 aciertos) hasta 5.00 (100% de aciertos).")
    add_bullet(doc, "Aprobación", "Nota mínima aprobatoria: 3.00 (50% de aciertos).")
    add_bullet(doc, "Sin gamificación", "La nota refleja estrictamente el conocimiento lingüístico demostrado.")

    # 5. Seguridad y privacidad en voz
    add_heading_with_spacing(doc, "5. Seguridad y privacidad en reconocimiento de voz", level=2)
    add_bullet(doc, "No almacenamiento de voz", "Las señales de audio captadas por el micrófono no se graban ni se envían a servidores externos; el navegador extrae el texto en tiempo real y descarta el buffer de audio de la memoria.")
    add_bullet(doc, "Cumplimiento Ley 1581 / Ley 1098", "Garantía total de protección de datos biométricos de menores de edad.")

    doc.save(filepath)
    print(f"Modelo de datos v0.3 guardado exitosamente en {filepath}")

generate_modelo_datos("Docs/VerbFlow — Modelo de Datos.docx")
