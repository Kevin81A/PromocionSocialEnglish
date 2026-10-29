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

def generate_politica_doc(filepath):
    doc = docx.Document()
    
    # Title
    p_title = doc.add_heading("VerbFlow — Política de Tratamiento de Datos Personales y Consentimiento Informado", level=1)
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    
    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(12)
    r_meta = p_meta.add_run("Versión 1.0 (Institucional) · Ley 1581 de 2012 / Ley 1098 de 2006 · I.E. Promoción Social (Palermo, Huila) · Octubre de 2026")
    r_meta.font.italic = True
    r_meta.font.size = Pt(9.5)
    r_meta.font.color.rgb = RGBColor(100, 116, 139)
    
    # 1. Identificación del Responsable y Marco Normativo
    add_heading_with_spacing(doc, "1. Identificación del Responsable y Marco Normativo", level=2)
    add_p(doc, "El presente documento establece la Política de Tratamiento y Protección de Datos Personales de la plataforma web educativa VerbFlow, en estricto cumplimiento del Régimen General de Protección de Datos Personales de la República de Colombia, constituido por la Ley Estatutaria 1581 de 2012, el Decreto Reglamentario 1377 de 2013, el Código de la Infancia y la Adolescencia (Ley 1098 de 2006) y la jurisprudencia de la Corte Constitucional (Sentencia C-748 de 2011).")
    
    h_resp = ["Identificación Institucional", "Detalle del Responsable"]
    d_resp = [
        ["Institución Educativa", "Institución Educativa Promoción Social"],
        ["Sede / Municipio", "Municipio de Palermo, Departamento del Huila, Colombia"],
        ["Naturaleza del Servicio", "Educación Media Oficial (Grados Décimo 10° y Once 11°)"],
        ["Plataforma Tecnológica", "VerbFlow — Sistema Web Privado de Aprendizaje y Evaluación del Idioma Inglés"],
        ["Área de Aplicación", "Área de Humanidades, Lengua Castellana e Idioma Extranjero (Inglés)"],
        ["Canal Oficial de Hábeas Data", "Buzón de Atención y Rectoría / Coordinación Académica I.E. Promoción Social"]
    ]
    add_table_data(doc, h_resp, d_resp, col_widths=[2.5, 4.5])

    # 2. Principios Rectores del Tratamiento
    add_heading_with_spacing(doc, "2. Principios Rectores del Tratamiento", level=2)
    add_p(doc, "El tratamiento de los datos recolectados en VerbFlow se rige por los siguientes principios legales:")
    add_bullet(doc, "Principio de Legalidad", "El tratamiento se sujeta a lo dispuesto en la Ley 1581 de 2012 y demás normas concordantes.")
    add_bullet(doc, "Principio de Finalidad", "Los datos se tratan con fines estrictamente académicos, pedagógicos y de evaluación formativa del inglés.")
    add_bullet(doc, "Principio de Libertad", "El registro requiere el consentimiento expreso, previo e informado del titular (mayor de edad) o de su representante legal (acudiente/padre de familia).")
    add_bullet(doc, "Principio de Minimización y Necesidad", "La plataforma solo recolecta los datos mínimos indispensables para individualizar al alumno y calificarlo. Se prohíbe exigir documentos de identidad, números telefónicos, domicilios o información socioeconómica sensible.")
    add_bullet(doc, "Principio de Seguridad y Confidencialidad", "Se implementan medidas técnicas y administrativas para evitar la adulteración, pérdida, consulta, uso o acceso no autorizado.")

    # 3. Tratamiento Especial de Datos de Menores de Edad
    add_heading_with_spacing(doc, "3. Tratamiento Especial de Datos de Menores de Edad (Ley 1098 de 2006)", level=2)
    add_p(doc, "En concordancia con el Artículo 47 del Código de la Infancia y la Adolescencia (Ley 1098 de 2006) y el Artículo 12 del Decreto 1377 de 2013, el tratamiento de datos de niños, niñas y adolescentes en VerbFlow cumple las siguientes condiciones imperativas:")
    add_bullet(doc, "Prevalencia de Derechos", "El tratamiento responde y respeta en todo momento el interés superior de los estudiantes menores de edad y asegura la primacía de sus derechos fundamentales.")
    add_bullet(doc, "Autorización del Acudiente", "En el formulario de registro se exige la manifestación expresa de que el registro cuenta con la debida autorización del padre, madre o acudiente legal debidamente facultado.")
    add_bullet(doc, "Prohibición de Perfilamiento Conductual y Gamificación Invasiva", "VerbFlow excluye intencionalmente sistemas de métricas comerciales, rankings competitivos públicos, acumulación de monedas virtuales o seguimiento de hábitos diarios para evitar la generación de estrés escolar y prevenir la creación de perfiles conductuales de menores.")

    # 4. Declaración Explícita sobre Voz y Datos Biométricos
    add_heading_with_spacing(doc, "4. Declaración Explícita sobre Reconocimiento de Voz y Biometría", level=2)
    add_p(doc, "La plataforma VerbFlow incorpora un Módulo Comunicativo y Oral que permite al estudiante interactuar mediante micrófono en diálogos contextuales y recibir retroalimentación fonética inmediata. Frente a esta funcionalidad, la I.E. Promoción Social declara de manera formal y categórica:")
    add_bullet(doc, "Procesamiento Exclusivo en la Memoria del Navegador (Client-Side Edge)", "El reconocimiento de voz se realiza mediante la Web Speech API nativa implementada en el navegador web del usuario. El análisis acústico y fonético ocurre en la memoria volátil (RAM) del computador o celular del alumno.")
    add_bullet(doc, "Cero Grabación o Almacenamiento de Voz", "La plataforma NO graba, NO almacena en bases de datos ni conserva en servidores ningún fragmento de audio, voz o patrón biométrico del estudiante. Una vez que la frase hablada es contrastada contra el modelo esperado, los búferes de audio son destruidos instantáneamente en memoria.")
    add_bullet(doc, "No Transmisión a Terceros", "Las señales de audio del micrófono no se transmiten a empresas de publicidad, modelos de entrenamiento externo de terceros ni a servicios pagos en la nube.")

    # 5. Inventario de Datos Recolectados y Finalidad
    add_heading_with_spacing(doc, "5. Inventario de Datos Personales Recolectados", level=2)
    h_datos = ["Dato Solicitado", "Carácter", "Finalidad Exclusiva de Uso"]
    d_datos = [
        ["Nombre Completo", "Obligatorio", "Identificación oficial en la lista de clase del docente para asentar calificaciones institucionales."],
        ["Nickname (Usuario)", "Obligatorio", "Credencial única de acceso a la plataforma sin necesidad de exponer el nombre en pantalla."],
        ["Contraseña Cifrada", "Obligatorio", "Autenticación segura (almacenada mediante hash unidireccional irreversible Argon2/PBKDF2)."],
        ["Grado Escolar (10° u 11°)", "Obligatorio (Estudiante)", "Segmentación del nivel curricular y asociación al plan de estudios de educación media."],
        ["Clase Asignada (clase_id)", "Obligatorio (Estudiante)", "Vinculación al grupo académico específico a cargo de un docente de inglés."],
        ["Correo Electrónico", "Opcional", "Canal técnico para restablecimiento autónomo de contraseña en caso de olvido."],
        ["Registro de Consentimiento", "Obligatorio", "Evidencia legal de aceptación de la presente política (fecha, hora, versión e IP)."],
        ["Respuestas y Calificaciones", "Obligatorio (Examen)", "Evaluación de desempeño académico en escala 1.0 a 5.0 para reportes oficiales de clase."]
    ]
    add_table_data(doc, h_datos, d_datos, col_widths=[2.0, 1.6, 3.4])

    # 6. Derechos de los Titulares (Hábeas Data)
    add_heading_with_spacing(doc, "6. Derechos de los Titulares (Derecho de Hábeas Data)", level=2)
    add_p(doc, "De conformidad con el Artículo 8 de la Ley 1581 de 2012, los estudiantes y sus representantes legales tienen los siguientes derechos:")
    add_bullet(doc, "Conocer y Consultar", "Conocer qué datos personales reposan en la plataforma y consultar su historial de notas en todo momento.")
    add_bullet(doc, "Actualizar y Rectificar", "Solicitar la corrección de nombres mal escritos o actualización de grado ante su docente titular o el Owner.")
    add_bullet(doc, "Revocar Consentimiento y Suprimir Datos (Derecho de Supresión)", "Solicitar la baja de la cuenta y eliminación de sus datos personales de la base de datos.")
    add_bullet(doc, "Protocolo Especial de Anonimización", "Por exigencia de la normatividad escolar del Ministerio de Educación Nacional (MEN), las calificaciones oficiales de exámenes no pueden destruirse durante el año escolar. Al solicitar la supresión de la cuenta, VerbFlow procede a la anonimización: se eliminan el nombre, nickname y correo, sustituyéndolos por identificadores genéricos anónimos, desvinculando para siempre la identidad del estudiante de los resultados numéricos.")

    # 7. Medidas de Seguridad de la Información
    add_heading_with_spacing(doc, "7. Medidas Técnicas y Administrativas de Seguridad", level=2)
    add_bullet(doc, "Cifrado en Tránsito (HTTPS / TLS)", "Todas las comunicaciones entre el navegador del estudiante y el servidor viajan cifradas con certificados de seguridad SSL/TLS.")
    add_bullet(doc, "Aislamiento Docente por RLS", "Mediante políticas de seguridad a nivel de filas (Row Level Security) en PostgreSQL, ningún docente puede ver, consultar o exportar notas de estudiantes pertenecientes a otros cursos.")
    add_bullet(doc, "Protección Antifraude de Exámenes (RNF-04)", "Las claves de respuestas correctas nunca residen en el código fuente de la página ni en el navegador durante una prueba de clase; se validan de forma confidencial en el servidor.")
    add_bullet(doc, "Control de Acceso del Owner", "Toda creación de cuentas con privilegios de Docente es verificada manualmente por el Owner antes de conceder permisos de creación de cursos o publicación de exámenes.")

    # 8. Texto Oficial de Consentimiento Informado en el Registro
    add_heading_with_spacing(doc, "8. Modelo de Consentimiento Informado para el Registro", level=2)
    add_p(doc, "El formulario de registro de la plataforma (`index.html`) presenta de forma obligatoria el siguiente texto de autorización que debe ser aceptado mediante casilla de verificación activa antes de procesar el registro:")
    
    h_box = ["Texto Legal de Autorización en Pantalla (Checkbox RF-02)"]
    d_box = [
        ["\"Declaro bajo la gravedad de juramento que soy mayor de edad (18 años) o que cuento con la debida autorización de mi padre, madre o acudiente legal para el tratamiento de mis datos personales en la plataforma VerbFlow de la Institución Educativa Promoción Social (Palermo, Huila). Conozco que mis datos serán tratados con fines exclusivamente pedagógicos y académicos del área de inglés, que las prácticas de voz con micrófono se ejecutan en mi dispositivo sin almacenar grabaciones de audio, y que puedo ejercer en cualquier momento mis derechos de hábeas data para consultar, rectificar o anonimizar mi cuenta conforme a la Ley 1581 de 2012 y la Política de Privacidad Institucional.\""]
    ]
    add_table_data(doc, h_box, d_box, col_widths=[7.0])

    # 9. Canales de Atención de PQR y Procedimiento de Hábeas Data
    add_heading_with_spacing(doc, "9. Canales de Atención de Consultas y Reclamos (PQR)", level=2)
    add_p(doc, "Para ejercer los derechos de hábeas data, los estudiantes o acudientes pueden acudir a los siguientes canales institucionales:")
    add_bullet(doc, "Atención Presencial", "Rectoría o Coordinación Académica de la Institución Educativa Promoción Social, Sede Principal, Municipio de Palermo, Huila.")
    add_bullet(doc, "Atención Virtual Institucional", "A través del docente titular del área de inglés o el canal de soporte técnico de la plataforma administrado por el Owner.")
    add_bullet(doc, "Tiempos de Respuesta", "Las consultas de información serán resueltas en un plazo máximo de diez (10) días hábiles; los reclamos de rectificación o supresión serán atendidos en un término no mayor a quince (15) días hábiles, conforme a los Artículos 14 y 15 de la Ley 1581 de 2012.")

    # 10. Vigencia y Actualizaciones
    add_heading_with_spacing(doc, "10. Vigencia y Actualizaciones de la Política", level=2)
    add_p(doc, "La presente política rige a partir de su publicación oficial en el mes de octubre de 2026 y permanecerá vigente durante el tiempo en que la plataforma VerbFlow preste sus servicios educativos en la Institución Educativa Promoción Social. Cualquier modificación sustancial en las finalidades del tratamiento será notificada previamente en la pantalla principal de la aplicación y requerirá una nueva aceptación informada por parte de los usuarios.")

    doc.save(filepath)
    print(f"Política de Tratamiento de Datos guardada exitosamente en {filepath}")

generate_politica_doc("Docs/VerbFlow — Política de Tratamiento de Datos Personales.docx")
