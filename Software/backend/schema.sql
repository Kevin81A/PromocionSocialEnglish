-- ==============================================================================
-- VerbFlow - Production Schema & Security Configuration (PostgreSQL / Supabase)
-- Versión Oficial: v0.3 (Enfoque Comunicativo, Diálogos y Escala Colombiana 1.0-5.0)
-- Institución Educativa Promoción Social • Palermo, Huila
-- 19 Tablas Relacionales bajo estricto cumplimiento Ley 1581 de 2012 y Ley 1098 de 2006
-- ==============================================================================

-- 1. EXTENSIONES REQUERIDAS
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";
CREATE EXTENSION IF NOT EXISTS "pgcrypto";

-- 2. TIPOS ENUMERADOS
DO $$ BEGIN
    CREATE TYPE user_role AS ENUM ('estudiante', 'docente', 'owner');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE account_status AS ENUM ('aprobado', 'pendiente', 'rechazado');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE evaluation_status AS ENUM ('borrador', 'publicada', 'cerrada');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE attempt_status AS ENUM ('en_progreso', 'finalizado', 'anulado');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE exam_security_event AS ENUM ('cambio_pestana', 'perdida_foco', 'intento_copia', 'bloqueo_pantalla');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

-- ==============================================================================
-- 3. DOMINIO INSTITUCIONAL Y CLASES
-- ==============================================================================

-- TABLA 13: clase (1 estudiante pertenece a 1 sola clase activa a la vez)
CREATE TABLE IF NOT EXISTS public.clase (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    codigo_clase VARCHAR(10) UNIQUE NOT NULL, -- Ej: VF-1102
    nombre VARCHAR(100) NOT NULL,            -- Ej: Grado 11-02 Inglés
    grado VARCHAR(10) NOT NULL,              -- Ej: 1102
    institucion VARCHAR(120) NOT NULL DEFAULT 'I.E. Promoción Social',
    docente_id UUID,                         -- Referencia a usuario docente
    activo BOOLEAN NOT NULL DEFAULT true,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- TABLA 1: usuario (perfiles vinculados a auth.users de Supabase)
CREATE TABLE IF NOT EXISTS public.usuario (
    id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
    nick VARCHAR(30) UNIQUE NOT NULL,
    nombre_completo VARCHAR(150) NOT NULL,
    institucion VARCHAR(120) NOT NULL DEFAULT 'I.E. Promoción Social',
    grado VARCHAR(10) DEFAULT '1102',
    rol user_role NOT NULL DEFAULT 'estudiante',
    estado account_status NOT NULL DEFAULT 'aprobado',
    clase_id UUID REFERENCES public.clase(id) ON DELETE SET NULL,
    consent_1581 BOOLEAN NOT NULL DEFAULT true,
    consent_timestamp TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    CONSTRAINT nick_format_check CHECK (nick ~* '^[a-zA-Z0-9_-]{3,30}$')
);

-- Llave foránea en clase hacia usuario docente
DO $$ BEGIN
    ALTER TABLE public.clase
    ADD CONSTRAINT fk_clase_docente FOREIGN KEY (docente_id) REFERENCES public.usuario(id) ON DELETE SET NULL;
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

-- TABLA 2: consentimiento (Trazabilidad legal Ley 1581 de 2012 & 1098 de 2006)
CREATE TABLE IF NOT EXISTS public.consentimiento (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    usuario_id UUID NOT NULL REFERENCES public.usuario(id) ON DELETE CASCADE,
    version_politica VARCHAR(10) NOT NULL DEFAULT 'v1.0',
    aceptado BOOLEAN NOT NULL DEFAULT true,
    ip_registro INET,
    fecha_registro TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- TABLA 3: solicitud_docente (Aprobación manual y estricta por el Owner 'thekeas')
CREATE TABLE IF NOT EXISTS public.solicitud_docente (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    usuario_id UUID NOT NULL REFERENCES public.usuario(id) ON DELETE CASCADE,
    email_institucional VARCHAR(150) NOT NULL,
    justificacion TEXT NOT NULL,
    estado account_status NOT NULL DEFAULT 'pendiente',
    revisado_por UUID REFERENCES public.usuario(id),
    fecha_revision TIMESTAMPTZ,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- TABLA 4: auditoria (Registro inmutable de acciones críticas)
CREATE TABLE IF NOT EXISTS public.auditoria (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    usuario_id UUID REFERENCES public.usuario(id) ON DELETE SET NULL,
    accion VARCHAR(60) NOT NULL,
    entidad VARCHAR(60) NOT NULL,
    entidad_id UUID,
    ip INET,
    datos_previos JSONB,
    datos_nuevos JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- ==============================================================================
-- 4. DOMINIO COMUNICATIVO Y PRÁCTICA ORAL (ROLEPLAY)
-- ==============================================================================

-- TABLA 5: escenario_dialogo (Escenarios de vida cotidiana para 10° y 11°)
CREATE TABLE IF NOT EXISTS public.escenario_dialogo (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    codigo VARCHAR(20) UNIQUE NOT NULL,       -- Ej: ESC-01
    titulo VARCHAR(120) NOT NULL,             -- Ej: The Detective's Alibi
    descripcion TEXT NOT NULL,
    nivel VARCHAR(10) NOT NULL DEFAULT 'B1',  -- A2, B1, B2
    tema_gramatical VARCHAR(60) NOT NULL,     -- Past Continuous
    contexto_emocional VARCHAR(100) NOT NULL, -- Misterioso, defensivo
    avatar_bot VARCHAR(50) DEFAULT '👮‍♂️',
    bot_nombre VARCHAR(80) DEFAULT 'Detective Jenkins',
    orden INT NOT NULL DEFAULT 1,
    activo BOOLEAN NOT NULL DEFAULT true,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- TABLA 6: linea_dialogo (Turnos interactivos y guías fonéticas para Web Speech API)
CREATE TABLE IF NOT EXISTS public.linea_dialogo (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    escenario_id UUID NOT NULL REFERENCES public.escenario_dialogo(id) ON DELETE CASCADE,
    orden_turno INT NOT NULL,
    hablante VARCHAR(15) NOT NULL CHECK (hablante IN ('bot', 'estudiante')),
    texto_ingles TEXT NOT NULL,
    texto_espanol TEXT NOT NULL,
    emocion VARCHAR(40),
    pista_fonetica TEXT,
    tolerancia_coincidencia NUMERIC(3,2) DEFAULT 0.80,
    palabras_clave TEXT[]
);

-- ==============================================================================
-- 5. DOMINIO GRAMATICAL Y TEORÍA PEDAGÓGICA
-- ==============================================================================

-- TABLA 7: tema (Los 7 tiempos verbales oficiales)
CREATE TABLE IF NOT EXISTS public.tema (
    id VARCHAR(40) PRIMARY KEY, -- 'present-continuous', 'past-simple', etc.
    nombre VARCHAR(80) NOT NULL,
    nombre_es VARCHAR(100) NOT NULL,
    categoria VARCHAR(40) NOT NULL,
    nivel VARCHAR(40) NOT NULL DEFAULT 'Básico - Intermedio',
    descripcion TEXT NOT NULL,
    color_hex VARCHAR(10) DEFAULT '#2563eb',
    icono VARCHAR(40) DEFAULT 'school',
    orden INT NOT NULL DEFAULT 1,
    activo BOOLEAN NOT NULL DEFAULT true
);

-- TABLA 8: formula (Estructuras afirmativas, negativas e interrogativas)
CREATE TABLE IF NOT EXISTS public.formula (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    tema_id VARCHAR(40) NOT NULL REFERENCES public.tema(id) ON DELETE CASCADE,
    estructura_tipo VARCHAR(20) NOT NULL CHECK (estructura_tipo IN ('afirmativa', 'negativa', 'interrogativa')),
    patron TEXT NOT NULL,
    explicacion TEXT NOT NULL
);

-- TABLA 9: audio_ejemplo (Frases modelo con síntesis de voz Web Speech)
CREATE TABLE IF NOT EXISTS public.audio_ejemplo (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    tema_id VARCHAR(40) NOT NULL REFERENCES public.tema(id) ON DELETE CASCADE,
    oracion_ingles TEXT NOT NULL,
    oracion_espanol TEXT NOT NULL,
    duracion_seg NUMERIC(4,2)
);

-- TABLA 10: pregunta (Banco de 1.050 preguntas clasificadas)
CREATE TABLE IF NOT EXISTS public.pregunta (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    tema_id VARCHAR(40) NOT NULL REFERENCES public.tema(id) ON DELETE CASCADE,
    nivel_dificultad VARCHAR(10) NOT NULL CHECK (nivel_dificultad IN ('easy', 'medium', 'hard')),
    enunciado TEXT NOT NULL,
    explicacion_pedagogica TEXT,
    activo BOOLEAN NOT NULL DEFAULT true
);

-- TABLA 11: opcion_pregunta (Para preguntas de opción múltiple)
CREATE TABLE IF NOT EXISTS public.opcion_pregunta (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    pregunta_id UUID NOT NULL REFERENCES public.pregunta(id) ON DELETE CASCADE,
    texto_opcion TEXT NOT NULL,
    es_correcta BOOLEAN NOT NULL DEFAULT false,
    orden INT NOT NULL DEFAULT 1
);

-- TABLA 12: respuesta_aceptada (Para respuestas de completación y orales)
CREATE TABLE IF NOT EXISTS public.respuesta_aceptada (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    pregunta_id UUID NOT NULL REFERENCES public.pregunta(id) ON DELETE CASCADE,
    texto_canonica TEXT NOT NULL,
    es_alternativa BOOLEAN NOT NULL DEFAULT false
);

-- ==============================================================================
-- 6. DOMINIO DE EVALUACIÓN OFICIAL (ESCALA COLOMBIANA 1.00 A 5.00)
-- ==============================================================================

-- TABLA 14: evaluacion (Creadas por docentes para sus clases)
CREATE TABLE IF NOT EXISTS public.evaluacion (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    clase_id UUID NOT NULL REFERENCES public.clase(id) ON DELETE CASCADE,
    docente_id UUID NOT NULL REFERENCES public.usuario(id),
    titulo VARCHAR(120) NOT NULL,
    descripcion TEXT,
    duracion_min INT NOT NULL DEFAULT 45,
    total_preguntas INT NOT NULL DEFAULT 10,
    intentos_max INT NOT NULL DEFAULT 1,
    fecha_inicio TIMESTAMPTZ NOT NULL,
    fecha_cierre TIMESTAMPTZ NOT NULL,
    estado evaluation_status NOT NULL DEFAULT 'borrador',
    detectar_cambio_pestana BOOLEAN NOT NULL DEFAULT true,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

-- TABLA 15: evaluacion_tema (Ponderación temática en la prueba)
CREATE TABLE IF NOT EXISTS public.evaluacion_tema (
    evaluacion_id UUID NOT NULL REFERENCES public.evaluacion(id) ON DELETE CASCADE,
    tema_id VARCHAR(40) NOT NULL REFERENCES public.tema(id) ON DELETE CASCADE,
    peso_porcentaje INT NOT NULL DEFAULT 100,
    PRIMARY KEY (evaluacion_id, tema_id)
);

-- TABLA 16: evaluacion_pregunta (Preguntas asignadas al examen)
CREATE TABLE IF NOT EXISTS public.evaluacion_pregunta (
    evaluacion_id UUID NOT NULL REFERENCES public.evaluacion(id) ON DELETE CASCADE,
    pregunta_id UUID NOT NULL REFERENCES public.pregunta(id) ON DELETE CASCADE,
    orden INT NOT NULL DEFAULT 1,
    PRIMARY KEY (evaluacion_id, pregunta_id)
);

-- TABLA 17: intento (Presentaciones de exámenes por estudiantes)
-- Escala institucional: 1.00 a 5.00 (Aprobatoria >= 3.00)
CREATE TABLE IF NOT EXISTS public.intento (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    evaluacion_id UUID NOT NULL REFERENCES public.evaluacion(id) ON DELETE CASCADE,
    estudiante_id UUID NOT NULL REFERENCES public.usuario(id) ON DELETE CASCADE,
    aciertos INT NOT NULL DEFAULT 0,
    total_preguntas INT NOT NULL DEFAULT 0,
    puntaje_obtenido NUMERIC(3,2) NOT NULL DEFAULT 1.00 CHECK (puntaje_obtenido >= 1.00 AND puntaje_obtenido <= 5.00),
    estado attempt_status NOT NULL DEFAULT 'en_progreso',
    fecha_inicio TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    fecha_fin TIMESTAMPTZ,
    anulado_motivo TEXT
);

-- TABLA 18: intento_respuesta (Trazabilidad pedagógica de cada ítem respondido)
CREATE TABLE IF NOT EXISTS public.intento_respuesta (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    intento_id UUID NOT NULL REFERENCES public.intento(id) ON DELETE CASCADE,
    pregunta_id UUID NOT NULL REFERENCES public.pregunta(id) ON DELETE CASCADE,
    respuesta_dada TEXT NOT NULL,
    es_correcta BOOLEAN NOT NULL DEFAULT false,
    tiempo_segundos INT DEFAULT 0
);

-- TABLA 19: evento_examen (Seguridad institucional y monitoreo anti-fraude)
CREATE TABLE IF NOT EXISTS public.evento_examen (
    id UUID PRIMARY KEY DEFAULT uuid_generate_v4(),
    intento_id UUID NOT NULL REFERENCES public.intento(id) ON DELETE CASCADE,
    tipo_evento exam_security_event NOT NULL,
    timestamp_evento TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    detalles JSONB
);

-- ==============================================================================
-- 7. PROCEDIMIENTOS Y FUNCIONES OFICIALES
-- ==============================================================================

-- Función para cálculo exacto de la nota institucional en Colombia (1.00 a 5.00)
CREATE OR REPLACE FUNCTION public.calcular_nota_colombiana(p_aciertos INT, p_total INT)
RETURNS NUMERIC(3,2) AS $$
BEGIN
    IF p_total IS NULL OR p_total <= 0 THEN
        RETURN 1.00;
    END IF;
    IF p_aciertos IS NULL OR p_aciertos <= 0 THEN
        RETURN 1.00;
    END IF;
    IF p_aciertos >= p_total THEN
        RETURN 5.00;
    END IF;
    RETURN ROUND(1.00 + 4.00 * (p_aciertos::NUMERIC / p_total::NUMERIC), 2);
END;
$$ LANGUAGE plpgsql IMMUTABLE;

-- Trigger para calificar intento automáticamente al finalizar
CREATE OR REPLACE FUNCTION public.fn_finalizar_intento()
RETURNS TRIGGER AS $$
BEGIN
    IF NEW.estado = 'finalizado' AND OLD.estado = 'en_progreso' THEN
        NEW.fecha_fin = NOW();
        NEW.puntaje_obtenido = public.calcular_nota_colombiana(NEW.aciertos, NEW.total_preguntas);
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_finalizar_intento ON public.intento;
CREATE TRIGGER trg_finalizar_intento
BEFORE UPDATE ON public.intento
FOR EACH ROW
EXECUTE FUNCTION public.fn_finalizar_intento();

-- ==============================================================================
-- 8. POLÍTICAS DE SEGURIDAD A NIVEL DE FILAS (RLS)
-- ==============================================================================

ALTER TABLE public.usuario ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.clase ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.solicitud_docente ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.auditoria ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.escenario_dialogo ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.linea_dialogo ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.tema ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.formula ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.pregunta ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.evaluacion ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.intento ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.intento_respuesta ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.evento_examen ENABLE ROW LEVEL SECURITY;

-- Políticas de lectura pública para contenido pedagógico de estudio
CREATE POLICY "Contenido pedagogico visible para usuarios autenticados"
ON public.tema FOR SELECT TO authenticated USING (activo = true);

CREATE POLICY "Escenarios visibles para usuarios autenticados"
ON public.escenario_dialogo FOR SELECT TO authenticated USING (activo = true);

CREATE POLICY "Lineas de dialogo visibles para usuarios autenticados"
ON public.linea_dialogo FOR SELECT TO authenticated USING (true);

CREATE POLICY "Formulas visibles para usuarios autenticados"
ON public.formula FOR SELECT TO authenticated USING (true);

-- Políticas para usuarios: cada uno lee y actualiza su propio perfil
CREATE POLICY "Usuarios ven su propio perfil"
ON public.usuario FOR SELECT TO authenticated
USING (auth.uid() = id OR EXISTS (
    SELECT 1 FROM public.usuario u WHERE u.id = auth.uid() AND u.rol IN ('docente', 'owner')
));

-- Docente ve intentos de su clase; estudiante ve solo sus propios intentos
CREATE POLICY "Estudiantes ven sus propios intentos"
ON public.intento FOR SELECT TO authenticated
USING (estudiante_id = auth.uid() OR EXISTS (
    SELECT 1 FROM public.evaluacion ev
    JOIN public.clase cl ON ev.clase_id = cl.id
    WHERE ev.id = public.intento.evaluacion_id AND cl.docente_id = auth.uid()
));

-- Inserción de eventos de examen por parte del estudiante en curso
CREATE POLICY "Estudiantes registran eventos de su propio examen"
ON public.evento_examen FOR INSERT TO authenticated
WITH CHECK (EXISTS (
    SELECT 1 FROM public.intento i WHERE i.id = intento_id AND i.estudiante_id = auth.uid()
));

-- ==============================================================================
-- 9. DATOS SEMILLA BÁSICOS (SEED DATA)
-- ==============================================================================

INSERT INTO public.tema (id, nombre, nombre_es, categoria, nivel, descripcion, color_hex, icono, orden)
VALUES
('present-simple', 'Present Simple', 'Presente Simple', 'simple', 'Básico', 'Rutinas diarias, verdades generales y hábitos.', '#2563eb', 'today', 1),
('present-continuous', 'Present Continuous', 'Presente Continuo', 'continuous', 'Básico - Intermedio', 'Acciones que ocurren ahora mismo o temporales.', '#0284c7', 'update', 2),
('past-simple', 'Past Simple', 'Pasado Simple', 'simple', 'Básico', 'Acciones completadas en un momento específico del pasado.', '#16a34a', 'history', 3),
('past-continuous', 'Past Continuous', 'Pasado Continuo', 'continuous', 'Intermedio', 'Acciones que estaban en progreso en un momento pasado.', '#d97706', 'pending_actions', 4),
('future-simple', 'Future Simple', 'Futuro Simple (Will)', 'simple', 'Básico', 'Decisiones espontáneas, predicciones y promesas.', '#9333ea', 'fast_forward', 5),
('future-continuous', 'Future Continuous', 'Futuro Continuo', 'continuous', 'Intermedio', 'Acciones que estarán en progreso en un momento del futuro.', '#ea580c', 'timelapse', 6),
('present-perfect', 'Present Perfect', 'Presente Perfecto', 'perfect', 'Intermedio', 'Acciones pasadas con relevancia en el presente o experiencias de vida.', '#0d9488', 'task_alt', 7)
ON CONFLICT (id) DO NOTHING;
