// ========================================================
// VerbFlow - Banco de Escenarios Conversacionales y Rol-Play
// Enfoque Comunicativo (CLT) y Reconocimiento de Voz
// Institución Educativa Promoción Social • Palermo, Huila
// ========================================================

const SCENARIOS_DATA = [
  {
    id: "detective-alibi",
    code: "ESC-01",
    title: "The Detective's Alibi",
    topic: "Past Continuous vs. Past Simple (Interrogatorio)",
    tenseCategory: "past-continuous",
    level: "A2 - B1",
    context: "Misterio en el colegio: El trofeo de química del laboratorio desapareció ayer a las 7:30 PM. El Oficial Jenkins está interrogando a los estudiantes de 11° para verificar sus coartadas.",
    emotion: "Misterioso / Inquisitivo",
    avatarBot: "👮‍♂️",
    botName: "Detective Jenkins",
    studentRole: "Estudiante sospechoso con coartada real",
    dialogues: [
      {
        turn: 1,
        speaker: "bot",
        textEn: "Good afternoon. I am Detective Jenkins. Where were you yesterday at 7:30 PM?",
        textEs: "Buenas tardes. Soy el Detective Jenkins. ¿Dónde estabas ayer a las 7:30 PM?",
        audioModel: "Good afternoon. I am Detective Jenkins. Where were you yesterday at 7:30 PM?",
        emotion: "serious",
        phoneticTip: "Presta atención al tono inquisitivo formal de la pregunta."
      },
      {
        turn: 2,
        speaker: "student",
        textEn: "I was doing my chemistry homework at home.",
        textEs: "Estaba haciendo mi tarea de química en casa.",
        audioModel: "I was doing my chemistry homework at home.",
        emotion: "defensive",
        phoneticTip: "'Chemistry' se pronuncia con sonido /k/ al inicio (/ˈkem.ɪ.stri/), nunca con /ch/. 'Was doing' suena suave: /wəz ˈduː.ɪŋ/.",
        keywords: ["i", "was", "doing", "chemistry", "homework", "at", "home"],
        tolerance: 0.8
      },
      {
        turn: 3,
        speaker: "bot",
        textEn: "Interesting. Was anyone with you while you were studying?",
        textEs: "Interesante. ¿Había alguien contigo mientras estabas estudiando?",
        audioModel: "Interesting. Was anyone with you while you were studying?",
        emotion: "suspicious",
        phoneticTip: "Nota la cláusula con 'while' que introduce una acción continua simultánea."
      },
      {
        turn: 4,
        speaker: "student",
        textEn: "Yes, my sister was watching TV in the living room.",
        textEs: "Sí, mi hermana estaba viendo televisión en la sala.",
        audioModel: "Yes, my sister was watching TV in the living room.",
        emotion: "confident",
        phoneticTip: "¡Cuidado!: 'Watching' tiene sonido /tʃ/ (como 'chango'), no lo confundas con 'washing' /ʃ/ (lavar).",
        keywords: ["yes", "sister", "was", "watching", "tv", "living", "room"],
        tolerance: 0.8
      },
      {
        turn: 5,
        speaker: "bot",
        textEn: "Did you hear any strange noise coming from the school street?",
        textEs: "¿Escuchaste algún ruido extraño que viniera de la calle del colegio?",
        audioModel: "Did you hear any strange noise coming from the school street?",
        emotion: "curious",
        phoneticTip: "Uso de 'Did' como auxiliar del Pasado Simple para una acción puntual interrumpida."
      },
      {
        turn: 6,
        speaker: "student",
        textEn: "No, while we were studying, it was raining heavily.",
        textEs: "No, mientras estudiábamos, estaba lloviendo fuertemente.",
        audioModel: "No, while we were studying, it was raining heavily.",
        emotion: "sincere",
        phoneticTip: "'Raining' /ˈreɪn.ɪŋ/. 'Heavily' /ˈhev.ɪ.li/ tiene acento en la primera sílaba.",
        keywords: ["no", "while", "were", "studying", "it", "was", "raining", "heavily"],
        tolerance: 0.8
      }
    ]
  },
  {
    id: "cafeteria-order",
    code: "ESC-02",
    title: "The School Cafeteria Order",
    topic: "Present Simple vs. Present Continuous (Rutina vs. Momento actual)",
    tenseCategory: "present-continuous",
    level: "A2",
    context: "Hora del recreo en la I.E. Promoción Social. Doña Carmen te atiende en la cafetería para tomar tu pedido y hablar de tus hábitos alimenticios.",
    emotion: "Alegre / Hambriento",
    avatarBot: "👩‍🍳",
    botName: "Doña Carmen (Cafetería)",
    studentRole: "Estudiante en fila de almuerzo",
    dialogues: [
      {
        turn: 1,
        speaker: "bot",
        textEn: "Hello! Welcome to the cafeteria. What do you usually have for lunch?",
        textEs: "¡Hola! Bienvenido a la cafetería. ¿Qué sueles almorzar normalmente?",
        audioModel: "Hello! Welcome to the cafeteria. What do you usually have for lunch?",
        emotion: "friendly",
        phoneticTip: "Pregunta sobre hábitos y rutinas usando el adverbio de frecuencia 'usually'."
      },
      {
        turn: 2,
        speaker: "student",
        textEn: "I usually eat chicken with rice and fresh salad.",
        textEs: "Normalmente como pollo con arroz y ensalada fresca.",
        audioModel: "I usually eat chicken with rice and fresh salad.",
        emotion: "casual",
        phoneticTip: "'Usually' se pronuncia con sonido /ʒ/: /ˈjuː.ʒu.ə.li/. 'Chicken' suena /ˈtʃɪk.ɪn/.",
        keywords: ["i", "usually", "eat", "chicken", "with", "rice", "salad"],
        tolerance: 0.8
      },
      {
        turn: 3,
        speaker: "bot",
        textEn: "Sounds healthy! And what are you drinking right now?",
        textEs: "¡Suena saludable! ¿Y qué estás tomando justo ahora?",
        audioModel: "Sounds healthy! And what are you drinking right now?",
        emotion: "curious",
        phoneticTip: "Cambio de Presente Simple a Presente Continuo con el marcador 'right now'."
      },
      {
        turn: 4,
        speaker: "student",
        textEn: "I am drinking cold passion fruit juice.",
        textEs: "Estoy tomando jugo frío de maracuyá.",
        audioModel: "I am drinking cold passion fruit juice.",
        emotion: "happy",
        phoneticTip: "'Juice' se pronuncia /dʒuːs/ (sin pronunciar la 'i'). 'Fruit' rima con 'boot': /fruːt/.",
        keywords: ["i", "am", "drinking", "cold", "passion", "fruit", "juice"],
        tolerance: 0.8
      }
    ]
  },
  {
    id: "late-homework",
    code: "ESC-03",
    title: "The Late Homework Excuse",
    topic: "Present Perfect vs. Past Simple (Acción reciente y causa)",
    tenseCategory: "present-perfect",
    level: "B1",
    context: "Inicio de la clase de inglés. La Licenciada Martha Lucía está revisando la entrega de proyectos y tú debes explicar tu situación con respeto y exactitud.",
    emotion: "Apenado / Respetuoso",
    avatarBot: "👩‍🏫",
    botName: "Teacher Martha",
    studentRole: "Estudiante rindiendo cuentas de su tarea",
    dialogues: [
      {
        turn: 1,
        speaker: "bot",
        textEn: "Good morning! Did you bring your English workbook today?",
        textEs: "¡Buenos días! ¿Trajiste tu cuaderno de inglés hoy?",
        audioModel: "Good morning! Did you bring your English workbook today?",
        emotion: "strict",
        phoneticTip: "Pregunta directa en Pasado Simple con auxiliar 'Did'."
      },
      {
        turn: 2,
        speaker: "student",
        textEn: "I am sorry teacher, I have lost my backpack.",
        textEs: "Lo siento profesora, he perdido mi morral.",
        audioModel: "I am sorry teacher, I have lost my backpack.",
        emotion: "apologetic",
        phoneticTip: "Usa Present Perfect 'have lost' porque la pérdida ocurrió en el pasado pero el efecto sigue vigente ahora. 'Lost' tiene 'o' corta: /lɒst/.",
        keywords: ["i", "am", "sorry", "teacher", "have", "lost", "backpack"],
        tolerance: 0.8
      },
      {
        turn: 3,
        speaker: "bot",
        textEn: "Oh no! Have you searched for it in the classroom?",
        textEs: "¡Oh no! ¿Lo has buscado en el salón de clases?",
        audioModel: "Oh no! Have you searched for it in the classroom?",
        emotion: "concerned",
        phoneticTip: "Verbo regular en participio pasado 'searched'. Nota que la '-ed' suena como /t/."
      },
      {
        turn: 4,
        speaker: "student",
        textEn: "Yes, I have already searched everywhere, but I couldn't find it.",
        textEs: "Sí, ya he buscado en todas partes, pero no pude encontrarlo.",
        audioModel: "Yes, I have already searched everywhere, but I could not find it.",
        emotion: "worried",
        phoneticTip: "'Searched' se pronuncia en una sola sílaba: /sɜːtʃt/. En 'couldn't', la letra 'l' es completamente muda: /ˈkʊd.ənt/.",
        keywords: ["yes", "have", "already", "searched", "everywhere", "could", "not", "find"],
        tolerance: 0.75
      }
    ]
  },
  {
    id: "tatacoa-trip",
    code: "ESC-04",
    title: "Weekend Trip to the Tatacoa Desert",
    topic: "Future Simple (will) vs. Going to (Planes y Predicciones)",
    tenseCategory: "future-simple",
    level: "A2 - B1",
    context: "Salida pedagógica del grado 11° por el departamento del Huila. Estás coordinando los preparativos de campamento y observación astronómica con tu compañero de grupo.",
    emotion: "Aventurero / Entusiasta",
    avatarBot: "⛺",
    botName: "Mateo (Compañero de 11°)",
    studentRole: "Líder de equipo de expedición escolar",
    dialogues: [
      {
        turn: 1,
        speaker: "bot",
        textEn: "Are you ready for our trip to the Tatacoa Desert next Saturday?",
        textEs: "¿Estás listo para nuestro viaje al Desierto de la Tatacoa el próximo sábado?",
        audioModel: "Are you ready for our trip to the Tatacoa Desert next Saturday?",
        emotion: "excited",
        phoneticTip: "Entonación ascendente para pregunta de confirmación entusiasta."
      },
      {
        turn: 2,
        speaker: "student",
        textEn: "Yes, we are going to camp under the stars.",
        textEs: "Sí, vamos a acampar bajo las estrellas.",
        audioModel: "Yes, we are going to camp under the stars.",
        emotion: "motivated",
        phoneticTip: "'Going to' suena fluido /ˌɡoʊ.ɪŋ tuː/. 'Stars' empieza directamente con el sonido /s/, ¡evita decir 'estars'! /stɑːrz/.",
        keywords: ["yes", "we", "are", "going", "to", "camp", "under", "stars"],
        tolerance: 0.8
      },
      {
        turn: 3,
        speaker: "bot",
        textEn: "What will the temperature be like during the night?",
        textEs: "¿Cómo estará la temperatura durante la noche?",
        audioModel: "What will the temperature be like during the night?",
        emotion: "curious",
        phoneticTip: "Pregunta sobre predicción meteorológica con 'will'."
      },
      {
        turn: 4,
        speaker: "student",
        textEn: "It will be very windy, so I will bring a warm jacket.",
        textEs: "Hará mucho viento, así que llevaré una chaqueta abrigada.",
        audioModel: "It will be very windy, so I will bring a warm jacket.",
        emotion: "responsible",
        phoneticTip: "'Windy' /ˈwɪn.di/. 'Warm' tiene sonido /ɔː/: /wɔːrm/, no /warm/.",
        keywords: ["it", "will", "be", "windy", "so", "bring", "warm", "jacket"],
        tolerance: 0.8
      }
    ]
  },
  {
    id: "lab-assistant-interview",
    code: "ESC-05",
    title: "Student Lab Assistant Interview",
    topic: "Present Simple & Past Experiences (Entrevista Formal)",
    tenseCategory: "present-simple",
    level: "B1",
    context: "Convocatoria escolar para el monitor del aula de sistemas e informática de la I.E. Promoción Social. El coordinador académico evalúa tus habilidades.",
    emotion: "Formal / Seguro",
    avatarBot: "👔",
    botName: "Coordinador Académico",
    studentRole: "Candidato a monitor estudiantil",
    dialogues: [
      {
        turn: 1,
        speaker: "bot",
        textEn: "Welcome to the interview. Why do you want to be our computer assistant?",
        textEs: "Bienvenido a la entrevista. ¿Por qué quieres ser nuestro asistente de computadores?",
        audioModel: "Welcome to the interview. Why do you want to be our computer assistant?",
        emotion: "professional",
        phoneticTip: "Pregunta abierta formal de motivación."
      },
      {
        turn: 2,
        speaker: "student",
        textEn: "Because I enjoy helping other students and I love technology.",
        textEs: "Porque disfruto ayudar a otros estudiantes y me encanta la tecnología.",
        audioModel: "Because I enjoy helping other students and I love technology.",
        emotion: "enthusiastic",
        phoneticTip: "'Enjoy' termina con diptongo /ɔɪ/. 'Technology' lleva el acento en la segunda sílaba: /tekˈnɒl.ə.dʒi/.",
        keywords: ["because", "enjoy", "helping", "other", "students", "love", "technology"],
        tolerance: 0.8
      },
      {
        turn: 3,
        speaker: "bot",
        textEn: "What did you do last year when the school network failed?",
        textEs: "¿Qué hiciste el año pasado cuando falló la red del colegio?",
        audioModel: "What did you do last year when the school network failed?",
        emotion: "inquisitive",
        phoneticTip: "Pregunta sobre experiencia pasada resolviendo problemas técnicos."
      },
      {
        turn: 4,
        speaker: "student",
        textEn: "I restarted the routers and checked the network cables.",
        textEs: "Reinicié los enrutadores y revisé los cables de red.",
        audioModel: "I restarted the routers and checked the network cables.",
        emotion: "confident",
        phoneticTip: "'Restarted' termina en /-ɪd/ porque el verbo termina en 't': /rɪˈstɑːtɪd/. Pero en 'checked', la terminación suena /t/: /tʃekt/.",
        keywords: ["i", "restarted", "routers", "checked", "network", "cables"],
        tolerance: 0.8
      }
    ]
  }
];

if (typeof window !== "undefined") {
  window.SCENARIOS_DATA = SCENARIOS_DATA;
}

if (typeof module !== "undefined" && module.exports) {
  module.exports = SCENARIOS_DATA;
}
