// ========================================================
// VerbFlow - Motor de Reconocimiento y Síntesis de Voz
// Enfoque Comunicativo (CLT) & Word-Diff Matcher
// Privacidad Total: Ejecución 100% en el Navegador del Cliente
// ========================================================

const SpeechEngine = (() => {
  const isSpeechRecognitionSupported = typeof window !== "undefined" && 
    Boolean(window.SpeechRecognition || window.webkitSpeechRecognition);

  const isSpeechSynthesisSupported = typeof window !== "undefined" && 
    Boolean(window.speechSynthesis);

  let recognition = null;
  let isListening = false;
  let selectedVoice = null;

  // Diccionario de Reglas Fonéticas para hispanohablantes aprendiendo inglés
  const PHONETIC_RULES = [
    {
      pattern: /\b(worked|watched|asked|cooked|stopped|talked)\b/i,
      tip: "Terminación '-ed' sorda: Se pronuncia como una sola sílaba con sonido /t/ final (ej: 'worked' suena /wɜːkt/, no 'work-ed')."
    },
    {
      pattern: /\b(lived|loved|called|played|opened|rained)\b/i,
      tip: "Terminación '-ed' sonora: Se pronuncia con sonido /d/ suave y una sola sílaba (ej: 'rained' suena /reɪnd/, no 'rein-ed')."
    },
    {
      pattern: /\b(started|decided|needed|visited|restarted)\b/i,
      tip: "Terminación '-ed' con vocal extra: Solo cuando el verbo termina en sonido /t/ o /d/ se pronuncia /-ɪd/ (ej: 'start-id')."
    },
    {
      pattern: /\b(watching|catch|chicken|teacher)\b/i,
      tip: "Sonido /tʃ/ ('ch' explosiva): No lo confundas con /ʃ/ ('sh' suave). 'Watching' = ver, 'washing' = lavar."
    },
    {
      pattern: /\b(stars|student|school|study|spanish|special)\b/i,
      tip: "Sonido 'S' inicial: Empieza directamente con /s/ como el silbido de una serpiente. ¡Evita poner una 'e' antes (no digas 'estars')!"
    },
    {
      pattern: /\b(chemistry|technology|architect)\b/i,
      tip: "Sonido 'ch' griego: 'Chemistry' y 'technology' suenan con sonido /k/ (/ˈkem.ɪ.stri/), no con sonido /ch/."
    },
    {
      pattern: /\b(could|would|should|half|talk|walk)\b/i,
      tip: "Letras mudas: La 'L' en 'could', 'would' y 'should' es completamente silenciosa (/kʊd/, /wʊd/, /ʃʊd/)."
    }
  ];

  // Inicializar voces del navegador
  function initVoices() {
    if (!isSpeechSynthesisSupported) return;
    const voices = window.speechSynthesis.getVoices();
    // Priorizar voz en inglés (EE. UU. o Reino Unido)
    selectedVoice = voices.find(v => v.lang.startsWith("en-US")) || 
                    voices.find(v => v.lang.startsWith("en-GB")) || 
                    voices.find(v => v.lang.startsWith("en")) || 
                    voices[0];
  }

  if (isSpeechSynthesisSupported) {
    initVoices();
    if (window.speechSynthesis.onvoiceschanged !== undefined) {
      window.speechSynthesis.onvoiceschanged = initVoices;
    }
  }

  // ----------------------------------------------------
  // SÍNTESIS DE VOZ (BOT HABLA EN INGLÉS)
  // ----------------------------------------------------
  function speak(text, { rate = 0.95, pitch = 1.0, onEnd = null } = {}) {
    if (!isSpeechSynthesisSupported) {
      console.warn("SpeechSynthesis no soportado en este navegador.");
      if (onEnd) onEnd();
      return;
    }

    // Cancelar cualquier audio anterior
    window.speechSynthesis.cancel();

    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = "en-US";
    utterance.rate = rate; // Ritmo ligeramente pausado para estudiantes
    utterance.pitch = pitch;

    if (selectedVoice) {
      utterance.voice = selectedVoice;
    }

    if (onEnd) {
      utterance.onend = onEnd;
      utterance.onerror = () => onEnd();
    }

    window.speechSynthesis.speak(utterance);
  }

  // ----------------------------------------------------
  // RECONOCIMIENTO DE VOZ DEL ESTUDIANTE
  // ----------------------------------------------------
  function createRecognitionInstance({ onResult, onError, onStart, onEnd }) {
    if (!isSpeechRecognitionSupported) return null;

    const SpeechRec = window.SpeechRecognition || window.webkitSpeechRecognition;
    const rec = new SpeechRec();
    rec.lang = "en-US";
    rec.continuous = false;
    rec.interimResults = true;
    rec.maxAlternatives = 3;

    rec.onstart = () => {
      isListening = true;
      if (onStart) onStart();
    };

    rec.onresult = (event) => {
      let interim = "";
      let finalTranscript = "";

      for (let i = event.resultIndex; i < event.results.length; ++i) {
        const item = event.results[i];
        if (item.isFinal) {
          finalTranscript += item[0].transcript;
        } else {
          interim += item[0].transcript;
        }
      }

      if (onResult) {
        onResult({
          transcript: (finalTranscript || interim).trim(),
          isFinal: Boolean(finalTranscript),
          confidence: event.results[0] && event.results[0][0] ? event.results[0][0].confidence : 0.85
        });
      }
    };

    rec.onerror = (event) => {
      console.warn("Speech Recognition Error:", event.error);
      isListening = false;
      if (onError) onError(event);
    };

    rec.onend = () => {
      isListening = false;
      if (onEnd) onEnd();
    };

    return rec;
  }

  function startListening(callbacks) {
    if (!isSpeechRecognitionSupported) {
      if (callbacks.onError) {
        callbacks.onError({ error: "not-supported", message: "Tu navegador no soporta reconocimiento de voz nativo." });
      }
      return false;
    }

    try {
      if (recognition && isListening) {
        recognition.stop();
      }
      recognition = createRecognitionInstance(callbacks);
      if (recognition) {
        recognition.start();
        return true;
      }
    } catch (e) {
      console.error("Error al iniciar micrófono:", e);
      if (callbacks.onError) callbacks.onError(e);
    }
    return false;
  }

  function stopListening() {
    if (recognition && isListening) {
      recognition.stop();
      isListening = false;
    }
  }

  // ----------------------------------------------------
  // WORD-DIFF MATCHER (ALGORITMO DE COMPARACIÓN PALABRA POR PALABRA)
  // ----------------------------------------------------
  function cleanToken(word) {
    return (word || "").toLowerCase().replace(/[^a-z0-9']/g, "").trim();
  }

  // Distancia de Levenshtein para tolerancia de pronunciación
  function levenshtein(a, b) {
    const matrix = [];
    for (let i = 0; i <= b.length; i++) matrix[i] = [i];
    for (let j = 0; j <= a.length; j++) matrix[0][j] = j;

    for (let i = 1; i <= b.length; i++) {
      for (let j = 1; j <= a.length; j++) {
        if (b.charAt(i - 1) === a.charAt(j - 1)) {
          matrix[i][j] = matrix[i - 1][j - 1];
        } else {
          matrix[i][j] = Math.min(
            matrix[i - 1][j - 1] + 1,
            Math.min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1)
          );
        }
      }
    }
    return matrix[b.length][a.length];
  }

  function similarityScore(w1, w2) {
    const clean1 = cleanToken(w1);
    const clean2 = cleanToken(w2);
    if (!clean1 || !clean2) return 0;
    if (clean1 === clean2) return 1.0;
    const maxLen = Math.max(clean1.length, clean2.length);
    const dist = levenshtein(clean1, clean2);
    return Math.max(0, (maxLen - dist) / maxLen);
  }

  function diffWords(targetSentence, spokenSentence) {
    const targetWords = (targetSentence || "").split(/\s+/).filter(Boolean);
    const spokenTokens = (spokenSentence || "")
      .split(/\s+/)
      .map(cleanToken)
      .filter(Boolean);

    let matchCount = 0;
    const results = [];
    const usedSpokenIndices = new Set();

    targetWords.forEach((targetWord) => {
      const cleanTarget = cleanToken(targetWord);
      let bestMatch = { index: -1, score: 0 };

      // Buscar si alguna palabra hablada coincide
      spokenTokens.forEach((spokenWord, idx) => {
        if (usedSpokenIndices.has(idx)) return;
        const score = similarityScore(cleanTarget, spokenWord);
        if (score > bestMatch.score) {
          bestMatch = { index: idx, score };
        }
      });

      // Si la similitud es >= 0.75, se considera acertada
      if (bestMatch.score >= 0.75) {
        usedSpokenIndices.add(bestMatch.index);
        matchCount++;
        results.push({
          word: targetWord,
          status: "matched",
          score: bestMatch.score,
          spokenAs: spokenTokens[bestMatch.index]
        });
      } else {
        results.push({
          word: targetWord,
          status: "mismatched",
          score: bestMatch.score,
          spokenAs: bestMatch.index !== -1 ? spokenTokens[bestMatch.index] : null
        });
      }
    });

    const accuracy = targetWords.length > 0 ? matchCount / targetWords.length : 0;

    // Calcular nota en escala institucional 1.00 a 5.00
    const institutionalScore = Number((1.0 + 4.0 * accuracy).toFixed(2));

    // Identificar sugerencias fonéticas aplicables
    const tips = [];
    PHONETIC_RULES.forEach(rule => {
      if (rule.pattern.test(targetSentence)) {
        tips.push(rule.tip);
      }
    });

    return {
      results,
      matchCount,
      totalWords: targetWords.length,
      accuracy,
      institutionalScore,
      passed: institutionalScore >= 3.0,
      tips
    };
  }

  return {
    isSpeechRecognitionSupported,
    isSpeechSynthesisSupported,
    speak,
    startListening,
    stopListening,
    diffWords,
    PHONETIC_RULES
  };
})();

if (typeof window !== "undefined") {
  window.SpeechEngine = SpeechEngine;
}
