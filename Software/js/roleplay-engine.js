// ========================================================
// VerbFlow - Gestor de Rol-Play Conversacional y Diálogos
// Coordina los turnos, emociones, retroalimentación y notas
// I.E. Promoción Social • Palermo, Huila
// ========================================================

const RoleplayEngine = (() => {
  let activeScenario = null;
  let currentTurnIndex = 0;
  let sessionScores = [];
  let isRecording = false;

  // Callbacks de UI registrados
  let uiListeners = {
    onScenarioLoaded: null,
    onTurnChanged: null,
    onBotSpeakingStart: null,
    onBotSpeakingEnd: null,
    onStudentListeningStart: null,
    onStudentInterimTranscript: null,
    onStudentTurnEvaluated: null,
    onScenarioCompleted: null
  };

  function setListeners(listeners) {
    uiListeners = { ...uiListeners, ...listeners };
  }

  function loadScenario(scenarioId) {
    const list = window.SCENARIOS_DATA || [];
    activeScenario = list.find(s => s.id === scenarioId) || list[0];
    currentTurnIndex = 0;
    sessionScores = [];

    if (uiListeners.onScenarioLoaded && activeScenario) {
      uiListeners.onScenarioLoaded(activeScenario);
    }

    // Iniciar el primer turno
    executeCurrentTurn();
    return activeScenario;
  }

  function getCurrentTurn() {
    if (!activeScenario || !activeScenario.dialogues) return null;
    return activeScenario.dialogues[currentTurnIndex] || null;
  }

  function executeCurrentTurn() {
    const turn = getCurrentTurn();
    if (!turn) {
      finishScenario();
      return;
    }

    if (uiListeners.onTurnChanged) {
      uiListeners.onTurnChanged(turn, currentTurnIndex, activeScenario.dialogues.length);
    }

    if (turn.speaker === "bot") {
      // El bot habla automáticamente
      if (uiListeners.onBotSpeakingStart) uiListeners.onBotSpeakingStart(turn);
      
      SpeechEngine.speak(turn.textEn, {
        rate: 0.9,
        onEnd: () => {
          if (uiListeners.onBotSpeakingEnd) uiListeners.onBotSpeakingEnd(turn);
          // Pasar automáticamente al turno del estudiante
          setTimeout(() => {
            currentTurnIndex++;
            executeCurrentTurn();
          }, 800);
        }
      });
    } else if (turn.speaker === "estudiante" || turn.speaker === "student") {
      // Turno del estudiante: listo para escuchar
      if (uiListeners.onTurnChanged) {
        uiListeners.onTurnChanged(turn, currentTurnIndex, activeScenario.dialogues.length);
      }
    }
  }

  function repeatBotAudio() {
    const turn = getCurrentTurn();
    if (turn && turn.speaker === "bot") {
      SpeechEngine.speak(turn.textEn, { rate: 0.9 });
    } else {
      // Si estamos en turno del estudiante, escuchar el modelo en inglés
      const current = getCurrentTurn();
      if (current) {
        SpeechEngine.speak(current.textEn, { rate: 0.85 });
      }
    }
  }

  function startStudentSpeaking() {
    const turn = getCurrentTurn();
    if (!turn || (turn.speaker !== "estudiante" && turn.speaker !== "student")) {
      console.warn("No es turno del estudiante.");
      return;
    }

    isRecording = true;
    if (uiListeners.onStudentListeningStart) {
      uiListeners.onStudentListeningStart();
    }

    const started = SpeechEngine.startListening({
      onStart: () => {
        isRecording = true;
      },
      onResult: (data) => {
        if (uiListeners.onStudentInterimTranscript) {
          uiListeners.onStudentInterimTranscript(data.transcript, data.isFinal);
        }
        if (data.isFinal) {
          evaluateStudentSpeech(data.transcript);
        }
      },
      onError: (err) => {
        isRecording = false;
        console.warn("Error en micrófono:", err);
      },
      onEnd: () => {
        isRecording = false;
      }
    });

    if (!started) {
      // Si el navegador no tiene soporte de micrófono, permitir modo manual simulado
      console.info("Simulando turno oral...");
    }
  }

  function stopStudentSpeaking() {
    SpeechEngine.stopListening();
    isRecording = false;
  }

  function evaluateStudentSpeech(spokenText) {
    const turn = getCurrentTurn();
    if (!turn) return;

    const evaluation = SpeechEngine.diffWords(turn.textEn, spokenText);
    sessionScores.push({
      turnIndex: currentTurnIndex,
      target: turn.textEn,
      spoken: spokenText,
      evaluation
    });

    if (uiListeners.onStudentTurnEvaluated) {
      uiListeners.onStudentTurnEvaluated(evaluation, spokenText);
    }
  }

  function nextTurn() {
    currentTurnIndex++;
    executeCurrentTurn();
  }

  function finishScenario() {
    // Calcular promedio general del escenario en escala 1.00 - 5.00
    let avgScore = 1.0;
    if (sessionScores.length > 0) {
      const sum = sessionScores.reduce((acc, curr) => acc + curr.evaluation.institutionalScore, 0);
      avgScore = Number((sum / sessionScores.length).toFixed(2));
    }

    const summary = {
      scenario: activeScenario,
      totalTurns: activeScenario ? activeScenario.dialogues.length : 0,
      studentTurns: sessionScores.length,
      averageScore: avgScore,
      passed: avgScore >= 3.0,
      turnHistory: sessionScores
    };

    if (uiListeners.onScenarioCompleted) {
      uiListeners.onScenarioCompleted(summary);
    }

    return summary;
  }

  return {
    setListeners,
    loadScenario,
    getCurrentTurn,
    repeatBotAudio,
    startStudentSpeaking,
    stopStudentSpeaking,
    evaluateStudentSpeech,
    nextTurn,
    finishScenario,
    getActiveScenario: () => activeScenario
  };
})();

if (typeof window !== "undefined") {
  window.RoleplayEngine = RoleplayEngine;
}
