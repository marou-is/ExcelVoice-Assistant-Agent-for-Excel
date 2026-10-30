import threading

import numpy as np
from faster_whisper import WhisperModel

from . import config

# Biases Whisper toward the vocabulary of spreadsheet commands.
VOCABULARY_HINT = (
    "Crée un tableau avec les colonnes id, nom du client, âge du client, date, "
    "montant, quantité, prix unitaire. Ajoute une colonne, une formule, un graphique, "
    "trie, filtre, mise en forme conditionnelle, feuille, cellule, ligne, total."
)


class Transcriber:
    def __init__(self):
        self._model: WhisperModel | None = None
        self._lock = threading.Lock()

    def load(self) -> None:
        with self._lock:
            if self._model is None:
                self._model = WhisperModel(
                    config.WHISPER_MODEL,
                    device=config.WHISPER_DEVICE,
                    compute_type=config.WHISPER_COMPUTE_TYPE,
                )

    def transcribe(self, audio: np.ndarray) -> str:
        if audio.size < config.SAMPLE_RATE // 2:
            return ""
        self.load()
        segments, _ = self._model.transcribe(
            audio,
            language=config.SPEECH_LANGUAGE,
            beam_size=5,
            vad_filter=True,
            initial_prompt=VOCABULARY_HINT,
        )
        return " ".join(segment.text.strip() for segment in segments).strip()
