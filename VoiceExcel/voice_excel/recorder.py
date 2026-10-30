import numpy as np
import sounddevice as sd

from . import config

LOOPBACK_KEYWORDS = ("stereo mix", "mixage stéréo", "loopback", "what u hear", "sound mapper", "primary sound")


def _is_loopback(name: str) -> bool:
    return any(keyword in name.lower() for keyword in LOOPBACK_KEYWORDS)


def pick_input_device() -> int:
    """Return the device index to record from, skipping loopback devices like Stereo Mix."""
    devices = sd.query_devices()
    inputs = [i for i, d in enumerate(devices) if d["max_input_channels"] > 0]

    if config.MIC_DEVICE:
        if config.MIC_DEVICE.isdigit():
            return int(config.MIC_DEVICE)
        for i in inputs:
            if config.MIC_DEVICE.lower() in devices[i]["name"].lower():
                return i
        raise RuntimeError(f"No input device matches MIC_DEVICE={config.MIC_DEVICE!r}")

    default = sd.default.device[0]
    if default is not None and default >= 0 and not _is_loopback(devices[default]["name"]):
        return default

    default_api = devices[default]["hostapi"] if default is not None and default >= 0 else 0
    candidates = [i for i in inputs if not _is_loopback(devices[i]["name"])]
    if not candidates:
        raise RuntimeError("No microphone found. Check Windows sound settings.")
    candidates.sort(key=lambda i: (devices[i]["hostapi"] != default_api, "micro" not in devices[i]["name"].lower()))
    return candidates[0]


class Recorder:
    """Push-to-talk microphone recorder producing 16 kHz mono float32 audio."""

    def __init__(self, sample_rate: int = config.SAMPLE_RATE):
        self.sample_rate = sample_rate
        self.device = pick_input_device()
        self.device_name = sd.query_devices(self.device)["name"]
        self._frames: list[np.ndarray] = []
        self._stream: sd.InputStream | None = None

    @property
    def is_recording(self) -> bool:
        return self._stream is not None

    def start(self) -> None:
        if self._stream is not None:
            return
        self._frames = []
        self._stream = sd.InputStream(
            device=self.device,
            samplerate=self.sample_rate,
            channels=1,
            dtype="float32",
            callback=self._on_audio,
        )
        self._stream.start()

    def stop(self) -> np.ndarray:
        if self._stream is None:
            return np.zeros(0, dtype=np.float32)
        self._stream.stop()
        self._stream.close()
        self._stream = None
        if not self._frames:
            return np.zeros(0, dtype=np.float32)
        return np.concatenate(self._frames)[:, 0]

    def _on_audio(self, indata, frames, time, status) -> None:
        self._frames.append(indata.copy())
