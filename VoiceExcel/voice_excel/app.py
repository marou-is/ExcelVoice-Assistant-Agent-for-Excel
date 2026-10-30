import queue
import threading
import tkinter as tk
import traceback
from datetime import datetime
from tkinter import messagebox, scrolledtext, ttk

from . import config
from .excel_bridge import ExcelBridge, ExcelError
from .recorder import Recorder
from .transcriber import Transcriber
from .vba_generator import GenerationError, VBAGenerator, find_risky_operations


class VoiceExcelApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.recorder = Recorder()
        self.transcriber = Transcriber()
        self.generator: VBAGenerator | None = None
        self._ui_queue: queue.Queue = queue.Queue()
        self._busy = False
        self.transcriber_ready = False

        root.title("Voice to Excel")
        root.geometry("820x720")
        self._build_widgets()
        self.root.after(100, self._drain_ui_queue)

        self._run_in_background(self._warm_up)

    # ---------- layout ----------

    def _build_widgets(self) -> None:
        pad = {"padx": 10, "pady": 4}

        top = ttk.Frame(self.root)
        top.pack(fill="x", **pad)
        self.talk_button = tk.Button(
            top, text="Hold to talk", width=18, height=2, bg="#2e7d32", fg="white",
            font=("Segoe UI", 11, "bold"), state="disabled",
        )
        self.talk_button.pack(side="left")
        self.talk_button.bind("<ButtonPress-1>", self._on_talk_press)
        self.talk_button.bind("<ButtonRelease-1>", self._on_talk_release)

        self.auto_run = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            top, text="Generate and run automatically after speaking", variable=self.auto_run
        ).pack(side="left", padx=12)

        self.status = tk.StringVar(value="Loading speech model...")
        ttk.Label(self.root, textvariable=self.status, font=("Segoe UI", 10, "italic")).pack(
            anchor="w", **pad
        )

        ttk.Label(self.root, text="Request (transcribed, editable):").pack(anchor="w", **pad)
        self.request_box = tk.Text(self.root, height=3, wrap="word", font=("Segoe UI", 11))
        self.request_box.pack(fill="x", **pad)

        actions = ttk.Frame(self.root)
        actions.pack(fill="x", **pad)
        self.generate_button = ttk.Button(actions, text="Generate VBA", command=self._on_generate)
        self.generate_button.pack(side="left")
        self.run_button = ttk.Button(actions, text="Run in Excel", command=self._on_run)
        self.run_button.pack(side="left", padx=8)

        ttk.Label(self.root, text="Generated VBA (editable):").pack(anchor="w", **pad)
        self.vba_box = scrolledtext.ScrolledText(self.root, height=18, font=("Consolas", 10))
        self.vba_box.pack(fill="both", expand=True, **pad)

        ttk.Label(self.root, text="Log:").pack(anchor="w", **pad)
        self.log_box = scrolledtext.ScrolledText(
            self.root, height=7, font=("Consolas", 9), state="disabled"
        )
        self.log_box.pack(fill="x", **pad)

    # ---------- thread helpers ----------

    def _ui(self, func, *args) -> None:
        self._ui_queue.put((func, args))

    def _drain_ui_queue(self) -> None:
        while not self._ui_queue.empty():
            func, args = self._ui_queue.get_nowait()
            func(*args)
        self.root.after(100, self._drain_ui_queue)

    def _run_in_background(self, target, *args) -> None:
        def wrapper():
            try:
                target(*args)
            except Exception as error:
                self._ui(self._log, "".join(traceback.format_exception(error)))
                self._ui(self._set_status, f"Error: {error}")
            finally:
                self._ui(self._set_busy, False)

        self._set_busy(True)
        threading.Thread(target=wrapper, daemon=True).start()

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.generate_button.config(state=state)
        self.run_button.config(state=state)
        if self.transcriber_ready:
            self.talk_button.config(state=state)

    def _set_status(self, text: str) -> None:
        self.status.set(text)

    def _log(self, text: str) -> None:
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.config(state="normal")
        self.log_box.insert("end", f"[{stamp}] {text}\n")
        self.log_box.see("end")
        self.log_box.config(state="disabled")

    def _set_text(self, widget: tk.Text, text: str) -> None:
        widget.delete("1.0", "end")
        widget.insert("1.0", text)

    def _request_text(self) -> str:
        return self.request_box.get("1.0", "end").strip()

    def _vba_text(self) -> str:
        return self.vba_box.get("1.0", "end").strip()

    # ---------- startup ----------

    def _warm_up(self) -> None:
        self._ui(self._log, f"Microphone: [{self.recorder.device}] {self.recorder.device_name}")
        self._ui(self._log, f"Loading Whisper '{config.WHISPER_MODEL}' (first run downloads it)...")
        self.transcriber.load()
        self.transcriber_ready = True
        self._ui(self._log, "Speech model ready.")
        try:
            self.generator = VBAGenerator()
        except GenerationError as error:
            self._ui(self._log, str(error))
        self._ui(self._set_status, "Ready. Hold the button and speak your request in French.")

    # ---------- voice ----------

    def _on_talk_press(self, _event) -> None:
        if self._busy or not self.transcriber_ready:
            return
        try:
            self.recorder.start()
        except Exception as error:
            self._log(f"Microphone error: {error}")
            return
        self.talk_button.config(text="Listening... release", bg="#c62828")
        self._set_status("Listening...")

    def _on_talk_release(self, _event) -> None:
        if not self.recorder.is_recording:
            return
        self.talk_button.config(text="Hold to talk", bg="#2e7d32")
        audio = self.recorder.stop()
        self._run_in_background(self._voice_pipeline, audio, self.auto_run.get())

    def _voice_pipeline(self, audio, auto: bool) -> None:
        seconds = audio.size / config.SAMPLE_RATE
        peak = float(abs(audio).max()) if audio.size else 0.0
        self._ui(self._log, f"Recorded {seconds:.1f}s, peak level {peak:.3f}")
        if seconds < 0.5:
            self._ui(self._set_status, "Too short. Keep holding the button while you speak.")
            return
        if peak < 0.005:
            self._ui(
                self._set_status,
                "The microphone recorded silence. Set MIC_DEVICE in .env or check Windows mic settings.",
            )
            return
        self._ui(self._set_status, "Transcribing...")
        text = self.transcriber.transcribe(audio)
        if not text:
            self._ui(self._set_status, "Didn't catch anything. Try again, a bit louder or longer.")
            return
        self._ui(self._set_text, self.request_box, text)
        self._ui(self._log, f"Heard: {text}")
        if auto:
            self._generate_and_maybe_run(text, run=True)
        else:
            self._ui(self._set_status, "Transcribed. Check the text, then click Generate VBA.")

    # ---------- generate / run ----------

    def _on_generate(self) -> None:
        request = self._request_text()
        if not request:
            messagebox.showinfo("Voice to Excel", "Say or type a request first.")
            return
        self._run_in_background(self._generate_and_maybe_run, request, False)

    def _on_run(self) -> None:
        code = self._vba_text()
        if not code:
            messagebox.showinfo("Voice to Excel", "There is no VBA to run yet.")
            return
        if not self._confirm_risks(code):
            return
        self._run_in_background(self._execute, self._request_text(), code)

    def _require_generator(self) -> VBAGenerator:
        if self.generator is None:
            self.generator = VBAGenerator()
        return self.generator

    def _generate_and_maybe_run(self, request: str, run: bool) -> None:
        generator = self._require_generator()
        self._ui(self._set_status, "Reading workbook and asking Gemini for VBA...")
        with ExcelBridge() as excel:
            context = excel.describe_workbook()
        code = generator.generate(request, context)
        self._ui(self._set_text, self.vba_box, code)
        self._ui(self._log, "VBA generated.")

        if not run:
            self._ui(self._set_status, "VBA ready. Review it, then click Run in Excel.")
            return
        risks = find_risky_operations(code)
        if risks:
            self._ui(self._log, "Auto-run paused, the code " + "; ".join(risks) + ".")
            self._ui(self._set_status, "Review the VBA: it does something sensitive. Click Run to confirm.")
            return
        self._execute(request, code)

    def _execute(self, request: str, code: str) -> None:
        attempts_left = config.AUTO_FIX_ATTEMPTS
        with ExcelBridge() as excel:
            while True:
                self._ui(self._set_status, "Running macro in Excel...")
                try:
                    excel.run_vba(code)
                except ExcelError as error:
                    self._ui(self._log, f"Excel error: {error}")
                    if attempts_left <= 0 or "Trust access" in str(error) or not request:
                        self._ui(self._set_status, f"Failed: {error}")
                        return
                    attempts_left -= 1
                    self._ui(self._set_status, "Macro failed, asking Gemini to fix it...")
                    code = self._require_generator().fix(
                        request, excel.describe_workbook(), code, str(error)
                    )
                    self._ui(self._set_text, self.vba_box, code)
                    if find_risky_operations(code):
                        self._ui(self._set_status, "Fixed code needs review before running. Click Run to confirm.")
                        return
                    continue
                break
        save_history(request, code)
        self._ui(self._log, "Done.")
        self._ui(self._set_status, "Done. Check Excel.")

    def _confirm_risks(self, code: str) -> bool:
        risks = find_risky_operations(code)
        if not risks:
            return True
        details = "\n".join(f"- {risk}" for risk in risks)
        return messagebox.askyesno(
            "Voice to Excel", f"This macro:\n{details}\n\nRun it anyway?", icon="warning"
        )


def save_history(request: str, code: str) -> None:
    config.HISTORY_DIR.mkdir(exist_ok=True)
    path = config.HISTORY_DIR / f"{datetime.now():%Y%m%d-%H%M%S}.bas"
    comment = "\n".join(f"' {line}" for line in request.splitlines())
    path.write_text(f"{comment}\n\n{code}\n", encoding="utf-8")


def main() -> None:
    root = tk.Tk()
    VoiceExcelApp(root)
    root.mainloop()
