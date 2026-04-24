# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import queue
import time
import re
import json
import webbrowser
from pathlib import Path
from collections import deque
import numpy as np
import pythoncom
import win32com.client
import ollama

# ── Kokoro constants ──────────────────────────────────────────────────────────

KOKORO_VOICES = [
    ("af_heart",    "Heart      — US Female"),
    ("af_bella",    "Bella      — US Female"),
    ("af_nicole",   "Nicole     — US Female"),
    ("af_sarah",    "Sarah      — US Female"),
    ("af_sky",      "Sky        — US Female"),
    ("af_nova",     "Nova       — US Female"),
    ("am_adam",     "Adam       — US Male"),
    ("am_michael",  "Michael    — US Male"),
    ("am_echo",     "Echo       — US Male"),
    ("am_eric",     "Eric       — US Male"),
    ("am_liam",     "Liam       — US Male"),
    ("bf_emma",     "Emma       — UK Female"),
    ("bf_isabella", "Isabella   — UK Female"),
    ("bm_george",   "George     — UK Male"),
    ("bm_lewis",    "Lewis      — UK Male"),
]

MODEL_DIR        = Path(__file__).parent / "models"
MODEL_ONNX       = MODEL_DIR / "kokoro-v1.0.onnx"
MODEL_VOICES_BIN = MODEL_DIR / "voices-v1.0.bin"

KOKORO_ONNX_URL   = "https://github.com/thewh1teagle/kokoro-onnx/releases/download/model-files-v1.0/kokoro-v1.0.onnx"
KOKORO_VOICES_URL = "https://github.com/thewh1teagle/kokoro-onnx/releases/download/model-files-v1.0/voices-v1.0.bin"

STT_MODELS    = ["tiny", "base", "small", "medium"]
SETTINGS_FILE = Path(__file__).parent / "settings.json"


# ── Tooltip ───────────────────────────────────────────────────────────────────

class Tooltip:
    def __init__(self, widget, text):
        self._widget = widget
        self._text   = text
        self._tip    = None
        widget.bind('<Enter>', self._show)
        widget.bind('<Leave>', self._hide)

    def _show(self, _=None):
        x = self._widget.winfo_rootx() + 20
        y = self._widget.winfo_rooty() + self._widget.winfo_height() + 6
        self._tip = tk.Toplevel(self._widget)
        self._tip.wm_overrideredirect(True)
        self._tip.wm_geometry(f'+{x}+{y}')
        tk.Label(
            self._tip, text=self._text,
            bg='#2a2a2a', fg='#ebebeb', font=('Segoe UI', 9),
            padx=10, pady=7, relief=tk.FLAT, wraplength=280, justify=tk.LEFT,
        ).pack()

    def _hide(self, _=None):
        if self._tip:
            self._tip.destroy()
            self._tip = None


# ── SAPI TTS engine ───────────────────────────────────────────────────────────

class SAPIEngine:
    _ASYNC = 1
    _PURGE = 2

    def __init__(self, on_ready):
        self._queue = queue.Queue()
        self._on_ready = on_ready
        self._live_rate = 0
        self._live_vol  = 100
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        pythoncom.CoInitialize()
        v = win32com.client.Dispatch("SAPI.SpVoice")
        tokens = v.GetVoices()
        self._tokens = [tokens.Item(i) for i in range(tokens.Count)]
        self._on_ready([t.GetDescription() for t in self._tokens])

        while True:
            try:
                cmd = self._queue.get(timeout=0.2)
            except queue.Empty:
                continue

            if cmd['action'] == 'stop':
                v.Speak('', self._PURGE | self._ASYNC)
            elif cmd['action'] == 'speak':
                idx = cmd['voice_idx']
                if 0 <= idx < len(self._tokens):
                    v.Voice = self._tokens[idx]
                v.Rate   = self._live_rate
                v.Volume = self._live_vol
                v.Speak('', self._PURGE | self._ASYNC)
                v.Speak(cmd['text'], self._ASYNC)

                stopped_early = False
                while v.Status.RunningState == 2:
                    time.sleep(0.04)
                    v.Rate   = self._live_rate
                    v.Volume = self._live_vol
                    if not self._queue.empty():
                        v.Speak('', self._PURGE | self._ASYNC)
                        stopped_early = True
                        break

                if not stopped_early and cmd.get('on_done'):
                    cmd['on_done']()

    def speak(self, text, voice_idx, rate, volume, on_done=None):
        self._live_rate = rate
        self._live_vol  = volume
        self._drain()
        self._queue.put({'action': 'speak', 'text': text, 'voice_idx': voice_idx,
                         'on_done': on_done})

    def update_live(self, rate, volume):
        self._live_rate = rate
        self._live_vol  = volume

    def stop(self):
        self._drain()
        self._queue.put({'action': 'stop'})

    def _drain(self):
        while not self._queue.empty():
            try:
                self._queue.get_nowait()
            except queue.Empty:
                break


# ── Kokoro TTS engine ─────────────────────────────────────────────────────────

class KokoroEngine:
    def __init__(self):
        self._kokoro   = None
        self._stop_evt = threading.Event()
        self._stream   = None

    @staticmethod
    def package_installed() -> bool:
        try:
            import kokoro_onnx  # noqa
            import sounddevice  # noqa
            return True
        except ImportError:
            return False

    @staticmethod
    def models_present() -> bool:
        return MODEL_ONNX.exists() and MODEL_VOICES_BIN.exists()

    def load(self):
        from kokoro_onnx import Kokoro
        self._kokoro = Kokoro(str(MODEL_ONNX), str(MODEL_VOICES_BIN))

    def speak(self, text, voice_code, get_speed, get_volume, on_done=None):
        import sounddevice as sd
        self._stop_evt.clear()
        sentences = [s.strip() for s in re.split(r'(?<=[.!?])\s+', text) if s.strip()] or [text]

        def run():
            for sent in sentences:
                if self._stop_evt.is_set():
                    return
                try:
                    samples, sr = self._kokoro.create(sent, voice=voice_code,
                                                      speed=get_speed(), lang="en-us")
                except Exception as exc:
                    print(f"Kokoro error: {exc}")
                    return
                if self._stop_evt.is_set():
                    return

                pos      = [0]
                done_evt = threading.Event()

                def callback(outdata, frames, _time, _status):
                    vol   = get_volume()
                    chunk = samples[pos[0]:pos[0] + frames]
                    n     = len(chunk)
                    if n < frames:
                        outdata[:n, 0]  = chunk * vol
                        outdata[n:, :]  = 0
                        raise sd.CallbackStop()
                    else:
                        outdata[:, 0] = chunk * vol
                    pos[0] += frames

                stream = sd.OutputStream(
                    samplerate=sr, channels=1, dtype='float32',
                    callback=callback,
                    finished_callback=lambda: done_evt.set(),
                )
                self._stream = stream
                stream.start()
                done_evt.wait()
                stream.close()
                self._stream = None

                if self._stop_evt.is_set():
                    return
            if on_done:
                on_done()

        threading.Thread(target=run, daemon=True).start()

    def stop(self):
        self._stop_evt.set()
        stream = self._stream
        if stream is not None:
            try:
                stream.stop()
            except Exception:
                pass

    def download_models(self, on_progress, on_done, on_error):
        import urllib.request
        MODEL_DIR.mkdir(parents=True, exist_ok=True)

        def fetch(url, dest, label):
            def report(count, block, total):
                if total > 0:
                    on_progress(f"Downloading {label}: {min(100, int(count * block * 100 / total))}%")
            urllib.request.urlretrieve(url, dest, reporthook=report)

        def run():
            try:
                fetch(KOKORO_VOICES_URL, MODEL_VOICES_BIN, "voices.bin")
                fetch(KOKORO_ONNX_URL,   MODEL_ONNX,        "kokoro.onnx (330 MB)")
                on_done()
            except Exception as exc:
                on_error(str(exc))

        threading.Thread(target=run, daemon=True).start()


# ── STT engine ────────────────────────────────────────────────────────────────

class STTEngine:
    def __init__(self):
        self._model       = None
        self._loaded_size: str | None = None
        self._stop_evt  = threading.Event()
        self._ready_evt = threading.Event()
        self._ready_evt.set()

    @staticmethod
    def package_installed() -> bool:
        import importlib.util
        return importlib.util.find_spec("faster_whisper") is not None

    def load(self, size: str, on_done, on_error):
        if self._loaded_size == size and self._model is not None:
            on_done()
            return

        def run():
            try:
                from faster_whisper import WhisperModel
                self._model = WhisperModel(size, device="cpu", compute_type="int8")
                self._loaded_size = size
                on_done()
            except Exception as exc:
                on_error(str(exc))

        threading.Thread(target=run, daemon=True).start()

    def start_continuous(self, on_interim, on_send, on_status, on_error, on_level=None):
        import sounddevice as sd
        RATE          = 16000
        CHUNK         = 1600   # 100 ms per chunk
        SILENCE_LIMIT = 20     # 20 × 100 ms = 2 s of silence triggers send
        INTERIM_EVERY = 20     # run interim transcription every ~2 s of speech

        self._stop_evt.clear()
        self._ready_evt.set()

        def run():
            on_status("Calibrating mic…")
            try:
                cal = sd.rec(RATE // 3, samplerate=RATE, channels=1, dtype='float32')
                sd.wait()
                threshold = max(0.015, float(np.sqrt(np.mean(cal ** 2))) * 4)
            except Exception as exc:
                on_error(str(exc))
                return

            while not self._stop_evt.is_set():
                self._ready_evt.wait()
                if self._stop_evt.is_set():
                    break

                on_status("Listening…")
                chunks       = []
                silence_cnt  = 0
                has_speech   = False
                speech_cnt   = 0
                interim_busy = [False]

                try:
                    with sd.InputStream(samplerate=RATE, channels=1,
                                        dtype='float32', blocksize=CHUNK) as stream:
                        while not self._stop_evt.is_set() and self._ready_evt.is_set():
                            data, _ = stream.read(CHUNK)
                            chunks.append(data.copy())
                            rms = float(np.sqrt(np.mean(data ** 2)))
                            if on_level:
                                on_level(rms)
                            if rms > threshold:
                                has_speech  = True
                                silence_cnt = 0
                                speech_cnt += 1
                                if speech_cnt % INTERIM_EVERY == 0 and not interim_busy[0]:
                                    interim_busy[0] = True
                                    snap = list(chunks)
                                    def _do_interim(snap=snap):
                                        try:
                                            audio = np.concatenate(snap).flatten()
                                            segs, _ = self._model.transcribe(
                                                audio, language="en", beam_size=1)
                                            txt = " ".join(s.text.strip() for s in segs).strip()
                                            if txt:
                                                on_interim(txt)
                                        except Exception:
                                            pass
                                        finally:
                                            interim_busy[0] = False
                                    threading.Thread(target=_do_interim, daemon=True).start()
                            elif has_speech:
                                silence_cnt += 1
                                if silence_cnt >= SILENCE_LIMIT:
                                    break
                except Exception as exc:
                    on_error(str(exc))
                    time.sleep(0.5)
                    continue

                if self._stop_evt.is_set():
                    break
                if not has_speech:
                    continue

                self._ready_evt.clear()
                on_status("Transcribing…")
                try:
                    audio = np.concatenate(chunks).flatten()
                    segs, _ = self._model.transcribe(audio, language="en", beam_size=5)
                    text = " ".join(s.text.strip() for s in segs).strip()
                    if text:
                        on_send(text)
                    else:
                        on_status("Didn't catch that")
                        self._ready_evt.set()
                except Exception as exc:
                    on_error(str(exc))
                    self._ready_evt.set()

        threading.Thread(target=run, daemon=True).start()

    def pause(self):
        self._ready_evt.clear()

    def resume(self):
        self._ready_evt.set()

    def stop(self):
        self._stop_evt.set()
        self._ready_evt.set()  # unblock any .wait()


# ── App ───────────────────────────────────────────────────────────────────────

class App:
    # Near-black minimalist palette
    BG      = "#111111"
    SURFACE = "#1a1a1a"
    BORDER  = "#272727"
    FG      = "#ebebeb"
    FG_DIM  = "#aaaaaa"
    MUTED   = "#555555"
    ACCENT  = "#ffffff"
    RED     = "#c0392b"
    YELLOW  = "#c8a84b"

    ENGINE_SAPI   = "Windows SAPI"
    ENGINE_KOKORO = "Kokoro (Neural)"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Ollama Voice")
        self.root.geometry("860x680")
        self.root.configure(bg=self.BG)
        self.root.minsize(640, 480)

        self.history: list[dict] = []
        self.sapi   = SAPIEngine(self._on_sapi_voices_ready)
        self.kokoro = KokoroEngine()
        self.stt    = STTEngine()
        self._active_engine     = self.ENGINE_KOKORO
        self._mic_active        = False
        self._stt_ready         = False
        self._settings_visible  = False
        self._sysprompt_visible = False
        self._saved_settings    = self._load_settings()
        self._wave_levels       = deque(maxlen=120)
        self._wave_paused       = False
        self._wave_animating    = False

        self._build_ui()
        self._toggle_settings()
        self._apply_saved_settings(self._saved_settings)
        self.root.after(300, self._first_run_popup)
        self._on_engine_change()
        self.rate_var.trace_add('write', lambda *_: self._on_slider_change())
        self.vol_var.trace_add('write',  lambda *_: self._on_slider_change())
        self.voice_cb.bind('<<ComboboxSelected>>', lambda _: self._save_settings())
        self.sys_prompt.bind('<FocusOut>', lambda _: self._save_settings())
        self._load_ollama_models()
        self.root.bind('<Escape>', lambda _: self._on_escape())

    # ── UI ───────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self._style_ttk()

        # ── Status bar (bottom-most, pack first) ──────────────────────────────
        self.status_var = tk.StringVar(value="Starting up…")
        self._status_lbl = tk.Label(
            self.root, textvariable=self.status_var,
            bg=self.BG, fg=self.MUTED, font=("Segoe UI", 9),
            anchor='w', padx=20, pady=6,
        )
        self._status_lbl.pack(fill=tk.X, side=tk.BOTTOM)

        # Waveform canvas – packed/unpacked dynamically when mic is active
        self._wave_canvas = tk.Canvas(
            self.root, bg=self.BG, height=36, highlightthickness=0)

        # ── Input area (side=BOTTOM, above status) ────────────────────────────
        self._inp_sep = tk.Frame(self.root, bg=self.BORDER, height=1)
        self._inp_sep.pack(fill=tk.X, side=tk.BOTTOM)

        inp_wrap = tk.Frame(self.root, bg=self.BG, padx=20, pady=14)
        inp_wrap.pack(fill=tk.X, side=tk.BOTTOM)

        # Bordered text input
        inp_border = tk.Frame(inp_wrap, bg=self.BORDER, padx=1, pady=1)
        inp_border.pack(fill=tk.X, pady=(0, 10))

        self.input = tk.Text(
            inp_border, height=3, bg=self.SURFACE, fg=self.FG,
            font=("Segoe UI", 11), relief=tk.FLAT,
            padx=14, pady=10, insertbackground=self.FG, wrap=tk.WORD,
        )
        self.input.pack(fill=tk.X)
        self.input.bind('<Return>', self._on_enter)

        # Action row: [whisper + mic]  ···  [clear · stop · Send →]
        action = tk.Frame(inp_wrap, bg=self.BG)
        action.pack(fill=tk.X)

        left = tk.Frame(action, bg=self.BG)
        left.pack(side=tk.LEFT)

        _wlbl = tk.Label(left, text="Whisper", bg=self.BG, fg=self.MUTED,
                         font=("Segoe UI", 9))
        _wlbl.pack(side=tk.LEFT, padx=(0, 6))
        self.stt_model_var = tk.StringVar(value="base")
        stt_cb = ttk.Combobox(left, textvariable=self.stt_model_var,
                               values=STT_MODELS, width=7, state='readonly')
        stt_cb.pack(side=tk.LEFT, padx=(0, 14))
        stt_cb.bind('<<ComboboxSelected>>', self._on_stt_model_change)
        _whisper_tip = (
            "Whisper is an offline speech-to-text model that transcribes your voice.\n\n"
            "tiny — fastest, least accurate\n"
            "base — good balance (recommended)\n"
            "small — more accurate, slower to load\n"
            "medium — most accurate, slowest"
        )
        Tooltip(_wlbl, _whisper_tip)
        Tooltip(stt_cb, _whisper_tip)

        self.mic_btn = tk.Button(
            left, text="⏺  Mic", command=self._on_mic_click,
            bg=self.BG, fg=self.MUTED, font=("Segoe UI", 10),
            relief=tk.FLAT, padx=6, pady=2, cursor='hand2',
            activebackground=self.BG, activeforeground=self.FG,
        )
        self.mic_btn.pack(side=tk.LEFT)

        right = tk.Frame(action, bg=self.BG)
        right.pack(side=tk.RIGHT)

        for label, cmd, color in [
            ("Clear",   self._clear_chat, self.MUTED),
            ("Stop",    self.stop_tts,    self.RED),
        ]:
            tk.Button(
                right, text=label, command=cmd,
                bg=self.BG, fg=color, font=("Segoe UI", 10),
                relief=tk.FLAT, padx=8, pady=2, cursor='hand2',
                activebackground=self.BG, activeforeground=self.FG,
            ).pack(side=tk.LEFT, padx=(0, 8))

        self.send_btn = tk.Button(
            right, text="Send  →", command=self.send,
            bg=self.ACCENT, fg=self.BG, font=("Segoe UI", 10, 'bold'),
            relief=tk.FLAT, padx=14, pady=4, cursor='hand2',
        )
        self.send_btn.pack(side=tk.LEFT)

        # ── Header (top) ──────────────────────────────────────────────────────
        header = tk.Frame(self.root, bg=self.SURFACE, height=48)
        header.pack(fill=tk.X, side=tk.TOP)
        header.pack_propagate(False)

        tk.Label(
            header, text="Ollama Voice",
            bg=self.SURFACE, fg=self.FG,
            font=("Segoe UI", 12, 'bold'),
        ).pack(side=tk.LEFT, padx=20)

        # Right side of header
        hdr_right = tk.Frame(header, bg=self.SURFACE)
        hdr_right.pack(side=tk.RIGHT, padx=12)

        self._gear_btn = tk.Button(
            hdr_right, text="⚙", command=self._toggle_settings,
            bg=self.SURFACE, fg=self.MUTED, font=("Segoe UI", 12),
            relief=tk.FLAT, padx=8, pady=6, cursor='hand2',
            activebackground=self.SURFACE, activeforeground=self.FG,
        )
        self._gear_btn.pack(side=tk.RIGHT, padx=(6, 0))

        self.voice_var = tk.StringVar()
        self.voice_cb  = ttk.Combobox(
            hdr_right, textvariable=self.voice_var,
            width=24, state='readonly',
        )
        self.voice_cb.pack(side=tk.RIGHT, pady=9, padx=(0, 10))

        self.model_var = tk.StringVar()
        self.model_cb  = ttk.Combobox(
            hdr_right, textvariable=self.model_var,
            width=20, state='readonly',
        )
        self.model_cb.pack(side=tk.RIGHT, pady=9)

        self._header_sep = tk.Frame(self.root, bg=self.BORDER, height=1)
        self._header_sep.pack(fill=tk.X, side=tk.TOP)

        # ── Middle section: settings + system prompt (between header and chat) ─
        self._middle = tk.Frame(self.root, bg=self.BG)
        self._middle.pack(fill=tk.X, side=tk.TOP)

        # Settings panel (hidden by default)
        self._settings_panel = tk.Frame(self._middle, bg=self.SURFACE, pady=10)
        sp = self._settings_panel

        self._slim_lbl(sp, "Engine").grid(row=0, column=0, sticky='w', padx=(16, 4))
        self.engine_var = tk.StringVar(value=self.ENGINE_KOKORO)
        self.engine_cb  = ttk.Combobox(sp, textvariable=self.engine_var,
                                        values=[self.ENGINE_SAPI, self.ENGINE_KOKORO],
                                        width=16, state='readonly')
        self.engine_cb.grid(row=0, column=1, padx=(0, 20))
        self.engine_cb.bind('<<ComboboxSelected>>', self._on_engine_change)

        self._slim_lbl(sp, "Speed").grid(row=0, column=2, sticky='w', padx=(0, 4))
        self.rate_var = tk.IntVar(value=175)
        ttk.Scale(sp, from_=60, to=350, variable=self.rate_var,
                  length=110).grid(row=0, column=3, padx=(0, 20))

        self._slim_lbl(sp, "Volume").grid(row=0, column=4, sticky='w', padx=(0, 4))
        self.vol_var = tk.IntVar(value=100)
        ttk.Scale(sp, from_=0, to=100, variable=self.vol_var,
                  length=110).grid(row=0, column=5, padx=(0, 20))

        self.kokoro_btn = tk.Button(
            sp, text="Setup Kokoro…", command=self._kokoro_setup,
            bg=self.BG, fg=self.YELLOW, font=("Segoe UI", 9),
            relief=tk.FLAT, padx=10, pady=4, cursor='hand2',
        )
        # placed at column=6 by _on_engine_change when needed

        tk.Frame(sp, bg=self.BORDER, height=1).grid(
            row=1, column=0, columnspan=7, sticky='ew', pady=(10, 0))

        # System prompt toggle (always visible in _middle)
        self._sp_toggle_row = sp_toggle_row = tk.Frame(self._middle, bg=self.BG)
        sp_toggle_row.pack(fill=tk.X)

        self._sp_btn = tk.Button(
            sp_toggle_row, text="▸  System prompt",
            command=self._toggle_sysprompt,
            bg=self.BG, fg=self.MUTED, font=("Segoe UI", 9),
            relief=tk.FLAT, padx=20, pady=7, cursor='hand2',
            anchor='w', activebackground=self.BG, activeforeground=self.FG,
        )
        self._sp_btn.pack(fill=tk.X)

        # System prompt panel (hidden by default)
        self._sysprompt_panel = tk.Frame(self._middle, bg=self.BG)
        self.sys_prompt = tk.Text(
            self._sysprompt_panel, height=2, bg=self.SURFACE, fg=self.FG_DIM,
            font=("Segoe UI", 10), relief=tk.FLAT,
            padx=20, pady=10, insertbackground=self.FG, wrap=tk.WORD, border=0,
        )
        self.sys_prompt.pack(fill=tk.X)
        self.sys_prompt.insert(
            '1.0', "You are a helpful assistant that speaks in a very casual tone. Please provide good context like a normal conversation in a clear succinct manner, maximum of 5 sentences.")
        tk.Frame(self._sysprompt_panel, bg=self.BORDER, height=1).pack(fill=tk.X)

        # Separator below middle section (always visible)
        tk.Frame(self.root, bg=self.BORDER, height=1).pack(fill=tk.X, side=tk.TOP)

        # ── Chat display ──────────────────────────────────────────────────────
        self.chat = scrolledtext.ScrolledText(
            self.root, wrap=tk.WORD, bg=self.BG, fg=self.FG,
            font=("Segoe UI", 11), state=tk.DISABLED, relief=tk.FLAT,
            padx=60, pady=24, insertbackground=self.FG, border=0,
        )
        self.chat.pack(fill=tk.BOTH, expand=True)

        self.chat.tag_configure('you_label',
            foreground=self.ACCENT, font=("Segoe UI", 9, 'bold'))
        self.chat.tag_configure('you_text',
            foreground=self.FG, font=("Segoe UI", 11),
            spacing1=2, spacing3=10)
        self.chat.tag_configure('ai_label',
            foreground=self.MUTED, font=("Segoe UI", 9))
        self.chat.tag_configure('ai_text',
            foreground=self.FG_DIM, font=("Segoe UI", 11),
            spacing1=2, spacing3=16)

    def _style_ttk(self):
        s = ttk.Style()
        s.theme_use('clam')
        s.configure('TCombobox',
                    fieldbackground=self.SURFACE, background=self.SURFACE,
                    foreground=self.FG, selectbackground=self.BORDER,
                    arrowcolor=self.MUTED, bordercolor=self.BORDER)
        s.map('TCombobox', fieldbackground=[('readonly', self.SURFACE)])
        s.configure('TScale', background=self.SURFACE,
                    troughcolor=self.BORDER, sliderlength=14)

    def _slim_lbl(self, parent, text):
        return tk.Label(parent, text=text, bg=self.SURFACE, fg=self.MUTED,
                        font=("Segoe UI", 9))

    # ── Toggle panels ─────────────────────────────────────────────────────────

    def _toggle_settings(self):
        if self._settings_visible:
            self._settings_panel.pack_forget()
            self._settings_visible = False
            self._gear_btn.config(fg=self.MUTED)
        else:
            self._settings_panel.pack(fill=tk.X, before=self._sp_toggle_row)
            self._settings_visible = True
            self._gear_btn.config(fg=self.FG)

    def _toggle_sysprompt(self):
        if self._sysprompt_visible:
            self._sysprompt_panel.pack_forget()
            self._sysprompt_visible = False
            self._sp_btn.config(text="▸  System prompt")
        else:
            self._sysprompt_panel.pack(fill=tk.X)
            self._sysprompt_visible = True
            self._sp_btn.config(text="▾  System prompt")

    # ── Escape key ───────────────────────────────────────────────────────────

    def _on_escape(self):
        if self._mic_active:
            self._stop_continuous()
        else:
            self.stop_tts()

    # ── Engine switching ─────────────────────────────────────────────────────

    def _on_engine_change(self, _=None):
        engine = self.engine_var.get()
        self._active_engine = engine
        self.stop_tts()

        if engine == self.ENGINE_SAPI:
            self.kokoro_btn.grid_remove()
            self._populate_sapi_voices()

        elif engine == self.ENGINE_KOKORO:
            if not KokoroEngine.package_installed():
                self.kokoro_btn.config(text="Install kokoro-onnx…")
                self.kokoro_btn.grid(row=0, column=6)
                self._status("Kokoro not installed — click 'Install kokoro-onnx…'")
                self.voice_cb['values'] = []
                if _ is not None:
                    messagebox.showwarning(
                        "Kokoro not installed",
                        "The kokoro-onnx and sounddevice packages are not installed.\n\n"
                        "Click 'Install kokoro-onnx…' in the settings panel to install them automatically, "
                        "or run:\n\n    pip install kokoro-onnx sounddevice",
                        parent=self.root,
                    )
            elif not KokoroEngine.models_present():
                self.kokoro_btn.config(text="Download models (330 MB)…")
                self.kokoro_btn.grid(row=0, column=6)
                self._status("Kokoro models missing — click 'Download models…'")
                self.voice_cb['values'] = []
            else:
                self.kokoro_btn.grid_remove()
                self._load_kokoro()

        self._save_settings()

    def _load_kokoro(self):
        self._status("Loading Kokoro…")
        def run():
            try:
                self.kokoro.load()
                names = [label for _, label in KOKORO_VOICES]
                self.root.after(0, lambda: self._populate_kokoro_voices(names))
                self.root.after(0, lambda: self._status("Kokoro ready"))
            except Exception as exc:
                self.root.after(0, lambda: self._status(f"Kokoro error: {exc}"))
        threading.Thread(target=run, daemon=True).start()

    def _kokoro_setup(self):
        if not KokoroEngine.package_installed():
            import subprocess, sys
            self._status("Installing kokoro-onnx and sounddevice…", highlight=True)
            def run():
                try:
                    subprocess.check_call(
                        [sys.executable, "-m", "pip", "install",
                         "kokoro-onnx", "sounddevice"],
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                    )
                    self.root.after(0, self._on_engine_change)
                except Exception as exc:
                    self.root.after(0, lambda: self._status(f"Install failed: {exc}"))
            threading.Thread(target=run, daemon=True).start()
        else:
            self.kokoro.download_models(
                on_progress=lambda m: self.root.after(
                    0, lambda msg=m: self._status(msg, highlight=True)),
                on_done=lambda: self.root.after(0, self._on_engine_change),
                on_error=lambda e: self.root.after(
                    0, lambda: self._status(f"Download error: {e}")),
            )

    # ── Voice dropdowns ───────────────────────────────────────────────────────

    def _on_sapi_voices_ready(self, names: list[str]):
        self._sapi_voice_names = names
        self.root.after(0, self._populate_sapi_voices)
        self.root.after(0, lambda: self._status("Ready"))
        self.root.after(200, self._load_stt_model)

    def _populate_sapi_voices(self):
        if self._active_engine != self.ENGINE_SAPI:
            return
        names = getattr(self, '_sapi_voice_names', [])
        self.voice_cb['values'] = names
        if names:
            saved = self._saved_settings.get('voice', '')
            self.voice_cb.current(names.index(saved) if saved in names else 0)

    def _populate_kokoro_voices(self, names: list[str]):
        self.voice_cb['values'] = names
        if names:
            saved = self._saved_settings.get('voice', '')
            self.voice_cb.current(names.index(saved) if saved in names else 0)

    # ── STT ──────────────────────────────────────────────────────────────────

    def _load_stt_model(self):
        if not STTEngine.package_installed():
            self.mic_btn.config(state=tk.DISABLED)
            return
        size = self.stt_model_var.get()
        self._stt_ready = False
        self._status(f"Loading Whisper ({size})…")
        self.mic_btn.config(state=tk.DISABLED)
        self.stt.load(
            size,
            on_done=lambda: self.root.after(0, self._on_stt_ready),
            on_error=lambda e: self.root.after(
                0, lambda: self._status(f"Whisper error: {e}")),
        )

    def _on_stt_ready(self):
        self._stt_ready = True
        self.mic_btn.config(state=tk.NORMAL)
        self._status("Ready")

    def _on_stt_model_change(self, _=None):
        self._load_stt_model()
        self._save_settings()

    def _on_mic_click(self):
        if self._mic_active:
            self._stop_continuous()
        else:
            self._start_continuous()

    def _start_continuous(self):
        if not self._stt_ready:
            self._status("Whisper not ready yet")
            return
        self._mic_active = True
        self._wave_paused = False
        self._wave_levels.clear()
        self._wave_canvas.pack(fill=tk.X, side=tk.BOTTOM, before=self._inp_sep)
        self.mic_btn.config(text="⏹  Stop Mic", fg=self.RED, activeforeground=self.RED)
        self.stop_tts()
        self.stt.start_continuous(
            on_interim=lambda t: self.root.after(0, lambda txt=t: self._on_interim(txt)),
            on_send=lambda t: self.root.after(0, lambda txt=t: self._on_send(txt)),
            on_status=lambda s: self.root.after(0, lambda msg=s: self._status(msg)),
            on_error=lambda e: self.root.after(0, lambda: self._on_mic_error(e)),
            on_level=lambda lvl: self._wave_levels.append(lvl),
        )
        if not self._wave_animating:
            self._wave_animating = True
            self.root.after(50, self._wave_draw)

    def _stop_continuous(self):
        self.stt.stop()
        self._mic_active = False
        self._wave_canvas.pack_forget()
        self.mic_btn.config(text="⏺  Mic", fg=self.MUTED, activeforeground=self.FG)
        self._status("Ready")

    def _on_interim(self, text: str):
        self._status(f"Hearing: {text[:80]}")

    def _on_send(self, text: str):
        self._wave_paused = True
        self.history.append({'role': 'user', 'content': text})
        self._chat_append("\nYou (voice)\n", 'you_label')
        self._chat_append(f"{text}\n",        'you_text')
        self._chat_append("\nassistant\n",    'ai_label')
        self.send_btn.config(state=tk.DISABLED)
        self._status("Waiting for response…")

        sys_text = self.sys_prompt.get('1.0', tk.END).strip()
        messages = ([{'role': 'system', 'content': sys_text}] if sys_text else []) + self.history
        model = self.model_var.get()
        if not model:
            self._status("Pick a model first")
            if self._mic_active:
                self.stt.resume()
            return
        threading.Thread(target=self._stream, args=(model, messages), daemon=True).start()

    def _on_mic_error(self, err: str):
        self._stop_continuous()
        print(f"[Mic error] {err}")
        self._status(f"Mic error: {err[:80]}")

    # ── Ollama ────────────────────────────────────────────────────────────────

    def _load_ollama_models(self):
        def fetch():
            try:
                names = [m.model for m in ollama.list().models]
                self.root.after(0, lambda: self._set_ollama_models(names))
            except Exception as exc:
                self.root.after(0, lambda: self._status(f"Ollama error: {exc}"))
        threading.Thread(target=fetch, daemon=True).start()

    def _set_ollama_models(self, names):
        self.model_cb['values'] = names
        if names:
            self.model_cb.current(0)

    def _on_enter(self, event):
        if not (event.state & 0x1):
            self.send()
            return 'break'

    def send(self):
        text = self.input.get('1.0', tk.END).strip()
        if not text:
            return
        model = self.model_var.get()
        if not model:
            self._status("Pick a model first")
            return

        self.input.delete('1.0', tk.END)
        self.history.append({'role': 'user', 'content': text})

        self._chat_append("\nYou\n",        'you_label')
        self._chat_append(f"{text}\n",      'you_text')
        self._chat_append("\nassistant\n",  'ai_label')

        self.send_btn.config(state=tk.DISABLED)
        self._status("Waiting for response…")
        self.stop_tts()

        sys_text = self.sys_prompt.get('1.0', tk.END).strip()
        messages = ([{'role': 'system', 'content': sys_text}] if sys_text else []) + self.history
        threading.Thread(target=self._stream, args=(model, messages), daemon=True).start()

    def _stream(self, model, messages):
        full = ""
        try:
            for chunk in ollama.chat(model=model, messages=messages, stream=True):
                token = chunk.message.content or ''
                full += token
                self.root.after(0, lambda t=token: self._chat_append(t, 'ai_text'))
            self.history.append({'role': 'assistant', 'content': full})
            self.root.after(0, lambda: self._chat_append("\n", 'ai_text'))
            self.root.after(0, lambda: self.send_btn.config(state=tk.NORMAL))
            self.root.after(0, lambda: self._status("Speaking…"))
            self.root.after(0, lambda: self._speak(full))
        except Exception as exc:
            self.root.after(0, lambda: self._status(f"Error: {exc}"))
            self.root.after(0, lambda: self.send_btn.config(state=tk.NORMAL))
            if self._mic_active:
                self.root.after(0, self.stt.resume)

    # ── TTS dispatch ──────────────────────────────────────────────────────────

    def _speak(self, text):
        clean = self._clean(text)
        idx   = self.voice_cb.current()

        def on_done():
            if self._mic_active:
                self._wave_paused = False
                self.root.after(0, lambda: self._status("Listening…"))
                self.stt.resume()
            else:
                self.root.after(0, lambda: self._status("Ready"))

        if self._active_engine == self.ENGINE_SAPI:
            if idx < 0:
                on_done()
                return
            self.sapi.speak(clean, voice_idx=idx,
                            rate=self._sapi_rate(self.rate_var.get()),
                            volume=self.vol_var.get(), on_done=on_done)
        elif self._active_engine == self.ENGINE_KOKORO:
            if idx < 0 or self.kokoro._kokoro is None:
                self._status("Kokoro not ready")
                on_done()
                return
            self.kokoro.speak(clean, voice_code=KOKORO_VOICES[idx][0],
                              get_speed=lambda: self.rate_var.get() / 175,
                              get_volume=lambda: self.vol_var.get() / 100,
                              on_done=on_done)

    def stop_tts(self):
        if self._active_engine == self.ENGINE_SAPI:
            self.sapi.stop()
        else:
            self.kokoro.stop()
        if self._mic_active:
            self.stt.resume()
            self._status("Listening…")
        else:
            self._status("Ready")

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _on_slider_change(self):
        self._save_settings()
        if self._active_engine == self.ENGINE_SAPI:
            self.sapi.update_live(
                rate=self._sapi_rate(self.rate_var.get()),
                volume=self.vol_var.get(),
            )

    @staticmethod
    def _sapi_rate(wpm: int) -> int:
        return max(-10, min(10, round((wpm - 175) / 17.5)))

    @staticmethod
    def _clean(text: str) -> str:
        text = re.sub(r'```[\s\S]*?```', ' ', text)
        text = re.sub(r'`[^`]*`', ' ', text)
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        text = re.sub(r'\*(.+?)\*', r'\1', text)
        text = re.sub(r'#{1,6}\s*', '', text)
        text = re.sub(r'\[(.+?)\]\(.+?\)', r'\1', text)
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r'[—–]', ', ', text)       # em/en dashes → brief comma pause
        text = re.sub(r'\.{2,}|…', '.', text)    # ellipses → single period
        text = re.sub(r';', ',', text)            # semicolons → lighter comma pause
        text = re.sub(r',(\s*,)+', ',', text)    # collapse repeated commas
        text = re.sub(r' {2,}', ' ', text)       # collapse extra spaces
        return text.strip()

    def _chat_append(self, text, tag):
        self.chat.config(state=tk.NORMAL)
        self.chat.insert(tk.END, text, tag)
        self.chat.see(tk.END)
        self.chat.config(state=tk.DISABLED)

    def _clear_chat(self):
        self.chat.config(state=tk.NORMAL)
        self.chat.delete('1.0', tk.END)
        self.chat.config(state=tk.DISABLED)
        self.history.clear()
        self.stop_tts()

    # ── Waveform ──────────────────────────────────────────────────────────────

    def _wave_draw(self):
        if not self._mic_active:
            self._wave_animating = False
            return

        c = self._wave_canvas
        c.delete('all')
        w = c.winfo_width()
        h = c.winfo_height()

        if w <= 1 or h <= 1:
            self.root.after(50, self._wave_draw)
            return

        mid = h // 2

        if self._wave_paused or not self._wave_levels:
            c.create_line(12, mid, w - 12, mid,
                          fill=self.MUTED, width=1, dash=(4, 6))
        else:
            bar_w = 3
            gap   = 2
            step  = bar_w + gap
            n_fit = max(1, (w - 24) // step)
            data  = list(self._wave_levels)[-n_fit:]
            x     = 12
            for lvl in data:
                bar_h = max(1, min(int(lvl * 160), mid - 2))
                fill  = self.FG if lvl > 0.08 else self.FG_DIM
                c.create_rectangle(x, mid - bar_h, x + bar_w, mid + bar_h,
                                   fill=fill, outline='')
                x += step

        self.root.after(50, self._wave_draw)

    # ── Settings persistence ──────────────────────────────────────────────────

    def _load_settings(self) -> dict:
        try:
            if SETTINGS_FILE.exists():
                return json.loads(SETTINGS_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
        return {}

    def _save_settings(self):
        try:
            SETTINGS_FILE.write_text(json.dumps({
                'engine':         self.engine_var.get(),
                'voice':          self.voice_var.get(),
                'rate':           self.rate_var.get(),
                'volume':         self.vol_var.get(),
                'stt_model':      self.stt_model_var.get(),
                'sys_prompt':     self.sys_prompt.get('1.0', tk.END).strip(),
                'first_run_done': self._saved_settings.get('first_run_done', False),
            }, indent=2), encoding='utf-8')
        except Exception:
            pass

    def _first_run_popup(self):
        if self._saved_settings.get('first_run_done'):
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("Welcome to Ollama Voice")
        dlg.configure(bg=self.SURFACE)
        dlg.resizable(False, False)
        dlg.grab_set()
        dlg.lift()
        dlg.focus_force()

        dlg.update_idletasks()
        w, h = 440, 300
        x = self.root.winfo_x() + (self.root.winfo_width()  - w) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - h) // 2
        dlg.geometry(f"{w}x{h}+{x}+{y}")

        tk.Label(dlg, text="Before you start",
                 bg=self.SURFACE, fg=self.FG,
                 font=("Segoe UI", 13, 'bold')).pack(pady=(24, 6))

        tk.Label(dlg, text="Make sure these are installed on your system:",
                 bg=self.SURFACE, fg=self.FG_DIM,
                 font=("Segoe UI", 10)).pack(pady=(0, 16))

        for label, url, display in [
            ("Python 3.9+", "https://www.python.org/downloads/", "python.org/downloads"),
            ("Ollama",      "https://ollama.com",                "ollama.com"),
        ]:
            row = tk.Frame(dlg, bg=self.SURFACE)
            row.pack(fill=tk.X, padx=48, pady=4)
            tk.Label(row, text=label, bg=self.SURFACE, fg=self.FG,
                     font=("Segoe UI", 10, 'bold'), width=12,
                     anchor='w').pack(side=tk.LEFT)
            lnk = tk.Label(row, text=display, bg=self.SURFACE, fg="#4a9eff",
                           font=("Segoe UI", 10, 'underline'), cursor='hand2')
            lnk.pack(side=tk.LEFT)
            lnk.bind('<Button-1>', lambda _, u=url: webbrowser.open(u))

        tk.Label(dlg,
                 text="Python packages install automatically when you run run.bat.",
                 bg=self.SURFACE, fg=self.MUTED,
                 font=("Segoe UI", 9)).pack(pady=(20, 0))

        def dismiss():
            self._saved_settings['first_run_done'] = True
            self._save_settings()
            dlg.destroy()

        tk.Button(dlg, text="Got it", command=dismiss,
                  bg=self.ACCENT, fg=self.BG,
                  font=("Segoe UI", 10, 'bold'),
                  relief=tk.FLAT, padx=24, pady=6,
                  cursor='hand2').pack(pady=20)

    def _apply_saved_settings(self, saved: dict):
        if 'rate' in saved:
            self.rate_var.set(saved['rate'])
        if 'volume' in saved:
            self.vol_var.set(saved['volume'])
        if 'stt_model' in saved:
            self.stt_model_var.set(saved['stt_model'])
        if 'sys_prompt' in saved:
            self.sys_prompt.delete('1.0', tk.END)
            self.sys_prompt.insert('1.0', saved['sys_prompt'])
        if 'engine' in saved:
            self.engine_var.set(saved['engine'])
            self._active_engine = saved['engine']

    def _status(self, msg: str, highlight: bool = False):
        self.status_var.set(msg)
        self._status_lbl.config(
            fg=self.FG if highlight else self.MUTED,
            font=("Segoe UI", 9, 'bold') if highlight else ("Segoe UI", 9),
        )


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
