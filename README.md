# Ollama Voice Reader

A local Windows desktop app that chats with Ollama and reads responses aloud. Supports neural TTS via Kokoro, microphone input with real-time transcription, and continuous hands-free conversation.

---

## Requirements

### Required

| Requirement | Notes |
|---|---|
| Windows 10/11 | SAPI TTS and mic input are Windows-only |
| Python 3.9+ | [python.org](https://www.python.org/downloads/) |
| [Ollama](https://ollama.com) | Must be installed and have at least one model pulled |
| At least one Ollama model | e.g. `ollama pull llama3` |

### Python packages (installed via pip)

| Package | Purpose |
|---|---|
| `ollama` | Talks to your local Ollama server |
| `pywin32` | Windows SAPI text-to-speech |
| `kokoro-onnx` | Neural TTS voices (optional but recommended) |
| `sounddevice` | Audio playback for Kokoro and mic input |
| `faster-whisper` | Offline speech-to-text for mic input |
| `numpy` | Audio processing |

### Kokoro model files (~330 MB, downloaded in-app)

Kokoro is the recommended voice engine. After installing the Python packages, the app will prompt you to download the model files the first time you select the Kokoro engine.

---

## Setup

**1. Install Ollama and pull a model**

Download Ollama from [ollama.com](https://ollama.com), then pull a model:

```bash
ollama pull llama3
```

**2. Install Python dependencies**

```bash
pip install -r requirements.txt
```

To also enable mic input (speech-to-text), install faster-whisper:

```bash
pip install faster-whisper
```

**3. Run the app**

Double-click `run.bat` — it checks whether Ollama is running and starts it automatically if not, then launches the app.

Or run manually:

```bash
python main.py
```

---

## How to Use

### Typing

Type your message in the input box and press **Enter** to send. Use **Shift+Enter** for a newline without sending.

### Microphone (hands-free conversation)

1. Click **Mic** to start listening. A waveform appears at the bottom while the mic is active.
2. Speak naturally. The app shows a live preview of what it's hearing in the status bar.
3. After **2 seconds of silence**, it automatically transcribes and sends your message.
4. The AI responds and speaks the reply aloud.
5. The mic resumes listening automatically after the response finishes — no need to click anything between turns.
6. Click **Stop Mic** (or press **Escape**) to end the mic session.

> The first time you use the mic, it downloads the Whisper speech-to-text model (choose the size in the Whisper dropdown — `base` is a good default).

### Kokoro Neural Voices (recommended)

1. Open settings (⚙ gear icon in the top-right).
2. Set **Engine** to `Kokoro (Neural)`.
3. If prompted, click **Download models (330 MB)** and wait for the download to finish.
4. Select a voice from the dropdown (15 voices across US and UK accents).

### Windows SAPI Voices (fallback)

Set **Engine** to `Windows SAPI`. Uses the voices installed on your system.

To add more Windows voices:

1. Open **Settings → Time & Language → Speech**
2. Under "Manage voices", click **Add voices**
3. Pick any language or accent

New voices appear in the Voice dropdown the next time you launch the app.

---

## Controls

| Control | Description |
|---|---|
| **Model** | Ollama model to use. Auto-populated from your installed models. |
| **Voice** | TTS voice. Options depend on the selected engine. |
| **⚙ Settings** | Expand to change engine, speed, and volume. |
| **System prompt** | Expand to edit the instructions sent to the model before your conversation. |
| **Mic** | Start continuous mic listening. Click again (or press Escape) to stop. |
| **Whisper** | Size of the speech-to-text model. Larger = more accurate, slower to load. |
| **Send / Enter** | Send the typed message. |
| **Stop** | Stop speech immediately. Mic resumes automatically if active. |
| **Clear** | Clear the chat history and start a new conversation. |

All settings (voice, engine, speed, volume, Whisper model, system prompt) are saved automatically and restored on next launch.

---

## Notes

- Everything runs locally — no internet connection is needed once Ollama, models, and Kokoro files are downloaded.
- Markdown is stripped from responses before they are read aloud, so the voice won't say "asterisk" or "backtick".
- Conversation history is kept for the session so the model has context across messages. Use **Clear** to start fresh.
- Press **Escape** at any time to stop speech or cancel the mic.
