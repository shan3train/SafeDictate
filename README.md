# SafeDictate

A Windows voice transcription app that lets you dictate text anywhere using a hotkey. Hold the hotkey, speak, release — your words are typed out instantly.

## Features

- Hold-to-record hotkey (default: `alt+1`)
- Powered by [faster-whisper](https://github.com/SYSTRAN/faster-whisper) (runs locally, no internet required)
- Switchable Whisper model sizes (tiny → large) without restarting
- Semi-transparent floating mini window, expands on hover to show settings
- Auto-detects microphone devices
- Saves settings to `config.ini`

## Requirements

- Windows 10/11
- [FFmpeg](https://ffmpeg.org/) in PATH or in a `bin/` folder next to the EXE
- Python 3.10+ (if running from source)

## Install (from source)

```bash
pip install -r requirements.txt
python transcription_v1.6.py
```

## Dependencies

See `requirements.txt`. Key packages: `faster-whisper`, `keyboard`, `pyperclip`.
