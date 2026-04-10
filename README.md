# SafeDictate

A Windows voice transcription app that lets you dictate text anywhere using a hotkey. Hold the hotkey, speak, release — your words are typed out instantly. No internet connection required.

## Features

- Hold-to-record hotkey (default: `alt+1`)
- Powered by [faster-whisper](https://github.com/SYSTRAN/faster-whisper) — runs fully locally
- Switchable Whisper model sizes (tiny → large) without restarting
- Semi-transparent floating mini window, expands on hover to show settings
- Auto-detects microphone devices
- Saves settings to `config.ini`

## Requirements

- Windows 10/11
- That's it — Python, FFmpeg, and all dependencies are bundled in the EXE

## Usage

1. Download and extract the release
2. Run `SafeDictate_v1.6.exe`
3. Hold your hotkey and speak
4. Text is typed wherever your cursor is

Models are downloaded automatically on first use when you select them in the app.
