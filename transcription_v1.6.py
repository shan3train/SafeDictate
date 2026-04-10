import os
import sys
import time
import signal
import threading
import subprocess
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox
import keyboard
import configparser
import re

# --- Fix for OpenMP Error #15 ---
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'

# ---------------- config ----------------
CONFIG_FILE = "config.ini"

MODEL_OPTIONS = [
    ("tiny",   "Tiny (~40MB) - Fastest"),
    ("base",   "Base (~150MB) - Balanced"),
    ("small",  "Small (~500MB) - Good"),
    ("medium", "Medium (~1.5GB) - Better"),
    ("large",  "Large (~3GB) - Best"),
]

HOTKEY_OPTIONS = [
    "alt+1", "alt+2", "alt+3", "alt+space",
    "ctrl+1", "ctrl+2", "ctrl+3", "ctrl+space",
    "ctrl+shift+space", "ctrl+alt+space",
    "f1", "f2", "f3", "f4", "f5", "f6",
]

ALPHA_MINI = 0.65
ALPHA_FULL = 1.0


def load_settings():
    cfg = configparser.ConfigParser()
    cfg["SETTINGS"] = {
        "model_size": "base",
        "sample_rate": "44100",
        "channels": "1",
        "device": "Microphone (USB Condenser Microphone)",
        "hotkey": "alt+1",
        "max_record_seconds": "30",
    }
    if os.path.exists(CONFIG_FILE):
        cfg.read(CONFIG_FILE)
    s = cfg["SETTINGS"]

    device_raw = s.get("device", "default").strip()
    mic_name = "default" if device_raw.isdigit() or not device_raw else device_raw

    return {
        "MODEL_SIZE":        s.get("model_size", "small"),
        "SAMPLE_RATE":       int(s.get("sample_rate", "44100")),
        "CHANNELS":          int(s.get("channels", "1")),
        "MIC_NAME":          mic_name,
        "HOTKEY":            s.get("hotkey", "alt+1"),
        "MAX_RECORD_SECONDS": int(s.get("max_record_seconds", "30")),
    }


def save_settings(settings):
    cfg = configparser.ConfigParser()
    cfg["SETTINGS"] = {
        "model_size":        settings["MODEL_SIZE"],
        "sample_rate":       str(settings["SAMPLE_RATE"]),
        "channels":          str(settings["CHANNELS"]),
        "device":            settings["MIC_NAME"],
        "hotkey":            settings["HOTKEY"],
        "max_record_seconds": str(settings["MAX_RECORD_SECONDS"]),
    }
    with open(CONFIG_FILE, "w") as f:
        cfg.write(f)


S = load_settings()


# --------------- paths ---------------
def get_ffmpeg_path():
    possible = [
        "bin/ffmpeg.exe",
        os.path.join(
            os.path.dirname(sys.executable if getattr(sys, "frozen", False) else __file__),
            "bin", "ffmpeg.exe",
        ),
        "ffmpeg",
    ]
    for p in possible:
        if os.path.exists(p):
            return p
    return "ffmpeg"


def get_models_dir():
    """Anchored to EXE / script location so it works in a frozen build."""
    base = os.path.dirname(sys.executable if getattr(sys, "frozen", False)
                           else os.path.abspath(__file__))
    return os.path.join(base, "models")


FFMPEG = get_ffmpeg_path()
APPLY_FILTERS = True
FILTERS = "highpass=f=200,lowpass=f=3000,acompressor=level_in=0.5:ratio=4:attack=20:release=100"


def ffmpeg_input_args():
    device = f"audio={S['MIC_NAME']}" if S["MIC_NAME"] != "default" else "audio=default"
    return ["-f", "dshow", "-i", device]


def get_audio_devices():
    try:
        result = subprocess.run(
            [FFMPEG, "-list_devices", "true", "-f", "dshow", "-i", "dummy"],
            capture_output=True, text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        devices = ["default"]
        for line in result.stderr.split("\n"):
            if "[dshow @" in line and "audio" in line.lower():
                m = re.search(r'"([^"]*)"', line)
                if m and m.group(1) not in devices:
                    devices.append(m.group(1))
        return devices
    except Exception:
        return ["default"]


# --------------- app ----------------
class DictateApp:
    def __init__(self, root):
        self.root = root
        self.model = None
        self.is_recording = False
        self.flash_on = False
        self.record_start_ts = None
        self._hotkey_stop = threading.Event()
        self._hotkey_thread = None
        self._expanded = False
        self.available_devices = ["default"]

        self._build_ui()
        threading.Thread(target=self._load_devices, daemon=True).start()
        # Heavy import happens inside _load_model — GUI appears immediately
        threading.Thread(target=self._load_model, daemon=True).start()

    # ------------------------------------------------------------------ UI --

    def _build_ui(self):
        self.root.title("SafeDictate v1.6")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", ALPHA_MINI)

        main = ttk.Frame(self.root, padding=10)
        main.pack(expand=True, fill="both")
        self._main = main

        # ── Status (always visible) ──────────────────────────────────────────
        self._status_frame = ttk.LabelFrame(main, text="Status", padding=10)
        self._status_frame.pack(fill="x")

        self.hotkey_label = ttk.Label(
            self._status_frame,
            text=f"Hold {S['HOTKEY']} to dictate",
            font=("Segoe UI", 11), anchor="center",
        )
        self.hotkey_label.pack(pady=2)

        self.status = ttk.Label(
            self._status_frame, text="Loading model…",
            foreground="blue", anchor="center",
        )
        self.status.pack(pady=2)

        self.recording_label = ttk.Label(
            self._status_frame, text="",
            foreground="red", font=("Segoe UI", 10, "bold"), anchor="center",
        )
        self.recording_label.pack()

        self.timer_label = ttk.Label(
            self._status_frame, text="",
            font=("Segoe UI", 10), anchor="center",
        )
        self.timer_label.pack(pady=2)

        # ── Header (collapsible) ─────────────────────────────────────────────
        self._header_frame = ttk.Frame(main)
        ttk.Label(
            self._header_frame, text="SafeDictate",
            font=("Segoe UI", 14, "bold"),
        ).pack()
        ttk.Label(
            self._header_frame, text="Voice Transcription",
            font=("Segoe UI", 8), foreground="gray",
        ).pack()

        # ── Settings (collapsible) ───────────────────────────────────────────
        self._settings_frame = ttk.LabelFrame(main, text="Settings", padding=10)

        ttk.Label(self._settings_frame, text="Model Size:").grid(row=0, column=0, sticky="w", pady=2)
        self.model_var = tk.StringVar(value=S["MODEL_SIZE"])
        self.model_combo = ttk.Combobox(
            self._settings_frame, textvariable=self.model_var,
            width=25, state="readonly",
        )
        self.model_combo["values"] = [f"{sz} - {desc}" for sz, desc in MODEL_OPTIONS]
        for i, (sz, _) in enumerate(MODEL_OPTIONS):
            if sz == S["MODEL_SIZE"]:
                self.model_combo.current(i)
                break
        self.model_combo.grid(row=0, column=1, sticky="ew", padx=(10, 0), pady=2)
        self.model_combo.bind("<<ComboboxSelected>>", self._on_model_change)

        ttk.Label(self._settings_frame, text="Microphone:").grid(row=1, column=0, sticky="w", pady=2)
        self.mic_var = tk.StringVar(value=S["MIC_NAME"])
        self.mic_combo = ttk.Combobox(self._settings_frame, textvariable=self.mic_var, width=25)
        self.mic_combo.grid(row=1, column=1, sticky="ew", padx=(10, 0), pady=2)
        self.mic_combo.bind("<<ComboboxSelected>>", self._on_mic_change)
        self.mic_combo.bind("<FocusIn>", self._refresh_devices)

        ttk.Label(self._settings_frame, text="Hotkey:").grid(row=2, column=0, sticky="w", pady=2)
        self.hotkey_var = tk.StringVar(value=S["HOTKEY"])
        self.hotkey_combo = ttk.Combobox(self._settings_frame, textvariable=self.hotkey_var, width=25)
        self.hotkey_combo["values"] = HOTKEY_OPTIONS
        self.hotkey_combo.grid(row=2, column=1, sticky="ew", padx=(10, 0), pady=2)
        self.hotkey_combo.bind("<<ComboboxSelected>>", self._on_hotkey_change)
        self.hotkey_combo.bind("<KeyRelease>", self._on_hotkey_change)

        ttk.Label(self._settings_frame, text="Max Recording (sec):").grid(row=3, column=0, sticky="w", pady=2)
        self.max_time_var = tk.StringVar(value=str(S["MAX_RECORD_SECONDS"]))
        self.max_time_spin = ttk.Spinbox(
            self._settings_frame, from_=5, to=120,
            textvariable=self.max_time_var, width=25,
        )
        self.max_time_spin.grid(row=3, column=1, sticky="ew", padx=(10, 0), pady=2)
        self.max_time_spin.bind("<KeyRelease>", self._on_max_time_change)
        self._settings_frame.columnconfigure(1, weight=1)

        # ── Stats (collapsible) ──────────────────────────────────────────────
        self._stats_frame = ttk.LabelFrame(main, text="Statistics", padding=10)
        self.last_record = ttk.Label(self._stats_frame, text="Last recording: —", font=("Segoe UI", 8))
        self.last_record.pack(anchor="w")
        self.last_transcribe = ttk.Label(self._stats_frame, text="Last transcribe: —", font=("Segoe UI", 8))
        self.last_transcribe.pack(anchor="w")

        # ── Snap geometry from natural collapsed height ───────────────────────
        self.root.update_idletasks()
        mini_h = self.root.winfo_reqheight()
        self._collapsed_geo = f"224x{mini_h}"
        self._expanded_geo = "420x480"
        self.root.geometry(self._collapsed_geo)

        # ── Hover bindings ───────────────────────────────────────────────────
        self._bind_hover(self.root)

    def _bind_hover(self, widget):
        widget.bind("<Enter>", self._on_mouse_enter)
        widget.bind("<Leave>", self._on_mouse_leave)
        for child in widget.winfo_children():
            self._bind_hover(child)

    def _on_mouse_enter(self, event=None):
        if not self._expanded:
            self._expand()

    def _on_mouse_leave(self, event=None):
        if not self._expanded:
            return
        x, y = self.root.winfo_pointerxy()
        wx, wy = self.root.winfo_rootx(), self.root.winfo_rooty()
        ww, wh = self.root.winfo_width(), self.root.winfo_height()
        if not (wx <= x <= wx + ww and wy <= y <= wy + wh):
            self._collapse()

    def _expand(self):
        self._expanded = True
        # Insert header + settings above status, stats below
        self._header_frame.pack(fill="x", pady=(0, 10), before=self._status_frame)
        self._settings_frame.pack(fill="x", pady=(0, 10), before=self._status_frame)
        self._stats_frame.pack(fill="x", pady=(10, 0))
        self.root.geometry(self._expanded_geo)
        self.root.attributes("-alpha", ALPHA_FULL)

    def _collapse(self):
        self._expanded = False
        self._header_frame.pack_forget()
        self._settings_frame.pack_forget()
        self._stats_frame.pack_forget()
        self.root.geometry(self._collapsed_geo)
        self.root.attributes("-alpha", ALPHA_MINI)

    # ----------------------------------------------------------- devices --

    def _load_devices(self):
        self.available_devices = get_audio_devices()
        self.root.after(0, self._update_device_list)

    def _update_device_list(self):
        self.mic_combo["values"] = self.available_devices
        if S["MIC_NAME"] not in self.available_devices:
            self.mic_combo["values"] = self.available_devices + [S["MIC_NAME"]]

    def _refresh_devices(self, event=None):
        threading.Thread(target=self._load_devices, daemon=True).start()

    # ---------------------------------------------------------- settings --

    def _on_model_change(self, event=None):
        selected = self.model_combo.get()
        model_size = selected.split(" - ")[0]
        if model_size != S["MODEL_SIZE"]:
            S["MODEL_SIZE"] = model_size
            save_settings(S)
            threading.Thread(target=self._load_model, daemon=True).start()

    def _on_mic_change(self, event=None):
        new_mic = self.mic_var.get()
        if new_mic != S["MIC_NAME"]:
            S["MIC_NAME"] = new_mic
            save_settings(S)
            self._set_status("Microphone updated", "green")

    def _on_hotkey_change(self, event=None):
        new_hotkey = self.hotkey_var.get().strip()
        if new_hotkey and new_hotkey != S["HOTKEY"]:
            S["HOTKEY"] = new_hotkey
            save_settings(S)
            self.hotkey_label.config(text=f"Hold {new_hotkey} to dictate")
            self._restart_hotkey_listener()

    def _on_max_time_change(self, event=None):
        try:
            new_time = int(self.max_time_var.get())
            if new_time != S["MAX_RECORD_SECONDS"]:
                S["MAX_RECORD_SECONDS"] = new_time
                save_settings(S)
        except ValueError:
            pass

    # ------------------------------------------------------------ model --

    def _load_model(self):
        """Load (or hot-swap) the Whisper model. Safe to call from any thread."""
        try:
            self._set_status(f"Loading {S['MODEL_SIZE']} model…", "blue")
            self.root.after(0, lambda: self.model_combo.config(state="disabled"))

            old_model = self.model
            self.model = None  # disable recording while swapping

            from faster_whisper import WhisperModel  # deferred import for fast GUI startup

            models_dir = get_models_dir()
            os.makedirs(models_dir, exist_ok=True)

            new_model = WhisperModel(
                S["MODEL_SIZE"], device="cpu", compute_type="int8",
                download_root=models_dir,
            )
            self.model = new_model
            del old_model

            self._set_status("Ready.", "green")
            self.root.after(0, lambda: self.model_combo.config(state="readonly"))

            if self._hotkey_thread is None or not self._hotkey_thread.is_alive():
                self._hotkey_stop = threading.Event()
                self._hotkey_thread = threading.Thread(
                    target=self._hotkey_loop, args=(self._hotkey_stop,), daemon=True,
                )
                self._hotkey_thread.start()

        except Exception as e:
            self._set_status("Model load failed", "red")
            self.root.after(0, lambda: self.model_combo.config(state="readonly"))
            self.root.after(0, lambda: messagebox.showerror("Model Error", str(e)))

    # --------------------------------------------------------- hotkey --

    def _restart_hotkey_listener(self):
        """Stop old listener cleanly and start a fresh one."""
        self._hotkey_stop.set()
        self._hotkey_stop = threading.Event()
        self._hotkey_thread = threading.Thread(
            target=self._hotkey_loop, args=(self._hotkey_stop,), daemon=True,
        )
        self._hotkey_thread.start()

    def _hotkey_loop(self, stop_event):
        """Recovers from transient hook errors instead of dying."""
        print(f"Listening for hotkey: {S['HOTKEY']}")
        while not stop_event.is_set():
            try:
                keyboard.wait(S["HOTKEY"])
                if stop_event.is_set():
                    break
                if self.model and not self.is_recording:
                    threading.Thread(target=self._record_then_transcribe, daemon=True).start()
            except Exception:
                if stop_event.is_set():
                    break
                time.sleep(0.5)  # brief pause then retry — never permanently die

    # --------------------------------------------------------- status --

    def _set_status(self, msg, color="blue"):
        self.root.after(0, lambda: self.status.config(text=msg, foreground=color))

    def _set_last_record(self, secs):
        self.root.after(0, lambda: self.last_record.config(text=f"Last recording: {secs:.2f}s"))

    def _set_last_transcribe(self, secs):
        self.root.after(0, lambda: self.last_transcribe.config(text=f"Last transcribe: {secs:.2f}s"))

    # ------------------------------------------------------- recording --

    def _start_record_ui(self):
        self.root.after(0, lambda: self.last_transcribe.config(text="Last transcribe: —"))
        self._set_status("Recording…")
        self.flash_on = True
        self.record_start_ts = time.time()
        self._update_timer()

        def flasher():
            while self.flash_on:
                self.root.after(0, lambda: self.recording_label.config(text="🔴 Recording…"))
                time.sleep(0.45)
                if not self.flash_on:
                    break
                self.root.after(0, lambda: self.recording_label.config(text=""))
                time.sleep(0.45)

        threading.Thread(target=flasher, daemon=True).start()

    def _stop_record_ui(self):
        self.flash_on = False
        self.root.after(0, lambda: self.recording_label.config(text=""))
        self.root.after(0, lambda: self.timer_label.config(text=""))

    def _update_timer(self):
        if self.flash_on and self.record_start_ts is not None:
            elapsed = time.time() - self.record_start_ts
            self.root.after(0, lambda e=elapsed: self.timer_label.config(text=f"{e:.1f}s"))
            self.root.after(100, self._update_timer)

    def _record_then_transcribe(self):
        self.is_recording = True
        self._start_record_ui()

        wav_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp:
                wav_path = tmp.name

            cmd = [FFMPEG] + ffmpeg_input_args()
            if APPLY_FILTERS:
                cmd += ["-af", FILTERS]
            cmd += [
                "-rtbufsize", "32M",
                "-thread_queue_size", "32",
                "-ar", str(S["SAMPLE_RATE"]),
                "-ac", str(S["CHANNELS"]),
                "-c:a", "pcm_s16le",
                "-f", "wav",
                "-flush_packets", "1",
                "-y", wav_path,
            ]

            creationflags = 0
            if sys.platform.startswith("win"):
                creationflags = subprocess.CREATE_NEW_PROCESS_GROUP | subprocess.CREATE_NO_WINDOW

            proc = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags,
            )

            start_ts = time.time()
            while (time.time() - start_ts) < S["MAX_RECORD_SECONDS"]:
                if not keyboard.is_pressed(S["HOTKEY"]):
                    break
                time.sleep(0.02)

            try:
                if proc.poll() is None:
                    try:
                        proc.stdin.write(b"q\n")
                        proc.stdin.flush()
                        proc.stdin.close()
                    except Exception:
                        pass
                    try:
                        proc.wait(timeout=0.5)
                    except subprocess.TimeoutExpired:
                        proc.terminate()
                        try:
                            proc.wait(timeout=0.5)
                        except subprocess.TimeoutExpired:
                            proc.kill()
            except Exception:
                pass

            self._set_last_record(time.time() - start_ts)

            try:
                if os.path.getsize(wav_path) <= 44:
                    self._set_status("Empty/invalid capture", "red")
                    try: os.remove(wav_path)
                    except Exception: pass
                    self._stop_record_ui()
                    self.is_recording = False
                    return
            except Exception:
                self._set_status("Capture failed", "red")
                self._stop_record_ui()
                self.is_recording = False
                return

        except FileNotFoundError:
            self._set_status("FFmpeg not found", "red")
            messagebox.showerror(
                "FFmpeg", "FFmpeg not found. Install it and add to PATH, or set a full path in the script."
            )
        except Exception as e:
            self._set_status("Recording failed", "red")
            messagebox.showerror("Recording Error", str(e))
        finally:
            self._stop_record_ui()
            self.is_recording = False

        if not wav_path or not os.path.exists(wav_path):
            return

        # transcribe
        self._set_status("Transcribing…")
        t0 = time.time()
        text = ""
        try:
            segments, _ = self.model.transcribe(wav_path, language="en", beam_size=5)
            text = "".join(s.text for s in segments).strip()
        except Exception as e:
            self._set_status("Transcription failed", "red")
            messagebox.showerror("Transcription Error", str(e))
        finally:
            try:
                os.remove(wav_path)
            except Exception:
                pass

        self._set_last_transcribe(time.time() - t0)

        if text:
            print("TEXT:", text)
            self._set_status("Done.", "green")
            try:
                import pyperclip
                pyperclip.copy(text)
                keyboard.write(text)
            except ImportError:
                keyboard.write(text)
        else:
            self._set_status("No text", "red")


def main():
    try:
        ffmpeg_path = "bin/ffmpeg.exe" if os.path.exists("bin/ffmpeg.exe") else "ffmpeg"
        subprocess.run(
            [ffmpeg_path, "-version"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True,
        )
    except Exception:
        pass

    root = tk.Tk()
    app = DictateApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
