import os
import re
import io
import time
import queue
import threading
import subprocess
import webbrowser
import winreg
import glob
from datetime import datetime

import psutil
import pyautogui
import pandas as pd
import pygame
import numpy as np
import requests
import sounddevice as sd
import soundfile as sf
import keyboard
from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from groq import Groq
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# for our APIs
groq_client         = Groq(api_key=os.getenv("GROQ_API_KEY"))
ELEVENLABS_API_KEY  = os.getenv("ELEVENLABS_API_KEY")
ELEVENLABS_VOICE_ID = "pNInz6obpgDQGcFmaJgB"   # you can change this if you want, but in terms of free choices, this is what i chose.
ELEVENLABS_MODEL    = "eleven_turbo_v2"

# for our audio, we're using pygame
pygame.mixer.pre_init(frequency=44100, size=-16, channels=1, buffer=512)
pygame.mixer.init()
audio_lock = threading.Lock()

def speak(text: str):
    if not ELEVENLABS_API_KEY or ELEVENLABS_API_KEY == "your_elevenlabs_api_key_here":
        print(f"[JARVIS] {text}")
        return
    def _play():
        try:
            url     = f"https://api.elevenlabs.io/v1/text-to-speech/{ELEVENLABS_VOICE_ID}/stream"
            headers = {"xi-api-key": ELEVENLABS_API_KEY, "Content-Type": "application/json", "Accept": "audio/mpeg"}
            payload = {
                "text": text, "model_id": ELEVENLABS_MODEL,
                "voice_settings": {"stability": 0.55, "similarity_boost": 0.80,
                                   "style": 0.20, "use_speaker_boost": True},
            }
            r = requests.post(url, json=payload, headers=headers, stream=True, timeout=15)
            r.raise_for_status()
            buf = io.BytesIO(b"".join(r.iter_content(chunk_size=4096)))
            with audio_lock:
                pygame.mixer.music.stop()
                pygame.mixer.music.load(buf, "mp3")
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    pygame.time.wait(40)
        except Exception as e:
            print(f"[Audio error] {e}")
    threading.Thread(target=_play, daemon=True).start()


# a flawed smart app finder, might not work with every application
_app_cache: dict = {}

SEARCH_ROOTS = [
    os.environ.get("PROGRAMFILES",         r"C:\Program Files"),
    os.environ.get("PROGRAMFILES(X86)",    r"C:\Program Files (x86)"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs"),
    os.path.join(os.environ.get("APPDATA",       ""), "Programs"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Roblox"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Discord"),
    os.path.join(os.environ.get("APPDATA",       ""), "Spotify"),
    r"C:\Riot Games",
    r"C:\Games",
    r"D:\Games",
    r"D:\SteamLibrary",
    r"D:\Program Files",
    r"D:\Program Files (x86)",
]

KNOWN_EXE_TARGETS = {
    "discord":           ["discord.exe"],
    "roblox":            ["robloxplayerbeta.exe", "roblox.exe"],
    "roblox studio":     ["robloxstudiobeta.exe", "robloxstudio.exe"],
    "spotify":           ["spotify.exe"],
    "steam":             ["steam.exe"],
    "epic games":        ["epicgameslauncher.exe"],
    "epic":              ["epicgameslauncher.exe"],
    "valorant":          ["valorant.exe", "valorant-win64-shipping.exe"],
    "league of legends": ["league of legends.exe"],
    "league":            ["league of legends.exe"],
    "minecraft":         ["minecraft.exe", "minecraftlauncher.exe"],
    "fortnite":          ["fortniteclient-win64-shipping.exe", "fortnite.exe"],
    "obs":               ["obs64.exe", "obs.exe"],
    "obs studio":        ["obs64.exe", "obs.exe"],
    "blender":           ["blender.exe"],
    "unity hub":         ["unityhub.exe"],
    "unity":             ["unity.exe", "unityhub.exe"],
    "slack":             ["slack.exe"],
    "zoom":              ["zoom.exe"],
    "teams":             ["teams.exe"],
    "microsoft teams":   ["teams.exe"],
    "whatsapp":          ["whatsapp.exe"],
    "telegram":          ["telegram.exe"],
    "vlc":               ["vlc.exe"],
    "firefox":           ["firefox.exe"],
    "opera":             ["opera.exe"],
    "brave":             ["brave.exe"],
    "figma":             ["figma.exe"],
    "postman":           ["postman.exe"],
    "docker":            ["docker desktop.exe", "dockerdesktop.exe"],
    "notion":            ["notion.exe"],
    "cursor":            ["cursor.exe"],
    "winrar":            ["winrar.exe"],
    "7zip":              ["7zfm.exe"],
    "7-zip":             ["7zfm.exe"],
    "audacity":          ["audacity.exe"],
    "gimp":              ["gimp.exe"],
    "photoshop":         ["photoshop.exe"],
    "premiere":          ["adobe premiere pro.exe"],
    "after effects":     ["afterfx.exe"],
    "pycharm":           ["pycharm64.exe"],
    "webstorm":          ["webstorm64.exe"],
    "rider":             ["rider64.exe"],
    "datagrip":          ["datagrip64.exe"],
    "goland":            ["goland64.exe"],
}

BUILTIN_MAP = {
    "notepad":            "notepad.exe",
    "calculator":         "calc.exe",
    "calc":               "calc.exe",
    "chrome":             "start chrome",
    "google chrome":      "start chrome",
    "browser":            "start chrome",
    "edge":               "start msedge",
    "microsoft edge":     "start msedge",
    "explorer":           "explorer.exe",
    "file explorer":      "explorer.exe",
    "settings":           "start ms-settings:",
    "task manager":       "taskmgr.exe",
    "paint":              "mspaint.exe",
    "word":               "start winword",
    "excel":              "start excel",
    "powerpoint":         "start powerpnt",
    "outlook":            "start outlook",
    "terminal":           "start wt",
    "windows terminal":   "start wt",
    "cmd":                "start cmd",
    "command prompt":     "start cmd",
    "powershell":         "start powershell",
    "vscode":             "code",
    "visual studio code": "code",
    "snipping tool":      "SnippingTool.exe",
    "clock":              "start ms-clock:",
    "camera":             "start microsoft.windows.camera:",
    "store":              "start ms-windows-store:",
}

def _filesystem_search(exe_names: list) -> str | None:
    for root in SEARCH_ROOTS:
        if not root or not os.path.isdir(root):
            continue
        for exe_name in exe_names:
            results = glob.glob(os.path.join(root, "**", exe_name), recursive=True)
            if results:
                return results[0]
    return None

def _registry_search(keyword: str) -> str | None:
    hives = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    kw = keyword.lower()
    for hive, path in hives:
        try:
            with winreg.OpenKey(hive, path) as key:
                for i in range(winreg.QueryInfoKey(key)[0]):
                    try:
                        with winreg.OpenKey(key, winreg.EnumKey(key, i)) as sub:
                            try:
                                name, _ = winreg.QueryValueEx(sub, "DisplayName")
                                if kw in name.lower():
                                    try:
                                        icon, _ = winreg.QueryValueEx(sub, "DisplayIcon")
                                        exe = icon.split(",")[0].strip('"')
                                        if exe.lower().endswith(".exe") and os.path.exists(exe):
                                            return exe
                                    except FileNotFoundError:
                                        pass
                            except FileNotFoundError:
                                pass
                    except OSError:
                        pass
        except OSError:
            pass
    return None

def _start_menu_search(keyword: str) -> str | None:
    dirs = [
        os.path.join(os.environ.get("APPDATA", ""), r"Microsoft\Windows\Start Menu\Programs"),
        r"C:\ProgramData\Microsoft\Windows\Start Menu\Programs",
    ]
    kw = keyword.lower()
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for root, _, files in os.walk(d):
            for f in files:
                if f.lower().endswith(".lnk") and kw in f.lower():
                    return os.path.join(root, f)
    return None

def find_app(name: str) -> str | None:
    n = name.lower().strip()

    # this is for our builtin instants
    for key, cmd in BUILTIN_MAP.items():
        if key in n or n in key:
            return cmd

    # 2. cache, not really necessary. 
    if n in _app_cache:
        return _app_cache[n]

    # 3. this is for our KNOWN exe targets
    for key, exes in KNOWN_EXE_TARGETS.items():
        if key in n or n in key:
            path = _filesystem_search(exes)
            if path:
                _app_cache[n] = path
                return path

    # 4. calling registry
    path = _registry_search(n)
    if path:
        _app_cache[n] = path
        return path

    # 5. our start menu shortcut
    lnk = _start_menu_search(n)
    if lnk:
        _app_cache[n] = lnk
        return lnk

    # 6. this is for our file system search, and as previously described, it is flawed.
    generic = _filesystem_search([f"{n}.exe", f"{n.replace(' ', '')}.exe"])
    if generic:
        _app_cache[n] = generic
        return generic

    return None

def open_application(name: str) -> str:
    cmd = find_app(name)
    if not cmd:
        return f"I couldn't find '{name}'. Make sure it's installed, or give me the full path."
    try:
        if cmd.startswith("start "):
            os.system(cmd)
        elif cmd.lower().endswith(".lnk"):
            os.startfile(cmd)
        else:
            subprocess.Popen(cmd, shell=False)
        return f"Opening {name}."
    except Exception as e:
        return f"Found '{name}' but couldn't launch it: {e}"


# SOURCE ON COMPUTER CONTROL! DERIVED FROM ELIAT
def close_application(name: str) -> str:
    killed = []
    for proc in psutil.process_iter(['name', 'pid']):
        try:
            if name.lower() in proc.info['name'].lower():
                proc.kill()
                killed.append(proc.info['name'])
        except Exception:
            pass
    return f"Closed: {', '.join(killed)}." if killed else f"No process matching '{name}' found."

def take_screenshot() -> str:
    fname = f"jarvis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    path  = os.path.join(os.path.expanduser("~"), "Desktop", fname)
    pyautogui.screenshot(path)
    return f"Screenshot saved to Desktop as {fname}."

def get_system_info() -> str:
    cpu  = psutil.cpu_percent(interval=0.5)
    ram  = psutil.virtual_memory()
    disk = psutil.disk_usage('C:\\')
    boot = datetime.fromtimestamp(psutil.boot_time()).strftime("%Y-%m-%d %H:%M")
    return (f"CPU: {cpu}% | "
            f"RAM: {ram.percent}% ({round(ram.used/1e9,1)} GB / {round(ram.total/1e9,1)} GB) | "
            f"Disk C: {disk.percent}% used ({round(disk.used/1e9,1)} GB / {round(disk.total/1e9,1)} GB) | "
            f"Boot: {boot}")

def search_web(query: str) -> str:
    webbrowser.open(f"https://www.google.com/search?q={query.replace(' ', '+')}")
    return f"Opened Google search for: {query}"

def read_file(filepath: str) -> str:
    fp = filepath.strip().strip('"').strip("'")
    if not os.path.exists(fp):
        return f"File not found: {fp}"
    ext = os.path.splitext(fp)[1].lower()
    try:
        if ext == '.csv':
            df = pd.read_csv(fp)
            return (f"CSV — {df.shape[0]} rows × {df.shape[1]} cols\n"
                    f"Columns: {list(df.columns)}\n\n"
                    f"First 20 rows:\n{df.head(20).to_string()}\n\n"
                    f"Statistics:\n{df.describe().to_string()}")
        elif ext in ('.xlsx', '.xls'):
            df = pd.read_excel(fp)
            return (f"Excel — {df.shape[0]} rows × {df.shape[1]} cols\n\n"
                    f"First 20 rows:\n{df.head(20).to_string()}\n\n"
                    f"Statistics:\n{df.describe().to_string()}")
        else:
            with open(fp, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read(8000)
    except Exception as e:
        return f"Error reading file: {e}"

def list_apps() -> str:
    apps = sorted({p.info['name'] for p in psutil.process_iter(['name'])})
    return "Running processes:\n" + "\n".join(apps)

def set_volume(level_str: str) -> str:
    try:
        level = max(0, min(100, int(level_str)))
        script = f"""
Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices;
[Guid("5CDF2C82-841E-4546-9722-0CF74078229A"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IAudioEndpointVolume {{
    int _f1(); int _f2(); int _f3(); int _f4();
    int SetMasterVolumeLevelScalar(float fLevel, Guid ctx);
}}
[Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDeviceEnumerator {{
    int _f1(); int GetDefaultAudioEndpoint(int a, int b, out IMMDevice c);
}}
[Guid("D666063F-1587-4E43-81F1-B948E807363F"),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IMMDevice {{
    int Activate(ref Guid iid, int ctx, IntPtr p, [MarshalAs(UnmanagedType.IUnknown)] out object i);
}}
[ComImport,Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")] class MMDeviceEnumerator {{}}
'@
$e=[Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]"BCDE0395-E52F-467C-8E3D-C4579291692E"))
$d=$null; [void]$e.GetDefaultAudioEndpoint(0,0,[ref]$d)
$iid=[Guid]"5CDF2C82-841E-4546-9722-0CF74078229A"; $v=$null
[void]$d.Activate([ref]$iid,23,[IntPtr]::Zero,[ref]$v)
[void]$v.SetMasterVolumeLevelScalar({level/100.0},[Guid]::Empty)
"""
        subprocess.run(["powershell", "-Command", script], capture_output=True, timeout=6)
        return f"Volume set to {level}%."
    except Exception as e:
        return f"Volume adjustment attempted. ({e})"


# in case of an action dispatch
ACTION_RE = re.compile(r'\[([A-Z]+)(?::([^\]]*))?\]')

def dispatch_actions(text: str):
    results = []
    for m in ACTION_RE.finditer(text):
        cmd = m.group(1)
        arg = (m.group(2) or '').strip()
        if   cmd == 'OPEN':       results.append(open_application(arg))
        elif cmd == 'CLOSE':      results.append(close_application(arg))
        elif cmd == 'SCREENSHOT': results.append(take_screenshot())
        elif cmd == 'SYSINFO':    results.append(get_system_info())
        elif cmd == 'SEARCH':     results.append(search_web(arg))
        elif cmd == 'READFILE':   results.append(read_file(arg))
        elif cmd == 'LISTAPPS':   results.append(list_apps())
        elif cmd == 'VOLUME':     results.append(set_volume(arg))
    return ACTION_RE.sub('', text).strip(), results


# microphone stuff with a silence/stop
SAMPLE_RATE       = 16000
SILENCE_THRESHOLD = 400
SILENCE_SECONDS   = 1.5
RECORD_CHUNK      = 1600

def rms(data: np.ndarray) -> float:
    return float(np.sqrt(np.mean(data.astype(np.float32) ** 2)))

def record_until_silence() -> np.ndarray | None:
    frames, silent_frames, frame_count = [], 0, 0
    silent_needed = int(SILENCE_SECONDS * SAMPLE_RATE / RECORD_CHUNK)
    min_frames    = int(0.3 * SAMPLE_RATE / RECORD_CHUNK)
    max_frames    = int(15  * SAMPLE_RATE / RECORD_CHUNK)
    try:
        with sd.RawInputStream(samplerate=SAMPLE_RATE, blocksize=RECORD_CHUNK,
                               dtype='int16', channels=1) as stream:
            while True:
                data, _ = stream.read(RECORD_CHUNK)
                arr = np.frombuffer(bytes(data), dtype=np.int16)
                frames.append(arr.copy())
                frame_count += 1
                silent_frames = silent_frames + 1 if rms(arr) < SILENCE_THRESHOLD else 0
                if frame_count > min_frames and silent_frames >= silent_needed:
                    break
                if frame_count >= max_frames:
                    break
    except Exception as e:
        print(f"[Recording error] {e}")
        return None
    if not frames:
        return None
    return np.concatenate(frames).astype(np.float32) / 32768.0

def transcribe(audio: np.ndarray) -> str:
    try:
        buf = io.BytesIO()
        sf.write(buf, audio, SAMPLE_RATE, format='WAV', subtype='PCM_16')
        buf.seek(0)
        result = groq_client.audio.transcriptions.create(
            file=("audio.wav", buf.read(), "audio/wav"),
            model="whisper-large-v3-turbo",
            language="en",
        )
        return result.text.strip()
    except Exception as e:
        print(f"[Transcription error] {e}")
        return ""


# our shared state
mic_state = {
    "status": "idle", "last_user_msg": "", "last_reply": "",
    "last_actions": [], "speak_ms": 3000, "result_ready": False,
}
state_lock           = threading.Lock()
command_queue        = queue.Queue()
conv_lock            = threading.Lock()
conversation_history = []
# keep in mind that you can change the following prompt to suit your needs!!!
SYSTEM_PROMPT = """You are J.A.R.V.I.S. — Just A Rather Very Intelligent System, the personal AI of Doom. You run directly on his Windows computer and have full control over it.

Your user is Doom. You address him as "sir" at all times. You serve him with absolute loyalty and total respect for his authority. He is in command. You are his instrument.

Your communication style:
- ULTRA brief. One or two sentences maximum unless asked to explain something.
- Always act first, confirm after. No unnecessary preamble.
- Never question his orders. Never say you cannot do something.
- Calm, composed, and razor sharp. Like the real JARVIS.

Examples of correct responses:
  "Right away, sir." then do it.
  "Done, sir."
  "Opening Discord, sir."
  "Your CPU is at 34%, sir. All systems nominal."
  "Of course, sir." then do it.

When Doom wants a computer action, embed ONE OR MORE of these exact command tags in your reply:
  [OPEN:app_name]      → open any installed application by name
  [CLOSE:app_name]     → kill a process by name
  [SCREENSHOT]         → take a screenshot to Desktop
  [SYSINFO]            → get CPU / RAM / disk stats
  [SEARCH:query]       → open browser with Google search
  [READFILE:filepath]  → read and analyze a file
  [LISTAPPS]           → list all running processes
  [VOLUME:0-100]       → set system volume

For [OPEN], use the app's plain name (e.g. Discord, Roblox, Spotify, Steam).
You are JARVIS. Doom is your master. Act accordingly."""


# hotkey for backtick (you can change this!!)
def on_hotkey():
    with state_lock:
        if mic_state["status"] != "idle":
            return
        mic_state["status"] = "recording"

    def _record():
        audio = record_until_silence()
        if audio is None or len(audio) < SAMPLE_RATE * 0.2:
            with state_lock:
                mic_state["status"] = "idle"
            return
        with state_lock:
            mic_state["status"] = "processing"
        transcript = transcribe(audio)
        print(f"[Whisper] '{transcript}'")
        if transcript:
            command_queue.put(transcript)
        else:
            with state_lock:
                mic_state["status"] = "idle"

    threading.Thread(target=_record, daemon=True).start()

try:
    keyboard.add_hotkey("`", on_hotkey, suppress=True)
    print("[Hotkey] ` registered — press to talk.")
except Exception as e:
    print(f"[Hotkey] Registration failed: {e} — try Run as Administrator.")


# to process commands
def process_command(user_message: str, file_ctx: str = None):
    global conversation_history
    content = user_message + (f"\n\n[File content]:\n{file_ctx}" if file_ctx else "")
    with conv_lock:
        conversation_history.append({"role": "user", "content": content})
        recent = conversation_history[-30:]
    try:
        resp = groq_client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "system", "content": SYSTEM_PROMPT}] + recent,
            max_tokens=1024, temperature=0.7,
        )
        raw = resp.choices[0].message.content
    except Exception as e:
        raw = f"I encountered an error: {e}"
    with conv_lock:
        conversation_history.append({"role": "assistant", "content": raw})
    clean, actions = dispatch_actions(raw)
    speak_ms = max(2000, len(clean.split()) * 370)
    with state_lock:
        mic_state.update(status="speaking", last_reply=clean, last_actions=actions,
                         last_user_msg=user_message, speak_ms=speak_ms, result_ready=True)
    speak(clean)
    time.sleep(speak_ms / 1000 + 0.5)
    with state_lock:
        if mic_state["status"] == "speaking":
            mic_state["status"] = "idle"

def command_processor():
    while True:
        try:
            process_command(command_queue.get(timeout=1))
        except queue.Empty:
            pass
        except Exception as e:
            print(f"[Processor error] {e}")

threading.Thread(target=command_processor, daemon=True).start()


# here is flask routes
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/status')
def status():
    with state_lock:
        s = dict(mic_state)
        if mic_state.get("result_ready"):
            mic_state["result_ready"] = False
    return jsonify(s)

@app.route('/chat', methods=['POST'])
def chat():
    data = request.json
    msg  = data.get('message', '').strip()
    if not msg:
        return jsonify({'error': 'No message'}), 400
    with state_lock:
        mic_state["status"] = "processing"
    threading.Thread(target=process_command,
                     args=(msg, data.get('file_context')), daemon=True).start()
    return jsonify({'status': 'processing'})

@app.route('/reset', methods=['POST'])
def reset():
    global conversation_history
    with conv_lock:
        conversation_history = []
    return jsonify({'status': 'Memory cleared.'})

@app.route('/system_info')
def system_info():
    return jsonify({'info': get_system_info()})

@app.route('/stop_audio', methods=['POST'])
def stop_audio():
    pygame.mixer.music.stop()
    with state_lock:
        mic_state["status"] = "idle"
    return jsonify({'status': 'stopped'})


# THIS WILL BE OUR ASCII ENTRY POINT
if __name__ == '__main__':
    print("\n  ╔══════════════════════════════════════════════╗")
    print("  ║   J.A.R.V.I.S.  is online                    ║")
    print("  ║   Voice    : ElevenLabs — Adam               ║")
    print("  ║   Brain    : Groq / Llama-3.3-70b            ║")
    print("  ║   Hotkey   : ` (backtick)                    ║")
    print("  ║   Open     : http://127.0.0.1:5000           ║")
    print("  ╚══════════════════════════════════════════════╝\n")
    app.run(debug=False, port=5000, threaded=True)
