import os
import time
import json
import base64
import hashlib
import secrets
import threading
import webbrowser
import urllib.parse
import subprocess
import datetime
import shutil
import ctypes
from ctypes import wintypes
from http.server import BaseHTTPRequestHandler, HTTPServer
import requests
import tkinter as tk
import customtkinter as ctk
try:
    import psutil
except:
    psutil = None
from pythonosc.udp_client import SimpleUDPClient

HELP_TEXT = """VRChat Spotify Status — Hilfe

Setup
1) Spotify: Button „Sign in to Spotify“ klicken und erlauben.
2) VRChat: In-Game Settings → OSC aktivieren. Ziel: 127.0.0.1 : 9000.
3) Auf „Start“ klicken. Chatbox-Sound und Modus sind automatisch aktiviert.

Anzeige & Template
- Platzhalter: {prefix} {title} {artist} {sep} {bar} {position} {duration} {elapsed} {remaining}
- Titel/Künstler in einer Zeile; Timestamp optional in eigener Zeile („Timestamp on 2nd line“).
- Progress style: ascii / unicode / hud (HUD-Style mit Zeiten links/rechts in der Bar).
- Progress bar: per „Progress bar“ ein-/ausblenden. Bar length = Länge.
- „Clamp long title/artist“ begrenzt Länge (Ellipsis), damit andere Infos sichtbar bleiben.
- Unicode benötigt „Strip non-ASCII“ aus.
- Max. Länge pro Chat-Nachricht: 144 Zeichen (längere Texte werden hart gekürzt).

Rotation
- Enable rotation → wechselt die Einträge im Intervall.
- Mode: standalone / prepend / append / twoline.

Uhrzeit
- „Show current time (extra line)“ → zusätzliche Zeile mit Uhrzeit (System-Zeitzone).
- 24h & Prefix konfigurierbar.

PC Specs (extra line)
- „Show PC specs“ → Zeile mit CPU/RAM/GPU.
- CPU% via psutil, RAM in GB+%, GPU via nvidia-smi (falls vorhanden), sonst n/a.
- Format: CPU 12% | RAM 8.2/32 GB (26%) | GPU 34%

AFK
- Anti-AFK: periodischer Jump.
- AFK Tagger: setzt z. B. [AFK] ans Ende der ersten Zeile, wenn du X Sekunden inaktiv bist (System-Idle).

Quick Tests
- Send Test / Typing 3s / Jump.

Dateien
- Config: spotify_vrchat_gui.json
- Tokens: spotify_tokens.json
(gleicher Ordner wie Script — portable mode)

Troubleshooting
- Nichts im Chat? In VRChat Chatbox einblenden und OSC aktivieren. IP/Port prüfen.
- 401/Refresh-Fehler: „Clear Tokens“ und neu bei Spotify anmelden.
- 400 bei Token: Redirect URI exakt wie angezeigt verwenden.
- Unicode/zu lange Texte: ggf. „Strip non-ASCII“ an, Bar-Länge und Clamp-Limits anpassen.

Chatbox-Sound
- „Chat sound“ schaltet den Sound bei Nachrichten an/aus.
"""

PORTABLE_MODE = True
CLIENT_ID_DEFAULT = ""
REDIRECT_HOST = "127.0.0.1"
REDIRECT_PORT = 57893
REDIRECT_URI = f"http://{REDIRECT_HOST}:{REDIRECT_PORT}/callback"
SCOPE = "user-read-playback-state"
MAX_MESSAGE_LEN = 144
CHATBOX_INPUT = "/chatbox/input"
CHATBOX_TYPING = "/chatbox/typing"
INPUT_JUMP = "/input/Jump"

APP_DEFAULTS = {
    "client_id": CLIENT_ID_DEFAULT,
    "save_client_id": True,
    "ip": "127.0.0.1",
    "port": 9000,
    "update_interval": 3,

    "bar_length": 20,
    "show_bar": True,
    "prefix": True,
    "prefix_text": "Spotify:",
    "sep_title_artist": " – ",
    "progress_style": "ascii",
    "show_title": True,
    "show_artist": True,
    "show_time": True,
    "time_mode": "both",
    "time_on_second_line": True,
    "ascii_only": True,
    "only_changes": True,
    "template": "{prefix} {title}{sep}{artist} {bar} {position}/{duration}",

    "rotation_enabled": True,
    "rotation_interval": 6,
    "rotation_mode": "twoline",
    "rotation_items": [
        {"text": "vibing in VRC"},
        {"text": "Spotify Status via OSC"}
    ],

    "show_clock_line": False,
    "clock_24h": True,
    "clock_prefix": "",

    "anti_afk_enabled": False,
    "anti_afk_interval": 240,

    "show_specs_line": False,
    "show_specs_cpu": True,
    "show_specs_ram": True,
    "show_specs_gpu": True,
    "ram_in_gb": True,

    "clamp_long": True,
    "max_title_len": 28,
    "max_artist_len": 28,

    "afk_tag_enabled": False,
    "afk_tag_after": 120,
    "afk_tag_text": "[AFK]",

    "chat_sound": True
}

def _data_dir():
    if PORTABLE_MODE:
        return os.path.dirname(os.path.abspath(__file__))
    p = os.path.join(os.getenv("LOCALAPPDATA", os.path.expanduser("~")), "VRChatSpotifyStatus")
    os.makedirs(p, exist_ok=True)
    return p

TOKEN_FILE = os.path.join(_data_dir(), "spotify_tokens.json")
CONFIG_FILE = os.path.join(_data_dir(), "spotify_vrchat_gui.json")

def b64u(b): return base64.urlsafe_b64encode(b).decode("ascii").rstrip("=")

def gen_pkce():
    v = b64u(secrets.token_bytes(64))
    d = hashlib.sha256(v.encode("ascii")).digest()
    c = b64u(d)
    return v, c

class AuthCodeServer(BaseHTTPRequestHandler):
    code_value = None
    def do_GET(self):
        p = urllib.parse.urlparse(self.path)
        if p.path != "/callback":
            self.send_response(404); self.end_headers(); return
        q = urllib.parse.parse_qs(p.query)
        code = q.get("code", [None])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(b"<html><body><h2>Spotify authorization complete.</h2>You can close this window.</body></html>")
        AuthCodeServer.code_value = code

def wait_for_code():
    httpd = HTTPServer((REDIRECT_HOST, REDIRECT_PORT), AuthCodeServer)
    while AuthCodeServer.code_value is None:
        httpd.handle_request()
    return AuthCodeServer.code_value

def token_store_load():
    if os.path.exists(TOKEN_FILE):
        try:
            with open(TOKEN_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return {}
    return {}

def token_store_save(tokens):
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(tokens, f)

def config_load():
    cfg = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except:
            cfg = {}
    out = dict(APP_DEFAULTS); out.update(cfg or {})
    for k, v in APP_DEFAULTS.items():
        if k not in out:
            out[k] = v
    return out

def config_save(cfg):
    data = dict(cfg)
    if not data.get("save_client_id", True):
        data["client_id"] = ""
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f)

def token_expired(tokens):
    if not tokens or "access_token" not in tokens or "expires_in" not in tokens or "obtained_at" not in tokens:
        return True
    return (int(time.time()) - int(tokens["obtained_at"])) >= int(tokens["expires_in"]) - 30

def raise_for_status_with_body(resp):
    try:
        resp.raise_for_status()
    except requests.HTTPError as e:
        try: body = resp.json()
        except:
            try: body = resp.text
            except: body = ""
        raise RuntimeError(f"{e} | Response: {body}")

def authorize_pkce(client_id):
    verifier, challenge = gen_pkce()
    params = {
        "client_id": client_id, "response_type": "code",
        "redirect_uri": REDIRECT_URI, "scope": SCOPE,
        "code_challenge_method": "S256", "code_challenge": challenge,
        "show_dialog": "true"
    }
    url = "https://accounts.spotify.com/authorize?" + urllib.parse.urlencode(params)
    webbrowser.open(url)
    code = wait_for_code()
    data = {
        "grant_type": "authorization_code", "code": code,
        "redirect_uri": REDIRECT_URI, "client_id": client_id,
        "code_verifier": verifier
    }
    r = requests.post("https://accounts.spotify.com/api/token", data=data, timeout=30)
    raise_for_status_with_body(r)
    tokens = r.json()
    tokens["client_id"] = client_id
    tokens["obtained_at"] = int(time.time())
    token_store_save(tokens)
    return tokens

def refresh_token(tokens):
    if not tokens or "refresh_token" not in tokens or "client_id" not in tokens:
        return tokens
    data = {"grant_type": "refresh_token", "refresh_token": tokens["refresh_token"], "client_id": tokens["client_id"]}
    r = requests.post("https://accounts.spotify.com/api/token", data=data, timeout=30)
    if r.status_code >= 400:
        try: j = r.json()
        except: j = {}
        err = str(j)
        if any(k in err for k in ("invalid_grant", "invalid_client", "invalid_request")):
            raise RuntimeError(f"Refresh failed: {err}")
        raise_for_status_with_body(r)
    j = r.json()
    tokens["access_token"] = j.get("access_token", tokens.get("access_token"))
    if "refresh_token" in j: tokens["refresh_token"] = j["refresh_token"]
    tokens["expires_in"] = j.get("expires_in", tokens.get("expires_in"))
    tokens["obtained_at"] = int(time.time())
    token_store_save(tokens)
    return tokens

def get_current_playback(access_token):
    h = {"Authorization": f"Bearer {access_token}"}
    r = requests.get("https://api.spotify.com/v1/me/player/currently-playing", headers=h, timeout=15)
    if r.status_code == 204: return None
    if r.status_code == 200: return r.json()
    if r.status_code == 401: return "unauthorized"
    return None

def ms_to_clock(ms):
    s = int(ms // 1000); m = s // 60; s = s % 60
    return f"{m}:{s:02d}"

def build_bar(position_ms, duration_ms, length_chars, style="ascii", ascii_only=True, inline_times=True):
    if duration_ms <= 0:
        core = "-" * length_chars
        if style == "hud" and not ascii_only and inline_times:
            return f"0:00 {core} 0:00"
        return "[" + core + "]"
    f = max(0.0, min(1.0, float(position_ms) / float(duration_ms)))
    filled = int(round(length_chars * f))
    if style == "hud":
        if ascii_only:
            bar = "#" * filled + "-" * (length_chars - filled)
            if inline_times:
                return f"{ms_to_clock(position_ms)} [{bar}] {ms_to_clock(duration_ms)}"
            return "[" + bar + "]"
        fill = "▉"
        empty = "░"
        bar = fill * filled + empty * (length_chars - filled)
        if inline_times:
            return f"{ms_to_clock(position_ms)} {bar} {ms_to_clock(duration_ms)}"
        return bar
    if style == "unicode" and not ascii_only:
        left = "│"; right = "│"; fill = "█"; empty = "░"
        return left + fill * filled + empty * (length_chars - filled) + right
    return "[" + "#" * filled + "-" * (length_chars - filled) + "]"

def clamp_ascii(s):
    try: s.encode("ascii"); return s
    except: return s.encode("ascii", "ignore").decode("ascii")

def trim_chatbox(s):
    return s if len(s) <= MAX_MESSAGE_LEN else s[:MAX_MESSAGE_LEN-1] + "…"

def detect_process_fallback(substrs):
    try:
        out = subprocess.check_output(["tasklist", "/fo", "csv", "/nh"], creationflags=0x08000000).decode("utf-8", "ignore").lower()
        for line in out.splitlines():
            for s in substrs:
                if s in line: return True
        return False
    except:
        return None

def detect_process_any(substrs):
    ls = [s.lower() for s in substrs]
    if psutil is not None:
        try:
            for p in psutil.process_iter(["name"]):
                n = (p.info.get("name") or "").lower()
                for s in ls:
                    if s in n: return True
            return False
        except:
            pass
    return detect_process_fallback(ls)

class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.UINT), ("dwTime", wintypes.DWORD)]

def get_idle_seconds():
    try:
        last = LASTINPUTINFO()
        last.cbSize = ctypes.sizeof(last)
        if ctypes.windll.user32.GetLastInputInfo(ctypes.byref(last)):
            tick = ctypes.windll.kernel32.GetTickCount()
            idle_ms = tick - last.dwTime
            return max(0.0, idle_ms / 1000.0)
    except Exception:
        pass
    return 0.0

def _gpu_from_nvidia_smi():
    try:
        exe = shutil.which("nvidia-smi")
        if not exe:
            return None
        out = subprocess.check_output(
            [exe, "--query-gpu=utilization.gpu,name", "--format=csv,noheader,nounits"],
            stderr=subprocess.STDOUT,
            timeout=1.5,
            creationflags=0x08000000
        ).decode("utf-8", "ignore").strip().splitlines()
        if not out:
            return None
        parts = [p.strip() for p in out[0].split(",")]
        util = float(parts[0]) if parts and parts[0] else None
        name = parts[1] if len(parts) > 1 else ""
        return {"util": util, "name": name}
    except Exception:
        return None

def read_specs():
    cpu = None; ram = None; gpu = None
    try:
        if psutil:
            cpu = psutil.cpu_percent(interval=None)
            vm = psutil.virtual_memory()
            ram = {"used": vm.used, "total": vm.total, "percent": vm.percent}
    except Exception:
        pass
    gpu = _gpu_from_nvidia_smi()
    return cpu, ram, gpu

def fmt_specs(cpu, ram, gpu, show_cpu, show_ram, show_gpu, ram_in_gb=True, ascii_only=True):
    parts = []
    try:
        if show_cpu and cpu is not None:
            parts.append(f"CPU {cpu:.0f}%")
        if show_ram and ram is not None:
            if ram_in_gb:
                used_gb = ram["used"] / (1024**3)
                total_gb = ram["total"] / (1024**3)
                parts.append(f"RAM {used_gb:.1f}/{total_gb:.0f} GB ({ram['percent']:.0f}%)")
            else:
                parts.append(f"RAM {ram['percent']:.0f}%")
        if show_gpu:
            if gpu and gpu.get("util") is not None:
                parts.append(f"GPU {gpu['util']:.0f}%")
            else:
                parts.append("GPU n/a")
    except Exception:
        pass
    s = " | ".join(parts)
    return clamp_ascii(s) if ascii_only else s

def shorten(s, limit):
    s = s or ""
    if limit <= 1 or len(s) <= limit: return s
    return s[:max(1, limit-1)] + "…"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("green")
        self.title("VRChat Spotify Status")
        self.geometry("1320x900")
        self.resizable(False, False)
        self.cfg = config_load()
        self.tokens = token_store_load()
        self.osc = None
        self.running = False
        self.worker = None
        self.last_message = ""
        self.last_track_id = ""
        self.last_item = None
        self.last_progress = 0
        self.last_duration = 0
        self.rot_idx = 0
        self.next_rotate_at = time.monotonic()
        self.current_rot_text = ""
        self.next_afk_at = time.monotonic()
        self._last_specs = ("", 0.0)
        try:
            if psutil: psutil.cpu_percent(interval=None)
        except Exception:
            pass
        self._build_ui()
        self._bind_autosave()
        self._update_status_loop()

    def get_int(self, var, default, lo=None, hi=None):
        try:
            s = str(var.get()).strip()
            if s == "":
                return default
            v = int(float(s))
        except:
            return default
        if lo is not None: v = max(lo, v)
        if hi is not None: v = min(hi, v)
        return v

    def _build_ui(self):
        grid = ctk.CTkFrame(self, corner_radius=12); grid.pack(fill="both", expand=True, padx=12, pady=12)

        left = ctk.CTkFrame(grid, width=450, corner_radius=12)
        left.grid(row=0, column=0, rowspan=2, sticky="nsw", padx=(12, 8), pady=12)

        ctk.CTkLabel(left, text="Connection", font=ctk.CTkFont(size=18, weight="bold")).pack(anchor="w", padx=12, pady=(12, 8))

        self.var_client_id = ctk.StringVar(value=self.cfg["client_id"])
        self.var_save_cid = ctk.BooleanVar(value=self.cfg["save_client_id"])
        self.var_ip = ctk.StringVar(value=self.cfg["ip"])
        self.var_port = ctk.StringVar(value=str(self.cfg["port"]))
        self.var_update = ctk.StringVar(value=str(self.cfg["update_interval"]))

        ctk.CTkLabel(left, text="Spotify Client ID").pack(anchor="w", padx=12)
        ctk.CTkEntry(left, textvariable=self.var_client_id).pack(fill="x", padx=12, pady=(0,6))
        top_row = ctk.CTkFrame(left); top_row.pack(fill="x", padx=12)
        ctk.CTkButton(top_row, text="Sign in to Spotify", command=self._on_spotify_login).pack(side="left")
        ctk.CTkCheckBox(top_row, text="Save ID", variable=self.var_save_cid).pack(side="left", padx=(8,0))

        ctk.CTkLabel(left, text=f"Redirect URI\n{REDIRECT_URI}", wraplength=420, justify="left").pack(anchor="w", padx=12, pady=(6,10))

        net = ctk.CTkFrame(left); net.pack(fill="x", padx=12, pady=(0,6))
        ctk.CTkLabel(net, text="VRChat IP").grid(row=0, column=0, sticky="w")
        ctk.CTkEntry(net, width=190, textvariable=self.var_ip).grid(row=0, column=1, padx=(6,12))
        ctk.CTkLabel(net, text="Port").grid(row=0, column=2)
        ctk.CTkEntry(net, width=120, textvariable=self.var_port).grid(row=0, column=3, padx=(6,0))

        upd_row = ctk.CTkFrame(left); upd_row.pack(fill="x", padx=12, pady=(6,10))
        ctk.CTkLabel(upd_row, text="Update interval (s)").pack(side="left")
        ctk.CTkEntry(upd_row, width=120, textvariable=self.var_update).pack(side="left", padx=(6,0))

        anti = ctk.CTkFrame(left); anti.pack(fill="x", padx=12, pady=(6,10))
        ctk.CTkLabel(anti, text="Anti-AFK").grid(row=0, column=0, sticky="w")
        self.var_afk_enabled = ctk.BooleanVar(value=self.cfg["anti_afk_enabled"])
        self.var_afk_interval = ctk.StringVar(value=str(self.cfg["anti_afk_interval"]))
        ctk.CTkCheckBox(anti, text="Enable", variable=self.var_afk_enabled).grid(row=0, column=1, padx=(12,6))
        ctk.CTkLabel(anti, text="Jump every (s)").grid(row=0, column=2, padx=(18,6))
        ctk.CTkEntry(anti, width=120, textvariable=self.var_afk_interval).grid(row=0, column=3)

        tests = ctk.CTkFrame(left); tests.pack(fill="x", padx=12, pady=(8,8))
        ctk.CTkLabel(tests, text="Quick tests", font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
        brow = ctk.CTkFrame(tests); brow.pack(fill="x", pady=(4,0))
        ctk.CTkButton(brow, text="Send Test", command=self._on_test, width=130).pack(side="left")
        ctk.CTkButton(brow, text="Typing 3s", command=self._on_typing_test, width=130).pack(side="left", padx=(8,0))
        ctk.CTkButton(brow, text="Jump", command=self._on_jump_test, width=130).pack(side="left", padx=(8,0))

        self.lbl_sp = ctk.CTkLabel(left, text="Spotify: ?"); self.lbl_vr = ctk.CTkLabel(left, text="VRChat: ?")
        self.lbl_pb = ctk.CTkLabel(left, text="Playback: ?"); self.lbl_auth = ctk.CTkLabel(left, text="Auth: ?")
        for w in (self.lbl_sp, self.lbl_vr, self.lbl_pb, self.lbl_auth): w.pack(anchor="w", padx=12)

        tabs = ctk.CTkTabview(grid, width=820, height=710, corner_radius=12)
        tabs.grid(row=0, column=1, sticky="nsew", padx=(8,12), pady=(12,8))
        tab_display = tabs.add("Display"); tab_rotate = tabs.add("Rotator"); tab_help = tabs.add("Help")

        self.var_prefix = ctk.BooleanVar(value=self.cfg["prefix"])
        self.var_prefix_text = ctk.StringVar(value=self.cfg["prefix_text"])
        self.var_sep = ctk.StringVar(value=self.cfg["sep_title_artist"])
        self.var_progress_style = ctk.StringVar(value=self.cfg["progress_style"])

        self.var_title = ctk.BooleanVar(value=self.cfg["show_title"])
        self.var_artist = ctk.BooleanVar(value=self.cfg["show_artist"])
        self.var_time = ctk.BooleanVar(value=self.cfg["show_time"])
        self.var_time_mode = ctk.StringVar(value=self.cfg["time_mode"])
        self.var_time_second_line = ctk.BooleanVar(value=self.cfg["time_on_second_line"])
        self.var_ascii = ctk.BooleanVar(value=self.cfg["ascii_only"])
        self.var_only_changes = ctk.BooleanVar(value=self.cfg["only_changes"])
        self.var_show_bar = ctk.BooleanVar(value=self.cfg["show_bar"])
        self.var_bar_len = ctk.StringVar(value=str(self.cfg["bar_length"]))
        self.var_template = ctk.StringVar(value=self.cfg["template"])
        self.var_preview = ctk.StringVar(value="")

        self.var_clamp_long = ctk.BooleanVar(value=self.cfg["clamp_long"])
        self.var_max_title = ctk.StringVar(value=str(self.cfg["max_title_len"]))
        self.var_max_artist = ctk.StringVar(value=str(self.cfg["max_artist_len"]))

        self.var_clock_line = ctk.BooleanVar(value=self.cfg["show_clock_line"])
        self.var_clock_24h = ctk.BooleanVar(value=self.cfg["clock_24h"])
        self.var_clock_prefix = ctk.StringVar(value=self.cfg["clock_prefix"])

        self.var_specs_line = ctk.BooleanVar(value=self.cfg["show_specs_line"])
        self.var_specs_cpu = ctk.BooleanVar(value=self.cfg["show_specs_cpu"])
        self.var_specs_ram = ctk.BooleanVar(value=self.cfg["show_specs_ram"])
        self.var_specs_gpu = ctk.BooleanVar(value=self.cfg["show_specs_gpu"])
        self.var_specs_ram_gb = ctk.BooleanVar(value=self.cfg["ram_in_gb"])

        self.var_afk_tag_enabled = ctk.BooleanVar(value=self.cfg["afk_tag_enabled"])
        self.var_afk_tag_after = ctk.StringVar(value=str(self.cfg["afk_tag_after"]))
        self.var_afk_tag_text = ctk.StringVar(value=self.cfg["afk_tag_text"])

        self.var_chat_sound = ctk.BooleanVar(value=self.cfg.get("chat_sound", True))

        look = ctk.CTkFrame(tab_display); look.pack(fill="x", padx=12, pady=(12,8))
        ctk.CTkCheckBox(look, text='Show prefix', variable=self.var_prefix).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkLabel(look, text="Prefix text").grid(row=0, column=1, sticky="e", padx=(18,6))
        ctk.CTkEntry(look, width=160, textvariable=self.var_prefix_text).grid(row=0, column=2, sticky="w")
        ctk.CTkLabel(look, text="Title–Artist sep").grid(row=1, column=1, sticky="e", padx=(18,6))
        ctk.CTkEntry(look, width=160, textvariable=self.var_sep).grid(row=1, column=2, sticky="w")
        ctk.CTkLabel(look, text="Progress style").grid(row=0, column=3, sticky="e", padx=(18,6))
        ctk.CTkOptionMenu(look, values=["ascii","unicode","hud"], variable=self.var_progress_style, width=130).grid(row=0, column=4, sticky="w")

        opt = ctk.CTkFrame(tab_display); opt.pack(fill="x", padx=12, pady=(12,8))
        ctk.CTkCheckBox(opt, text="Title", variable=self.var_title).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(opt, text="Artist", variable=self.var_artist).grid(row=0, column=1, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(opt, text="Show timestamp", variable=self.var_time).grid(row=0, column=2, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(opt, text="Timestamp on 2nd line", variable=self.var_time_second_line).grid(row=0, column=3, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(opt, text="Progress bar", variable=self.var_show_bar).grid(row=0, column=4, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(opt, text="Strip non-ASCII", variable=self.var_ascii).grid(row=1, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(opt, text="Only send on change", variable=self.var_only_changes).grid(row=1, column=1, padx=6, pady=4, sticky="w")
        ctk.CTkLabel(opt, text="Bar length").grid(row=1, column=2, sticky="e", padx=(18,6))
        ctk.CTkEntry(opt, width=110, textvariable=self.var_bar_len).grid(row=1, column=3, sticky="w")
        ctk.CTkCheckBox(opt, text="Chat sound", variable=self.var_chat_sound).grid(row=1, column=4, padx=6, pady=4, sticky="w")

        clampf = ctk.CTkFrame(tab_display); clampf.pack(fill="x", padx=12, pady=(6,8))
        ctk.CTkCheckBox(clampf, text="Clamp long title/artist", variable=self.var_clamp_long).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkLabel(clampf, text="Max title").grid(row=0, column=1, sticky="e", padx=(18,6))
        ctk.CTkEntry(clampf, width=110, textvariable=self.var_max_title).grid(row=0, column=2, sticky="w")
        ctk.CTkLabel(clampf, text="Max artist").grid(row=0, column=3, sticky="e", padx=(18,6))
        ctk.CTkEntry(clampf, width=110, textvariable=self.var_max_artist).grid(row=0, column=4, sticky="w")

        clockf = ctk.CTkFrame(tab_display); clockf.pack(fill="x", padx=12, pady=(6,8))
        ctk.CTkCheckBox(clockf, text="Show current time (extra line)", variable=self.var_clock_line).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(clockf, text="24h", variable=self.var_clock_24h).grid(row=0, column=1, padx=(18,6), sticky="w")
        ctk.CTkLabel(clockf, text="Clock prefix").grid(row=0, column=2, padx=(18,6), sticky="e")
        ctk.CTkEntry(clockf, width=140, textvariable=self.var_clock_prefix).grid(row=0, column=3, sticky="w")

        specsf = ctk.CTkFrame(tab_display); specsf.pack(fill="x", padx=12, pady=(6,8))
        ctk.CTkCheckBox(specsf, text="Show PC specs (extra line)", variable=self.var_specs_line).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(specsf, text="CPU", variable=self.var_specs_cpu).grid(row=0, column=1, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(specsf, text="RAM", variable=self.var_specs_ram).grid(row=0, column=2, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(specsf, text="GPU", variable=self.var_specs_gpu).grid(row=0, column=3, padx=6, pady=4, sticky="w")
        ctk.CTkCheckBox(specsf, text="RAM in GB", variable=self.var_specs_ram_gb).grid(row=0, column=4, padx=6, pady=4, sticky="w")

        afkf = ctk.CTkFrame(tab_display); afkf.pack(fill="x", padx=12, pady=(6,8))
        ctk.CTkCheckBox(afkf, text="AFK tagger enabled", variable=self.var_afk_tag_enabled).grid(row=0, column=0, padx=6, pady=4, sticky="w")
        ctk.CTkLabel(afkf, text="AFK after (s)").grid(row=0, column=1, sticky="e", padx=(18,6))
        ctk.CTkEntry(afkf, width=120, textvariable=self.var_afk_tag_after).grid(row=0, column=2, sticky="w")
        ctk.CTkLabel(afkf, text="AFK tag text").grid(row=0, column=3, sticky="e", padx=(18,6))
        ctk.CTkEntry(afkf, width=160, textvariable=self.var_afk_tag_text).grid(row=0, column=4, sticky="w")

        ctk.CTkLabel(tab_display, text="Template").pack(anchor="w", padx=12, pady=(6,0))
        ctk.CTkEntry(tab_display, textvariable=self.var_template).pack(fill="x", padx=12, pady=(0,6))
        r1 = ctk.CTkFrame(tab_display); r1.pack(fill="x", padx=12, pady=(0,6))
        ctk.CTkButton(r1, text="Reset Template", command=self._reset_template, width=150).pack(side="left")
        ctk.CTkLabel(tab_display, text="Preview").pack(anchor="w", padx=12)
        ctk.CTkLabel(tab_display, textvariable=self.var_preview, wraplength=790, justify="left").pack(fill="x", padx=12, pady=(0,12))

        self.var_rot_enabled = ctk.BooleanVar(value=self.cfg["rotation_enabled"])
        self.var_rot_interval = ctk.StringVar(value=str(self.cfg["rotation_interval"]))
        self.var_rot_mode = ctk.StringVar(value=self.cfg.get("rotation_mode", "twoline"))

        rot_top = ctk.CTkFrame(tab_rotate); rot_top.pack(fill="x", padx=12, pady=(12,6))
        ctk.CTkCheckBox(rot_top, text="Enable rotation", variable=self.var_rot_enabled).grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(rot_top, text="Interval (s)").grid(row=0, column=1, padx=(18,6))
        ctk.CTkEntry(rot_top, width=120, textvariable=self.var_rot_interval).grid(row=0, column=2, padx=(0,12))
        ctk.CTkLabel(rot_top, text="Mode").grid(row=0, column=3, padx=(18,6))
        ctk.CTkOptionMenu(rot_top, values=["standalone","prepend","append","twoline"], variable=self.var_rot_mode, width=160).grid(row=0, column=4)

        mid = ctk.CTkFrame(tab_rotate); mid.pack(fill="both", expand=True, padx=12, pady=6)
        leftpane = ctk.CTkFrame(mid); leftpane.pack(side="left", fill="both", expand=True, padx=(0,6))
        rightpane = ctk.CTkFrame(mid, width=360); rightpane.pack(side="left", fill="y", padx=(6,0))

        self.listbox = tk.Listbox(leftpane, height=18, activestyle="dotbox")
        self.listbox.pack(fill="both", expand=True, padx=8, pady=8)
        self.listbox.bind("<<ListboxSelect>>", lambda e: self._load_selected_rotation_item())

        ctled = ctk.CTkFrame(rightpane); ctled.pack(fill="x", padx=8, pady=(12,6))
        ctk.CTkLabel(ctled, text="Text / Template").pack(anchor="w")
        self.entry_item_text = ctk.CTkTextbox(ctled, height=160)
        self.entry_item_text.pack(fill="x")
        self.entry_item_text.bind("<KeyRelease>", self._on_rot_text_edit)

        btns = ctk.CTkFrame(ctled); btns.pack(fill="x", pady=6)
        ctk.CTkButton(btns, text="Add", command=self._rot_add).pack(side="left", padx=(0,6))
        ctk.CTkButton(btns, text="Delete", command=self._rot_delete).pack(side="left")

        order = ctk.CTkFrame(rightpane); order.pack(fill="x", padx=8, pady=(0,8))
        ctk.CTkButton(order, text="↑ Move Up", command=self._rot_up).pack(fill="x", pady=3)
        ctk.CTkButton(order, text="↓ Move Down", command=self._rot_down).pack(fill="x", pady=3)

        self._build_help_tab(tab_help)

        bottom = ctk.CTkFrame(grid, height=180, corner_radius=12)
        bottom.grid(row=1, column=1, sticky="nsew", padx=(8,12), pady=(8,12))
        control = ctk.CTkFrame(bottom); control.pack(fill="x", padx=12, pady=(10,4))
        ctk.CTkButton(control, text="Start", command=self._on_start, width=120).pack(side="left")
        ctk.CTkButton(control, text="Stop", command=self._on_stop, width=120).pack(side="left", padx=(8,0))
        ctk.CTkButton(control, text="Clear Tokens", command=self._clear_tokens).pack(side="left", padx=(16,0))
        ctk.CTkButton(control, text="Reset Config", command=self._reset_config).pack(side="left", padx=(8,0))

        self.txt_log = tk.Text(bottom, height=10); self.txt_log.pack(fill="both", expand=True, padx=12, pady=(4,10))

        self.rotation_items = list(self.cfg.get("rotation_items", []))
        self._refresh_rot_list()
        self._update_preview()

    def _build_help_tab(self, tab):
        box = ctk.CTkTextbox(tab, height=620)
        box.pack(fill="both", expand=True, padx=12, pady=12)
        box.insert("1.0", HELP_TEXT)
        box.configure(state="disabled")
        bar = ctk.CTkFrame(tab)
        bar.pack(fill="x", padx=12, pady=(0,12))
        ctk.CTkButton(bar, text="Copy help", command=self._copy_help_text, width=120).pack(side="left")
        ctk.CTkButton(bar, text="Open data folder", command=self._open_data_dir, width=160).pack(side="left", padx=(8,0))

    def _copy_help_text(self):
        try:
            self.clipboard_clear()
            self.clipboard_append(HELP_TEXT)
            self._log("Help copied to clipboard")
        except Exception as e:
            self._log(f"Clipboard error: {e}")

    def _open_data_dir(self):
        try:
            path = _data_dir()
            os.makedirs(path, exist_ok=True)
            subprocess.Popen(["explorer", path], shell=True)
        except Exception as e:
            self._log(f"Open folder error: {e}")

    def _log(self, s):
        self.txt_log.insert("end", time.strftime("[%H:%M:%S] ") + s + "\n")
        self.txt_log.see("end")

    def _bind_autosave(self):
        def save(*_):
            self._save_config()
            self._update_preview()
        for v in (
            self.var_client_id, self.var_ip, self.var_time_mode, self.var_template,
            self.var_rot_mode, self.var_port, self.var_update, self.var_bar_len,
            self.var_rot_interval, self.var_prefix_text, self.var_sep,
            self.var_progress_style, self.var_clock_prefix, self.var_afk_interval,
            self.var_max_title, self.var_max_artist, self.var_afk_tag_after,
            self.var_afk_tag_text
        ):
            v.trace_add("write", save)
        for v in (
            self.var_save_cid, self.var_prefix, self.var_title, self.var_artist, self.var_time,
            self.var_time_second_line, self.var_ascii, self.var_only_changes, self.var_rot_enabled,
            self.var_clock_line, self.var_clock_24h, self.var_afk_enabled, self.var_show_bar,
            self.var_specs_line, self.var_specs_cpu, self.var_specs_ram, self.var_specs_gpu,
            self.var_specs_ram_gb, self.var_clamp_long, self.var_afk_tag_enabled,
            self.var_chat_sound
        ):
            v.trace_add("write", save)

    def _save_config(self):
        cfg = {
            "client_id": self.var_client_id.get().strip(),
            "save_client_id": bool(self.var_save_cid.get()),
            "ip": self.var_ip.get().strip(),
            "port": self.get_int(self.var_port, self.cfg.get("port", 9000), 1, 65535),
            "update_interval": self.get_int(self.var_update, self.cfg.get("update_interval", 3), 1, 120),

            "bar_length": self.get_int(self.var_bar_len, self.cfg.get("bar_length", 20), 4, 60),
            "show_bar": bool(self.var_show_bar.get()),
            "prefix": bool(self.var_prefix.get()),
            "prefix_text": self.var_prefix_text.get(),
            "sep_title_artist": self.var_sep.get(),
            "progress_style": self.var_progress_style.get(),
            "show_title": bool(self.var_title.get()),
            "show_artist": bool(self.var_artist.get()),
            "show_time": bool(self.var_time.get()),
            "time_mode": self.var_time_mode.get(),
            "time_on_second_line": bool(self.var_time_second_line.get()),
            "ascii_only": bool(self.var_ascii.get()),
            "only_changes": bool(self.var_only_changes.get()),
            "template": self.var_template.get(),

            "rotation_enabled": bool(self.var_rot_enabled.get()),
            "rotation_interval": self.get_int(self.var_rot_interval, self.cfg.get("rotation_interval", 6), 1, 3600),
            "rotation_mode": self.var_rot_mode.get(),
            "rotation_items": self.rotation_items,

            "show_clock_line": bool(self.var_clock_line.get()),
            "clock_24h": bool(self.var_clock_24h.get()),
            "clock_prefix": self.var_clock_prefix.get(),

            "anti_afk_enabled": bool(self.var_afk_enabled.get()),
            "anti_afk_interval": self.get_int(self.var_afk_interval, self.cfg.get("anti_afk_interval", 240), 5, 3600),

            "show_specs_line": bool(self.var_specs_line.get()),
            "show_specs_cpu": bool(self.var_specs_cpu.get()),
            "show_specs_ram": bool(self.var_specs_ram.get()),
            "show_specs_gpu": bool(self.var_specs_gpu.get()),
            "ram_in_gb": bool(self.var_specs_ram_gb.get()),

            "clamp_long": bool(self.var_clamp_long.get()),
            "max_title_len": self.get_int(self.var_max_title, self.cfg.get("max_title_len", 28), 6, 80),
            "max_artist_len": self.get_int(self.var_max_artist, self.cfg.get("max_artist_len", 28), 6, 80),

            "afk_tag_enabled": bool(self.var_afk_tag_enabled.get()),
            "afk_tag_after": self.get_int(self.var_afk_tag_after, self.cfg.get("afk_tag_after", 120), 10, 36000),
            "afk_tag_text": self.var_afk_tag_text.get().strip() or "[AFK]",

            "chat_sound": bool(self.var_chat_sound.get())
        }
        config_save(cfg); self.cfg = cfg

    def _reset_config(self):
        try:
            if os.path.exists(CONFIG_FILE): os.remove(CONFIG_FILE)
            self.cfg = config_load()
            self.var_client_id.set(self.cfg["client_id"]); self.var_save_cid.set(self.cfg["save_client_id"])
            self.var_ip.set(self.cfg["ip"]); self.var_port.set(str(self.cfg["port"]))
            self.var_update.set(str(self.cfg["update_interval"])); self.var_bar_len.set(str(self.cfg["bar_length"]))
            self.var_show_bar.set(self.cfg["show_bar"])
            self.var_prefix.set(self.cfg["prefix"]); self.var_prefix_text.set(self.cfg["prefix_text"])
            self.var_sep.set(self.cfg["sep_title_artist"]); self.var_progress_style.set(self.cfg["progress_style"])
            self.var_title.set(self.cfg["show_title"]); self.var_artist.set(self.cfg["show_artist"])
            self.var_time.set(self.cfg["show_time"]); self.var_time_mode.set(self.cfg["time_mode"])
            self.var_time_second_line.set(self.cfg["time_on_second_line"])
            self.var_ascii.set(self.cfg["ascii_only"]); self.var_only_changes.set(self.cfg["only_changes"])
            self.var_template.set(self.cfg["template"])
            self.var_rot_enabled.set(self.cfg["rotation_enabled"]); self.var_rot_interval.set(str(self.cfg["rotation_interval"]))
            self.var_rot_mode.set(self.cfg["rotation_mode"])
            self.var_clock_line.set(self.cfg["show_clock_line"])
            self.var_clock_24h.set(self.cfg["clock_24h"]); self.var_clock_prefix.set(self.cfg["clock_prefix"])
            self.var_afk_enabled.set(self.cfg["anti_afk_enabled"]); self.var_afk_interval.set(str(self.cfg["anti_afk_interval"]))
            self.var_specs_line.set(self.cfg["show_specs_line"])
            self.var_specs_cpu.set(self.cfg["show_specs_cpu"]); self.var_specs_ram.set(self.cfg["show_specs_ram"]); self.var_specs_gpu.set(self.cfg["show_specs_gpu"])
            self.var_specs_ram_gb.set(self.cfg["ram_in_gb"])
            self.var_clamp_long.set(self.cfg["clamp_long"])
            self.var_max_title.set(str(self.cfg["max_title_len"])); self.var_max_artist.set(str(self.cfg["max_artist_len"]))
            self.var_afk_tag_enabled.set(self.cfg["afk_tag_enabled"])
            self.var_afk_tag_after.set(str(self.cfg["afk_tag_after"]))
            self.var_afk_tag_text.set(self.cfg["afk_tag_text"])
            self.var_chat_sound.set(self.cfg["chat_sound"])
            self.rotation_items = list(self.cfg["rotation_items"])
            self._refresh_rot_list(); self._update_preview(); self._save_config()
            self._log("Config reset")
        except Exception as e:
            self._log(f"Reset config error: {e}")

    def _clear_tokens(self):
        try:
            if os.path.exists(TOKEN_FILE): os.remove(TOKEN_FILE)
            self.tokens = {}; self.lbl_auth.configure(text="Auth: required")
            self._log("Tokens cleared")
        except Exception as e:
            self._log(f"Clear tokens error: {e}")

    def _on_spotify_login(self):
        try:
            cid = self.var_client_id.get().strip()
            if not cid: self._log("Client ID missing"); return
            self.tokens = authorize_pkce(cid); self.lbl_auth.configure(text="Auth: ok")
            self._log("Spotify authorized")
        except Exception as e:
            self.lbl_auth.configure(text="Auth: failed"); self._log(f"Auth error: {e}")

    def _ensure_osc(self):
        if self.osc is None:
            port = self.get_int(self.var_port, self.cfg.get("port", 9000), 1, 65535)
            self.osc = SimpleUDPClient(self.var_ip.get().strip(), port)

    def _send_chatbox_raw(self, text):
        self._ensure_osc()
        try:
            play_sound = bool(self.var_chat_sound.get())
            self.osc.send_message(CHATBOX_INPUT, [text, True, play_sound])
        except:
            self.osc.send_message(CHATBOX_INPUT, [text, True])

    def _send_chatbox(self, text):
        try:
            self._send_chatbox_raw(text); return True
        except Exception as e:
            self._log(f"Send error: {e}"); return False

    def _send_typing(self, value):
        try:
            self._ensure_osc(); self.osc.send_message(CHATBOX_TYPING, [bool(value)]); return True
        except Exception as e:
            self._log(f"Typing error: {e}"); return False

    def _send_jump(self):
        try:
            self._ensure_osc(); self.osc.send_message(INPUT_JUMP, 1); time.sleep(0.08); self.osc.send_message(INPUT_JUMP, 0); return True
        except Exception as e:
            self._log(f"Jump error: {e}"); return False

    def _on_test(self):
        main, time_line = self._render_spotify_lines(self.last_item, self.last_progress, self.last_duration)
        final = self._compose_full(main, time_line)
        if not final: final = "Test"
        if self._send_chatbox(final): self._log("Test sent")

    def _on_typing_test(self):
        if self._send_typing(True):
            self._log("Typing on")
            def off():
                self._send_typing(False)
                self._log("Typing off")
            threading.Timer(3.0, off).start()

    def _on_jump_test(self):
        if self._send_jump(): self._log("Jump sent")

    def _reset_template(self):
        self.var_template.set(APP_DEFAULTS["template"]); self._save_config(); self._update_preview()

    def _apply_clamp(self, title, artist):
        if not self.var_clamp_long.get():
            return title, artist
        mt = self.get_int(self.var_max_title, self.cfg.get("max_title_len", 28), 6, 80)
        ma = self.get_int(self.var_max_artist, self.cfg.get("max_artist_len", 28), 6, 80)
        return shorten(title, mt), shorten(artist, ma)

    def _render_spotify_lines(self, item, progress_ms, duration_ms):
        tpl = self.var_template.get().strip() or APP_DEFAULTS["template"]
        prefix_text = (self.var_prefix_text.get() if self.var_prefix.get() else "").strip()
        sep = self.var_sep.get()
        if item is None:
            main = tpl
            main = main.replace("{prefix}", prefix_text)
            main = main.replace("{title}","").replace("{artist}","")
            main = main.replace("{sep}", "" if not sep else sep)
            main = main.replace("{bar}","")
            main = main.replace("{position}","").replace("{duration}","").replace("{elapsed}","").replace("{remaining}","")
            main = " ".join(main.split())
            main = clamp_ascii(main) if self.var_ascii.get() else main
            return trim_chatbox(main), ""
        raw_title = item.get("name","")
        raw_artist = ", ".join([a.get("name","") for a in item.get("artists",[])])
        title, artist = self._apply_clamp(raw_title, raw_artist)
        ps = self.var_progress_style.get()
        show_bar = self.var_show_bar.get()
        inline_times_requested = (ps == "hud") and self.var_time.get()
        if show_bar:
            bar = build_bar(
                progress_ms, duration_ms,
                self.get_int(self.var_bar_len, 20, 4, 60),
                ps, self.var_ascii.get(),
                inline_times=inline_times_requested
            )
        else:
            bar = ""
        position = ms_to_clock(progress_ms); duration = ms_to_clock(duration_ms)
        elapsed = position; remaining = ms_to_clock(max(0, duration_ms - progress_ms))
        main = tpl
        main = main.replace("{prefix}", prefix_text)
        main = main.replace("{title}", title if self.var_title.get() else "")
        main = main.replace("{artist}", artist if self.var_artist.get() else "")
        main = main.replace("{sep}", sep if (self.var_title.get() and self.var_artist.get()) else "")
        main = main.replace("{bar}", bar)
        if self.var_time.get() and not self.var_time_second_line.get() and not inline_times_requested:
            main = main.replace("{position}", position).replace("{duration}", duration)
            main = main.replace("{elapsed}", elapsed if self.var_time_mode.get() in ("elapsed","both") else "")
            main = main.replace("{remaining}", ("-" + remaining) if self.var_time_mode.get() in ("remaining","both") else "")
        else:
            main = main.replace("{position}","").replace("{duration}","").replace("{elapsed}","").replace("{remaining}","")
        main = " ".join(main.split())
        main = clamp_ascii(main) if self.var_ascii.get() else main
        main = trim_chatbox(main)
        time_line = ""
        if self.var_time.get() and self.var_time_second_line.get() and not inline_times_requested:
            if self.var_time_mode.get() == "elapsed":
                time_line = elapsed
            elif self.var_time_mode.get() == "remaining":
                time_line = "-" + remaining
            else:
                time_line = f"{elapsed} / {duration}"
            time_line = clamp_ascii(time_line) if self.var_ascii.get() else time_line
            time_line = trim_chatbox(time_line)
        return main, time_line

    def _render_rotation_item(self, txt):
        t = (txt or "").strip()
        if not t: return ""
        backup = self.var_template.get()
        try:
            self.var_template.set(t)
            m, _ = self._render_spotify_lines(self.last_item, self.last_progress, self.last_duration)
            return m
        finally:
            self.var_template.set(backup)

    def _clock_line(self):
        if not self.var_clock_line.get(): return ""
        now = datetime.datetime.now()
        fmt = "%H:%M:%S" if self.var_clock_24h.get() else "%I:%M:%S %p"
        prefix = (self.var_clock_prefix.get() or "").strip()
        s = f"{prefix} {now.strftime(fmt)}".strip() if prefix else now.strftime(fmt)
        s = clamp_ascii(s) if self.var_ascii.get() else s
        return trim_chatbox(s)

    def _specs_line(self):
        if not self.var_specs_line.get():
            return ""
        now = time.monotonic()
        cached, ts = self._last_specs
        if now - ts < 1.0 and cached:
            return cached
        cpu, ram, gpu = read_specs()
        line = fmt_specs(cpu, ram, gpu,
                         self.var_specs_cpu.get(), self.var_specs_ram.get(), self.var_specs_gpu.get(),
                         self.var_specs_ram_gb.get(), self.var_ascii.get())
        line = trim_chatbox(line)
        self._last_specs = (line, now)
        return line

    def _afk_tag_if_needed(self, text):
        if not self.var_afk_tag_enabled.get():
            return text
        try:
            idle_s = get_idle_seconds()
            if idle_s >= self.get_int(self.var_afk_tag_after, self.cfg.get("afk_tag_after", 120), 10, 36000):
                tag = self.var_afk_tag_text.get().strip() or "[AFK]"
                candidate = (text + " " + tag).strip()
                return trim_chatbox(candidate)
        except Exception:
            pass
        return text

    def _compose_full(self, spotify_main, spotify_time_line):
        rot_mode = self.var_rot_mode.get()
        txts = []
        base_line = spotify_main
        base_line = self._afk_tag_if_needed(base_line)

        if self.var_rot_enabled.get() and self.current_rot_text:
            if rot_mode == "standalone":
                txts.append(self.current_rot_text)
            elif rot_mode == "prepend":
                line = f"{self.current_rot_text} {base_line}".strip()
                txts.append(trim_chatbox(line))
            elif rot_mode == "append":
                line = f"{base_line} {self.current_rot_text}".strip()
                txts.append(trim_chatbox(line))
            else:
                txts.append(self.current_rot_text)
                txts.append(base_line)
        else:
            txts.append(base_line)

        if self.var_time_second_line.get() and spotify_time_line:
            txts.append(spotify_time_line)

        specs = self._specs_line()
        if specs:
            txts.append(specs)

        clock = self._clock_line()
        if clock:
            txts.append(clock)
        return "\n".join([t for t in txts if t]).strip()

    def _update_preview(self):
        m, tline = self._render_spotify_lines(self.last_item, self.last_progress, self.last_duration)
        self.var_preview.set(self._compose_full(m, tline))

    def _on_start(self):
        if self.running: return
        try:
            self._ensure_osc()
            self.running = True
            self.last_message = ""; self.last_track_id = ""; self.rot_idx = 0
            self.next_rotate_at = time.monotonic(); self.current_rot_text = ""
            now = time.monotonic()
            afk_iv = max(5, self.get_int(self.var_afk_interval, self.cfg.get("anti_afk_interval", 240), 5, 3600))
            self.next_afk_at = now + afk_iv if self.var_afk_enabled.get() else now + 10**9
            self.worker = threading.Thread(target=self._loop, daemon=True); self.worker.start()
            self._log("Updater started")
        except Exception as e:
            self._log(f"Start error: {e}")

    def _on_stop(self):
        self.running = False; self._log("Updater stopped")

    def _loop(self):
        while self.running:
            try:
                if not self.tokens or token_expired(self.tokens):
                    try:
                        self.tokens = refresh_token(self.tokens)
                    except Exception as e:
                        self.lbl_auth.configure(text="Auth: required"); self._log(f"Token refresh failed: {e}")
                        time.sleep(max(1, self.get_int(self.var_update, 3, 1, 120))); continue

                pb = get_current_playback(self.tokens.get("access_token",""))
                if pb == "unauthorized":
                    self.lbl_auth.configure(text="Auth: required"); self._log("Access revoked or expired")
                    time.sleep(max(1, self.get_int(self.var_update, 3, 1, 120))); continue

                if not pb or not pb.get("item"):
                    self.lbl_pb.configure(text="Playback: none")
                    self.last_item = None; self.last_progress = 0; self.last_duration = 0
                else:
                    item = pb["item"]
                    self.last_item = item
                    self.last_progress = pb.get("progress_ms", 0)
                    self.last_duration = item.get("duration_ms", 0)
                    self.lbl_pb.configure(text="Playback: playing" if pb.get("is_playing", False) else "Playback: paused")

                spotify_main, time_line = self._render_spotify_lines(self.last_item, self.last_progress, self.last_duration)

                now = time.monotonic()
                rotated = False
                if self.var_rot_enabled.get() and len(self.rotation_items) > 0 and now >= self.next_rotate_at:
                    self.next_rotate_at = now + max(1, self.get_int(self.var_rot_interval, 6, 1, 3600))
                    it = self.rotation_items[self.rot_idx % len(self.rotation_items)]
                    self.rot_idx += 1
                    self.current_rot_text = self._render_rotation_item(it.get("text",""))
                    final_text = self._compose_full(spotify_main, time_line)
                    if final_text and self._send_chatbox(final_text):
                        self.last_message = final_text; rotated = True

                track_id = (self.last_item or {}).get("id","")
                combined = self._compose_full(spotify_main, time_line)
                if not rotated and (not self.var_only_changes.get() or combined != self.last_message or track_id != self.last_track_id):
                    if self._send_chatbox(combined):
                        self.last_message = combined; self.last_track_id = track_id

                if self.var_afk_enabled.get() and now >= self.next_afk_at:
                    if self._send_jump():
                        self._log("Anti-AFK jump")
                    afk_iv = max(5, self.get_int(self.var_afk_interval, self.cfg.get("anti_afk_interval", 240), 5, 3600))
                    self.next_afk_at = now + afk_iv

                self._update_preview()

            except Exception as e:
                self._log(f"Loop error: {e}")

            time.sleep(max(1, self.get_int(self.var_update, 3, 1, 120)))

    def _refresh_rot_list(self):
        self.listbox.delete(0, "end")
        for i, it in enumerate(self.rotation_items):
            txt = it.get("text","").replace("\n"," ")
            if len(txt) > 80: txt = txt[:77]+"…"
            self.listbox.insert("end", f"{i+1:02d} {txt}")

    def _load_selected_rotation_item(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]; it = self.rotation_items[idx]
        self.entry_item_text.delete("1.0","end"); self.entry_item_text.insert("1.0", it.get("text",""))

    def _on_rot_text_edit(self, _event=None):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        text = self.entry_item_text.get("1.0","end").strip()
        self.rotation_items[idx] = {"text": text}
        self._refresh_rot_list()
        self.listbox.select_set(idx)
        self._save_config()

    def _rot_add(self):
        text = self.entry_item_text.get("1.0","end").strip()
        if not text: return
        self.rotation_items.append({"text": text})
        self._refresh_rot_list(); self.listbox.select_clear(0, "end"); self.listbox.select_set(len(self.rotation_items)-1)
        self._save_config()

    def _rot_delete(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]; self.rotation_items.pop(idx)
        self._refresh_rot_list(); self._save_config()

    def _rot_up(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx == 0: return
        self.rotation_items[idx-1], self.rotation_items[idx] = self.rotation_items[idx], self.rotation_items[idx-1]
        self._refresh_rot_list(); self.listbox.select_set(idx-1)
        self._save_config()

    def _rot_down(self):
        sel = self.listbox.curselection()
        if not sel: return
        idx = sel[0]
        if idx >= len(self.rotation_items)-1: return
        self.rotation_items[idx+1], self.rotation_items[idx] = self.rotation_items[idx], self.rotation_items[idx+1]
        self._refresh_rot_list(); self.listbox.select_set(idx+1)
        self._save_config()

    def _update_status_loop(self):
        sp = detect_process_any(["spotify.exe","spotify"])
        self.lbl_sp.configure(text="Spotify: active" if sp else "Spotify: not found" if sp is False else "Spotify: unknown")
        vr = detect_process_any(["vrchat.exe","vrchat","vrchatclient.exe"])
        self.lbl_vr.configure(text="VRChat: active" if vr else "VRChat: not found" if vr is False else "VRChat: unknown")
        if not self.tokens:
            self.lbl_auth.configure(text="Auth: required")
        else:
            self.lbl_auth.configure(text="Auth: renew" if token_expired(self.tokens) else "Auth: ok")
        self.after(1200, self._update_status_loop)

def main():
    app = App()
    def on_close():
        try: app._save_config()
        except: pass
        app.running = False; app.destroy()
    app.protocol("WM_DELETE_WINDOW", on_close)
    app.mainloop()

if __name__ == "__main__":
    main()
