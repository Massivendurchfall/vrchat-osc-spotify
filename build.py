import sys
import os
import subprocess
import shutil
import importlib.util

def is_installed(modname):
    return importlib.util.find_spec(modname) is not None

def pip_install(pkgs):
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade"] + pkgs)

required = {
    "pyinstaller": "PyInstaller",
    "customtkinter": "customtkinter",
    "requests": "requests",
    "pythonosc": "python-osc",
    "psutil": "psutil",
    "certifi": "certifi",
    "urllib3": "urllib3",
    "idna": "idna",
    "charset_normalizer": "charset-normalizer"
}

to_install = []
for mod, pip_name in required.items():
    if not is_installed(mod):
        to_install.append(pip_name)

if to_install:
    pip_install(to_install)

import certifi
from PyInstaller import __main__ as pyimain
import PyInstaller

entry = sys.argv[1] if len(sys.argv) > 1 else "script.py"
if not os.path.isfile(entry):
    print(f"Entry script not found: {entry}")
    sys.exit(1)

name = "VRChatSpotifyStatus"
workpath = os.path.abspath("build_artifacts")
distpath = os.path.abspath("dist_artifacts")

if os.path.isdir(workpath):
    shutil.rmtree(workpath, ignore_errors=True)
if os.path.isdir(distpath):
    shutil.rmtree(distpath, ignore_errors=True)

cert_path = certifi.where()
add_data_sep = ";" if os.name == "nt" else ":"
add_data = f"{cert_path}{add_data_sep}certifi"

pyi_version = tuple(int(x) for x in PyInstaller.__version__.split(".")[:2])
collect_flag = "--collect-all"
if pyi_version < (6, 0):
    collect_flag = "--collect-data"

hidden_imports = [
    "tkinter",
    "customtkinter",
    "pythonosc",
    "psutil",
    "requests",
    "urllib3",
    "idna",
    "charset_normalizer",
    "certifi",
    "http.server",          # robust: explizit aufnehmen
    "socketserver"
]

args = [
    entry,
    "--noconfirm",
    "--clean",
    "--onefile",
    "--windowed",
    "--name", name,
    "--distpath", distpath,
    "--workpath", workpath,
    collect_flag, "customtkinter",
    "--add-data", add_data
]

for hi in hidden_imports:
    args += ["--hidden-import", hi]

pyimain.run(args)

exe_path = os.path.join(distpath, f"{name}.exe")
if os.path.isfile(exe_path):
    print(f"Built: {exe_path}")
else:
    print("Build finished, but .exe not found. Check PyInstaller output.")
