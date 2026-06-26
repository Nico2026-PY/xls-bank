r"""
Abrir_Procesador_Bancario - Launcher / Actualizador por GitHub Releases
=======================================================================

Qué hace:
- Revisa la última Release del repositorio configurado.
- Descarga el asset Procesador_Bancario_Windows.zip.
- Instala cada versión en %LOCALAPPDATA%\Procesador_Bancario\versions\vX.X.X.
- Abre Procesador_Bancario.exe.
- Si falla internet/GitHub, intenta abrir la última versión local instalada.

Repo público: no hace falta token.
Repo privado: crear variable de entorno GITHUB_TOKEN con permisos de lectura.
"""

from __future__ import annotations

import ctypes
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except Exception:  # pragma: no cover
    tk = None
    messagebox = None
    ttk = None


# ============================================================
# CONFIGURACIÓN PRINCIPAL
# ============================================================

APP_DISPLAY_NAME = "Procesador Bancario"

# Cambiar estos 2 valores por tu GitHub.
# Ejemplo:
# GITHUB_OWNER = "nicolas"
# GITHUB_REPO = "procesador-bancario"
GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"

# Nombre exacto del archivo que se adjunta en cada GitHub Release.
ASSET_NAME = "Procesador_Bancario_Windows.zip"

# Nombre exacto del .exe que debe existir dentro del zip.
APP_EXE_NAME = "Procesador_Bancario.exe"

# Instalación local por usuario. No requiere permisos de administrador.
LOCAL_ROOT = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Procesador_Bancario"
VERSIONS_DIR = LOCAL_ROOT / "versions"
CURRENT_FILE = LOCAL_ROOT / "current_version.txt"
LOG_FILE = LOCAL_ROOT / "launcher.log"
LOCK_FILE = LOCAL_ROOT / "update.lock"

# Seguridad de actualización.
OPEN_LOCAL_IF_UPDATE_FAILS = True
NETWORK_TIMEOUT = 45
KEEP_LAST_VERSIONS = 3


# ============================================================
# MODELO
# ============================================================

@dataclass
class ReleaseAsset:
    name: str
    asset_id: int
    size: int
    browser_download_url: str


@dataclass
class ReleaseInfo:
    tag_name: str
    name: str
    html_url: str
    asset: ReleaseAsset


# ============================================================
# UTILIDADES
# ============================================================

def log(message: str) -> None:
    LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    with LOG_FILE.open("a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")


def log_exception(prefix: str) -> None:
    log(prefix)
    log(traceback.format_exc())


def normalize_version(version: str) -> tuple[int, ...]:
    """Convierte v0.2.11 / 0.2.11 / v0.2.11-beta en tuplas comparables."""
    v = (version or "0.0.0").strip().lower()
    if v.startswith("v"):
        v = v[1:]
    parts: list[int] = []
    for chunk in v.replace("-", ".").split("."):
        digits = "".join(ch for ch in chunk if ch.isdigit())
        if digits == "":
            break
        parts.append(int(digits))
    return tuple(parts or [0])


def is_newer(remote: str, local: str) -> bool:
    r = normalize_version(remote)
    l = normalize_version(local)
    max_len = max(len(r), len(l))
    r = r + (0,) * (max_len - len(r))
    l = l + (0,) * (max_len - len(l))
    return r > l


def safe_version_dir_name(version: str) -> str:
    return (version or "0.0.0").replace("/", "_").replace("\\", "_").strip()


def read_local_version() -> str:
    if not CURRENT_FILE.exists():
        return "0.0.0"
    return CURRENT_FILE.read_text(encoding="utf-8").strip() or "0.0.0"


def write_local_version(version: str) -> None:
    LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
    CURRENT_FILE.write_text(version.strip(), encoding="utf-8")


def get_headers(accept: str = "application/vnd.github+json") -> dict[str, str]:
    headers = {
        "Accept": accept,
        "User-Agent": f"{APP_DISPLAY_NAME.replace(' ', '-')}-Launcher",
        "X-GitHub-Api-Version": "2022-11-28",
    }
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


def http_get_json(url: str) -> dict:
    req = urllib.request.Request(url, headers=get_headers())
    with urllib.request.urlopen(req, timeout=NETWORK_TIMEOUT) as response:
        return json.loads(response.read().decode("utf-8"))


def find_latest_release() -> ReleaseInfo:
    if "TU_USUARIO_GITHUB" in GITHUB_OWNER or "TU_REPOSITORIO" in GITHUB_REPO:
        raise RuntimeError(
            "Falta configurar GITHUB_OWNER y GITHUB_REPO dentro de launcher.py."
        )

    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
    data = http_get_json(api_url)
    assets = data.get("assets", [])

    selected = None
    for asset in assets:
        if asset.get("name") == ASSET_NAME:
            selected = asset
            break

    if not selected:
        available = ", ".join(a.get("name", "") for a in assets) or "sin assets"
        raise RuntimeError(
            f"No se encontró el asset '{ASSET_NAME}' en la última release. "
            f"Assets disponibles: {available}"
        )

    return ReleaseInfo(
        tag_name=data.get("tag_name", "0.0.0"),
        name=data.get("name", ""),
        html_url=data.get("html_url", ""),
        asset=ReleaseAsset(
            name=selected["name"],
            asset_id=int(selected["id"]),
            size=int(selected.get("size", 0)),
            browser_download_url=selected["browser_download_url"],
        ),
    )


def download_asset(asset: ReleaseAsset, destination: Path, progress: Callable[[str, Optional[int]], None]) -> None:
    """Descarga público por browser_download_url o privado por API usando GITHUB_TOKEN."""
    token = os.environ.get("GITHUB_TOKEN", "").strip()

    if token:
        url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/assets/{asset.asset_id}"
        headers = get_headers(accept="application/octet-stream")
    else:
        url = asset.browser_download_url
        headers = {"User-Agent": f"{APP_DISPLAY_NAME.replace(' ', '-')}-Launcher"}

    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=NETWORK_TIMEOUT) as response:
        total = int(response.headers.get("Content-Length") or asset.size or 0)
        downloaded = 0
        destination.parent.mkdir(parents=True, exist_ok=True)

        with destination.open("wb") as f:
            while True:
                chunk = response.read(1024 * 256)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                pct = int(downloaded * 100 / total) if total else None
                progress("Descargando actualización...", pct)


def install_zip(zip_path: Path, version: str, progress: Callable[[str, Optional[int]], None]) -> Path:
    version_name = safe_version_dir_name(version)
    version_dir = VERSIONS_DIR / version_name
    temp_extract = VERSIONS_DIR / f"_tmp_{version_name}_{int(time.time())}"

    if temp_extract.exists():
        shutil.rmtree(temp_extract, ignore_errors=True)
    temp_extract.mkdir(parents=True, exist_ok=True)

    progress("Instalando actualización...", None)
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(temp_extract)

    exe_candidates = list(temp_extract.rglob(APP_EXE_NAME))
    if not exe_candidates:
        shutil.rmtree(temp_extract, ignore_errors=True)
        raise RuntimeError(
            f"El zip no contiene '{APP_EXE_NAME}'. Revisá el nombre del ejecutable dentro del zip."
        )

    if version_dir.exists():
        shutil.rmtree(version_dir, ignore_errors=True)
    temp_extract.rename(version_dir)

    write_local_version(version)
    cleanup_old_versions()
    return version_dir


def cleanup_old_versions() -> None:
    try:
        if not VERSIONS_DIR.exists():
            return
        dirs = [p for p in VERSIONS_DIR.iterdir() if p.is_dir() and not p.name.startswith("_tmp_")]
        dirs.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        for old in dirs[KEEP_LAST_VERSIONS:]:
            shutil.rmtree(old, ignore_errors=True)
    except Exception:
        log_exception("No se pudieron limpiar versiones anteriores")


def get_current_app_exe() -> Optional[Path]:
    version = read_local_version()
    version_dir = VERSIONS_DIR / safe_version_dir_name(version)
    candidates = list(version_dir.rglob(APP_EXE_NAME))
    return candidates[0] if candidates else None


def launch_app(exe_path: Path) -> None:
    log(f"Abriendo app: {exe_path}")
    subprocess.Popen([str(exe_path)], cwd=str(exe_path.parent), close_fds=True)


def show_error(title: str, message: str) -> None:
    log(f"ERROR - {title}: {message}")
    if messagebox:
        messagebox.showerror(title, message)
    else:
        print(f"{title}: {message}", file=sys.stderr)


class SimpleLock:
    def __enter__(self):
        LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
        if LOCK_FILE.exists():
            try:
                age = time.time() - LOCK_FILE.stat().st_mtime
                if age < 120:
                    raise RuntimeError("Ya hay una actualización en curso. Esperá unos segundos y abrí de nuevo.")
            except FileNotFoundError:
                pass
        LOCK_FILE.write_text(str(os.getpid()), encoding="utf-8")
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            LOCK_FILE.unlink(missing_ok=True)
        except Exception:
            pass


# ============================================================
# FLUJO PRINCIPAL
# ============================================================

def run_update_flow(progress: Callable[[str, Optional[int]], None]) -> None:
    LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
    VERSIONS_DIR.mkdir(parents=True, exist_ok=True)

    with SimpleLock():
        progress("Revisando actualizaciones...", None)
        local_version = read_local_version()
        log(f"Versión local: {local_version}")

        try:
            release = find_latest_release()
            remote_version = release.tag_name
            log(f"Versión GitHub: {remote_version}")

            current_exe = get_current_app_exe()
            if is_newer(remote_version, local_version) or not current_exe:
                progress(f"Nueva versión encontrada: {remote_version}", None)
                with tempfile.TemporaryDirectory() as td:
                    zip_path = Path(td) / ASSET_NAME
                    download_asset(release.asset, zip_path, progress)
                    install_zip(zip_path, remote_version, progress)
            else:
                progress("Ya tenés la última versión.", 100)

        except Exception:
            log_exception("Falló la búsqueda/instalación de actualización")
            if not OPEN_LOCAL_IF_UPDATE_FAILS:
                raise

        exe = get_current_app_exe()
        if not exe or not exe.exists():
            raise RuntimeError(
                "No se encontró una versión local instalada. "
                "Creá una Release en GitHub con el zip correcto y volvé a abrir el launcher."
            )

        progress("Abriendo aplicación...", 100)
        launch_app(exe)


# ============================================================
# INTERFAZ DEL LAUNCHER
# ============================================================

class LauncherWindow:
    def __init__(self) -> None:
        if not tk:
            raise RuntimeError("Tkinter no está disponible")

        self.root = tk.Tk()
        self.root.title(f"{APP_DISPLAY_NAME} - Actualizador")
        self.root.geometry("500x220")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f7fb")

        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("procesador.bancario.launcher")
        except Exception:
            pass

        self.title_label = tk.Label(
            self.root,
            text=APP_DISPLAY_NAME,
            font=("Segoe UI", 18, "bold"),
            bg="#f5f7fb",
            fg="#12315f",
        )
        self.title_label.pack(pady=(28, 8))

        self.status_var = tk.StringVar(value="Preparando actualizador...")
        self.status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg="#f5f7fb",
            fg="#4f5f75",
        )
        self.status_label.pack(pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            self.root,
            orient="horizontal",
            mode="indeterminate",
            length=380,
        )
        self.progress_bar.pack(pady=(0, 12))
        self.progress_bar.start(12)

        self.footer_label = tk.Label(
            self.root,
            text="Actualización automática por GitHub Releases",
            font=("Segoe UI", 8),
            bg="#f5f7fb",
            fg="#8190a5",
        )
        self.footer_label.pack(side="bottom", pady=14)

    def progress(self, message: str, percent: Optional[int] = None) -> None:
        def update_ui():
            self.status_var.set(message if percent is None else f"{message} {percent}%")
            if percent is None:
                try:
                    self.progress_bar.configure(mode="indeterminate")
                    self.progress_bar.start(12)
                except Exception:
                    pass
            else:
                try:
                    self.progress_bar.stop()
                    self.progress_bar.configure(mode="determinate", maximum=100, value=percent)
                except Exception:
                    pass
            self.root.update_idletasks()
        self.root.after(0, update_ui)

    def start(self) -> None:
        def worker() -> None:
            try:
                run_update_flow(self.progress)
                time.sleep(0.6)
                self.root.after(0, self.root.destroy)
            except urllib.error.HTTPError as e:
                msg = f"GitHub respondió error {e.code}. Revisá repo, release, permisos o token."
                log_exception(msg)
                self.root.after(0, lambda: show_error("Error de actualización", msg))
                self.root.after(0, self.root.destroy)
            except Exception as e:
                msg = str(e)
                log_exception("Error general del launcher")
                self.root.after(0, lambda: show_error("Error de actualización", msg))
                self.root.after(0, self.root.destroy)

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()
        self.root.mainloop()


def main() -> None:
    log("Launcher iniciado")
    if tk:
        LauncherWindow().start()
    else:
        run_update_flow(lambda msg, pct=None: print(msg if pct is None else f"{msg} {pct}%"))


if __name__ == "__main__":
    main()
