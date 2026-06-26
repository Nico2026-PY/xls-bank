r"""
XlsBank - Launcher / Actualizador profesional por GitHub Releases
=================================================================

Objetivo:
- Instalar y actualizar la app SIN permisos de administrador.
- Descargar la ultima version publicada en GitHub Releases.
- Abrir la ultima version local si no hay internet.
- Evitar descargar si ya esta actualizado.
- Mantener backups/versiones anteriores.
- Validar el ZIP antes de instalar.
- Guardar logs para diagnostico.

Repo publico: no necesita token.
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
# CONFIGURACION PRINCIPAL
# ============================================================

APP_DISPLAY_NAME = "XlsBank"
APP_SUBTITLE = "Procesador Bancario"

GITHUB_OWNER = "Nico2026-PY"
GITHUB_REPO = "xls-bank"

# IMPORTANTE: mantener estos nombres mientras no quieras cambiar el launcher
ASSET_NAME = "Procesador_Bancario_Windows.zip"
APP_EXE_NAME = "Procesador_Bancario.exe"

# Instalacion por usuario: NO requiere administrador
LOCAL_ROOT = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Procesador_Bancario"
VERSIONS_DIR = LOCAL_ROOT / "versions"
BACKUP_DIR = LOCAL_ROOT / "backups"
CURRENT_FILE = LOCAL_ROOT / "current_version.txt"
LAST_GOOD_FILE = LOCAL_ROOT / "last_good_version.txt"
LOCK_FILE = LOCAL_ROOT / "update.lock"

LAUNCHER_ROOT = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Procesador_Bancario_Launcher"
LOGS_DIR = LAUNCHER_ROOT / "logs"
SETTINGS_FILE = LAUNCHER_ROOT / "launcher_settings.json"

# Comportamiento
DEFAULT_CHANNEL = "stable"  # stable o beta
OPEN_LOCAL_IF_UPDATE_FAILS = True
NETWORK_TIMEOUT = 45
KEEP_LAST_VERSIONS = 3
KEEP_LAST_BACKUPS = 3


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
    prerelease: bool
    asset: ReleaseAsset


@dataclass
class LauncherSettings:
    channel: str = DEFAULT_CHANNEL


# ============================================================
# LOGS Y CONFIGURACION
# ============================================================

def today_log_file() -> Path:
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    return LOGS_DIR / f"launcher_{time.strftime('%Y%m%d')}.log"


def log(message: str) -> None:
    try:
        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with today_log_file().open("a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


def log_exception(prefix: str) -> None:
    log(prefix)
    log(traceback.format_exc())


def load_settings() -> LauncherSettings:
    LAUNCHER_ROOT.mkdir(parents=True, exist_ok=True)

    # Variable de entorno opcional para pruebas: XLSBANK_CHANNEL=beta
    env_channel = os.environ.get("XLSBANK_CHANNEL", "").strip().lower()
    if env_channel in {"stable", "beta"}:
        return LauncherSettings(channel=env_channel)

    if not SETTINGS_FILE.exists():
        return LauncherSettings()

    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        channel = str(data.get("channel", DEFAULT_CHANNEL)).strip().lower()
        if channel not in {"stable", "beta"}:
            channel = DEFAULT_CHANNEL
        return LauncherSettings(channel=channel)
    except Exception:
        log_exception("No se pudo leer launcher_settings.json")
        return LauncherSettings()


def save_settings(settings: LauncherSettings) -> None:
    LAUNCHER_ROOT.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(
        json.dumps({"channel": settings.channel}, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


# ============================================================
# UTILIDADES DE VERSION
# ============================================================

def normalize_version(version: str) -> tuple[int, ...]:
    """Convierte v0.2.11 / 0.2.11 / v0.2.12-beta en tuplas comparables."""
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
    return r + (0,) * (max_len - len(r)) > l + (0,) * (max_len - len(l))


def safe_version_dir_name(version: str) -> str:
    return (version or "0.0.0").replace("/", "_").replace("\\", "_").strip()


def read_text_file(path: Path, default: str = "") -> str:
    if not path.exists():
        return default
    try:
        return path.read_text(encoding="utf-8").strip() or default
    except Exception:
        return default


def read_local_version() -> str:
    return read_text_file(CURRENT_FILE, "0.0.0")


def read_last_good_version() -> str:
    return read_text_file(LAST_GOOD_FILE, "")


def write_local_version(version: str) -> None:
    LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
    CURRENT_FILE.write_text(version.strip(), encoding="utf-8")


def write_last_good_version(version: str) -> None:
    LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
    LAST_GOOD_FILE.write_text(version.strip(), encoding="utf-8")


# ============================================================
# GITHUB RELEASES
# ============================================================

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


def http_get_json(url: str) -> object:
    req = urllib.request.Request(url, headers=get_headers())
    with urllib.request.urlopen(req, timeout=NETWORK_TIMEOUT) as response:
        return json.loads(response.read().decode("utf-8"))


def release_from_api_data(data: dict) -> Optional[ReleaseInfo]:
    assets = data.get("assets", []) or []
    selected = None
    for asset in assets:
        if asset.get("name") == ASSET_NAME:
            selected = asset
            break
    if not selected:
        return None

    return ReleaseInfo(
        tag_name=data.get("tag_name", "0.0.0"),
        name=data.get("name", ""),
        html_url=data.get("html_url", ""),
        prerelease=bool(data.get("prerelease", False)),
        asset=ReleaseAsset(
            name=selected["name"],
            asset_id=int(selected["id"]),
            size=int(selected.get("size", 0)),
            browser_download_url=selected["browser_download_url"],
        ),
    )


def find_latest_release(channel: str) -> ReleaseInfo:
    if "TU_USUARIO_GITHUB" in GITHUB_OWNER or "TU_REPOSITORIO" in GITHUB_REPO:
        raise RuntimeError("Falta configurar GITHUB_OWNER y GITHUB_REPO dentro de launcher.py.")

    channel = (channel or DEFAULT_CHANNEL).lower().strip()

    if channel == "stable":
        api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
        data = http_get_json(api_url)
        if not isinstance(data, dict):
            raise RuntimeError("Respuesta inesperada de GitHub Releases.")
        release = release_from_api_data(data)
        if not release:
            assets = data.get("assets", []) or []
            available = ", ".join(a.get("name", "") for a in assets) or "sin assets"
            raise RuntimeError(f"No se encontró el asset '{ASSET_NAME}'. Assets disponibles: {available}")
        return release

    # Canal beta: toma la release mas nueva, incluyendo pre-release, que tenga el asset esperado.
    api_url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases?per_page=20"
    data = http_get_json(api_url)
    if not isinstance(data, list):
        raise RuntimeError("Respuesta inesperada de GitHub Releases.")

    for item in data:
        if not isinstance(item, dict) or item.get("draft"):
            continue
        release = release_from_api_data(item)
        if release:
            return release

    raise RuntimeError(f"No se encontró ninguna release beta/estable con el asset '{ASSET_NAME}'.")


def download_asset(asset: ReleaseAsset, destination: Path, progress: Callable[[str, Optional[int]], None]) -> None:
    token = os.environ.get("GITHUB_TOKEN", "").strip()
    if token:
        url = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/assets/{asset.asset_id}"
        headers = get_headers(accept="application/octet-stream")
    else:
        url = asset.browser_download_url
        headers = {"User-Agent": f"{APP_DISPLAY_NAME.replace(' ', '-')}-Launcher"}

    log(f"Descargando asset: {asset.name} ({asset.size} bytes)")
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


# ============================================================
# VALIDACION, BACKUP E INSTALACION
# ============================================================

def validate_zip(zip_path: Path) -> None:
    if not zip_path.exists() or zip_path.stat().st_size == 0:
        raise RuntimeError("El ZIP descargado está vacío o no existe.")

    if not zipfile.is_zipfile(zip_path):
        raise RuntimeError("El archivo descargado no es un ZIP válido.")

    with zipfile.ZipFile(zip_path, "r") as zf:
        names = zf.namelist()

    has_exe = any(Path(name).name.lower() == APP_EXE_NAME.lower() for name in names)
    has_version = any(Path(name).name.lower() == "version.txt" for name in names)
    has_assets = any("assets/" in name.replace("\\", "/").lower() for name in names)

    missing = []
    if not has_exe:
        missing.append(APP_EXE_NAME)
    if not has_version:
        missing.append("VERSION.txt")
    if not has_assets:
        missing.append("assets/")

    if missing:
        raise RuntimeError("El ZIP no tiene la estructura esperada. Falta: " + ", ".join(missing))


def get_version_dir(version: str) -> Path:
    return VERSIONS_DIR / safe_version_dir_name(version)


def find_app_exe_in_dir(base_dir: Path) -> Optional[Path]:
    candidates = list(base_dir.rglob(APP_EXE_NAME)) if base_dir.exists() else []
    return candidates[0] if candidates else None


def get_current_app_exe() -> Optional[Path]:
    return find_app_exe_in_dir(get_version_dir(read_local_version()))


def get_last_good_app_exe() -> Optional[Path]:
    version = read_last_good_version()
    if not version:
        return None
    return find_app_exe_in_dir(get_version_dir(version))


def backup_current_version() -> None:
    current_version = read_local_version()
    current_dir = get_version_dir(current_version)
    if current_version in {"", "0.0.0"} or not current_dir.exists():
        return

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    backup_path = BACKUP_DIR / safe_version_dir_name(current_version)
    if backup_path.exists():
        shutil.rmtree(backup_path, ignore_errors=True)

    try:
        shutil.copytree(current_dir, backup_path)
        log(f"Backup creado: {backup_path}")
        cleanup_backups()
    except Exception:
        log_exception("No se pudo crear backup de la versión actual")


def cleanup_backups() -> None:
    try:
        if not BACKUP_DIR.exists():
            return
        dirs = [p for p in BACKUP_DIR.iterdir() if p.is_dir()]
        dirs.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        for old in dirs[KEEP_LAST_BACKUPS:]:
            shutil.rmtree(old, ignore_errors=True)
    except Exception:
        log_exception("No se pudieron limpiar backups anteriores")


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


def install_zip(zip_path: Path, version: str, progress: Callable[[str, Optional[int]], None]) -> Path:
    validate_zip(zip_path)

    version_name = safe_version_dir_name(version)
    version_dir = get_version_dir(version)
    temp_extract = VERSIONS_DIR / f"_tmp_{version_name}_{int(time.time())}"

    if temp_extract.exists():
        shutil.rmtree(temp_extract, ignore_errors=True)
    temp_extract.mkdir(parents=True, exist_ok=True)

    progress("Instalando actualización...", None)
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(temp_extract)

    exe_path = find_app_exe_in_dir(temp_extract)
    if not exe_path:
        shutil.rmtree(temp_extract, ignore_errors=True)
        raise RuntimeError(f"El ZIP no contiene '{APP_EXE_NAME}'.")

    backup_current_version()

    if version_dir.exists():
        shutil.rmtree(version_dir, ignore_errors=True)
    temp_extract.rename(version_dir)

    write_local_version(version)
    write_last_good_version(version)
    cleanup_old_versions()
    log(f"Instalación correcta: {version_dir}")
    return version_dir


def restore_last_good_if_needed() -> Optional[Path]:
    exe = get_current_app_exe()
    if exe and exe.exists():
        return exe

    last_good_exe = get_last_good_app_exe()
    if last_good_exe and last_good_exe.exists():
        write_local_version(read_last_good_version())
        log("Restaurada última versión buena.")
        return last_good_exe

    return None


# ============================================================
# APERTURA DE APP Y LOCK
# ============================================================

def launch_app(exe_path: Path) -> None:
    if not exe_path.exists():
        raise RuntimeError(f"No se encontró el ejecutable: {exe_path}")
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
                if age < 180:
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

def run_update_flow(progress: Callable[[str, Optional[int]], None], *, force_reinstall: bool = False) -> None:
    LOCAL_ROOT.mkdir(parents=True, exist_ok=True)
    VERSIONS_DIR.mkdir(parents=True, exist_ok=True)
    settings = load_settings()

    with SimpleLock():
        local_version = read_local_version()
        progress(f"Buscando actualizaciones... Canal: {settings.channel}", None)
        log(f"Launcher iniciado | canal={settings.channel} | local={local_version} | force={force_reinstall}")

        try:
            release = find_latest_release(settings.channel)
            remote_version = release.tag_name
            current_exe = get_current_app_exe()
            log(f"Versión GitHub: {remote_version} | prerelease={release.prerelease}")

            if force_reinstall:
                progress(f"Reinstalando versión {remote_version}...", None)
                should_install = True
            elif is_newer(remote_version, local_version) or not current_exe:
                progress(f"Nueva versión encontrada: {remote_version}", None)
                should_install = True
            else:
                progress(f"Ya tenés la última versión ({local_version}).", 100)
                should_install = False

            if should_install:
                with tempfile.TemporaryDirectory() as td:
                    zip_path = Path(td) / ASSET_NAME
                    download_asset(release.asset, zip_path, progress)
                    progress("Validando paquete descargado...", None)
                    validate_zip(zip_path)
                    install_zip(zip_path, remote_version, progress)

        except Exception as exc:
            log_exception("Falló la búsqueda/instalación de actualización")
            if OPEN_LOCAL_IF_UPDATE_FAILS:
                progress("No se pudo actualizar. Abriendo última versión instalada...", None)
            else:
                raise exc

        exe = restore_last_good_if_needed()
        if not exe or not exe.exists():
            raise RuntimeError(
                "No se encontró una versión local instalada. "
                "Creá una Release en GitHub con el ZIP correcto y volvé a abrir el launcher."
            )

        progress(f"Abriendo {APP_DISPLAY_NAME}...", 100)
        launch_app(exe)


# ============================================================
# INTERFAZ DEL LAUNCHER
# ============================================================

class LauncherWindow:
    def __init__(self) -> None:
        if not tk:
            raise RuntimeError("Tkinter no está disponible")

        self.settings = load_settings()
        self.running = False
        self.root = tk.Tk()
        self.root.title(f"{APP_DISPLAY_NAME} - Actualizador")
        self.root.geometry("560x300")
        self.root.resizable(False, False)
        self.root.configure(bg="#f5f7fb")

        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("xlsbank.launcher")
        except Exception:
            pass

        self.title_label = tk.Label(
            self.root,
            text=APP_DISPLAY_NAME,
            font=("Segoe UI", 22, "bold"),
            bg="#f5f7fb",
            fg="#12315f",
        )
        self.title_label.pack(pady=(24, 2))

        self.subtitle_label = tk.Label(
            self.root,
            text=APP_SUBTITLE,
            font=("Segoe UI", 10),
            bg="#f5f7fb",
            fg="#64748b",
        )
        self.subtitle_label.pack(pady=(0, 14))

        self.status_var = tk.StringVar(value="Preparando actualizador...")
        self.status_label = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bg="#f5f7fb",
            fg="#334155",
        )
        self.status_label.pack(pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            self.root,
            orient="horizontal",
            mode="indeterminate",
            length=430,
        )
        self.progress_bar.pack(pady=(0, 12))
        self.progress_bar.start(12)

        self.info_var = tk.StringVar(value=self._info_text())
        self.info_label = tk.Label(
            self.root,
            textvariable=self.info_var,
            font=("Segoe UI", 8),
            bg="#f5f7fb",
            fg="#8190a5",
        )
        self.info_label.pack(pady=(0, 10))

        btn_frame = tk.Frame(self.root, bg="#f5f7fb")
        btn_frame.pack(pady=(0, 8))

        self.reinstall_btn = ttk.Button(btn_frame, text="Reinstalar última versión", command=self.reinstall)
        self.reinstall_btn.grid(row=0, column=0, padx=5)

        self.channel_btn = ttk.Button(btn_frame, text=self._channel_button_text(), command=self.toggle_channel)
        self.channel_btn.grid(row=0, column=1, padx=5)

        self.logs_btn = ttk.Button(btn_frame, text="Abrir logs", command=self.open_logs_folder)
        self.logs_btn.grid(row=0, column=2, padx=5)

        self.footer_label = tk.Label(
            self.root,
            text="Actualización automática por GitHub Releases · Instalación sin administrador",
            font=("Segoe UI", 8),
            bg="#f5f7fb",
            fg="#94a3b8",
        )
        self.footer_label.pack(side="bottom", pady=12)

    def _info_text(self) -> str:
        return f"Local: {read_local_version()}  |  Canal: {self.settings.channel}  |  Repo: {GITHUB_OWNER}/{GITHUB_REPO}"

    def _channel_button_text(self) -> str:
        return "Canal: Estable" if self.settings.channel == "stable" else "Canal: Beta"

    def set_running(self, value: bool) -> None:
        self.running = value
        state = "disabled" if value else "normal"
        self.reinstall_btn.configure(state=state)
        self.channel_btn.configure(state=state)

    def progress(self, message: str, percent: Optional[int] = None) -> None:
        def update_ui() -> None:
            self.status_var.set(message if percent is None else f"{message} {percent}%")
            self.info_var.set(self._info_text())
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

    def start_update_thread(self, *, force_reinstall: bool = False) -> None:
        if self.running:
            return
        self.set_running(True)

        def worker() -> None:
            try:
                run_update_flow(self.progress, force_reinstall=force_reinstall)
                time.sleep(0.7)
                self.root.after(0, self.root.destroy)
            except urllib.error.HTTPError as e:
                msg = f"GitHub respondió error {e.code}. Revisá repo, release, permisos o token."
                log_exception(msg)
                self.root.after(0, lambda: show_error("Error de actualización", msg))
                self.root.after(0, lambda: self.set_running(False))
            except Exception as e:
                msg = str(e)
                log_exception("Error general del launcher")
                self.root.after(0, lambda: show_error("Error de actualización", msg))
                self.root.after(0, lambda: self.set_running(False))

        threading.Thread(target=worker, daemon=True).start()

    def reinstall(self) -> None:
        self.start_update_thread(force_reinstall=True)

    def toggle_channel(self) -> None:
        self.settings.channel = "beta" if self.settings.channel == "stable" else "stable"
        save_settings(self.settings)
        self.channel_btn.configure(text=self._channel_button_text())
        self.info_var.set(self._info_text())
        messagebox.showinfo(
            "Canal actualizado",
            f"Canal seleccionado: {self.settings.channel}\n\nCerrá y abrí de nuevo para usar este canal, o tocá Reinstalar última versión.",
        )

    def open_logs_folder(self) -> None:
        LOGS_DIR.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(LOGS_DIR))  # type: ignore[attr-defined]
        except Exception:
            messagebox.showinfo("Logs", str(LOGS_DIR))

    def start(self, *, force_reinstall: bool = False) -> None:
        self.root.after(600, lambda: self.start_update_thread(force_reinstall=force_reinstall))
        self.root.mainloop()


# ============================================================
# MAIN
# ============================================================

def parse_args() -> dict[str, bool]:
    args = {a.lower().strip() for a in sys.argv[1:]}

    if "--beta" in args:
        save_settings(LauncherSettings(channel="beta"))
    if "--stable" in args:
        save_settings(LauncherSettings(channel="stable"))

    return {
        "force_reinstall": "--reinstall" in args or "/reinstall" in args,
    }


def main() -> None:
    parsed = parse_args()
    log("Launcher ejecutado")
    if tk:
        LauncherWindow().start(force_reinstall=parsed["force_reinstall"])
    else:
        run_update_flow(
            lambda msg, pct=None: print(msg if pct is None else f"{msg} {pct}%"),
            force_reinstall=parsed["force_reinstall"],
        )


if __name__ == "__main__":
    main()
