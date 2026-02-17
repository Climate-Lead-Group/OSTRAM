# run.py
# -*- coding: utf-8 -*-
"""
Runner para DVC en Windows con manejo de entorno Conda y parche temporal del dvc.yaml.

Características:
- Backup de dvc.yaml, reemplazo temporal de 'fecha' -> YYYY-MM-DD (cualquier aparición).
- Si el entorno Conda existe, NO lo recrea.
- Si el entorno existe, verifica dependencias e instala las faltantes:
    * conda-forge: pandas, numpy, openpyxl, pyyaml, xlsxwriter
    * pip: dvc, otoole
  (instala 'pip' en el entorno si hiciera falta).
- Inicializa repo DVC si falta (.dvc/).
- Ejecuta 'dvc pull' solo si hay remoto configurado.
- Ejecuta 'dvc repro'.
- Restaura el dvc.yaml desde el backup y ELIMINA el archivo .bak (siempre).
"""

import argparse
import datetime as dt
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
import json

# ---------- Config por defecto ----------
ENV_NAME_DEFAULT = "OG-MOMF-env"
ENV_FILE_DEFAULT = "environment.yaml"
DVC_FILE_DEFAULT = "dvc.yaml"

# Dependencias a verificar/instalar
CONDA_DEPS = {
    # módulo_python: paquete_conda
    "pandas": "pandas",
    "numpy": "numpy",
    "openpyxl": "openpyxl",
    "yaml": "pyyaml",          # PyYAML se importa como 'yaml'
    "xlsxwriter": "xlsxwriter"
}
PIP_DEPS = {
    # módulo_python: paquete_pip
    "dvc": "dvc",
    "otoole": "otoole>=1.1.1",
}

# ---------- Utilidades shell ----------
def run(cmd: str) -> None:
    # Set PYTHONHASHSEED for deterministic hash-based operations
    env = os.environ.copy()
    env['PYTHONHASHSEED'] = '0'
    subprocess.check_call(cmd, shell=True, env=env)

def check_tool_available(tool: str) -> None:
    try:
        subprocess.check_call(f"{tool} --version", shell=True,
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:
        raise RuntimeError(
            f"Requisito '{tool}' no encontrado en PATH. "
            f"Abre una Anaconda/Miniconda Prompt o instala la herramienta. Error original: {exc}"
        )

# ---------- Manejo de entorno Conda ----------
def env_exists(name: str) -> bool:
    """
    Devuelve True si existe un entorno conda cuyo directorio final se llama 'name'.
    Ej.: .../envs/OG-MOMF-env  -> name == 'OG-MOMF-env'
    Usa 'conda env list --json' y hace fallback al parseo de texto.
    """
    target = name.lower()

    # 1) Camino principal: JSON
    try:
        out = subprocess.check_output(
            ["conda", "env", "list", "--json"],
            text=True,
            stderr=subprocess.STDOUT
        )
        data = json.loads(out)
        envs = data.get("envs", []) or []
        return any(Path(p).name.lower() == target for p in envs)
    # Si conda es muy viejo o algo falla, pasamos al fallback
    except Exception:
        pass

    # 2) Fallback: parseo de texto de 'conda env list'
    try:
        txt = subprocess.check_output(
            ["conda", "env", "list"],
            text=True,
            stderr=subprocess.STDOUT
        )
        for line in txt.splitlines():
            line = line.strip()
            if not line or line.startswith(("#", "conda environments:")):
                continue
            # líneas típicas:
            # base                  *  C:\Users\...\anaconda3
            # OG-MOMF-env              C:\Users\...\envs\OG-MOMF-env
            parts = line.split()
            if not parts:
                continue
            # Si la segunda columna es '*', la primera es el nombre
            cand = parts[0].lower()
            if cand == target:
                return True
        return False
    except Exception:
        return False


def guess_env_name_from_yaml(env_file: str) -> str | None:
    p = Path(env_file)
    if not p.exists():
        return None
    try:
        # Parseo sencillo: buscar línea 'name: ...'
        for line in p.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.lower().startswith("name:"):
                val = line.split(":", 1)[1].strip().strip("'\"")
                return val or None
    except Exception:
        pass
    return None

def create_env_if_missing(env_name: str, env_file: str) -> None:
    if env_exists(env_name):
        print(f"Conda env '{env_name}' ya existe. No se recrea.")
        return
    print(f"Creando Conda env '{env_name}' desde {env_file} …")
    run(f"conda env create -n {env_name} -f {env_file} -y")

def ensure_pip_available(env_name: str) -> None:
    try:
        run(f"conda run -n {env_name} python -m pip --version")
    except subprocess.CalledProcessError:
        print("pip no encontrado en el entorno. Instalando 'pip' en el env…")
        run(f"conda install -n {env_name} pip -y")

def module_present(env_name: str, module: str) -> bool:
    code = (
        "import importlib,sys;"
        f"sys.exit(0) if importlib.util.find_spec('{module}') else sys.exit(1)"
    )
    try:
        run(f'conda run -n {env_name} python -c "{code}"')
        return True
    except subprocess.CalledProcessError:
        return False

def ensure_deps(env_name: str) -> None:
    """
    Verifica módulos en el entorno y los instala si faltan.
    - Conda (conda-forge) para stack de datos.
    - Pip para dvc/otoole.
    """
    # Asegura pip dentro del entorno si lo necesitaremos
    need_pip = any(not module_present(env_name, m) for m in list(PIP_DEPS.keys()))
    if need_pip:
        ensure_pip_available(env_name)

    # Conda deps
    missing_conda = [pkg for mod, pkg in CONDA_DEPS.items() if not module_present(env_name, mod)]
    if missing_conda:
        pkgs = " ".join(missing_conda)
        print(f"Instalando conda deps que faltan: {missing_conda}")
        run(f"conda install -n {env_name} -c conda-forge -y {pkgs}")

    # Pip deps
    missing_pip = [pkg for mod, pkg in PIP_DEPS.items() if not module_present(env_name, mod)]
    if missing_pip:
        for spec in missing_pip:
            print(f"Instalando pip dep que falta: {spec}")
            run(f"conda run -n {env_name} python -m pip install -U {spec}")

# ---------- DVC ----------
def is_dvc_repo() -> bool:
    return (Path(".dvc").is_dir())

def is_git_repo() -> bool:
    """Check if current directory is inside a Git repository."""
    return (Path(".git").is_dir())

def ensure_dvc_repo(env_name: str) -> None:
    if is_dvc_repo():
        print("DVC repository detected (.dvc/ found).")
        return

    # Check if we're in a Git repo to decide how to init DVC
    if is_git_repo():
        print("No hay repo DVC. Ejecutando `dvc init`…")
        run(f"conda run -n {env_name} dvc init")
    else:
        print("No hay repo Git. Ejecutando `dvc init --no-scm`…")
        run(f"conda run -n {env_name} dvc init --no-scm")

    if not is_dvc_repo():
        raise RuntimeError("Fallo al inicializar DVC (no se creó .dvc).")

def has_dvc_remote(env_name: str) -> bool:
    try:
        out = subprocess.check_output(f"conda run -n {env_name} dvc remote list",
                                      shell=True, stderr=subprocess.STDOUT)
        return bool(out.decode("utf-8", errors="ignore").strip())
    except subprocess.CalledProcessError:
        return False

def dvc_command(env_name: str, args: str) -> None:
    run(f"conda run -n {env_name} dvc {args}")

# ---------- Backup / parche dvc.yaml ----------
def backup_file(src: Path) -> Path:
    if not src.exists():
        raise FileNotFoundError(f"No se encontró {src} para hacer backup.")
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = src.with_suffix(src.suffix + f".bak.{ts}")
    shutil.copy2(src, bak)
    print(f"Backup creado: {bak.name}")
    return bak

def restore_and_delete_backup(backup_path: Path, target: Path) -> None:
    if backup_path and backup_path.exists():
        shutil.copy2(backup_path, target)
        print(f"Restaurado {target.name} desde backup: {backup_path.name}")
        try:
            backup_path.unlink()  # borrar el .bak
            print(f"Backup eliminado: {backup_path.name}")
        except Exception as e:
            print(f"Advertencia: no se pudo borrar el backup ({e})")
    else:
        print("Backup no encontrado; nada que restaurar/borrar.")

def patch_fecha_anywhere(dvc_path: Path, date_stamp: str) -> int:
    """
    Reemplaza TODAS las apariciones literales de 'fecha' por la fecha (YYYY-MM-DD).
    No usa regex, así también cubre '..._fecha.csv' (con guión bajo).
    """
    text = dvc_path.read_text(encoding="utf-8")
    count = text.count("fecha")
    if count:
        dvc_path.write_text(text.replace("fecha", date_stamp), encoding="utf-8", newline="\n")
    return count

# ---------- Main ----------
def main():
    parser = argparse.ArgumentParser(description="Runner de DVC con entorno Conda y parche temporal de dvc.yaml")
    parser.add_argument("--env-name", default=None, help="Nombre del entorno Conda (si no se pasa, intenta leer del YAML).")
    parser.add_argument("--env-file", default=ENV_FILE_DEFAULT, help="Ruta del environment.yaml.")
    parser.add_argument("--dvc-file", default=DVC_FILE_DEFAULT, help="Ruta del dvc.yaml a parchear.")
    parser.add_argument("--date", default=None, help="Fecha YYYY-MM-DD (por defecto hoy).")
    args = parser.parse_args()

    # Determinar env_name
    env_name = args.env_name or guess_env_name_from_yaml(args.env_file) or ENV_NAME_DEFAULT
    env_file = args.env_file
    dvc_file = Path(args.dvc_file).resolve()

    # Fecha
    if args.date:
        try:
            dt.date.fromisoformat(args.date)
        except ValueError:
            raise SystemExit("Formato de --date inválido. Usa YYYY-MM-DD (ej. 2025-08-21).")
        date_stamp = args.date
    else:
        date_stamp = dt.date.today().isoformat()

    print(f"Usando entorno: {env_name}")
    print(f"Usando date stamp: {date_stamp}")
    print(f"dvc.yaml: {dvc_file}")

    # Requisitos base
    check_tool_available("conda")

    backup_path = None
    try:
        # 1) Backup + parcheo de 'fecha' antes de cualquier cosa
        backup_path = backup_file(dvc_file)
        replaced = patch_fecha_anywhere(dvc_file, date_stamp)
        if replaced:
            print(f"Parche aplicado: {replaced} ocurrencia(s) de 'fecha' cambiadas por '{date_stamp}'.")
        else:
            print("No se encontraron ocurrencias de 'fecha' en dvc.yaml.")

        # 2) Entorno: crear si falta; si existe, NO recrear.
        create_env_if_missing(env_name, env_file)

        # 3) Verificar/instalar dependencias dentro del entorno
        ensure_deps(env_name)

        # 4) Asegurar repo DVC
        ensure_dvc_repo(env_name)

        # 5) Pull solo si hay remoto
        if has_dvc_remote(env_name):
            print("📥 dvc pull…")
            dvc_command(env_name, "pull")
        else:
            print("ℹ️ Sin remoto DVC configurado. Se omite `dvc pull`.")

        # 6) Reproducir pipeline
        print("🔄 dvc repro…")
        start_time = dt.datetime.now()
        dvc_command(env_name, "repro")
        end_time = dt.datetime.now()

        # Calculate and display duration
        duration = end_time - start_time
        total_seconds = int(duration.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        duration_str = []
        if hours > 0:
            duration_str.append(f"{hours}h")
        if minutes > 0 or hours > 0:
            duration_str.append(f"{minutes}m")
        duration_str.append(f"{seconds}s")

        print(f"✅ ¡Pipeline completado en {' '.join(duration_str)}!")

    finally:
        # 7) Restaurar dvc.yaml y borrar backup
        if backup_path:
            restore_and_delete_backup(backup_path, dvc_file)

if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Falla de comando (exit {e.returncode}): {e.cmd}", file=sys.stderr)
        sys.exit(e.returncode)
    except Exception as e:
        print(f"\n❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
