# run.py
# -*- coding: utf-8 -*-
"""
Ejecutor DVC para Windows con gestión de entorno Conda y parcheo temporal de dvc.yaml.

Autor: Climate Lead Group, Andrey Salazar-Vargas

Funcionalidades:
- Respaldo de dvc.yaml, reemplazo temporal de 'DATEPLACEHOLDER' -> YYYY-MM-DD (cualquier ocurrencia).
- Si el entorno Conda ya existe, NO se recrea.
- Si el entorno existe, verifica dependencias e instala las faltantes:
    * conda-forge: pandas, numpy, openpyxl, pyyaml, xlsxwriter
    * pip: dvc, otoole
  (instala 'pip' en el entorno si es necesario).
- Inicializa el repositorio DVC si no existe (.dvc/).
- Ejecuta 'dvc pull' solo si hay un remoto configurado.
- Ejecuta 'dvc repro'.
- Restaura dvc.yaml desde el respaldo y ELIMINA el archivo .bak (siempre).
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
ENV_NAME_DEFAULT = "OSTRAM-env"
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
    # Fijar PYTHONHASHSEED para operaciones determinísticas basadas en hash
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
            f"Abra un Anaconda/Miniconda Prompt o instale la herramienta. Error original: {exc}"
        )

# ---------- Gestión de entorno Conda ----------
def env_exists(name: str) -> bool:
    """
    Retorna True si existe un entorno conda cuyo directorio final coincide con 'name'.
    Ej.: .../envs/OSTRAM-env  -> name == 'OSTRAM-env'
    Usa 'conda env list --json' con respaldo de parseo de texto.
    """
    target = name.lower()

    # 1) Ruta principal: JSON
    try:
        out = subprocess.check_output(
            ["conda", "env", "list", "--json"],
            text=True,
            stderr=subprocess.STDOUT
        )
        data = json.loads(out)
        envs = data.get("envs", []) or []
        return any(Path(p).name.lower() == target for p in envs)
    # Si conda es muy antiguo o algo falla, recurrir al parseo de texto
    except Exception:
        pass

    # 2) Respaldo: parseo de texto de 'conda env list'
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
            # OSTRAM-env              C:\Users\...\envs\OSTRAM-env
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
        # Parseo simple: buscar línea 'name: ...'
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
        print(f"El entorno Conda '{env_name}' ya existe. No se recrea.")
        return
    print(f"Creando entorno Conda '{env_name}' desde {env_file} …")
    run(f"conda env create -n {env_name} -f {env_file} -y")

def ensure_pip_available(env_name: str) -> None:
    try:
        run(f"conda run -n {env_name} python -m pip --version")
    except subprocess.CalledProcessError:
        print("pip no encontrado en el entorno. Instalando 'pip' en el entorno…")
        run(f"conda install -n {env_name} pip -y")

def module_present(env_name: str, module: str) -> bool:
    code = (
        "import importlib.util,sys;"
        f"sys.exit(0) if importlib.util.find_spec('{module}') else sys.exit(1)"
    )
    try:
        run(f'conda run -n {env_name} python -c "{code}"')
        return True
    except subprocess.CalledProcessError:
        return False

def ensure_deps(env_name: str) -> None:
    """
    Verifica módulos en el entorno e instala los faltantes.
    - Conda (conda-forge) para el stack de datos.
    - Pip para dvc/otoole.
    """
    # Asegurar que pip esté disponible en el entorno si lo necesitamos
    need_pip = any(not module_present(env_name, m) for m in list(PIP_DEPS.keys()))
    if need_pip:
        ensure_pip_available(env_name)

    # Dependencias Conda
    missing_conda = [pkg for mod, pkg in CONDA_DEPS.items() if not module_present(env_name, mod)]
    if missing_conda:
        pkgs = " ".join(missing_conda)
        print(f"Instalando dependencias conda faltantes: {missing_conda}")
        run(f"conda install -n {env_name} -c conda-forge -y {pkgs}")

    # Dependencias Pip
    missing_pip = [pkg for mod, pkg in PIP_DEPS.items() if not module_present(env_name, mod)]
    if missing_pip:
        for spec in missing_pip:
            print(f"Instalando dependencia pip faltante: {spec}")
            run(f"conda run -n {env_name} python -m pip install -U {spec}")

# ---------- DVC ----------
def is_dvc_repo() -> bool:
    return (Path(".dvc").is_dir())

def is_git_repo() -> bool:
    """Verifica si el directorio actual está dentro de un repositorio Git."""
    return (Path(".git").is_dir())

def ensure_dvc_repo(env_name: str) -> None:
    if is_dvc_repo():
        print("Repositorio DVC detectado (.dvc/ encontrado).")
        return

    # Verificar si estamos en un repo Git para decidir cómo inicializar DVC
    if is_git_repo():
        print("No se encontró repo DVC. Ejecutando `dvc init`…")
        run(f"conda run -n {env_name} dvc init")
    else:
        print("No se encontró repo Git. Ejecutando `dvc init --no-scm`…")
        run(f"conda run -n {env_name} dvc init --no-scm")

    if not is_dvc_repo():
        raise RuntimeError("Error al inicializar DVC (.dvc no fue creado).")

def has_dvc_remote(env_name: str) -> bool:
    try:
        out = subprocess.check_output(f"conda run -n {env_name} dvc remote list",
                                      shell=True, stderr=subprocess.STDOUT)
        return bool(out.decode("utf-8", errors="ignore").strip())
    except subprocess.CalledProcessError:
        return False

def dvc_command(env_name: str, args: str) -> None:
    run(f"conda run -n {env_name} dvc {args}")

# ---------- Respaldo / parcheo de dvc.yaml ----------
def backup_file(src: Path) -> Path:
    if not src.exists():
        raise FileNotFoundError(f"{src} no encontrado para respaldo.")
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = src.with_suffix(src.suffix + f".bak.{ts}")
    shutil.copy2(src, bak)
    print(f"Respaldo creado: {bak.name}")
    return bak

def restore_and_delete_backup(backup_path: Path, target: Path) -> None:
    if backup_path and backup_path.exists():
        shutil.copy2(backup_path, target)
        print(f"Restaurado {target.name} desde respaldo: {backup_path.name}")
        try:
            backup_path.unlink()  # eliminar el .bak
            print(f"Respaldo eliminado: {backup_path.name}")
        except Exception as e:
            print(f"Advertencia: no se pudo eliminar el respaldo ({e})")
    else:
        print("Respaldo no encontrado; nada que restaurar/eliminar.")

def patch_date_placeholder(dvc_path: Path, date_stamp: str) -> int:
    """
    Reemplaza TODAS las ocurrencias literales de 'DATEPLACEHOLDER' con la fecha (YYYY-MM-DD).
    No usa regex, así que también cubre '..._DATEPLACEHOLDER.csv' (con guión bajo).
    """
    text = dvc_path.read_text(encoding="utf-8")
    count = text.count("DATEPLACEHOLDER")
    if count:
        dvc_path.write_text(text.replace("DATEPLACEHOLDER", date_stamp), encoding="utf-8", newline="\n")
    return count

# ---------- Main ----------
def main():
    parser = argparse.ArgumentParser(description="Ejecutor DVC con entorno Conda y parcheo temporal de dvc.yaml")
    parser.add_argument("--env-name", default=None, help="Nombre del entorno Conda (si no se proporciona, intenta leerlo del YAML).")
    parser.add_argument("--env-file", default=ENV_FILE_DEFAULT, help="Ruta a environment.yaml.")
    parser.add_argument("--dvc-file", default=DVC_FILE_DEFAULT, help="Ruta a dvc.yaml a parchear.")
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
            raise SystemExit("Formato de --date inválido. Use YYYY-MM-DD (ej. 2025-08-21).")
        date_stamp = args.date
    else:
        date_stamp = dt.date.today().isoformat()

    print(f"Usando entorno: {env_name}")
    print(f"Usando fecha: {date_stamp}")
    print(f"dvc.yaml: {dvc_file}")

    # Requisitos base
    check_tool_available("conda")

    backup_path = None
    try:
        # 1) Respaldo + parcheo de 'DATEPLACEHOLDER' antes de todo lo demás
        backup_path = backup_file(dvc_file)
        replaced = patch_date_placeholder(dvc_file, date_stamp)
        if replaced:
            print(f"Parche aplicado: {replaced} ocurrencia(s) de 'DATEPLACEHOLDER' reemplazada(s) con '{date_stamp}'.")
        else:
            print("No se encontraron ocurrencias de 'DATEPLACEHOLDER' en dvc.yaml.")

        # 2) Entorno: crear si no existe; si ya existe, NO recrear.
        create_env_if_missing(env_name, env_file)

        # 3) Verificar/instalar dependencias dentro del entorno
        ensure_deps(env_name)

        # 4) Asegurar repo DVC
        ensure_dvc_repo(env_name)

        # 5) Pull solo si hay un remoto configurado
        if has_dvc_remote(env_name):
            print("📥 dvc pull…")
            dvc_command(env_name, "pull")
        else:
            print("ℹ️ No hay remoto DVC configurado. Omitiendo `dvc pull`.")

        # 6) Reproducir pipeline
        print("🔄 dvc repro…")
        start_time = dt.datetime.now()
        dvc_command(env_name, "repro")
        end_time = dt.datetime.now()

        # Calcular y mostrar duración
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

        print(f"✅ Pipeline completado en {' '.join(duration_str)}!")

    finally:
        # 7) Restaurar dvc.yaml y eliminar respaldo
        if backup_path:
            restore_and_delete_backup(backup_path, dvc_file)

if __name__ == "__main__":
    try:
        main()
    except subprocess.CalledProcessError as e:
        print(f"\n❌ Fallo de comando (exit {e.returncode}): {e.cmd}", file=sys.stderr)
        sys.exit(e.returncode)
    except Exception as e:
        print(f"\n❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
