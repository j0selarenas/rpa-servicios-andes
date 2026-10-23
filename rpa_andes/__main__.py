"""Línea de comandos.

    python -m rpa_andes ejecutar --entrada solicitudes.xlsx
    python -m rpa_andes reprocesar --fecha 2026-10-23
    python -m rpa_andes snapshots

Código de salida: 0 = OK (con o sin cuarentena), 1 = error de sistema, 2 = configuración.
"""
from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

from .config import load_config
from .landing import read_manifest
from .pipeline import execute


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="python -m rpa_andes",
                                     description="Robot de solicitudes CLP a UF y USD (Servicios Andes SpA)")
    parser.add_argument("--config", type=Path, default=Path("config.toml"))
    commands = parser.add_subparsers(dest="comando", required=True)
    run = commands.add_parser("ejecutar", help="captura el SII (o reutiliza la captura de hoy) y procesa el Excel")
    run.add_argument("--entrada", type=Path, required=True, help="Excel de solicitudes")
    replay = commands.add_parser("reprocesar", help="reconstruye bronce, silver y gold desde landing, sin red")
    replay.add_argument("--fecha", type=date.fromisoformat, required=True, help="AAAA-MM-DD")
    replay.add_argument("--snapshot", help="solo si hay más de un snapshot ese día")
    commands.add_parser("snapshots", help="lista los snapshots guardados en landing")
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
    except (OSError, ValueError) as exc:
        print(f"Error de configuración: {exc}", file=sys.stderr)
        return 2

    if args.comando == "snapshots":
        folder = config["rutas"]["datos"] / "landing" / "manifiestos"
        for path in sorted(folder.glob("*/*.json")):
            manifest = read_manifest(path)
            print(f"{manifest['fecha_extraccion']}  {manifest['snapshot_id']}  capturado {manifest['creado_en']}")
        return 0

    if args.comando == "ejecutar" and not args.entrada.exists():
        print(f"No existe el archivo de entrada: {args.entrada}", file=sys.stderr)
        return 2

    console = logging.StreamHandler()
    console.setFormatter(logging.Formatter("%(asctime)s %(levelname)-7s %(message)s", "%H:%M:%S"))
    logging.getLogger("rpa_andes").addHandler(console)

    today = datetime.now(ZoneInfo(config["negocio"]["zona_horaria"])).date()
    day = today if args.comando == "ejecutar" else args.fecha
    summary = execute(config, args.comando, day, excel_path=getattr(args, "entrada", None),
                      snapshot=getattr(args, "snapshot", None))
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 1 if summary["estado"] == "ERROR_SISTEMA" else 0


if __name__ == "__main__":
    sys.exit(main())
