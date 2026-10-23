"""Capa landing: los archivos tal como llegaron, sin modificar ni borrar.

Estructura:
    landing/solicitudes/<fecha>/<hash>_solicitudes.xlsx
    landing/sii/<fecha>/<hash>_uf2026.html
    landing/manifiestos/<fecha>/<hash_excel>.json

El manifiesto amarra un Excel con las páginas del SII usadas para procesarlo.
Ese par (fecha + Excel) es el snapshot, la unidad de idempotencia y reproceso.
"""
from __future__ import annotations

import hashlib
import json
from datetime import date, datetime, timezone
from pathlib import Path


class IntegrityError(RuntimeError):
    """Un archivo de landing no coincide con su hash o falta."""


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def write_once(path: Path, data: bytes) -> None:
    """Landing nunca se sobrescribe. Si el archivo ya existe debe ser idéntico."""
    if path.exists():
        if path.read_bytes() != data:
            raise IntegrityError(f"Se intentó sobrescribir un archivo de landing: {path}")
        return
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_bytes(data)
    tmp.rename(path)


def save_source(data_dir: Path, source: str, day: date, filename: str, data: bytes) -> dict:
    digest = sha256(data)
    relative = Path("landing", source, day.isoformat(), f"{digest[:12]}_{filename}")
    write_once(data_dir / relative, data)
    return {"ruta": relative.as_posix(), "sha256": digest, "bytes": len(data)}


def snapshot_id(day: date, excel_sha: str) -> str:
    return f"{day.isoformat()}_{excel_sha[:12]}"


def manifest_path(data_dir: Path, day: date, excel_sha: str) -> Path:
    return data_dir / "landing" / "manifiestos" / day.isoformat() / f"{excel_sha[:12]}.json"


def write_manifest(data_dir: Path, day: date, excel: dict, sii: list[dict]) -> dict:
    manifest = {
        "snapshot_id": snapshot_id(day, excel["sha256"]),
        "fecha_extraccion": day.isoformat(),
        "creado_en": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "excel": excel,
        "sii": sii,
    }
    data = json.dumps(manifest, ensure_ascii=False, indent=2).encode("utf-8")
    write_once(manifest_path(data_dir, day, excel["sha256"]), data)
    return manifest


def read_manifest(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def find_manifest(data_dir: Path, day: date, snapshot: str | None = None) -> dict:
    """Busca el snapshot de una fecha. Si hay más de uno, hay que indicar cuál."""
    folder = data_dir / "landing" / "manifiestos" / day.isoformat()
    manifests = [read_manifest(p) for p in sorted(folder.glob("*.json"))]
    if snapshot:
        manifests = [m for m in manifests if m["snapshot_id"] == snapshot]
    if not manifests:
        raise IntegrityError(f"No hay snapshots guardados para {day.isoformat()}")
    if len(manifests) > 1:
        ids = ", ".join(m["snapshot_id"] for m in manifests)
        raise IntegrityError(f"Hay {len(manifests)} snapshots para {day.isoformat()} ({ids}); use --snapshot")
    return manifests[0]


def verify_sources(data_dir: Path, manifest: dict) -> None:
    """Antes de procesar, cada archivo del manifiesto debe existir y calzar con su hash."""
    for record in [manifest["excel"], *manifest["sii"]]:
        path = data_dir / record["ruta"]
        if not path.exists():
            raise IntegrityError(f"Falta un archivo de landing: {record['ruta']}")
        if sha256(path.read_bytes()) != record["sha256"]:
            raise IntegrityError(f"El archivo de landing fue modificado: {record['ruta']}")
