"""Lectura de config.toml. Si falta algo, el robot no parte."""
from __future__ import annotations

import tomllib
from pathlib import Path
from zoneinfo import ZoneInfo

REQUIRED = {
    "rutas": ("datos", "base_datos", "logs", "salidas"),
    "entrada": ("hoja",),
    "sii": ("url_uf", "url_dolar"),
    "navegador": ("canal", "headless", "timeout_ms", "intentos", "espera_inicial_s"),
    "negocio": ("zona_horaria", "dias_max_arrastre_usd", "decimales_uf", "decimales_usd"),
}


def load_config(path: Path | str) -> dict:
    path = Path(path).resolve()
    with path.open("rb") as f:
        config = tomllib.load(f)

    for section, keys in REQUIRED.items():
        missing = [k for k in keys if k not in config.get(section, {})]
        if missing:
            raise ValueError(f"config.toml: faltan {', '.join(missing)} en [{section}]")

    # Las rutas se resuelven contra la carpeta del config, no contra el cwd.
    config["raiz"] = path.parent
    for key in REQUIRED["rutas"]:
        config["rutas"][key] = (path.parent / config["rutas"][key]).resolve()

    for key in ("url_uf", "url_dolar"):
        url = config["sii"][key]
        if not url.startswith("https://www.sii.cl/") or "{anio}" not in url:
            raise ValueError(f"config.toml: sii.{key} debe ser una URL https://www.sii.cl/ con {{anio}}")

    ZoneInfo(config["negocio"]["zona_horaria"])
    if config["navegador"]["intentos"] < 1:
        raise ValueError("config.toml: navegador.intentos debe ser 1 o más")
    return config
