"""Cargar variables de .env para credenciales de vTiger."""
import os
from pathlib import Path

def load_env():
    """Lee .env y carga las variables de entorno."""
    env_file = Path(__file__).parent.parent / ".env"
    if env_file.exists():
        with open(env_file) as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    key, value = line.split("=", 1)
                    os.environ[key.strip()] = value.strip()

load_env()
