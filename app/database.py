from __future__ import annotations

from contextlib import contextmanager
from pathlib import Path
import sqlite3
import sys


def get_base_dir() -> Path:
    """
    Ruta base portable:
    - En desarrollo/Spyder: carpeta raíz del proyecto.
    - En .exe/PyInstaller: carpeta donde está el ejecutable.
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    # app/database.py -> app/ -> raíz del proyecto
    return Path(__file__).resolve().parent.parent


DB_PATH = get_base_dir() / "app.db"


@contextmanager
def db_cursor():
    """
    Crea la carpeta si hace falta y abre conexión SQLite.
    Si app.db no existe, SQLite lo crea automáticamente.
    Luego services.initialize_database() crea las tablas base.
    """
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")

    try:
        cur = conn.cursor()
        yield conn, cur
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
