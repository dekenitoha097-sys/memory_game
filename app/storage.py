from __future__ import annotations

from pathlib import Path
import sqlite3
from typing import Any


class Storage:
    def __init__(self, db_path: Path) -> None:
        self.db_path = db_path
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        return connection

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS scores (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    player_name TEXT NOT NULL,
                    grid TEXT NOT NULL,
                    theme TEXT NOT NULL,
                    duration_seconds INTEGER NOT NULL,
                    moves INTEGER NOT NULL,
                    errors INTEGER NOT NULL,
                    played_at TEXT NOT NULL
                )
                """
            )
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )

    def get_setting(self, key: str, default: str) -> str:
        with self._connect() as conn:
            row = conn.execute("SELECT value FROM settings WHERE key = ?", (key,)).fetchone()
            if row is None:
                return default
            return row["value"]

    def set_setting(self, key: str, value: str) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO settings (key, value)
                VALUES (?, ?)
                ON CONFLICT(key) DO UPDATE SET value = excluded.value
                """,
                (key, value),
            )

    def save_score(
        self,
        *,
        player_name: str,
        grid: str,
        theme: str,
        duration_seconds: int,
        moves: int,
        errors: int,
        played_at: str,
    ) -> None:
        with self._connect() as conn:
            conn.execute(
                """
                INSERT INTO scores (
                    player_name,
                    grid,
                    theme,
                    duration_seconds,
                    moves,
                    errors,
                    played_at
                )
                VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                (player_name, grid, theme, duration_seconds, moves, errors, played_at),
            )

    def fetch_top_scores(self, limit: int = 12) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                SELECT
                    player_name,
                    grid,
                    theme,
                    duration_seconds,
                    moves,
                    errors,
                    played_at
                FROM scores
                ORDER BY duration_seconds ASC, moves ASC, errors ASC, id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
        return [dict(row) for row in rows]
