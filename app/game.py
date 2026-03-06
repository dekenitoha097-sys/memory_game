from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import ctypes
import random
import time
import tkinter as tk
from tkinter import messagebox, ttk

from PIL import Image, ImageDraw, ImageOps, ImageTk, UnidentifiedImageError

from .storage import Storage

try:
    import winsound
except ImportError:  # pragma: no cover - not available on non-Windows systems
    winsound = None

try:
    import pygame
except ImportError:  # pragma: no cover - optional dependency
    pygame = None


BASE_DIR = Path(__file__).resolve().parent.parent
IMAGES_ROOT = BASE_DIR / "assets" / "images"
DB_PATH = BASE_DIR / "data" / "memory_game.db"
MUSIC_PATH = BASE_DIR / "assets" / "audio" / "happy-kids-music-307326.mp3"
SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}


@dataclass
class Card:
    key: str
    image_path: Path
    is_revealed: bool = False
    is_matched: bool = False


class MemoryGameApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Memory Game Pro")
        self.geometry("1220x820")
        self.minsize(1100, 720)
        self.configure(bg="#0B1220")

        self.storage = Storage(DB_PATH)
        self.themes = self._discover_themes()
        if not self.themes:
            messagebox.showerror(
                "Assets manquants",
                "Aucune image valide detectee dans assets/images/.",
            )
            self.destroy()
            raise SystemExit(1)

        self.grid_options = ["4x4", "4x5", "5x6", "6x6"]

        default_theme = next(iter(self.themes.keys()))
        grid_default = self.storage.get_setting("grid", "4x4")
        theme_default = self.storage.get_setting("theme", default_theme)
        player_default = self.storage.get_setting("player_name", "Joueur")
        sound_default = self.storage.get_setting("sound_enabled", "1") == "1"
        flip_default = self.storage.get_setting("flip_delay_ms", "850")
        volume_default = self.storage.get_setting("music_volume", "0.45")

        if grid_default not in self.grid_options:
            grid_default = "4x4"
        if theme_default not in self.themes:
            theme_default = default_theme

        self.player_var = tk.StringVar(value=player_default)
        self.grid_var = tk.StringVar(value=grid_default)
        self.theme_var = tk.StringVar(value=theme_default)
        self.sound_var = tk.BooleanVar(value=sound_default)
        try:
            volume_value = float(volume_default)
        except ValueError:
            volume_value = 0.45
        volume_value = max(0.0, min(1.0, volume_value))
        self.volume_var = tk.DoubleVar(value=volume_value)
        self.volume_slider_var = tk.IntVar(value=int(volume_value * 100))

        flip_ms = int(flip_default) if flip_default.isdigit() else 850
        flip_ms = max(400, min(1500, flip_ms))
        self.flip_delay_var = tk.IntVar(value=flip_ms)

        self.status_var = tk.StringVar(value="Pret. Choisis tes options puis clique sur Nouvelle Partie.")

        self.cards: list[Card] = []
        self.card_buttons: list[tk.Button] = []
        self.card_faces: dict[int, ImageTk.PhotoImage] = {}
        self.card_back: ImageTk.PhotoImage | None = None

        self.rows = 4
        self.cols = 4
        self.total_pairs = 0
        self.card_size = 120

        self.first_choice: int | None = None
        self.second_choice: int | None = None
        self.locked_board = False

        self.moves = 0
        self.errors = 0
        self.matches = 0
        self.started_at = 0.0

        self.timer_job: str | None = None
        self.flip_job: str | None = None

        self.stats_value_labels: dict[str, tk.Label] = {}
        self.leaderboard: ttk.Treeview | None = None
        self.board_frame: tk.Frame | None = None
        self.music_loaded = False
        self.music_backend = "none"

        self._configure_styles()
        self._build_layout()
        self._init_music()
        self._populate_leaderboard()

        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(100, self.start_new_game)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            "Treeview",
            background="#0F172A",
            fieldbackground="#0F172A",
            foreground="#E2E8F0",
            rowheight=26,
            bordercolor="#0F172A",
        )
        style.configure(
            "Treeview.Heading",
            background="#1F2937",
            foreground="#F8FAFC",
            font=("Segoe UI", 10, "bold"),
            bordercolor="#1F2937",
        )
        style.map(
            "Treeview",
            background=[("selected", "#334155")],
            foreground=[("selected", "#FFFFFF")],
        )
        style.configure(
            "TCombobox",
            fieldbackground="#1F2937",
            background="#1F2937",
            foreground="#E2E8F0",
            selectbackground="#334155",
            selectforeground="#E2E8F0",
        )

    def _build_layout(self) -> None:
        top = tk.Frame(self, bg="#111827", padx=18, pady=14)
        top.pack(fill="x")

        title = tk.Label(
            top,
            text="Memory Game Pro",
            bg="#111827",
            fg="#F8FAFC",
            font=("Segoe UI", 30, "bold"),
        )
        title.pack(anchor="w")

        subtitle = tk.Label(
            top,
            text="Interface modernisee, grille stable, et sauvegarde locale des scores.",
            bg="#111827",
            fg="#93C5FD",
            font=("Segoe UI", 11),
        )
        subtitle.pack(anchor="w", pady=(0, 10))

        controls = tk.Frame(top, bg="#111827")
        controls.pack(fill="x")

        tk.Label(controls, text="Joueur", bg="#111827", fg="#E2E8F0", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 6))
        tk.Entry(
            controls,
            textvariable=self.player_var,
            width=16,
            bg="#1F2937",
            fg="#E2E8F0",
            insertbackground="#E2E8F0",
            relief="flat",
            font=("Segoe UI", 11),
        ).grid(row=0, column=1, padx=(0, 14), pady=2)

        tk.Label(controls, text="Grille", bg="#111827", fg="#E2E8F0", font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w", padx=(0, 6))
        grid_menu = tk.OptionMenu(controls, self.grid_var, *self.grid_options)
        grid_menu.configure(
            bg="#1F2937",
            fg="#F8FAFC",
            activebackground="#334155",
            activeforeground="#F8FAFC",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#475569",
            font=("Segoe UI", 10, "bold"),
            width=5,
        )
        grid_menu["menu"].configure(bg="#1F2937", fg="#F8FAFC", activebackground="#0EA5E9", activeforeground="#FFFFFF", font=("Segoe UI", 10))
        grid_menu.grid(row=0, column=3, padx=(0, 14), pady=2)

        tk.Label(controls, text="Theme", bg="#111827", fg="#E2E8F0", font=("Segoe UI", 10, "bold")).grid(row=0, column=4, sticky="w", padx=(0, 6))
        theme_menu = tk.OptionMenu(controls, self.theme_var, *list(self.themes.keys()))
        theme_menu.configure(
            bg="#1F2937",
            fg="#F8FAFC",
            activebackground="#334155",
            activeforeground="#F8FAFC",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#475569",
            font=("Segoe UI", 10, "bold"),
            width=12,
        )
        theme_menu["menu"].configure(bg="#1F2937", fg="#F8FAFC", activebackground="#0EA5E9", activeforeground="#FFFFFF", font=("Segoe UI", 10))
        theme_menu.grid(row=0, column=5, padx=(0, 14), pady=2)

        tk.Label(controls, text="Flip (ms)", bg="#111827", fg="#E2E8F0", font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky="w", padx=(0, 6))
        tk.Scale(
            controls,
            from_=400,
            to=1500,
            resolution=50,
            orient="horizontal",
            variable=self.flip_delay_var,
            length=180,
            bg="#111827",
            fg="#E2E8F0",
            troughcolor="#1F2937",
            highlightthickness=0,
        ).grid(row=0, column=7, padx=(0, 14))

        tk.Checkbutton(
            controls,
            text="Son",
            variable=self.sound_var,
            command=self._on_sound_toggle,
            bg="#111827",
            fg="#E2E8F0",
            selectcolor="#1F2937",
            activebackground="#111827",
            activeforeground="#E2E8F0",
            font=("Segoe UI", 10, "bold"),
        ).grid(row=0, column=8, padx=(0, 14))

        tk.Label(controls, text="Volume", bg="#111827", fg="#E2E8F0", font=("Segoe UI", 10, "bold")).grid(row=0, column=9, sticky="w", padx=(0, 6))
        tk.Scale(
            controls,
            from_=0,
            to=100,
            resolution=5,
            orient="horizontal",
            variable=self.volume_slider_var,
            command=self._on_volume_change,
            length=130,
            bg="#111827",
            fg="#E2E8F0",
            troughcolor="#1F2937",
            highlightthickness=0,
        ).grid(row=0, column=10, padx=(0, 14))

        tk.Button(
            controls,
            text="Nouvelle Partie",
            command=self.start_new_game,
            bg="#0EA5E9",
            fg="#FFFFFF",
            activebackground="#0284C7",
            activeforeground="#FFFFFF",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=8,
            cursor="hand2",
        ).grid(row=0, column=11, padx=(0, 8))

        tk.Label(
            controls,
            textvariable=self.status_var,
            bg="#111827",
            fg="#CBD5E1",
            font=("Segoe UI", 10),
            anchor="w",
        ).grid(row=1, column=0, columnspan=12, sticky="we", pady=(10, 0))

        main = tk.Frame(self, bg="#0B1220", padx=16, pady=16)
        main.pack(fill="both", expand=True)
        main.grid_rowconfigure(0, weight=1)
        main.grid_columnconfigure(0, weight=4)
        main.grid_columnconfigure(1, weight=2)

        board_shell = tk.Frame(main, bg="#111827", padx=14, pady=14)
        board_shell.grid(row=0, column=0, sticky="nsew", padx=(0, 12))

        self.board_frame = tk.Frame(board_shell, bg="#111827")
        self.board_frame.pack(expand=True, fill="both")

        side = tk.Frame(main, bg="#111827", padx=12, pady=12)
        side.grid(row=0, column=1, sticky="nsew")

        tk.Label(side, text="Statistiques", bg="#111827", fg="#F8FAFC", font=("Segoe UI", 16, "bold")).pack(anchor="w")
        stats = tk.Frame(side, bg="#111827")
        stats.pack(fill="x", pady=(8, 14))

        for key, label in (
            ("time", "Temps"),
            ("moves", "Coups"),
            ("errors", "Erreurs"),
            ("accuracy", "Precision"),
            ("matches", "Paires"),
        ):
            row = tk.Frame(stats, bg="#111827")
            row.pack(fill="x", pady=2)
            tk.Label(row, text=f"{label}:", bg="#111827", fg="#94A3B8", font=("Segoe UI", 11, "bold")).pack(side="left")
            value = tk.Label(row, text="0", bg="#111827", fg="#E2E8F0", font=("Segoe UI", 11))
            value.pack(side="right")
            self.stats_value_labels[key] = value

        tk.Label(side, text="Top Scores", bg="#111827", fg="#F8FAFC", font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(8, 8))

        columns = ("player", "grid", "theme", "time", "moves", "errors", "date")
        self.leaderboard = ttk.Treeview(side, columns=columns, show="headings", height=17)
        self.leaderboard.heading("player", text="Joueur")
        self.leaderboard.heading("grid", text="Grille")
        self.leaderboard.heading("theme", text="Theme")
        self.leaderboard.heading("time", text="Temps")
        self.leaderboard.heading("moves", text="Coups")
        self.leaderboard.heading("errors", text="Erreurs")
        self.leaderboard.heading("date", text="Date")

        self.leaderboard.column("player", width=120, anchor="center")
        self.leaderboard.column("grid", width=55, anchor="center")
        self.leaderboard.column("theme", width=90, anchor="center")
        self.leaderboard.column("time", width=70, anchor="center")
        self.leaderboard.column("moves", width=60, anchor="center")
        self.leaderboard.column("errors", width=60, anchor="center")
        self.leaderboard.column("date", width=120, anchor="center")

        scrollbar = ttk.Scrollbar(side, orient="vertical", command=self.leaderboard.yview)
        self.leaderboard.configure(yscrollcommand=scrollbar.set)
        self.leaderboard.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def _discover_themes(self) -> dict[str, list[Path]]:
        themes: dict[str, list[Path]] = {}
        if not IMAGES_ROOT.exists():
            return themes

        prettify = {
            "animals": "Animaux",
            "topics": "Themes",
            "vehicles": "Vehicules",
        }

        for folder in sorted(IMAGES_ROOT.iterdir()):
            if not folder.is_dir():
                continue
            if folder.name.lower() == "ui":
                continue
            images = [
                path
                for path in sorted(folder.iterdir())
                if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
            ]
            if images:
                theme_name = prettify.get(folder.name.lower(), folder.name.capitalize())
                themes[theme_name] = images

        all_cards = sorted({path for paths in themes.values() for path in paths})
        if all_cards:
            themes["Mixte"] = all_cards

        return themes

    def start_new_game(self) -> None:
        parsed = self._parse_grid(self.grid_var.get())
        if parsed is None:
            messagebox.showerror("Grille invalide", "Format attendu: LxC (ex: 4x4) et nombre total pair.")
            return
        self.rows, self.cols = parsed
        self.total_pairs = (self.rows * self.cols) // 2

        theme = self.theme_var.get()
        theme_pool = self.themes.get(theme, [])
        if len(theme_pool) < self.total_pairs:
            mixed_pool = self.themes.get("Mixte", [])
            if len(mixed_pool) < self.total_pairs:
                messagebox.showerror("Images insuffisantes", "Pas assez d'images pour cette taille de grille.")
                return
            theme_pool = mixed_pool
            self.theme_var.set("Mixte")
            self.status_var.set("Theme trop petit pour la grille choisie. Bascule automatique vers Mixte.")

        selected_images = random.sample(theme_pool, self.total_pairs)

        deck = [Card(key=str(path.resolve()), image_path=path) for path in selected_images for _ in range(2)]
        random.shuffle(deck)
        self.cards = deck

        self.first_choice = None
        self.second_choice = None
        self.locked_board = False
        self.moves = 0
        self.errors = 0
        self.matches = 0
        self.started_at = time.monotonic()

        self._prepare_card_images()
        self._render_board()
        self._persist_settings()
        self._start_timer()
        self._refresh_stats()
        self._play_feedback(success=True)
        self.status_var.set("Partie en cours. Bonne chance.")

    def _prepare_card_images(self) -> None:
        max_width = 840
        max_height = 600
        computed = min(max_width // self.cols, max_height // self.rows) - 12
        self.card_size = max(82, min(150, computed))

        self.card_back = self._build_card_back(self.card_size)
        self.card_faces.clear()
        for index, card in enumerate(self.cards):
            self.card_faces[index] = self._load_card_face(card.image_path, self.card_size)

    def _build_card_back(self, size: int) -> ImageTk.PhotoImage:
        back_path = IMAGES_ROOT / "ui" / "cartes.png"
        if back_path.exists():
            with Image.open(back_path) as image:
                card_back = ImageOps.fit(image.convert("RGB"), (size, size), Image.Resampling.LANCZOS)
        else:
            card_back = Image.new("RGB", (size, size), "#0EA5E9")
            draw = ImageDraw.Draw(card_back)
            draw.rectangle([6, 6, size - 6, size - 6], outline="#BAE6FD", width=3)
            draw.text((size // 2 - 8, size // 2 - 12), "?", fill="#F8FAFC")

        card_back = ImageOps.expand(card_back, border=2, fill="#1E293B")
        return ImageTk.PhotoImage(card_back)

    def _load_card_face(self, path: Path, size: int) -> ImageTk.PhotoImage:
        try:
            with Image.open(path) as image:
                face = ImageOps.fit(image.convert("RGB"), (size, size), Image.Resampling.LANCZOS)
        except (FileNotFoundError, UnidentifiedImageError, OSError):
            face = Image.new("RGB", (size, size), "#7F1D1D")
            draw = ImageDraw.Draw(face)
            draw.rectangle([4, 4, size - 4, size - 4], outline="#FCA5A5", width=3)
            draw.text((size // 2 - 28, size // 2 - 8), "IMAGE", fill="#FEE2E2")

        face = ImageOps.expand(face, border=2, fill="#1E293B")
        return ImageTk.PhotoImage(face)

    def _render_board(self) -> None:
        if self.board_frame is None or self.card_back is None:
            return

        for widget in self.board_frame.winfo_children():
            widget.destroy()

        self.card_buttons.clear()

        for row in range(self.rows):
            self.board_frame.grid_rowconfigure(row, weight=1)
        for col in range(self.cols):
            self.board_frame.grid_columnconfigure(col, weight=1)

        for index, _card in enumerate(self.cards):
            row = index // self.cols
            col = index % self.cols
            button = tk.Button(
                self.board_frame,
                image=self.card_back,
                command=lambda i=index: self._on_card_clicked(i),
                relief="flat",
                bd=0,
                highlightthickness=0,
                bg="#1E293B",
                activebackground="#334155",
                cursor="hand2",
            )
            button.grid(row=row, column=col, padx=6, pady=6, sticky="nsew")
            self.card_buttons.append(button)

    def _on_card_clicked(self, index: int) -> None:
        if self.locked_board:
            return

        if index == self.first_choice:
            return

        card = self.cards[index]
        if card.is_revealed or card.is_matched:
            return

        self._reveal(index)

        if self.first_choice is None:
            self.first_choice = index
            return

        self.second_choice = index
        self.moves += 1

        first_card = self.cards[self.first_choice]
        second_card = self.cards[self.second_choice]

        if first_card.key == second_card.key:
            first_card.is_matched = True
            second_card.is_matched = True
            self._mark_matched(self.first_choice)
            self._mark_matched(self.second_choice)

            self.matches += 1
            self._play_feedback(success=True)
            self.first_choice = None
            self.second_choice = None
            self._refresh_stats()

            if self.matches == self.total_pairs:
                self._handle_victory()
        else:
            self.errors += 1
            self.locked_board = True
            self._play_feedback(success=False)
            self._refresh_stats()
            delay = int(self.flip_delay_var.get())
            self.flip_job = self.after(delay, self._flip_back_mismatch)

    def _reveal(self, index: int) -> None:
        card = self.cards[index]
        card.is_revealed = True
        button = self.card_buttons[index]
        button.configure(image=self.card_faces[index], bg="#1E3A8A", activebackground="#1E3A8A")

    def _mark_matched(self, index: int) -> None:
        button = self.card_buttons[index]
        button.configure(state="disabled", bg="#14532D", activebackground="#14532D", cursor="arrow")

    def _hide(self, index: int) -> None:
        card = self.cards[index]
        card.is_revealed = False
        button = self.card_buttons[index]
        if self.card_back is not None:
            button.configure(image=self.card_back, bg="#1E293B", activebackground="#334155")

    def _flip_back_mismatch(self) -> None:
        if self.first_choice is None or self.second_choice is None:
            self.locked_board = False
            self.flip_job = None
            return

        if not self.cards[self.first_choice].is_matched:
            self._hide(self.first_choice)
        if not self.cards[self.second_choice].is_matched:
            self._hide(self.second_choice)

        self.first_choice = None
        self.second_choice = None
        self.locked_board = False
        self.flip_job = None

    def _handle_victory(self) -> None:
        elapsed = self._current_elapsed_seconds()
        self._cancel_timer()

        player_name = self.player_var.get().strip() or "Joueur"
        grid = self.grid_var.get()
        theme = self.theme_var.get()

        self.storage.save_score(
            player_name=player_name,
            grid=grid,
            theme=theme,
            duration_seconds=elapsed,
            moves=self.moves,
            errors=self.errors,
            played_at=datetime.now().strftime("%Y-%m-%d %H:%M"),
        )

        self._populate_leaderboard()
        self._refresh_stats()
        self.status_var.set("Bravo ! Partie terminee et score enregistre.")

        accuracy = self._accuracy_percent()
        messagebox.showinfo(
            "Victoire",
            (
                f"Excellent, {player_name} !\n\n"
                f"Temps: {self._format_seconds(elapsed)}\n"
                f"Coups: {self.moves}\n"
                f"Erreurs: {self.errors}\n"
                f"Precision: {accuracy:.1f}%"
            ),
        )

    def _start_timer(self) -> None:
        self._cancel_timer()
        self._tick_timer()

    def _tick_timer(self) -> None:
        self._refresh_stats()
        self.timer_job = self.after(250, self._tick_timer)

    def _cancel_timer(self) -> None:
        if self.timer_job is not None:
            self.after_cancel(self.timer_job)
            self.timer_job = None

    def _refresh_stats(self) -> None:
        elapsed = self._current_elapsed_seconds()
        self.stats_value_labels["time"].configure(text=self._format_seconds(elapsed))
        self.stats_value_labels["moves"].configure(text=str(self.moves))
        self.stats_value_labels["errors"].configure(text=str(self.errors))
        self.stats_value_labels["accuracy"].configure(text=f"{self._accuracy_percent():.1f}%")
        self.stats_value_labels["matches"].configure(text=f"{self.matches}/{self.total_pairs}")

    def _populate_leaderboard(self) -> None:
        if self.leaderboard is None:
            return

        for item in self.leaderboard.get_children():
            self.leaderboard.delete(item)

        for row in self.storage.fetch_top_scores(limit=18):
            self.leaderboard.insert(
                "",
                "end",
                values=(
                    row["player_name"],
                    row["grid"],
                    row["theme"],
                    self._format_seconds(int(row["duration_seconds"])),
                    row["moves"],
                    row["errors"],
                    row["played_at"],
                ),
            )

    def _persist_settings(self) -> None:
        self.storage.set_setting("player_name", self.player_var.get().strip() or "Joueur")
        self.storage.set_setting("grid", self.grid_var.get())
        self.storage.set_setting("theme", self.theme_var.get())
        self.storage.set_setting("sound_enabled", "1" if self.sound_var.get() else "0")
        self.storage.set_setting("flip_delay_ms", str(int(self.flip_delay_var.get())))
        self.storage.set_setting("music_volume", f"{self.volume_var.get():.2f}")

    def _accuracy_percent(self) -> float:
        if self.moves == 0:
            return 0.0
        return (self.matches / self.moves) * 100.0

    def _current_elapsed_seconds(self) -> int:
        if self.started_at <= 0:
            return 0
        return int(time.monotonic() - self.started_at)

    @staticmethod
    def _format_seconds(total_seconds: int) -> str:
        minutes, seconds = divmod(max(total_seconds, 0), 60)
        hours, minutes = divmod(minutes, 60)
        return f"{hours:02}:{minutes:02}:{seconds:02}"

    @staticmethod
    def _parse_grid(text: str) -> tuple[int, int] | None:
        try:
            rows_text, cols_text = text.lower().split("x")
            rows = int(rows_text)
            cols = int(cols_text)
        except ValueError:
            return None

        if rows <= 0 or cols <= 0:
            return None
        if (rows * cols) % 2 != 0:
            return None
        return rows, cols

    def _play_feedback(self, *, success: bool) -> None:
        if not self.sound_var.get():
            return
        if winsound is None:
            return
        try:
            winsound.MessageBeep(winsound.MB_ICONASTERISK if success else winsound.MB_ICONHAND)
        except RuntimeError:
            pass

    def _init_music(self) -> None:
        if not MUSIC_PATH.exists():
            self.status_var.set("Fichier MP3 introuvable: musique indisponible.")
            return

        # Priority on native Windows audio (MCI), no external dependency required.
        open_cmd = f'open "{MUSIC_PATH.resolve()}" type mpegvideo alias memory_music'
        if self._mci_send(open_cmd):
            self.music_loaded = True
            self.music_backend = "mci"
            self._apply_music_state()
            return

        # Optional fallback for non-Windows environments.
        if pygame is not None:
            try:
                pygame.mixer.init()
                pygame.mixer.music.load(str(MUSIC_PATH))
                self.music_loaded = True
                self.music_backend = "pygame"
                self._apply_music_state()
                return
            except pygame.error:
                self.music_loaded = False

        self.music_loaded = False
        self.music_backend = "none"
        self.status_var.set("Impossible de demarrer la musique.")

    @staticmethod
    def _mci_send(command: str) -> bool:
        try:
            result = ctypes.windll.winmm.mciSendStringW(command, None, 0, None)
            return result == 0
        except Exception:
            return False

    def _apply_music_state(self) -> None:
        if not self.music_loaded:
            return

        volume = max(0.0, min(1.0, self.volume_var.get()))

        if self.music_backend == "mci":
            self._mci_send(f"setaudio memory_music volume to {int(volume * 1000)}")
            if self.sound_var.get():
                # resume if paused, play repeat if not already running
                if not self._mci_send("resume memory_music"):
                    self._mci_send("play memory_music repeat")
            else:
                self._mci_send("pause memory_music")
            return

        if self.music_backend == "pygame" and pygame is not None:
            pygame.mixer.music.set_volume(volume)
            if self.sound_var.get():
                if not pygame.mixer.music.get_busy():
                    pygame.mixer.music.play(-1)
                else:
                    pygame.mixer.music.unpause()
            else:
                pygame.mixer.music.pause()

    def _on_sound_toggle(self) -> None:
        self._apply_music_state()
        self._persist_settings()

    def _on_volume_change(self, value: str) -> None:
        try:
            self.volume_var.set(float(value) / 100.0)
        except ValueError:
            return
        self._apply_music_state()
        self._persist_settings()

    def _on_close(self) -> None:
        self._persist_settings()
        self._cancel_timer()
        if self.flip_job is not None:
            self.after_cancel(self.flip_job)
            self.flip_job = None
        if self.music_loaded:
            if self.music_backend == "mci":
                self._mci_send("stop memory_music")
                self._mci_send("close memory_music")
            elif self.music_backend == "pygame" and pygame is not None:
                pygame.mixer.music.stop()
                pygame.mixer.quit()
        self.destroy()
