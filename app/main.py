from .game import MemoryGameApp


def main() -> None:
    app = MemoryGameApp()
    app.mainloop()


if __name__ == "__main__":
    main()
