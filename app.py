import tkinter as tk

from ui.desktop_app import DesktopApp


def main() -> None:
    root = tk.Tk()
    app = DesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
