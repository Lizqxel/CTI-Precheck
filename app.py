import tkinter as tk
import logging

from ui.desktop_app import DesktopApp


def main() -> None:
    logging.getLogger("urllib3.connectionpool").setLevel(logging.ERROR)
    root = tk.Tk()
    app = DesktopApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
