"""Application bootstrap for the Python TLIC port."""

import ctypes
from pathlib import Path
import sys
import tkinter as tk

from .main_form import MainForm


def _resource_path(*parts: str) -> Path:
    base_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return base_dir.joinpath("Resources", *parts)


APP_ICON_PATH = _resource_path("assets", "app.ico")
APP_ICON_16_PATH = _resource_path("assets", "app_16.ico")
APP_ICON_32_PATH = _resource_path("assets", "app_32.ico")
APP_ICON_48_PATH = _resource_path("assets", "app_48.ico")
APP_ICON_64_PATH = _resource_path("assets", "app_64.ico")
APP_ICON_128_PATH = _resource_path("assets", "app_128.ico")
APP_ICON_256_PATH = _resource_path("assets", "app_256.ico")
SPLASH_IMAGE_PATH = _resource_path("assets", "splash.png")
SPLASH_FAILSAFE_MS = 8000


def _pyi_splash_update(text: str) -> None:
    try:
        import pyi_splash  # type: ignore

        try:
            pyi_splash.update_text(text)
        except Exception:
            pass
    except Exception:
        pass


def _close_pyinstaller_splash() -> None:
    try:
        import pyi_splash  # type: ignore

        pyi_splash.close()
    except Exception:
        pass


def _has_pyinstaller_splash() -> bool:
    try:
        import pyi_splash  # type: ignore

        return bool(pyi_splash)
    except Exception:
        return False


def _set_windows_app_id() -> None:
    if sys.platform != "win32":
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("tlic.desktop.app")
    except Exception:
        # Keep startup resilient if the platform API is unavailable.
        pass


def _set_app_icon(root: tk.Tk) -> None:
    # On Windows prefer true ICO resource for both taskbar and titlebar.
    if sys.platform == "win32":
        small_ico = APP_ICON_16_PATH if APP_ICON_16_PATH.exists() else APP_ICON_PATH
        big_ico = APP_ICON_256_PATH if APP_ICON_256_PATH.exists() else APP_ICON_PATH

        if APP_ICON_PATH.exists():
            try:
                root.iconbitmap(str(APP_ICON_PATH))
            except tk.TclError:
                pass

        try:
            user32 = ctypes.windll.user32
            small_handle = None
            big_handle = None

            if small_ico.exists():
                small_handle = user32.LoadImageW(
                    None,
                    str(small_ico),
                    1,  # IMAGE_ICON
                    16,
                    16,
                    0x00000010,  # LR_LOADFROMFILE
                )
            if big_ico.exists():
                big_handle = user32.LoadImageW(
                    None,
                    str(big_ico),
                    1,  # IMAGE_ICON
                    256,
                    256,
                    0x00000010,  # LR_LOADFROMFILE
                )

            hwnd = root.winfo_id()
            if big_handle:
                user32.SendMessageW(hwnd, 0x0080, 1, big_handle)  # WM_SETICON, ICON_BIG
            if small_handle:
                user32.SendMessageW(hwnd, 0x0080, 0, small_handle)  # WM_SETICON, ICON_SMALL

            if small_handle or big_handle:
                root._win_icon_small_handle = small_handle
                root._win_icon_big_handle = big_handle
                return
        except Exception:
            pass

    # Cross-platform fallback path.
    for icon_path in (APP_ICON_256_PATH, APP_ICON_128_PATH, APP_ICON_64_PATH, APP_ICON_PATH, SPLASH_IMAGE_PATH):
        if not icon_path.exists():
            continue
        try:
            app_icon = tk.PhotoImage(file=str(icon_path))
            root.iconphoto(True, app_icon)
            root._app_icon_image = app_icon
            return
        except tk.TclError:
            continue


def _show_splash(root: tk.Tk):
    if not SPLASH_IMAGE_PATH.exists():
        return None

    try:
        splash = tk.Toplevel(root)
        splash.overrideredirect(True)
        splash.attributes("-topmost", True)
        splash.configure(background="#ff00ff")

        splash_image = tk.PhotoImage(file=str(SPLASH_IMAGE_PATH))
        splash_label = tk.Label(
            splash,
            image=splash_image,
            borderwidth=0,
            highlightthickness=0,
            background="#ff00ff",
        )
        splash_label.image = splash_image
        splash_label.pack()
        if sys.platform == "win32":
            try:
                splash.wm_attributes("-transparentcolor", "#ff00ff")
            except tk.TclError:
                pass

        splash.update_idletasks()
        width = splash.winfo_width()
        height = splash.winfo_height()
        x = (splash.winfo_screenwidth() // 2) - (width // 2)
        y = (splash.winfo_screenheight() // 2) - (height // 2)
        splash.geometry(f"{width}x{height}+{x}+{y}")
        root._splash_window = splash

        def _close_splash() -> None:
            current = getattr(root, "_splash_window", None)
            if current is not None:
                try:
                    current.destroy()
                except tk.TclError:
                    pass
                root._splash_window = None
            root.deiconify()

        # Failsafe close in case startup logic never reaches the ready callback.
        root.after(SPLASH_FAILSAFE_MS, _close_splash)
        return _close_splash
    except tk.TclError:
        # If splash image fails to load, continue with normal startup.
        return None


def main() -> int:
    _set_windows_app_id()
    _pyi_splash_update("Starting...")

    # Root window setup mirrors the C# app’s "big desktop utility" footprint.
    root = tk.Tk()
    root.withdraw()
    root.title("TLIC (Python Port) - v1.0")
    # Start large enough to show all panes without manual resize.
    root.geometry("1440x900")
    # Guardrail so controls/plot/output stay usable.
    root.minsize(1180, 720)
    _set_app_icon(root)
    use_pyi_splash = _has_pyinstaller_splash()
    close_splash = None if use_pyi_splash else _show_splash(root)
    _pyi_splash_update("Loading UI...")

    def _finish_startup() -> None:
        # MainForm owns all tab-level application behavior.
        app = MainForm(root)
        app.pack(fill="both", expand=True)
        root._main_form = app
        _set_app_icon(root)
        if close_splash is not None:
            close_splash()
        _close_pyinstaller_splash()
        root.deiconify()

    if close_splash is not None:
        # Ensure splash paints immediately before heavier UI setup.
        root.update_idletasks()
        root.update()
    root.after(10, _finish_startup)

    root.mainloop()
    return 0
