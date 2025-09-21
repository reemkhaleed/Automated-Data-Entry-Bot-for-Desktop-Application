import os
import time
import logging
import traceback
from typing import List
import requests
from pywinauto.application import Application
from pywinauto import keyboard, timings


API_URL = "https://jsonplaceholder.typicode.com/posts"
POST_COUNT = 10
KEYBOARD_PAUSE = 0.008  
SAVE_TIMEOUT = 8  

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

def desktop_tjm_dir() -> str:
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    out = os.path.join(desktop, "tjm-project")
    os.makedirs(out, exist_ok=True)
    return out


def fetch_posts(n: int) -> List[dict]:
    logging.info(f"Fetching posts from {API_URL}")
    r = requests.get(API_URL, timeout=10)
    r.raise_for_status()
    posts = r.json()
    logging.info(f"Got {len(posts)} posts from API")
    return posts[:n]

def start_notepad() -> Application:
    app = Application(backend="win32")
    try:
        proc = app.start(r"C:\Windows\System32\notepad.exe")
        return app
    except Exception as e:
        logging.error("Failed to start Notepad: %s", e)
        raise

def find_notepad_window(app: Application, timeout: float = 8):
    try:
        timings.wait_until(timeout, 0.25, lambda: any("Notepad" in w.window_text() or "Untitled" in w.window_text() for w in app.windows()))
        for w in app.windows():
            title = w.window_text()
            if "Notepad" in title or "Untitled" in title:
                return w
    except Exception:
        pass
    try:
        return app.window(class_name="Notepad")
    except Exception as e:
        logging.warning("Could not find Notepad window by heuristics: %s", e)
        return None

def type_into_notepad(window, text: str):
    if window is None:
        raise RuntimeError("Notepad window is None")
    try:
        edit = window.child_window(class_name="Edit")
        edit.set_focus()
        edit.type_keys(text, with_spaces=True, with_newlines=True, pause=KEYBOARD_PAUSE)
    except Exception as e:
        logging.warning("Edit control typing failed (%s). Falling back to keyboard.send_keys", e)
        window.set_focus()
        keyboard.send_keys(text, pause=KEYBOARD_PAUSE, with_spaces=True)


def save_via_notepad(app: Application, window, full_path: str) -> bool:

    try:

        try:
            window.menu_select("File->Save As...")
        except Exception:

            keyboard.send_keys("^s")

        save_dlg = None
        start = time.time()
        while time.time() - start < SAVE_TIMEOUT:

            for w in app.windows():
                t = w.window_text().lower()
                if "save as" in t or "save as" == t.strip() or "save" == t.strip() and "save as" in t:
                    save_dlg = w
                    break
            if save_dlg:
                break
            time.sleep(0.2)

        if not save_dlg:
            logging.warning("Save As dialog not detected within timeout.")
            return False

        save_dlg.set_focus()
        try:
            fname_edit = save_dlg.child_window(class_name="Edit")
            fname_edit.set_edit_text(full_path)
        except Exception as e:
            logging.warning("Could not set Save As filename via control: %s. Will try typing the path.", e)
            save_dlg.set_focus()
            keyboard.send_keys(full_path, pause=KEYBOARD_PAUSE)

        keyboard.send_keys("{ENTER}")
        time.sleep(0.5)

        start = time.time()
        while time.time() - start < 3:
            for w in app.windows():
                title = w.window_text().lower()
                if "confirm" in title and "save" in title or title.strip().startswith("confirm"):
                    try:

                        if hasattr(w, "Yes"):
                            w.Yes.click_input()
                        else:

                            w.set_focus()
                            keyboard.send_keys("{ENTER}")
                    except Exception:
                        try:
                            keyboard.send_keys("{ENTER}")
                        except Exception:
                            pass
                    return True
            time.sleep(0.2)

        return True
    except Exception:
        logging.error("Exception while saving: %s", traceback.format_exc())
        return False


def close_notepad(window):
    try:
        window.close()
    except Exception:
        try:
            keyboard.send_keys("%{F4}")  
        except Exception:
            pass

def write_fallback_file(full_path: str, content: str):
    try:
        with open(full_path, "w", encoding="utf-8") as f:
            f.write(content)
        logging.info("WROTE (fallback) %s", full_path)
    except Exception:
        logging.error("Fallback write failed: %s", traceback.format_exc())

def compose_post_text(post: dict) -> str:
    title = post.get("title", "").strip()
    body = post.get("body", "").strip()
    lines = []
    lines.append(title.upper())
    lines.append("")  
    lines.append(body)
    lines.append("")  
    lines.append(f"(source: jsonplaceholder.typicode.com/posts/{post.get('id')})")
    return "\n".join(lines)


def main():
    out_dir = desktop_tjm_dir()
    logging.info("Output directory: %s", out_dir)

    try:
        posts = fetch_posts(POST_COUNT)
    except Exception as e:
        logging.error("Failed to fetch posts: %s", e)
        return

    for p in posts:
        pid = p.get("id", "unknown")
        filename = f"post {pid}.txt"
        full_path = os.path.join(out_dir, filename)
        content = compose_post_text(p)
        logging.info("Processing post %s -> %s", pid, full_path)

        app = None
        notepad_window = None
        saved_ok = False

        try:
            app = start_notepad()
            time.sleep(0.4)
            notepad_window = find_notepad_window(app, timeout=6)
            if notepad_window is None:
                raise RuntimeError("Notepad window not found after start.")

            type_into_notepad(notepad_window, content)
            time.sleep(0.25)

            saved_ok = save_via_notepad(app, notepad_window, full_path)
            if not saved_ok:
                logging.warning("Save via Notepad failed for %s", full_path)

            close_notepad(notepad_window)
            time.sleep(0.2)
        except Exception as e:
            logging.error("Error handling post %s: %s", pid, e)
            logging.debug(traceback.format_exc())
            saved_ok = False

            try:
                if notepad_window:
                    close_notepad(notepad_window)
            except Exception:
                pass

        if not saved_ok:
            logging.info("Falling back to writing file directly for post %s", pid)
            write_fallback_file(full_path, content)

        time.sleep(0.5)

    logging.info("Done. Files saved to: %s", out_dir)


if __name__ == "__main__":

    main()
