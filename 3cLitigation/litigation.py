import os
import subprocess
import sys
import time
from pathlib import Path

import pandas as pd
import pyautogui


CPJ_EXECUTABLE_PATH = os.getenv("CPJ_EXECUTABLE_PATH", r"C:\path\to\cpj3cclient.exe")
CPJ_USERNAME = os.getenv("CPJ_USERNAME", "")
CPJ_PASSWORD = os.getenv("CPJ_PASSWORD", "")

INPUT_EXCEL_FILENAME = os.getenv("INPUT_EXCEL_FILENAME", "input.xlsx")
LITIGATION_ID_COLUMN = os.getenv("LITIGATION_ID_COLUMN", "LitigationID")

OUTPUT_ROOT_DIRNAME = os.getenv("OUTPUT_ROOT_DIRNAME", "documents_by_litigation_id")
ASSETS_DIRNAME = os.getenv("ASSETS_DIRNAME", "assets")
SAVE_BUTTON_IMAGE = os.getenv("SAVE_BUTTON_IMAGE", "save.png")

DOWNLOADS_DIR = Path(os.getenv("DOWNLOADS_DIR", str(Path.home() / "Downloads")))

proc = None


def log(msg: str):
    now = time.strftime("%H:%M:%S")
    print(f"[{now}] {msg}", flush=True)


def get_run_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "executable"):
        return Path(sys.executable).resolve().parent
    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd().resolve()


def _sort_key_litigation(x: str):
    s = (x or "").strip()
    if s.isdigit():
        return (0, int(s))
    return (1, s.lower(), s)


def read_litigation_ids() -> list[str]:
    base_dir = get_run_base_dir()
    excel_path = base_dir / INPUT_EXCEL_FILENAME

    log(f"Reading Excel: {excel_path}")

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    df = pd.read_excel(excel_path, engine="openpyxl")

    if LITIGATION_ID_COLUMN not in df.columns:
        raise ValueError(
            f"Column '{LITIGATION_ID_COLUMN}' not found. Available: {list(df.columns)}"
        )

    series = df[LITIGATION_ID_COLUMN].dropna()

    ids: list[str] = []
    for v in series.tolist():
        if isinstance(v, float) and v.is_integer():
            ids.append(str(int(v)))
        else:
            s = str(v).strip()
            if s:
                ids.append(s)

    seen = set()
    out: list[str] = []
    for x in ids:
        if x not in seen:
            out.append(x)
            seen.add(x)

    out.sort(key=_sort_key_litigation)
    log(f"{len(out)} litigation ids loaded")

    return out


def launch_app():
    global proc

    log("Launching application")

    if not os.path.exists(CPJ_EXECUTABLE_PATH):
        raise FileNotFoundError("Executable not found. Set CPJ_EXECUTABLE_PATH.")

    proc = subprocess.Popen(CPJ_EXECUTABLE_PATH)
    log("Application started")


def do_login():
    if not CPJ_USERNAME or not CPJ_PASSWORD:
        raise ValueError("Missing credentials. Set CPJ_USERNAME and CPJ_PASSWORD.")

    log("Starting login")

    time.sleep(4)

    pyautogui.write(CPJ_USERNAME, interval=0.05)
    pyautogui.press("tab")
    time.sleep(0.5)

    pyautogui.write(CPJ_PASSWORD, interval=0.05)
    pyautogui.press("tab")
    time.sleep(0.5)

    pyautogui.press("enter")

    time.sleep(10)

    pyautogui.press("esc")
    time.sleep(0.5)

    pyautogui.press("n")
    time.sleep(1)

    pyautogui.press("f8")
    time.sleep(5)


def ensure_output_folder(litigation_id: str) -> Path:
    base_dir = get_run_base_dir()

    root = base_dir / OUTPUT_ROOT_DIRNAME
    root.mkdir(parents=True, exist_ok=True)

    folder = root / str(litigation_id)
    folder.mkdir(parents=True, exist_ok=True)

    return folder


def _get_save_image_path(img_name: str) -> Path:
    base_dir = get_run_base_dir()
    img_path = base_dir / ASSETS_DIRNAME / img_name
    if not img_path.exists():
        raise FileNotFoundError(f"Image not found: {img_path}")
    return img_path


def is_save_button_visible(
    img_name: str = SAVE_BUTTON_IMAGE,
    confidence: float = 0.80,
    region=None,
) -> bool:
    img_path = _get_save_image_path(img_name)

    try:
        pos = pyautogui.locateCenterOnScreen(
            str(img_path),
            confidence=confidence,
            grayscale=True,
            region=region,
        )
    except Exception:
        pos = None

    return bool(pos)


def click_save_button(
    img_name: str = SAVE_BUTTON_IMAGE,
    timeout_sec: int = 25,
    confidence: float = 0.80,
    region=None,
) -> bool:
    img_path = _get_save_image_path(img_name)

    start = time.time()
    while (time.time() - start) < timeout_sec:
        try:
            pos = pyautogui.locateCenterOnScreen(
                str(img_path),
                confidence=confidence,
                grayscale=True,
                region=region,
            )
        except Exception:
            pos = None

        if pos:
            pyautogui.moveTo(pos.x, pos.y, duration=0.1)
            pyautogui.click()
            return True

        time.sleep(0.3)

    return False


def start_download():
    time.sleep(1)
    pyautogui.press("up", presses=2, interval=0.2)
    time.sleep(0.5)
    pyautogui.press("enter")
    pyautogui.press("enter")


def _unique_destination(dest: Path) -> Path:
    if not dest.exists():
        return dest

    stem = dest.stem
    suffix = dest.suffix

    for i in range(1, 1000):
        candidate = dest.with_name(f"{stem}_{i}{suffix}")
        if not candidate.exists():
            return candidate

    return dest.with_name(f"{stem}_{int(time.time())}{suffix}")


def download_and_move_to_folder(dest_folder: Path, timeout_sec: int = 180) -> Path:
    if not DOWNLOADS_DIR.exists():
        raise FileNotFoundError(f"Downloads folder not found: {DOWNLOADS_DIR}")

    t0 = time.time()
    start_download()

    target = None

    while (time.time() - t0) < timeout_sec:
        files = [p for p in DOWNLOADS_DIR.iterdir() if p.is_file()]
        candidates = []

        for p in files:
            name = p.name.lower()

            if name == "desktop.ini":
                continue
            if name.endswith((".crdownload", ".part", ".tmp")):
                continue

            try:
                st = p.stat()
            except Exception:
                continue

            if st.st_mtime >= (t0 - 1):
                candidates.append(p)

        if candidates:
            target = max(candidates, key=lambda x: x.stat().st_mtime)
            break

        time.sleep(0.5)

    if not target:
        raise TimeoutError("Downloaded file not detected within the timeout.")

    prev_size = -1
    stable = 0
    stability_start = time.time()

    while (time.time() - stability_start) < 60:
        try:
            size_now = target.stat().st_size
        except Exception:
            size_now = -1

        if size_now == prev_size and size_now > 0:
            stable += 1
            if stable >= 2:
                break
        else:
            stable = 0
            prev_size = size_now

        time.sleep(0.7)

    dest_folder.mkdir(parents=True, exist_ok=True)
    final_path = _unique_destination(dest_folder / target.name)

    for _ in range(40):
        try:
            os.replace(str(target), str(final_path))
            return final_path
        except PermissionError:
            time.sleep(0.5)

    raise PermissionError("Could not move the file (likely still in use).")


def prep_next_item():
    pyautogui.press("esc")
    time.sleep(0.2)
    pyautogui.press("esc")
    time.sleep(0.3)

    pyautogui.press("f8")
    time.sleep(6)

    pyautogui.hotkey("ctrl", "i")
    time.sleep(0.8)

    pyautogui.press("tab")
    time.sleep(0.3)


def process_litigation_id(litigation_id: str) -> bool:
    time.sleep(0.3)

    pyautogui.write(litigation_id, interval=0.03)
    time.sleep(0.4)

    pyautogui.press("enter")
    time.sleep(0.5)

    pyautogui.hotkey("shift", "enter")
    time.sleep(1)

    pyautogui.hotkey("ctrl", "m")
    time.sleep(0.6)

    if not is_save_button_visible():
        return False

    pyautogui.press("tab")
    pyautogui.press("m")
    time.sleep(1)

    if not click_save_button():
        return False

    folder = ensure_output_folder(litigation_id)
    download_and_move_to_folder(folder)

    return True


def close_app():
    global proc

    if proc and proc.poll() is None:
        try:
            proc.terminate()
            proc.wait(timeout=10)
        except subprocess.TimeoutExpired:
            proc.kill()

    subprocess.run(
        ["taskkill", "/F", "/IM", "cpj3cclient.exe"],
        capture_output=True,
        text=True,
    )

    sys.exit(0)


def main():
    ids = read_litigation_ids()

    launch_app()
    do_login()

    pyautogui.hotkey("ctrl", "i")
    time.sleep(1)
    pyautogui.press("tab")
    time.sleep(0.3)

    for litigation_id in ids:
        process_litigation_id(litigation_id)
        prep_next_item()

    close_app()


if __name__ == "__main__":
    main()