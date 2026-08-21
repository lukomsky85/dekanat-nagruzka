"""
Проверка обновлений и самообновление приложения.

Работает по простому JSON-манифесту, размещённому по URL. Манифест может
лежать где угодно - на GitHub (как raw-файл в репозитории), на внутреннем
сервере вуза, в любой папке с статической раздачей файлов. Формат:

{
  "version": "1.1.0",
  "url": "https://.../UchebnayaNagruzka.exe",
  "notes": "Что изменилось в этой версии (необязательно)"
}

"version" сравнивается с текущей версией приложения (constants.VERSION).
"url" должен указывать напрямую на .exe для скачивания (например,
GitHub Release asset: https://github.com/<user>/<repo>/releases/download/vX.Y.Z/App.exe).
"""

import json
import os
import sys
import subprocess
import tempfile
import urllib.request
import urllib.error

USER_AGENT = "UchebnayaNagruzkaUpdater/1.0"


def _parse_version(v: str):
    """'v1.2.10' / '1.2.10' -> (1, 2, 10) для корректного сравнения версий
    (обычное сравнение строк даёт неверный результат, напр. '1.9' > '1.10')."""
    parts = []
    for p in str(v).strip().lstrip("vV").split("."):
        digits = "".join(ch for ch in p if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def is_newer(remote_version: str, current_version: str) -> bool:
    return _parse_version(remote_version) > _parse_version(current_version)


def check_for_update(manifest_url: str, current_version: str, timeout: int = 6):
    """
    Возвращает dict {"version", "url", "notes"}, если на сервере версия новее
    текущей, иначе None (в т.ч. если манифест недоступен или некорректен -
    в этом случае вызывающий код сам решает, показывать ли ошибку).
    Поднимает исключение при сетевой/HTTP ошибке.
    """
    if not manifest_url:
        return None

    req = urllib.request.Request(manifest_url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        raw = resp.read().decode("utf-8")
    data = json.loads(raw)

    remote_version = str(data.get("version", "")).strip()
    if not remote_version:
        return None

    if is_newer(remote_version, current_version):
        return {
            "version": remote_version,
            "url": data.get("url", "").strip(),
            "notes": (data.get("notes") or "").strip(),
        }
    return None


def download_update(url: str, dest_path: str, progress_callback=None, timeout: int = 30):
    """Скачивает файл по url в dest_path. progress_callback(downloaded, total)
    вызывается по мере получения данных (total может быть 0, если сервер
    не прислал Content-Length)."""
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        total = int(resp.headers.get("Content-Length", 0) or 0)
        downloaded = 0
        chunk_size = 262144  # 256 KB
        with open(dest_path, "wb") as f:
            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                f.write(chunk)
                downloaded += len(chunk)
                if progress_callback:
                    progress_callback(downloaded, total)
    return dest_path


def is_frozen() -> bool:
    """True, если запущено как собранный PyInstaller .exe."""
    return bool(getattr(sys, "frozen", False))


def apply_update_and_restart(new_exe_path: str):
    """
    Заменяет текущий исполняемый файл новым и перезапускает приложение.
    Доступно только для собранного .exe (PyInstaller). Т.к. Windows не
    позволяет перезаписать работающий exe напрямую, создаётся .bat-скрипт,
    который дожидается завершения текущего процесса, подменяет файл и
    запускает новую версию.

    Вызывающая сторона должна сразу после вызова этой функции завершить
    приложение (функция сама не вызывает os._exit - это остаётся на
    усмотрение GUI-кода, чтобы он мог корректно закрыть соединения с БД).
    """
    if not is_frozen():
        raise RuntimeError(
            "Самообновление доступно только для собранного .exe "
            "(при запуске из исходников python замените файлы вручную)."
        )

    current_exe = sys.executable
    current_dir = os.path.dirname(current_exe)
    exe_name = os.path.basename(current_exe)
    backup_path = os.path.join(current_dir, exe_name + ".old")

    bat_path = os.path.join(tempfile.gettempdir(), "un_update.bat")
    # В .bat сознательно не используем кириллицу в самом теле скрипта -
    # только в значениях путей (которые Windows обработает через свою
    # текущую кодовую страницу), чтобы не ловить проблемы с кодировкой .bat.
    bat_content = f"""@echo off
setlocal EnableDelayedExpansion
set "NEWEXE={new_exe_path}"
set "CUREXE={current_exe}"
set "BACKUP={backup_path}"
set "LOG={backup_path}.update_log.txt"
set /a WAITED=0

echo Ожидание завершения {exe_name}... > "%LOG%"

:waitloop
tasklist /FI "IMAGENAME eq {exe_name}" 2>NUL | find /I "{exe_name}" >NUL
if "%ERRORLEVEL%"=="0" (
    set /a WAITED+=1
    if !WAITED! GEQ 30 (
        echo Таймаут ожидания (30 сек) - процесс не завершился, пробую заменить файл принудительно. >> "%LOG%"
        goto forceupdate
    )
    timeout /t 1 /nobreak >NUL
    goto waitloop
)

:forceupdate
if exist "%BACKUP%" del /F /Q "%BACKUP%"
move /Y "%CUREXE%" "%BACKUP%" >> "%LOG%" 2>&1
if not exist "%CUREXE%" (
    move /Y "%NEWEXE%" "%CUREXE%" >> "%LOG%" 2>&1
    echo Файл обновлён успешно. >> "%LOG%"
) else (
    echo [ОШИБКА] Не удалось переименовать текущий exe (возможно, всё ещё занят). >> "%LOG%"
    echo Обновление НЕ применено. Старая версия сохранена. >> "%LOG%"
    move /Y "%BACKUP%" "%CUREXE%" >> "%LOG%" 2>&1
    goto end
)

start "" "%CUREXE%"
del /F /Q "%BACKUP%" >NUL 2>&1

:end
del /F /Q "%~f0"
"""
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat_content)

    # Windows-переносы строк (CRLF) - важно для корректного разбора .bat
    with open(bat_path, "rb") as f:
        data = f.read()
    data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    with open(bat_path, "wb") as f:
        f.write(data)

    creationflags = 0
    if os.name == "nt":
        creationflags = subprocess.CREATE_NO_WINDOW | getattr(subprocess, "DETACHED_PROCESS", 0)

    subprocess.Popen(
        ["cmd.exe", "/c", bat_path],
        creationflags=creationflags,
        close_fds=True,
    )
