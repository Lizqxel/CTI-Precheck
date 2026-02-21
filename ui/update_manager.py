import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, Optional

import requests
from packaging import version
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from core.settings_store import (
    SETTINGS_PATH,
    append_update_history,
    load_update_settings,
    save_update_settings,
)
from version import APP_NAME, GITHUB_OWNER, GITHUB_REPO, VERSION

try:
    import tkinter as tk
    from tkinter import messagebox, ttk
except Exception:
    tk = Any
    ttk = Any
    messagebox = None

LogCallback = Callable[[str], None]


class UpdateManager:
    def __init__(self, root: Any, log_callback: Optional[LogCallback] = None) -> None:
        self.root = root
        self.log_callback = log_callback
        self._download_cancelled = threading.Event()
        self._progress_dialog: Any = None
        self._progress_var: Any = None
        self._progress_label_var: Any = None

        self._session = requests.Session()
        retry_strategy = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=1.0,
            status_forcelist=(429, 500, 502, 503, 504),
            allowed_methods=frozenset(["GET"]),
        )
        adapter = HTTPAdapter(max_retries=retry_strategy)
        self._session.mount("https://", adapter)
        self._session.mount("http://", adapter)

    def check_for_updates(self, interactive: bool = False, auto: bool = False) -> None:
        worker = threading.Thread(target=self._check_for_updates_worker, args=(interactive, auto), daemon=True)
        worker.start()

    def _check_for_updates_worker(self, interactive: bool, auto: bool) -> None:
        settings = load_update_settings(SETTINGS_PATH)

        if GITHUB_OWNER.startswith("REPLACE_WITH") or GITHUB_REPO.startswith("REPLACE_WITH"):
            msg = "更新チェックをスキップ: version.py の GitHub リポジトリ設定が未完了です"
            self._log(msg)
            if interactive:
                self._show_info("更新チェック", msg)
            return

        if auto and not self._should_auto_check(settings):
            return

        checked_at = self._utc_now_iso()
        status = "failed"
        latest_version = ""
        message = ""

        try:
            release_info = self._fetch_latest_release(settings)
            latest_tag = str(release_info.get("tag_name", "")).lstrip("v")
            latest_version = latest_tag

            if not latest_tag:
                raise RuntimeError("最新リリースの tag_name が取得できませんでした")

            save_update_settings(
                {
                    **settings,
                    "last_checked_at": checked_at,
                    "last_latest_version": latest_tag,
                },
                SETTINGS_PATH,
            )

            if version.parse(latest_tag) <= version.parse(VERSION):
                status = "up-to-date"
                message = f"最新です（現在: {VERSION} / 最新: {latest_tag}）"
                self._log(f"更新チェック結果: {message}")
                if interactive:
                    self._show_info("更新チェック", message)
                return

            skipped_version = str(settings.get("skipped_version", ""))
            if auto and skipped_version and skipped_version == latest_tag:
                self._log(f"更新通知をスキップ（ユーザーが {latest_tag} をスキップ済み）")
                status = "skipped"
                message = f"{latest_tag} は通知スキップ設定"
                return

            body = str(release_info.get("body", "") or "")
            prompt = (
                f"新しいバージョン {latest_tag} が見つかりました。\n"
                f"現在のバージョン: {VERSION}\n\n"
                "更新をダウンロードして適用しますか？"
            )
            if body.strip():
                prompt += f"\n\n--- Release Note ---\n{body[:1400]}"

            choice = self._ask_update_choice(prompt, latest_tag)
            if choice == "skip":
                save_update_settings({**settings, "skipped_version": latest_tag}, SETTINGS_PATH)
                status = "skipped"
                message = f"{latest_tag} をスキップしました"
                self._log(message)
                return
            if choice != "yes":
                status = "cancelled"
                message = "ユーザーが更新をキャンセルしました"
                self._log(message)
                return

            self._download_cancelled.clear()
            asset = self._select_exe_asset(release_info)
            downloaded_path = self._download_asset_with_progress(asset)
            self._verify_sha256(release_info, asset, downloaded_path)
            self._apply_update(downloaded_path, str(asset.get("name", "")))

            status = "applied"
            message = f"更新 {latest_tag} を適用しました"
            self._log(message)

        except Exception as exc:
            message = str(exc)
            self._log(f"更新チェック失敗: {message}")
            if interactive:
                self._show_error("更新チェック", f"更新処理に失敗しました\n{message}")
        finally:
            refreshed = load_update_settings(SETTINGS_PATH)
            save_update_settings(
                {
                    **refreshed,
                    "last_checked_at": checked_at,
                    "last_result": {
                        "status": status,
                        "message": message,
                        "current_version": VERSION,
                        "latest_version": latest_version,
                        "checked_at": checked_at,
                    },
                },
                SETTINGS_PATH,
            )
            append_update_history(
                {
                    "status": status,
                    "message": message,
                    "current_version": VERSION,
                    "latest_version": latest_version,
                    "checked_at": checked_at,
                },
                SETTINGS_PATH,
            )

    def _should_auto_check(self, settings: Dict[str, Any]) -> bool:
        interval_hours = int(settings.get("auto_check_interval_hours", 24))
        if interval_hours <= 0:
            return True

        last_checked_at = str(settings.get("last_checked_at", "")).strip()
        if not last_checked_at:
            return True

        try:
            last_dt = datetime.fromisoformat(last_checked_at.replace("Z", "+00:00"))
        except Exception:
            return True

        now_dt = datetime.now(timezone.utc)
        elapsed_hours = (now_dt - last_dt).total_seconds() / 3600
        return elapsed_hours >= interval_hours

    def _fetch_latest_release(self, settings: Dict[str, Any]) -> Dict[str, Any]:
        channel = str(settings.get("channel", "stable")).strip().lower() or "stable"
        headers: Dict[str, str] = {
            "Accept": "application/vnd.github+json",
            "User-Agent": f"{APP_NAME}/{VERSION}",
        }

        etag = str(settings.get("etag", "")).strip()
        if etag:
            headers["If-None-Match"] = etag

        base = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}"
        if channel == "prerelease":
            url = f"{base}/releases"
            response = self._session.get(url, timeout=15, headers=headers)
            if response.status_code == 304:
                cached = settings.get("cached_release")
                if isinstance(cached, dict) and cached:
                    return cached
                raise RuntimeError("304 を受信しましたがキャッシュがありません")
            response.raise_for_status()
            data = response.json()
            if not isinstance(data, list) or not data:
                raise RuntimeError("リリース情報が取得できませんでした")
            release = next((r for r in data if not r.get("draft", False)), None)
            if release is None:
                raise RuntimeError("有効なリリースが見つかりませんでした")
        else:
            url = f"{base}/releases/latest"
            response = self._session.get(url, timeout=15, headers=headers)
            if response.status_code == 304:
                cached = settings.get("cached_release")
                if isinstance(cached, dict) and cached:
                    return cached
                raise RuntimeError("304 を受信しましたがキャッシュがありません")
            response.raise_for_status()
            release = response.json()

        new_settings = {
            **settings,
            "etag": response.headers.get("ETag", ""),
            "cached_release": release,
            "channel": channel,
        }
        save_update_settings(new_settings, SETTINGS_PATH)
        return release

    def _select_exe_asset(self, release_info: Dict[str, Any]) -> Dict[str, Any]:
        assets = release_info.get("assets", [])
        if not isinstance(assets, list) or not assets:
            raise RuntimeError("リリースに assets がありません")

        latest_tag = str(release_info.get("tag_name", "")).lstrip("v")
        exact_name = f"{APP_NAME}-{latest_tag}.exe"
        for asset in assets:
            name = str(asset.get("name", ""))
            if name == exact_name:
                return asset

        prefix = f"{APP_NAME}-"
        candidates = [
            asset
            for asset in assets
            if str(asset.get("name", "")).startswith(prefix) and str(asset.get("name", "")).endswith(".exe")
        ]
        if len(candidates) == 1:
            return candidates[0]

        if not candidates:
            raise RuntimeError(f"更新用 EXE が見つかりません（期待: {exact_name}）")

        names = ", ".join(str(a.get("name", "")) for a in candidates)
        raise RuntimeError(f"更新対象 EXE が一意に特定できません: {names}")

    def _download_asset_with_progress(self, asset: Dict[str, Any]) -> Path:
        name = str(asset.get("name", "")).strip()
        download_url = str(asset.get("browser_download_url", "")).strip()
        if not download_url:
            raise RuntimeError("asset の download URL が見つかりません")

        temp_dir = Path(tempfile.mkdtemp(prefix="cti_update_"))
        target_path = temp_dir / name

        self._show_progress_dialog(name)
        self._log(f"更新ファイルをダウンロードします: {name}")

        try:
            response = self._session.get(download_url, stream=True, timeout=30)
            response.raise_for_status()

            total = int(response.headers.get("Content-Length") or 0)
            downloaded = 0
            with target_path.open("wb") as f:
                for chunk in response.iter_content(chunk_size=1024 * 256):
                    if self._download_cancelled.is_set():
                        raise RuntimeError("ダウンロードをキャンセルしました")
                    if not chunk:
                        continue
                    f.write(chunk)
                    downloaded += len(chunk)
                    self._update_progress(downloaded, total)

            self._log(f"ダウンロード完了: {target_path}")
            return target_path
        finally:
            self._close_progress_dialog()

    def _verify_sha256(self, release_info: Dict[str, Any], asset: Dict[str, Any], file_path: Path) -> None:
        file_name = str(asset.get("name", "")).strip()
        expected_hash = self._find_expected_sha256(release_info, file_name)
        if not expected_hash:
            raise RuntimeError("SHA256 がリリース情報に見つかりませんでした")

        actual_hash = self._sha256_file(file_path)
        if actual_hash.lower() != expected_hash.lower():
            raise RuntimeError("SHA256 検証に失敗しました")

        self._log("SHA256 検証に成功しました")

    def _find_expected_sha256(self, release_info: Dict[str, Any], file_name: str) -> str:
        assets = release_info.get("assets", [])
        if isinstance(assets, list):
            checksum_asset = next(
                (
                    a
                    for a in assets
                    if str(a.get("name", "")).lower() in ("checksums.txt", "sha256sums.txt")
                    or "checksum" in str(a.get("name", "")).lower()
                ),
                None,
            )
            if checksum_asset:
                checksum_url = str(checksum_asset.get("browser_download_url", "")).strip()
                if checksum_url:
                    content = self._session.get(checksum_url, timeout=15).text
                    parsed = self._parse_checksum_lines(content)
                    if file_name in parsed:
                        return parsed[file_name]

        body = str(release_info.get("body", "") or "")
        if body:
            parsed = self._parse_checksum_lines(body)
            if file_name in parsed:
                return parsed[file_name]

            escaped = re.escape(file_name)
            patterns = [
                rf"(?im){escaped}\s*[:=]\s*([a-f0-9]{{64}})",
                rf"(?im)([a-f0-9]{{64}})\s+\*?{escaped}",
            ]
            for pattern in patterns:
                match = re.search(pattern, body)
                if match:
                    return match.group(1)

        return ""

    def _parse_checksum_lines(self, text: str) -> Dict[str, str]:
        result: Dict[str, str] = {}
        for line in text.splitlines():
            cleaned = line.strip()
            if not cleaned:
                continue

            match = re.match(r"(?i)^([a-f0-9]{64})\s+\*?(.+)$", cleaned)
            if match:
                result[match.group(2).strip()] = match.group(1)
                continue

            match = re.match(r"(?i)^(.+?)\s*[:=]\s*([a-f0-9]{64})$", cleaned)
            if match:
                result[match.group(1).strip()] = match.group(2)

        return result

    def _apply_update(self, downloaded_exe: Path, asset_name: str) -> None:
        if not getattr(sys, "frozen", False):
            self._show_info(
                "更新ファイル取得完了",
                f"開発実行中のため自動差し替えは行いません。\nダウンロード先: {downloaded_exe}",
            )
            return

        current_exe = Path(sys.executable)
        latest_name = asset_name.strip() or downloaded_exe.name
        launch_exe = current_exe.parent / latest_name
        replace_in_place = launch_exe.resolve() == current_exe.resolve()

        staged_new = current_exe.with_suffix(".new.exe")
        backup_exe = current_exe.with_suffix(".bak.exe")

        if replace_in_place:
            shutil.copy2(downloaded_exe, staged_new)

        bat_path = downloaded_exe.parent / "apply_update.bat"
        pid = os.getpid()
        bat_content = self._build_update_bat(
            current_exe=current_exe,
            launch_exe=launch_exe,
            downloaded_exe=downloaded_exe,
            staged_new_exe=staged_new,
            backup_exe=backup_exe,
            pid=pid,
            replace_in_place=replace_in_place,
        )
        bat_path.write_text(bat_content, encoding="utf-8")

        creationflags = 0
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            creationflags |= int(subprocess.CREATE_NO_WINDOW)

        startupinfo = None
        if hasattr(subprocess, "STARTUPINFO"):
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startupinfo.wShowWindow = 0

        subprocess.Popen(
            ["cmd", "/d", "/q", "/c", str(bat_path)],
            creationflags=creationflags,
            startupinfo=startupinfo,
        )
        self._log("更新を適用します。アプリを自動で再起動します")
        self.root.after(150, self.root.destroy)

    def _build_update_bat(
        self,
        current_exe: Path,
        launch_exe: Path,
        downloaded_exe: Path,
        staged_new_exe: Path,
        backup_exe: Path,
        pid: int,
        replace_in_place: bool,
    ) -> str:
        replace_mode = "1" if replace_in_place else "0"
        return (
            "@echo off\n"
            "setlocal EnableDelayedExpansion\n"
            f"set \"CURRENT={current_exe}\"\n"
            f"set \"LAUNCH={launch_exe}\"\n"
            f"set \"DOWNLOADED={downloaded_exe}\"\n"
            f"set \"STAGED_NEW={staged_new_exe}\"\n"
            f"set \"BACKUP={backup_exe}\"\n"
            f"set \"PID={pid}\"\n"
            f"set \"REPLACE_MODE={replace_mode}\"\n"
            "for /L %%i in (1,1,90) do (\n"
            "  tasklist /FI \"PID eq %PID%\" | findstr /I \"%PID%\" >nul\n"
            "  if errorlevel 1 goto replace\n"
            "  timeout /t 1 /nobreak >nul\n"
            ")\n"
            ":replace\n"
            "timeout /t 3 /nobreak >nul\n"
            "set \"REPLACED=0\"\n"
            "if \"%REPLACE_MODE%\"==\"1\" (\n"
            "  for /L %%i in (1,1,30) do (\n"
            "    if exist \"%BACKUP%\" del /f /q \"%BACKUP%\" >nul 2>nul\n"
            "    if exist \"%CURRENT%\" move /y \"%CURRENT%\" \"%BACKUP%\" >nul 2>nul\n"
            "    if exist \"%STAGED_NEW%\" move /y \"%STAGED_NEW%\" \"%CURRENT%\" >nul 2>nul\n"
            "    if exist \"%CURRENT%\" (\n"
            "      set \"REPLACED=1\"\n"
            "      goto launch\n"
            "    )\n"
            "    timeout /t 1 /nobreak >nul\n"
            "  )\n"
            ") else (\n"
            "  for /L %%i in (1,1,30) do (\n"
            "    if exist \"%DOWNLOADED%\" copy /y \"%DOWNLOADED%\" \"%LAUNCH%\" >nul 2>nul\n"
            "    if exist \"%LAUNCH%\" (\n"
            "      set \"REPLACED=1\"\n"
            "      goto launch\n"
            "    )\n"
            "    timeout /t 1 /nobreak >nul\n"
            "  )\n"
            ")\n"
            ":launch\n"
            "if \"%REPLACED%\"==\"1\" (\n"
            "  timeout /t 5 /nobreak >nul\n"
            "  if \"%REPLACE_MODE%\"==\"1\" (\n"
            "    start \"\" /D \"%~dp0\" \"%CURRENT%\"\n"
            "  ) else (\n"
            "    if /I NOT \"%CURRENT%\"==\"%LAUNCH%\" if exist \"%CURRENT%\" del /f /q \"%CURRENT%\" >nul 2>nul\n"
            "    start \"\" /D \"%~dp0\" \"%LAUNCH%\"\n"
            "  )\n"
            ")\n"
            "if exist \"%BACKUP%\" del /f /q \"%BACKUP%\" >nul 2>nul\n"
            "if exist \"%DOWNLOADED%\" del /f /q \"%DOWNLOADED%\" >nul 2>nul\n"
            "if exist \"%STAGED_NEW%\" del /f /q \"%STAGED_NEW%\" >nul 2>nul\n"
            "timeout /t 1 /nobreak >nul\n"
            "del /f /q \"%~f0\"\n"
        )

    def _sha256_file(self, file_path: Path) -> str:
        hash_obj = hashlib.sha256()
        with file_path.open("rb") as f:
            for chunk in iter(lambda: f.read(1024 * 256), b""):
                hash_obj.update(chunk)
        return hash_obj.hexdigest()

    def _show_progress_dialog(self, file_name: str) -> None:
        def _create() -> None:
            if self._progress_dialog and self._progress_dialog.winfo_exists():
                self._progress_dialog.destroy()

            dialog = tk.Toplevel(self.root)
            dialog.title("アップデートをダウンロード中")
            dialog.geometry("420x140")
            dialog.resizable(False, False)
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.protocol("WM_DELETE_WINDOW", self._download_cancelled.set)

            frame = ttk.Frame(dialog, padding=12)
            frame.pack(fill=tk.BOTH, expand=True)

            label = ttk.Label(frame, text=f"{file_name} をダウンロード中")
            label.pack(anchor="w")

            self._progress_var = tk.DoubleVar(value=0.0)
            bar = ttk.Progressbar(frame, variable=self._progress_var, maximum=100)
            bar.pack(fill=tk.X, pady=(10, 6))

            self._progress_label_var = tk.StringVar(value="0%")
            ttk.Label(frame, textvariable=self._progress_label_var).pack(anchor="w")
            ttk.Button(frame, text="キャンセル", command=self._download_cancelled.set).pack(anchor="e", pady=(8, 0))

            self._progress_dialog = dialog

        self._run_on_ui(_create)

    def _update_progress(self, downloaded: int, total: int) -> None:
        def _update() -> None:
            if not self._progress_dialog or not self._progress_dialog.winfo_exists():
                return
            if not self._progress_var or not self._progress_label_var:
                return

            if total > 0:
                percent = (downloaded / total) * 100
                self._progress_var.set(percent)
                self._progress_label_var.set(f"{percent:.1f}% ({downloaded}/{total} bytes)")
            else:
                self._progress_var.set(0.0)
                self._progress_label_var.set(f"{downloaded} bytes")

        self.root.after(0, _update)

    def _close_progress_dialog(self) -> None:
        def _close() -> None:
            if self._progress_dialog and self._progress_dialog.winfo_exists():
                self._progress_dialog.destroy()
            self._progress_dialog = None
            self._progress_var = None
            self._progress_label_var = None

        self.root.after(0, _close)

    def _ask_update_choice(self, message: str, latest_version: str) -> str:
        result: Dict[str, str] = {"value": "no"}

        def _prompt() -> None:
            yes = messagebox.askyesno("アップデート", message)
            if yes:
                result["value"] = "yes"
                return

            skip = messagebox.askyesno("アップデート", f"バージョン {latest_version} の通知をスキップしますか？")
            result["value"] = "skip" if skip else "no"

        self._run_on_ui(_prompt)
        return result["value"]

    def _show_info(self, title: str, message: str) -> None:
        self._run_on_ui(lambda: messagebox.showinfo(title, message))

    def _show_error(self, title: str, message: str) -> None:
        self._run_on_ui(lambda: messagebox.showerror(title, message))

    def _run_on_ui(self, func: Callable[[], None]) -> None:
        done = threading.Event()

        def wrapped() -> None:
            try:
                func()
            finally:
                done.set()

        self.root.after(0, wrapped)
        done.wait()

    def _log(self, message: str) -> None:
        if self.log_callback:
            self.root.after(0, lambda: self.log_callback(message))

    def _utc_now_iso(self) -> str:
        return datetime.now(timezone.utc).replace(microsecond=0).isoformat()
