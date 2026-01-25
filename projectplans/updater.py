from __future__ import annotations

import json
import shutil
import sys
import tempfile
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from pathlib import Path

from PyQt6.QtCore import QObject, QThread, QTimer, pyqtSignal
from PyQt6.QtWidgets import QProgressDialog, QMessageBox

APP_USER_AGENT = "ProjectPlans-Updater"
VERSION_FILE_NAME = "version.txt"
UPDATE_CONFIG_NAME = "update_config.json"
DEFAULT_VERSION = "0.0.0"
DEFAULT_CHECK_INTERVAL_DAYS = 7

DEFAULT_UPDATE_CONFIG = {
    "github": {
        "owner": "your-username",
        "repo": "your-repo",
        "token": "your-personal-access-token",
        "enabled": False,
        "private": True,
    },
    "auto_check": False,
    "check_interval_days": DEFAULT_CHECK_INTERVAL_DAYS,
    "update_instructions": "Edit this file to enable updates.",
}


@dataclass
class UpdateConfig:
    owner: str
    repo: str
    token: str
    enabled: bool
    private: bool
    auto_check: bool
    interval_days: int
    path: Path

    def is_ready(self) -> bool:
        return bool(self.enabled and self.owner and self.repo and self.token)


class ReleaseCheckWorker(QObject):
    finished = pyqtSignal(object, str)

    def __init__(self, config: UpdateConfig) -> None:
        super().__init__()
        self._config = config

    def run(self) -> None:
        release = None
        error = None
        try:
            release = _fetch_latest_release(self._config)
        except UpdateCheckError as exc:
            error = str(exc)
        except Exception as exc:  # pragma: no cover - safety net for unexpected errors
            error = f"Update check failed: {exc}"
        self.finished.emit(release, error or "")


class DownloadWorker(QObject):
    progress = pyqtSignal(int)
    status = pyqtSignal(str)
    finished = pyqtSignal(object, str)

    def __init__(self, url: str, token: str, target_path: Path) -> None:
        super().__init__()
        self._url = url
        self._token = token
        self._target_path = target_path

    def run(self) -> None:
        try:
            import requests
        except ImportError as exc:
            self.finished.emit(None, f"Missing dependency: {exc}")
            return
        self.status.emit("Connecting...")
        headers = {
            "Authorization": f"token {self._token}",
            "Accept": "application/octet-stream",
            "User-Agent": APP_USER_AGENT,
        }
        try:
            response = requests.get(self._url, headers=headers, stream=True)
        except Exception as exc:
            self.finished.emit(None, f"Download failed: {exc}")
            return
        if response.status_code != 200:
            self.finished.emit(
                None,
                f"Download failed (HTTP {response.status_code}).",
            )
            return
        total = response.headers.get("Content-Length")
        try:
            total_bytes = int(total) if total else 0
        except ValueError:
            total_bytes = 0
        if not total_bytes:
            self.progress.emit(-1)
        bytes_read = 0
        self.status.emit("Downloading...")
        try:
            with self._target_path.open("wb") as handle:
                for chunk in response.iter_content(chunk_size=8192):
                    if not chunk:
                        continue
                    handle.write(chunk)
                    if total_bytes:
                        bytes_read += len(chunk)
                        percent = int((bytes_read / total_bytes) * 100)
                        self.progress.emit(max(0, min(100, percent)))
        except Exception as exc:
            self.finished.emit(None, f"Download failed: {exc}")
            return
        self.finished.emit(self._target_path, "")


class UpdateCheckError(RuntimeError):
    pass


class UpdateManager(QObject):
    def __init__(self, parent, settings) -> None:
        super().__init__(parent)
        self._parent = parent
        self._settings = settings
        self._check_thread: QThread | None = None
        self._check_worker: ReleaseCheckWorker | None = None
        self._download_thread: QThread | None = None
        self._download_worker: DownloadWorker | None = None
        self._progress_dialog: QProgressDialog | None = None
        self._download_temp_dir: tempfile.TemporaryDirectory | None = None
        self._pending_release: dict | None = None
        self._pending_config: UpdateConfig | None = None
        self._current_version = DEFAULT_VERSION
        self._check_is_auto = False

    def schedule_auto_check(self) -> None:
        QTimer.singleShot(2000, self._auto_check)

    def manual_check(self) -> None:
        if self._check_thread and self._check_thread.isRunning():
            QMessageBox.information(self._parent, "Updates", "An update check is already running.")
            return
        if not _ensure_dependencies(interactive=True, parent=self._parent):
            return
        config, created = _load_update_config(_app_dir(), interactive=True, parent=self._parent)
        if created:
            return
        if not config.enabled:
            QMessageBox.information(
                self._parent,
                "Updates Disabled",
                f"Updates are disabled.\nEdit {config.path} to enable them.",
            )
            return
        if not config.owner or not config.repo or not config.token:
            QMessageBox.warning(
                self._parent,
                "Update Configuration",
                f"Update configuration is incomplete.\nEdit {config.path} and set owner, repo, and token.",
            )
            return
        self._current_version = _read_current_version(_app_dir())
        self._pending_config = config
        self._check_is_auto = False
        self._start_release_check(config)

    def _auto_check(self) -> None:
        if self._check_thread and self._check_thread.isRunning():
            return
        if not _ensure_dependencies(interactive=False, parent=self._parent):
            return
        config, _created = _load_update_config(_app_dir(), interactive=False, parent=self._parent)
        if not config.enabled or not config.auto_check:
            return
        if not config.owner or not config.repo or not config.token:
            return
        if not self._auto_check_due(config.interval_days):
            return
        self._current_version = _read_current_version(_app_dir())
        self._pending_config = config
        self._check_is_auto = True
        self._start_release_check(config)

    def _auto_check_due(self, interval_days: int) -> bool:
        interval_days = max(1, int(interval_days))
        value = self._settings.value("Updates/LastCheck", "")
        if value:
            try:
                last_check = datetime.fromisoformat(str(value))
            except ValueError:
                last_check = None
        else:
            last_check = None
        if last_check is not None and last_check.tzinfo is None:
            last_check = last_check.replace(tzinfo=timezone.utc)
        if last_check is None:
            return True
        return datetime.now(timezone.utc) - last_check >= timedelta(days=interval_days)

    def _start_release_check(self, config: UpdateConfig) -> None:
        self._progress_dialog = QProgressDialog(
            "Checking for updates...", "Cancel", 0, 0, self._parent
        )
        self._progress_dialog.setWindowTitle("Updates")
        self._progress_dialog.setAutoClose(False)
        self._progress_dialog.canceled.connect(self._progress_dialog.close)
        self._progress_dialog.show()

        thread = QThread(self._parent)
        worker = ReleaseCheckWorker(config)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._handle_release_check)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._check_thread = thread
        self._check_worker = worker
        thread.start()

    def _handle_release_check(self, release: dict | None, error: str) -> None:
        if self._progress_dialog:
            self._progress_dialog.close()
            self._progress_dialog = None
        if self._check_is_auto:
            self._settings.setValue(
                "Updates/LastCheck", datetime.now(timezone.utc).isoformat()
            )
            self._settings.sync()
        if error:
            if not self._check_is_auto:
                QMessageBox.warning(self._parent, "Update Check", error)
            return
        if not release:
            if not self._check_is_auto:
                QMessageBox.warning(self._parent, "Update Check", "No release data found.")
            return
        latest_version, version_error = _release_version(release)
        if version_error:
            if not self._check_is_auto:
                QMessageBox.warning(self._parent, "Update Check", version_error)
            return
        is_newer, compare_error = _compare_versions(latest_version, self._current_version)
        if compare_error:
            if not self._check_is_auto:
                QMessageBox.warning(self._parent, "Update Check", compare_error)
            return
        if not is_newer:
            if not self._check_is_auto:
                QMessageBox.information(self._parent, "Updates", "You are up to date.")
            return
        if self._check_is_auto:
            self._show_update_available_status(latest_version)
            return
        self._pending_release = release
        self._prompt_install(latest_version)

    def _prompt_install(self, latest_version: str) -> None:
        release = self._pending_release or {}
        body = (release.get("body") or "").strip() or "No release notes provided."
        message = (
            f"Update available.\n\nCurrent version: {self._current_version}\n"
            f"Latest version: {latest_version}\n\nInstall now?"
        )
        prompt = QMessageBox(self._parent)
        prompt.setWindowTitle("Update Available")
        prompt.setText(message)
        prompt.setDetailedText(body)
        prompt.setIcon(QMessageBox.Icon.Question)
        prompt.setStandardButtons(
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if prompt.exec() != QMessageBox.StandardButton.Yes:
            return
        self._start_download(latest_version)

    def _start_download(self, latest_version: str) -> None:
        if self._download_thread and self._download_thread.isRunning():
            QMessageBox.information(self._parent, "Updates", "Download already in progress.")
            return
        if self._pending_config is None:
            QMessageBox.warning(self._parent, "Updates", "Update configuration is missing.")
            return
        release = self._pending_release or {}
        assets = release.get("assets", []) if isinstance(release.get("assets"), list) else []
        asset = _select_asset(assets)
        if asset is None:
            QMessageBox.warning(self._parent, "Updates", "No usable release assets found.")
            return
        asset_name = asset.get("name", "update")
        asset_url = asset.get("url")
        if not asset_url:
            QMessageBox.warning(self._parent, "Updates", "Selected asset has no download URL.")
            return
        ext = _asset_extension(asset_name)
        self._download_temp_dir = tempfile.TemporaryDirectory()
        target_path = Path(self._download_temp_dir.name) / f"update_v{latest_version}.{ext}"

        self._progress_dialog = QProgressDialog(
            "Downloading update...", "Cancel", 0, 100, self._parent
        )
        self._progress_dialog.setWindowTitle("Download Update")
        self._progress_dialog.setAutoClose(False)
        self._progress_dialog.canceled.connect(self._progress_dialog.close)
        self._progress_dialog.show()

        thread = QThread(self._parent)
        worker = DownloadWorker(asset_url, self._pending_config.token, target_path)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.status.connect(self._handle_download_status)
        worker.progress.connect(self._handle_download_progress)
        worker.finished.connect(lambda path, err: self._handle_download_finished(path, err, asset_name, latest_version))
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self._download_thread = thread
        self._download_worker = worker
        thread.start()

    def _handle_download_status(self, text: str) -> None:
        if self._progress_dialog:
            self._progress_dialog.setLabelText(text)

    def _handle_download_progress(self, value: int) -> None:
        if self._progress_dialog is None:
            return
        if value < 0:
            self._progress_dialog.setRange(0, 0)
        else:
            if self._progress_dialog.maximum() == 0:
                self._progress_dialog.setRange(0, 100)
            self._progress_dialog.setValue(value)

    def _handle_download_finished(
        self, path: Path | None, error: str, asset_name: str, latest_version: str
    ) -> None:
        if self._progress_dialog:
            self._progress_dialog.close()
            self._progress_dialog = None
        try:
            if error:
                QMessageBox.warning(self._parent, "Download Failed", error)
                return
            if path is None or not Path(path).exists():
                QMessageBox.warning(self._parent, "Download Failed", "Download did not produce a file.")
                return
            success = self._install_update(Path(path), asset_name, latest_version)
            if success:
                self._restart_application()
        finally:
            if self._download_temp_dir:
                self._download_temp_dir.cleanup()
                self._download_temp_dir = None

    def _install_update(self, asset_path: Path, asset_name: str, latest_version: str) -> bool:
        lower_name = asset_name.lower()
        if lower_name.endswith(".zip"):
            return self._install_zip_update(asset_path, latest_version)
        if lower_name.endswith(".pyw") or lower_name.endswith(".py"):
            return self._install_script_update(asset_path, latest_version)
        if lower_name.endswith(".exe") or lower_name.endswith(".tar.gz"):
            return self._install_executable_update(asset_path, latest_version)
        QMessageBox.warning(
            self._parent,
            "Update Failed",
            f"Unsupported update asset type: {asset_name}",
        )
        return False

    def _install_zip_update(self, asset_path: Path, latest_version: str) -> bool:
        app_dir = _app_dir()
        with tempfile.TemporaryDirectory() as temp_dir:
            try:
                with zipfile.ZipFile(asset_path, "r") as archive:
                    archive.extractall(temp_dir)
            except Exception as exc:
                QMessageBox.warning(self._parent, "Update Failed", f"Could not extract ZIP.\n{exc}")
                return False
            manifest_path = Path(temp_dir) / "update_manifest.json"
            if manifest_path.exists():
                return self._install_manifest_update(manifest_path, Path(temp_dir), app_dir, latest_version)
            return self._install_zip_script_update(Path(temp_dir), latest_version)

    def _install_manifest_update(
        self,
        manifest_path: Path,
        source_root: Path,
        app_dir: Path,
        latest_version: str,
    ) -> bool:
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception as exc:
            QMessageBox.warning(self._parent, "Update Failed", f"Invalid manifest.\n{exc}")
            return False
        entries = manifest.get("files", [])
        if not isinstance(entries, list):
            QMessageBox.warning(self._parent, "Update Failed", "Manifest files list is invalid.")
            return False
        backup_dir = app_dir / "backups" / f"update_backup_{datetime.now():%Y%m%d_%H%M%S}"
        updated = 0
        errors: list[str] = []
        for entry in entries:
            if not isinstance(entry, dict):
                errors.append("Manifest entry is not an object.")
                continue
            source = entry.get("source")
            target = entry.get("target")
            if not source or not target:
                errors.append("Manifest entry missing source or target.")
                continue
            if Path(source).is_absolute():
                errors.append(f"Manifest source must be relative: {source}")
                continue
            if Path(target).is_absolute():
                errors.append(f"Manifest target must be relative: {target}")
                continue
            backup = bool(entry.get("backup", True))
            required = bool(entry.get("required", False))
            source_path = (source_root / source).resolve()
            if not source_path.exists():
                if required:
                    errors.append(f"Missing required source: {source}")
                continue
            target_path = (app_dir / target).resolve()
            if not target_path.is_relative_to(app_dir.resolve()):
                errors.append(f"Target escapes app directory: {target}")
                continue
            if backup and target_path.exists():
                backup_target = backup_dir / target
                try:
                    backup_target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(target_path, backup_target)
                except Exception as exc:
                    errors.append(f"Backup failed for {target}: {exc}")
            try:
                if source_path.is_dir():
                    if target_path.exists():
                        if target_path.is_dir():
                            shutil.rmtree(target_path)
                        else:
                            target_path.unlink()
                    shutil.copytree(source_path, target_path)
                else:
                    target_path.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source_path, target_path)
                updated += 1
            except Exception as exc:
                errors.append(f"Install failed for {target}: {exc}")

        _write_version_file(app_dir, latest_version, errors)

        success = not errors or updated > 0
        if errors:
            self._show_warning("Update completed with warnings.", "\n".join(errors))
        else:
            QMessageBox.information(self._parent, "Update Complete", "Update installed successfully.")
        return success

    def _install_zip_script_update(self, source_root: Path, latest_version: str) -> bool:
        script_path = _main_script_path()
        if not script_path or not script_path.exists():
            QMessageBox.warning(
                self._parent,
                "Update Failed",
                "Unable to locate the running script for ZIP update.",
            )
            return False
        target_name = script_path.name
        candidates = [
            path
            for path in source_root.rglob("*")
            if path.is_file()
            and path.name == target_name
            and path.suffix.lower() in (".py", ".pyw")
        ]
        if not candidates:
            QMessageBox.warning(
                self._parent,
                "Update Failed",
                f"Could not find {target_name} inside the ZIP.",
            )
            return False
        source_path = candidates[0]
        backup_path = script_path.with_name(f"{script_path.name}.backup")
        try:
            shutil.copy2(script_path, backup_path)
            shutil.copy2(source_path, script_path)
        except Exception as exc:
            QMessageBox.warning(self._parent, "Update Failed", f"Script update failed.\n{exc}")
            return False
        errors: list[str] = []
        _write_version_file(_app_dir(), latest_version, errors)
        if errors:
            self._show_warning("Update completed with warnings.", "\n".join(errors))
        else:
            QMessageBox.information(self._parent, "Update Complete", "Update installed successfully.")
        return True

    def _install_script_update(self, asset_path: Path, latest_version: str) -> bool:
        script_path = _main_script_path()
        if not script_path or not script_path.exists():
            QMessageBox.warning(
                self._parent,
                "Update Failed",
                "Unable to locate the running script for update.",
            )
            return False
        backup_path = script_path.with_name(f"{script_path.name}.backup")
        try:
            shutil.copy2(script_path, backup_path)
            shutil.copy2(asset_path, script_path)
        except Exception as exc:
            QMessageBox.warning(self._parent, "Update Failed", f"Script update failed.\n{exc}")
            return False
        _write_version_file(_app_dir(), latest_version, ignore_errors=True)
        QMessageBox.information(self._parent, "Update Complete", "Update installed successfully.")
        return True

    def _install_executable_update(self, asset_path: Path, latest_version: str) -> bool:
        if not getattr(sys, "frozen", False):
            QMessageBox.warning(
                self._parent,
                "Update Failed",
                "Executable updates are only supported for compiled builds.",
            )
            return False
        exe_path = Path(sys.executable).resolve()
        backup_path = exe_path.with_name(f"{exe_path.name}.backup")
        try:
            shutil.copy2(exe_path, backup_path)
        except Exception as exc:
            QMessageBox.warning(self._parent, "Update Failed", f"Backup failed.\n{exc}")
            return False
        QMessageBox.information(
            self._parent,
            "Manual Update Required",
            "The update has been downloaded.\n"
            f"Replace {exe_path} with:\n{asset_path}\n"
            "The application will now restart.",
        )
        _write_version_file(_app_dir(), latest_version, ignore_errors=True)
        return True

    def _restart_application(self) -> None:
        maybe_save = getattr(self._parent, "maybe_save", None)
        if callable(maybe_save) and not maybe_save():
            return
        self._close_open_resources()
        command = _restart_command()
        if not command:
            QMessageBox.warning(self._parent, "Restart Failed", "Unable to restart the application.")
            return
        try:
            import subprocess

            subprocess.Popen(command)
        except Exception as exc:
            QMessageBox.warning(self._parent, "Restart Failed", f"Could not restart.\n{exc}")
            return
        self._parent.close()

    def _close_open_resources(self) -> None:
        close_method = getattr(self._parent, "close_open_resources", None)
        if callable(close_method):
            close_method()

    def _show_update_available_status(self, latest_version: str) -> None:
        status = self._parent.statusBar()
        status.showMessage(
            f"Update available: {latest_version}. Use Help > Check for Updates.",
            10000,
        )

    def _show_warning(self, message: str, details: str) -> None:
        dialog = QMessageBox(self._parent)
        dialog.setWindowTitle("Update Warnings")
        dialog.setText(message)
        dialog.setIcon(QMessageBox.Icon.Warning)
        dialog.setDetailedText(details)
        dialog.exec()


def _ensure_dependencies(*, interactive: bool, parent) -> bool:
    try:
        import requests  # noqa: F401
        from packaging.version import Version  # noqa: F401
    except ImportError as exc:
        if interactive:
            QMessageBox.warning(
                parent,
                "Updates",
                f"Missing dependency: {exc}.\nInstall 'requests' and 'packaging' to use updates.",
            )
        return False
    return True


def _load_update_config(
    app_dir: Path, *, interactive: bool, parent
) -> tuple[UpdateConfig, bool]:
    config_path = app_dir / UPDATE_CONFIG_NAME
    created = False
    if not config_path.exists():
        config_path.write_text(json.dumps(DEFAULT_UPDATE_CONFIG, indent=2), encoding="utf-8")
        created = True
        if interactive:
            QMessageBox.information(
                parent,
                "Update Configuration",
                f"Created {config_path}.\nEdit it to enable updates.",
            )
    if not config_path.exists():
        return _default_config(config_path), created
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        if interactive:
            QMessageBox.warning(
                parent,
                "Update Configuration",
                f"{config_path} is malformed. Updates are disabled.",
            )
        return _default_config(config_path), created
    github = raw.get("github") if isinstance(raw, dict) else {}
    owner = str(github.get("owner", "") or "").strip()
    repo = str(github.get("repo", "") or "").strip()
    token = str(github.get("token", "") or "").strip()
    enabled = bool(github.get("enabled", False))
    private = bool(github.get("private", True))
    auto_check = bool(raw.get("auto_check", False))
    interval_days = raw.get("check_interval_days", DEFAULT_CHECK_INTERVAL_DAYS)
    try:
        interval_days = int(interval_days)
    except (TypeError, ValueError):
        interval_days = DEFAULT_CHECK_INTERVAL_DAYS
    if interval_days < 1:
        interval_days = DEFAULT_CHECK_INTERVAL_DAYS
    config = UpdateConfig(
        owner=owner,
        repo=repo,
        token=token,
        enabled=enabled,
        private=private,
        auto_check=auto_check,
        interval_days=interval_days,
        path=config_path,
    )
    return config, created


def _default_config(config_path: Path) -> UpdateConfig:
    return UpdateConfig(
        owner="",
        repo="",
        token="",
        enabled=False,
        private=True,
        auto_check=False,
        interval_days=DEFAULT_CHECK_INTERVAL_DAYS,
        path=config_path,
    )


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    script_path = _main_script_path()
    if script_path:
        return script_path.resolve().parent
    return Path.cwd()


def _main_script_path() -> Path | None:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    argv_path = Path(sys.argv[0])
    if argv_path.is_file():
        return argv_path.resolve()
    spec = getattr(sys.modules.get("__main__"), "__spec__", None)
    if spec and getattr(spec, "origin", None):
        origin = Path(spec.origin)
        if origin.is_file():
            return origin.resolve()
    return None


def _restart_command() -> list[str]:
    if getattr(sys, "frozen", False):
        return [sys.executable]
    spec = getattr(sys.modules.get("__main__"), "__spec__", None)
    if spec and getattr(spec, "name", None):
        return [sys.executable, "-m", spec.name]
    script_path = _main_script_path()
    if script_path:
        return [sys.executable, str(script_path)]
    return []


def _read_current_version(app_dir: Path) -> str:
    version_path = app_dir / VERSION_FILE_NAME
    if not version_path.exists():
        return DEFAULT_VERSION
    try:
        return version_path.read_text(encoding="utf-8").strip() or DEFAULT_VERSION
    except Exception:
        return DEFAULT_VERSION


def _write_version_file(app_dir: Path, version: str, errors: list[str] | None = None, ignore_errors: bool = False) -> None:
    version_path = app_dir / VERSION_FILE_NAME
    try:
        version_path.write_text(version, encoding="utf-8")
    except Exception as exc:
        if ignore_errors:
            return
        if errors is not None:
            errors.append(f"Failed to write version.txt: {exc}")


def _fetch_latest_release(config: UpdateConfig) -> dict:
    try:
        import requests
    except ImportError as exc:
        raise UpdateCheckError(f"Missing dependency: {exc}") from exc
    url = f"https://api.github.com/repos/{config.owner}/{config.repo}/releases/latest"
    headers = {
        "Authorization": f"Bearer {config.token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": APP_USER_AGENT,
        "X-GitHub-Api-Version": "2022-11-28",
    }
    try:
        response = requests.get(url, headers=headers, timeout=15)
    except Exception as exc:
        raise UpdateCheckError(f"Update check failed: {exc}") from exc
    if response.status_code == 404:
        if config.private:
            raise UpdateCheckError("No releases found.")
        raise UpdateCheckError("Repository not found.")
    if response.status_code == 401:
        raise UpdateCheckError("Authentication failed. Check your token.")
    if response.status_code == 403:
        raise UpdateCheckError("Insufficient permissions to access the repository.")
    if response.status_code != 200:
        raise UpdateCheckError(
            f"Update check failed (HTTP {response.status_code})."
        )
    try:
        return response.json()
    except Exception as exc:
        raise UpdateCheckError(f"Invalid response from GitHub.\n{exc}") from exc


def _select_asset(assets: list[dict]) -> dict | None:
    if not assets:
        return None
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if name.endswith(".pyw") or name.endswith(".py"):
            return asset
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if name.endswith(".zip"):
            return asset
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if name.endswith(".exe") or name.endswith(".tar.gz"):
            return asset
    keywords = ("app", "main", "project", "source")
    for asset in assets:
        name = str(asset.get("name", "")).lower()
        if any(keyword in name for keyword in keywords):
            return asset
    return assets[0]


def _asset_extension(name: str) -> str:
    lowered = name.lower()
    if lowered.endswith(".tar.gz"):
        return "tar.gz"
    suffix = Path(name).suffix.lower().lstrip(".")
    return suffix or "bin"


def _release_version(release: dict) -> tuple[str | None, str | None]:
    tag = str(release.get("tag_name", "")).strip()
    if not tag:
        return None, "Release tag is missing."
    if tag[:1] in ("v", "V"):
        tag = tag[1:]
    return tag, None


def _compare_versions(latest: str, current: str) -> tuple[bool | None, str | None]:
    try:
        from packaging.version import Version, InvalidVersion
    except ImportError:
        return None, "Missing dependency: packaging."
    try:
        latest_version = Version(latest)
        current_version = Version(current)
    except InvalidVersion as exc:
        return None, f"Invalid version format: {exc}"
    return latest_version > current_version, None
