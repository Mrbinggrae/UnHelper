from __future__ import annotations

import base64
from dataclasses import dataclass

from PySide6.QtCore import QSettings

try:
    import win32crypt
except ImportError:  # pragma: no cover - UnHelper is distributed for Windows.
    win32crypt = None


class CredentialError(RuntimeError):
    """Raised when WMS credentials cannot be stored or restored safely."""


@dataclass(frozen=True)
class WMSCredentials:
    wms_id: str = ""
    password: str = ""

    @property
    def is_complete(self) -> bool:
        return bool(self.wms_id and self.password)


class WMSCredentialStore:
    """Persist a WMS ID and a Windows-DPAPI-protected password in QSettings."""

    ID_KEY = "wms_id"
    PASSWORD_KEY = "wms_password_dpapi"
    LEGACY_PASSWORD_KEYS = ("wms_password", "wms_pw")
    _DPAPI_DESCRIPTION = "UnHelper WMS credential"

    def __init__(self, settings: QSettings | None = None):
        self.settings = settings if settings is not None else QSettings("Mrbinggrae", "UnHelper")

    def load(self) -> WMSCredentials:
        wms_id = self._setting_text(self.ID_KEY).strip()

        if self.settings.contains(self.PASSWORD_KEY):
            # A protected value always wins. Remove obsolete plaintext copies
            # before attempting decryption so they cannot linger indefinitely.
            self._remove_legacy_passwords()
            self._sync("저장된 WMS 계정 정보를 정리하지 못했습니다.")
            password = self._unprotect(self._setting_text(self.PASSWORD_KEY))
            return WMSCredentials(wms_id=wms_id, password=password)

        legacy_password, legacy_found = self._read_legacy_password()
        if not legacy_found:
            return WMSCredentials(wms_id=wms_id, password="")

        # Protect first. If Windows DPAPI is temporarily unavailable, do not
        # destroy the only recoverable copy of the user's password.
        protected = self._protect(legacy_password) if legacy_password else ""
        if protected:
            self.settings.setValue(self.PASSWORD_KEY, protected)
        else:
            self.settings.remove(self.PASSWORD_KEY)
        self._remove_legacy_passwords()
        self._sync("기존 WMS 비밀번호를 안전한 저장 방식으로 전환하지 못했습니다.")
        return WMSCredentials(wms_id=wms_id, password=legacy_password)

    def save(self, credentials: WMSCredentials) -> None:
        wms_id = str(credentials.wms_id or "").strip()
        password = str(credentials.password or "")

        # Encrypt before changing QSettings. This avoids leaving a partially
        # updated ID behind when DPAPI rejects the password operation.
        protected = self._protect(password) if password else ""

        self.settings.setValue(self.ID_KEY, wms_id)
        if protected:
            self.settings.setValue(self.PASSWORD_KEY, protected)
        else:
            self.settings.remove(self.PASSWORD_KEY)
        self._remove_legacy_passwords()
        self._sync("WMS 계정 정보를 저장하지 못했습니다.")

    def clear_password(self) -> None:
        self.settings.remove(self.PASSWORD_KEY)
        self._remove_legacy_passwords()
        self._sync("저장된 WMS 비밀번호를 지우지 못했습니다.")

    def _setting_text(self, key: str) -> str:
        value = self.settings.value(key, "")
        return "" if value is None else str(value)

    def _read_legacy_password(self) -> tuple[str, bool]:
        found = False
        password = ""
        for key in self.LEGACY_PASSWORD_KEYS:
            if not self.settings.contains(key):
                continue
            found = True
            candidate = self._setting_text(key)
            if not password and candidate:
                password = candidate
        return password, found

    def _remove_legacy_passwords(self) -> None:
        for key in self.LEGACY_PASSWORD_KEYS:
            self.settings.remove(key)

    def _sync(self, message: str) -> None:
        self.settings.sync()
        if self.settings.status() != QSettings.Status.NoError:
            raise CredentialError(message)

    @classmethod
    def _protect(cls, password: str) -> str:
        if win32crypt is None:
            raise CredentialError(
                "이 Windows 환경에서는 WMS 비밀번호를 안전하게 보호할 수 없습니다."
            )
        try:
            encrypted = win32crypt.CryptProtectData(
                password.encode("utf-8"),
                cls._DPAPI_DESCRIPTION,
                None,
                None,
                None,
                getattr(win32crypt, "CRYPTPROTECT_UI_FORBIDDEN", 0x1),
            )
            return base64.b64encode(encrypted).decode("ascii")
        except Exception:
            raise CredentialError(
                "WMS 비밀번호를 안전하게 보호하지 못했습니다. Windows 계정 상태를 확인한 뒤 다시 저장해 주세요."
            ) from None

    @classmethod
    def _unprotect(cls, protected_value: str) -> str:
        if win32crypt is None:
            raise CredentialError(
                "이 Windows 환경에서는 저장된 WMS 비밀번호를 불러올 수 없습니다."
            )
        try:
            encrypted = base64.b64decode(protected_value.encode("ascii"), validate=True)
            if not encrypted:
                raise ValueError("empty protected value")
            _description, decrypted = win32crypt.CryptUnprotectData(
                encrypted,
                None,
                None,
                None,
                getattr(win32crypt, "CRYPTPROTECT_UI_FORBIDDEN", 0x1),
            )
            return decrypted.decode("utf-8")
        except Exception:
            raise CredentialError(
                "저장된 WMS 비밀번호를 불러오지 못했습니다. 설정에서 비밀번호를 다시 입력해 주세요."
            ) from None
