from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

from PySide6.QtCore import QSettings

from Modules.Common.Credentials import (
    CredentialError,
    WMSCredentials,
    WMSCredentialStore,
)


class WMSCredentialStoreTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.settings_path = Path(self.temp_dir.name) / "credentials.ini"
        self.settings = self._new_settings()

    def tearDown(self) -> None:
        self.settings.sync()
        del self.settings
        self.temp_dir.cleanup()

    def _new_settings(self) -> QSettings:
        return QSettings(str(self.settings_path), QSettings.Format.IniFormat)

    def test_roundtrip_uses_real_ini_qsettings_and_windows_dpapi(self) -> None:
        expected = WMSCredentials(wms_id="wms-user", password="S3cret!/한글")
        WMSCredentialStore(self.settings).save(expected)

        reloaded = self._new_settings()
        actual = WMSCredentialStore(reloaded).load()

        self.assertEqual(actual, expected)
        self.assertTrue(actual.is_complete)

    def test_plaintext_password_keys_and_value_are_absent_after_save(self) -> None:
        password = "plain-password-must-not-remain"
        WMSCredentialStore(self.settings).save(
            WMSCredentials(wms_id="safe-id", password=password)
        )

        self.assertEqual(self.settings.value("wms_id"), "safe-id")
        self.assertTrue(self.settings.contains("wms_password_dpapi"))
        self.assertFalse(self.settings.contains("wms_password"))
        self.assertFalse(self.settings.contains("wms_pw"))
        self.settings.sync()
        self.assertNotIn(password, self.settings_path.read_text(encoding="utf-8"))

    def test_blank_password_clears_protected_and_legacy_values_but_keeps_id(self) -> None:
        store = WMSCredentialStore(self.settings)
        store.save(WMSCredentials(wms_id="saved-user", password="initial-secret"))
        self.settings.setValue("wms_password", "obsolete-secret")
        self.settings.setValue("wms_pw", "obsolete-secret-2")

        store.save(WMSCredentials(wms_id="saved-user", password=""))

        self.assertFalse(self.settings.contains("wms_password_dpapi"))
        self.assertFalse(self.settings.contains("wms_password"))
        self.assertFalse(self.settings.contains("wms_pw"))
        self.assertEqual(store.load(), WMSCredentials(wms_id="saved-user", password=""))

    def test_legacy_plaintext_password_is_migrated_to_dpapi(self) -> None:
        self.settings.setValue("wms_id", "legacy-user")
        self.settings.setValue("wms_password", "legacy/password")
        self.settings.setValue("wms_pw", "legacy/password")
        self.settings.sync()

        credentials = WMSCredentialStore(self.settings).load()

        self.assertEqual(
            credentials,
            WMSCredentials(wms_id="legacy-user", password="legacy/password"),
        )
        self.assertTrue(self.settings.contains("wms_password_dpapi"))
        self.assertFalse(self.settings.contains("wms_password"))
        self.assertFalse(self.settings.contains("wms_pw"))
        self.settings.sync()
        self.assertNotIn("legacy/password", self.settings_path.read_text(encoding="utf-8"))

        reloaded = self._new_settings()
        self.assertEqual(WMSCredentialStore(reloaded).load(), credentials)

    def test_corrupt_protected_data_raises_friendly_error_without_credentials(self) -> None:
        self.settings.setValue("wms_id", "private-user-id")
        self.settings.setValue("wms_password_dpapi", "not-valid-base64!!!")
        self.settings.sync()

        with self.assertRaises(CredentialError) as raised:
            WMSCredentialStore(self.settings).load()

        message = str(raised.exception)
        self.assertIn("비밀번호", message)
        self.assertNotIn("private-user-id", message)
        self.assertNotIn("not-valid-base64", message)

    def test_dpapi_decrypt_failure_does_not_expose_credentials(self) -> None:
        credentials = WMSCredentials(
            wms_id="private-user-id",
            password="private-password",
        )
        store = WMSCredentialStore(self.settings)
        store.save(credentials)

        with mock.patch(
            "Modules.Common.Credentials.win32crypt.CryptUnprotectData",
            side_effect=RuntimeError("private-user-id private-password"),
        ):
            with self.assertRaises(CredentialError) as raised:
                store.load()

        message = str(raised.exception)
        self.assertNotIn(credentials.wms_id, message)
        self.assertNotIn(credentials.password, message)

    def test_dpapi_encrypt_failure_does_not_expose_or_partially_save_credentials(self) -> None:
        credentials = WMSCredentials(
            wms_id="private-user-id",
            password="private-password",
        )
        with mock.patch(
            "Modules.Common.Credentials.win32crypt.CryptProtectData",
            side_effect=RuntimeError("private-user-id private-password"),
        ):
            with self.assertRaises(CredentialError) as raised:
                WMSCredentialStore(self.settings).save(credentials)

        message = str(raised.exception)
        self.assertNotIn(credentials.wms_id, message)
        self.assertNotIn(credentials.password, message)
        self.assertFalse(self.settings.contains("wms_id"))
        self.assertFalse(self.settings.contains("wms_password_dpapi"))


if __name__ == "__main__":
    unittest.main()
