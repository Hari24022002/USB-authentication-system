import unittest
from unittest.mock import patch

import app as usb_app


class USBAppSmokeTests(unittest.TestCase):
    def setUp(self):
        usb_app.app.config.update(TESTING=True)
        usb_app.DRIVE_LETTER = None
        usb_app.DRIVE_SERIAL = None
        usb_app.DEVICE_AUTHORIZED = False
        usb_app.STATUS_MESSAGE = "Waiting for USB insertion"
        usb_app.BROWSER_OPENED = False
        self.original_allowlist = usb_app.ALLOWED_USB_SERIALS
        self.client = usb_app.app.test_client()

    def tearDown(self):
        usb_app.ALLOWED_USB_SERIALS = self.original_allowlist

    def test_waiting_page_renders(self):
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"USB Authentication", response.data)
        self.assertIn(b"Waiting for USB insertion", response.data)

    def test_details_requires_authentication(self):
        response = self.client.get("/details")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/", response.headers["Location"])

    @patch("app.webbrowser.open")
    @patch("app.win32api.GetVolumeInformation", return_value=("TEST", 0x1234, 0, 0, "NTFS"))
    @patch("app.is_removable_drive", return_value=True)
    def test_allowlisted_serial_is_authorized(self, _removable, _volume, browser):
        usb_app.ALLOWED_USB_SERIALS = {"1234"}
        usb_app.process_usb_drive("E:\\\\")
        self.assertTrue(usb_app.DEVICE_AUTHORIZED)
        self.assertIn("Authorized USB device", usb_app.STATUS_MESSAGE)
        browser.assert_called_once()

    @patch("app.webbrowser.open")
    @patch("app.win32api.GetVolumeInformation", return_value=("TEST", 0x1234, 0, 0, "NTFS"))
    @patch("app.is_removable_drive", return_value=True)
    def test_non_allowlisted_serial_is_rejected(self, _removable, _volume, browser):
        usb_app.ALLOWED_USB_SERIALS = {"ffff"}
        usb_app.process_usb_drive("E:\\\\")
        self.assertFalse(usb_app.DEVICE_AUTHORIZED)
        self.assertIn("Unauthorized USB device", usb_app.STATUS_MESSAGE)
        browser.assert_called_once()


if __name__ == "__main__":
    unittest.main()
