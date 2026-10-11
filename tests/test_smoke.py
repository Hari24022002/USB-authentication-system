import unittest

import app as usb_app


class USBAppSmokeTests(unittest.TestCase):
    def setUp(self):
        usb_app.app.config.update(TESTING=True)
        self.client = usb_app.app.test_client()

    def test_waiting_page_renders(self):
        usb_app.DRIVE_LETTER = None
        usb_app.DEVICE_AUTHORIZED = False
        response = self.client.get("/")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b"USB Authentication", response.data)
        self.assertIn(b"Waiting for USB insertion", response.data)

    def test_details_requires_authentication(self):
        response = self.client.get("/details")
        self.assertEqual(response.status_code, 302)
        self.assertIn("/", response.headers["Location"])


if __name__ == "__main__":
    unittest.main()
