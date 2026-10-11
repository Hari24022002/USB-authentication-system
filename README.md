# USB Authentication System — Windows Prototype

A Windows-specific Python/Flask prototype that monitors removable drives, compares the Windows volume serial against a configured allowlist, and displays drive details only after device authorization and a local password check.

> **Security note:** Windows volume serials are not cryptographic device identities and can be copied or spoofed. Treat this as an educational access-control prototype, not high-assurance authentication.

## Current functionality

- Detect removable drives using Windows APIs and WMI
- Monitor for USB insertion with polling fallback
- Compare the detected volume serial with `USB_ALLOWED_SERIALS`
- Open a local interface at `http://127.0.0.1:5005`
- Require a configured password before showing drive details
- Display volume name, volume serial, file system, capacity, free space, and top-level entries
- Write logs to the project-local `logs/app.log`

## Requirements

- Windows 10/11
- Python 3.10 or newer
- A removable USB drive for hardware testing

## Setup (PowerShell)

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
$env:USB_APP_SECRET = [guid]::NewGuid().ToString() + [guid]::NewGuid().ToString()
$env:USB_AUTH_PASSWORD = "choose-a-local-password"
# First run: leave this empty to identify the drive serial.
$env:USB_ALLOWED_SERIALS = ""
python app.py
```

Connect the USB drive and read its volume serial from the status page or `logs/app.log`. Stop the app, then configure the serial and restart:

```powershell
$env:USB_ALLOWED_SERIALS = "0x1234ABCD"
python app.py
```

Replace `0x1234ABCD` with the actual serial shown for your test USB. Multiple allowed serials can be comma-separated. Serial values are normalized to lowercase and the optional `0x` prefix is ignored. Keep the terminal open while the app runs; set the environment variables again in each new PowerShell session.

The server listens on localhost only and debug mode is disabled. `.env.example` documents the settings, but the app does not automatically load a `.env` file.

## Files

```text
app.py
templates/
  index.html
  details.html
requirements.txt
.env.example
README.md
```

## Testing checklist

1. Start with `USB_ALLOWED_SERIALS` empty and confirm the status page reports the detected serial.
2. Configure that serial and restart.
3. Confirm the allowed drive reveals the password form.
4. Enter the configured password and confirm drive details load.
5. Connect a different removable drive and confirm it is marked unauthorized and cannot access details.
6. Try a wrong password and confirm the details page remains protected.
7. Restart the app and confirm the monitoring loop still detects insertions.

## Security and limitations

- Volume serial allowlisting is a useful demo gate but not a cryptographic device identity.
- The password is compared with an environment-provided value; this is a local prototype, not a multi-user identity system.
- The details page lists top-level file names and sizes. Use a test USB drive and avoid exposing private contents.
- Do not expose this app to the public internet. The project uses Windows-specific APIs and is not suitable for a generic cloud host without redesign.

## Author

Harish T · Python Backend Developer

Repository: https://github.com/Hari24022002/USB-authentication-system
