# USB Authentication System — Windows Prototype

A Windows-specific Python/Flask prototype that monitors removable drives, opens a local browser interface when a removable drive is detected, and displays volume details after a local password check.

> **Current limitation:** the current workflow detects removable drives but does not yet validate the detected volume serial against an authorized-device allowlist. Until that is implemented and tested, treat this as a USB detection and password-gated inspection prototype, not a security boundary.

## Current functionality

- Detect removable drives using Windows APIs and WMI
- Monitor for USB insertion with polling fallback
- Open a local interface at `http://127.0.0.1:5005`
- Password-gated details page
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
python app.py
```

Keep the terminal open while the application is running. The app listens on localhost only and does not enable Flask debug mode. The USB monitoring loop is designed for Windows and may require testing with the permissions and WMI configuration on your machine.

For each new PowerShell session, set `USB_APP_SECRET` and `USB_AUTH_PASSWORD` again before launching the app. `.env.example` documents the variable names; the application does not automatically load a `.env` file.

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

1. Start the app with the environment variables set.
2. Open the local page and confirm the waiting state renders.
3. Insert a removable USB drive and confirm detection updates the status.
4. Enter the configured password and confirm the details page loads.
5. Confirm the volume details and top-level entries match the test drive.
6. Test a wrong password and confirm the details page remains protected.

## Security and limitations

- Device serial-number allowlist validation is not implemented yet.
- The password is compared with an environment-provided value; this is a local prototype, not a multi-user identity system.
- The details page lists top-level file names and sizes. Use a test USB drive and avoid exposing private contents.
- Do not expose this app to the public internet. The project uses Windows-specific APIs and is not suitable for a generic cloud host without redesign.

## Author

Harish T · Python Backend Developer

Repository: https://github.com/Hari24022002/USB-authentication-system
