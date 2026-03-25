from flask import Flask, render_template, request, redirect, url_for, session
import win32com.client
import win32api
import os
import time
import threading
import logging
import pythoncom
import wmi
import webbrowser
import secrets

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)

# Configure logging
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', filename='C:\\DSkey\\app.log')
logger = logging.getLogger(__name__)
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(console_handler)

# Configuration
PASSWORD = "Admin@1234"  # Hardcoded password
STATUS_MESSAGE = "Waiting for USB insertion"  # Global status message
DRIVE_LETTER = None  # Detected USB drive letter
PROCESSED_DRIVES = set()  # Track processed drives
BROWSER_OPENED = False  # Track if browser has been opened for the current USB

def is_removable_drive(drive_letter):
    """Check if the drive is removable."""
    try:
        drive_type = win32api.GetDriveType(drive_letter)
        return drive_type == win32api.DRIVE_REMOVABLE
    except Exception as e:
        logger.error(f"Error checking drive type for {drive_letter}: {e}")
        return False

def check_existing_drives():
    """Check for already inserted USB drives."""
    try:
        drives = win32api.GetLogicalDriveStrings().split('\0')[:-1]
        for drive in drives:
            if drive not in PROCESSED_DRIVES and is_removable_drive(drive):
                logger.debug(f"Checking existing drive: {drive}")
                PROCESSED_DRIVES.add(drive)
                return drive
        return None
    except Exception as e:
        logger.error(f"Error checking existing drives: {e}")
        return None

def poll_usb_drives():
    """Poll for USB drives as a fallback if WMI fails."""
    try:
        c = wmi.WMI()
        for disk in c.Win32_LogicalDisk(DriveType=2):
            drive_letter = disk.DeviceID + "\\"
            if drive_letter not in PROCESSED_DRIVES:
                logger.debug(f"Polling detected drive: {drive_letter}")
                PROCESSED_DRIVES.add(drive_letter)
                return drive_letter
        return None
    except Exception as e:
        logger.error(f"Error polling USB drives: {e}")
        return None

def process_usb_drive(drive_letter):
    """Process a detected USB drive and open the application."""
    global STATUS_MESSAGE, DRIVE_LETTER, BROWSER_OPENED
    if is_removable_drive(drive_letter):
        DRIVE_LETTER = drive_letter
        STATUS_MESSAGE = f"USB detected: {drive_letter}. Opening application."
        logger.info(STATUS_MESSAGE)
        if not BROWSER_OPENED:
            try:
                webbrowser.open('http://localhost:5005/')
                BROWSER_OPENED = True
                logger.info("Web application opened in browser")
            except Exception as e:
                STATUS_MESSAGE = f"Error opening browser: {e}"
                logger.error(STATUS_MESSAGE)
    else:
        STATUS_MESSAGE = "Detected drive is not removable."
        logger.warning(STATUS_MESSAGE)

def monitor_usb():
    """Monitor USB device insertion using Windows WMI with fallback polling."""
    global STATUS_MESSAGE, BROWSER_OPENED
    pythoncom.CoInitialize()
    
    # Check for existing USB drives at startup
    existing_drive = check_existing_drives()
    if existing_drive:
        process_usb_drive(existing_drive)

    # Monitor for new USB insertions
    try:
        wmi_obj = win32com.client.GetObject("winmgmts:")
        watcher = wmi_obj.ExecNotificationQuery(
            "SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_LogicalDisk'"
        )
        logger.info("Monitoring for USB insertion via WMI...")
        while True:
            try:
                event = watcher.NextEvent(10000)
                drive = event.TargetInstance
                drive_letter = drive.DeviceID + "\\"
                if drive.DriveType == 2 and drive_letter not in PROCESSED_DRIVES:
                    PROCESSED_DRIVES.add(drive_letter)
                    BROWSER_OPENED = False  # Reset for new USB insertion
                    process_usb_drive(drive_letter)
                else:
                    logger.debug(f"Ignoring non-removable or processed drive: {drive_letter}")
            except Exception as e:
                if "Timed out" in str(e):
                    logger.debug("WMI timeout, falling back to polling")
                    drive_letter = poll_usb_drives()
                    if drive_letter:
                        BROWSER_OPENED = False  # Reset for new USB insertion
                        process_usb_drive(drive_letter)
                else:
                    logger.error(f"Error in USB monitoring loop: {e}")
                time.sleep(1)
    except Exception as e:
        STATUS_MESSAGE = f"Failed to initialize USB monitoring: {e}"
        logger.error(STATUS_MESSAGE)
        logger.info("Falling back to USB polling...")
        while True:
            drive_letter = poll_usb_drives()
            if drive_letter:
                BROWSER_OPENED = False  # Reset for new USB insertion
                process_usb_drive(drive_letter)
            time.sleep(5)
    finally:
        pythoncom.CoUninitialize()

@app.route('/', methods=['GET', 'POST'])
def index():
    """Render the login page if USB is detected, else waiting message."""
    global STATUS_MESSAGE, DRIVE_LETTER
    error = None
    if request.method == 'POST':
        entered_pass = request.form.get('password')
        if entered_pass == PASSWORD:
            session['authenticated'] = True
            return redirect(url_for('details'))
        else:
            error = "Incorrect password"
    if DRIVE_LETTER:
        return render_template('index.html', status=STATUS_MESSAGE, show_form=True, error=error)
    else:
        return render_template('index.html', status=STATUS_MESSAGE, show_form=False, error=None)

@app.route('/details')
def details():
    """Display USB drive details and data if authenticated."""
    if not session.get('authenticated'):
        return redirect(url_for('index'))
    if not DRIVE_LETTER:
        return "No USB drive detected.", 400
    
    try:
        # Get volume information
        volume_name, serial_num, max_comp_len, flags, fs_name = win32api.GetVolumeInformation(DRIVE_LETTER)
        
        # Get disk space
        free_user, total_size, free_total = win32api.GetDiskFreeSpaceEx(DRIVE_LETTER)
        
        # List top-level files and directories
        items = []
        for item in os.listdir(DRIVE_LETTER):
            full_path = os.path.join(DRIVE_LETTER, item)
            if os.path.isdir(full_path):
                items.append(f"Directory: {item}")
            else:
                size = os.path.getsize(full_path)
                items.append(f"File: {item} (Size: {size:,} bytes)")
        
        details = {
            'drive_letter': DRIVE_LETTER,
            'volume_name': volume_name,
            'serial_number': hex(serial_num),
            'file_system': fs_name,
            'total_size': f"{total_size:,} bytes",
            'free_space': f"{free_user:,} bytes",
            'items': items
        }
        
        return render_template('details.html', details=details)
    except Exception as e:
        logger.error(f"Error retrieving drive details: {e}")
        return f"Error retrieving drive details: {e}", 500

if __name__ == '__main__':
    # Start USB monitoring in a separate thread
    usb_thread = threading.Thread(target=monitor_usb, daemon=True)
    usb_thread.start()
    
    # Start Flask server
    app.run(host='0.0.0.0', port=5005, debug=True)
    
