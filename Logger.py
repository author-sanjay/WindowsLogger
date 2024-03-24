import cv2
import time
import os
import win32gui
import win32console
import win32con
import pyscreenshot
import wmi

win = win32console.GetConsoleWindow()
win32gui.ShowWindow(win, win32con.SW_HIDE)


def check_camera_lid_status():
    try:
        c = wmi.WMI()
        for camera in c.Win32_PnPEntity():
            if 'camera' in str(camera.caption).lower():
                # Check if the camera status is "OK" (indicating it's functional)
                if camera.status == "OK":
                    return "Camera is ready and lid is open"
                else:
                    return "Camera lid is closed or there's an issue with the camera"
        return "Camera device not found"
    except Exception as e:
        return f"Error occurred while checking camera status: {str(e)}"

# Function to get the active window title
def get_active_window():
    if os.name == 'nt':
        window = win32gui.GetForegroundWindow()
        title = win32gui.GetWindowText(window)
        return title
    else:
        return "Unsupported OS"

# Function to write to log file
def write_to_log(log_file, text):
    with open(log_file, 'a') as file:
        file.write(text + '\n')

# Function to capture image from webcam
def capture_image(image_dir):
    try:
        if(check_camera_lid_status()):
            cap = cv2.VideoCapture(0, cv2.CAP_DSHOW)
            ret, frame = cap.read()
            if ret:
                image_file = os.path.join(image_dir, f"webcam_image_{int(time.time())}.jpg")
                cv2.imwrite(image_file, frame)
                cap.release()
                return image_file
            else:
                return None
        else:
            print("Camera lid close")
    except Exception as e:
        print("Error capturing webcam image:", e)
        return None

# Function to capture screenshot

def contains_sensitive_words(text):
    forbidden_words = ["adult", "murder", "violence", "disturbing", "sex", "porn"]
    text_lower = text.lower()  # Convert text to lowercase
    for word in forbidden_words:
        word_lower = word.lower()  # Convert forbidden word to lowercase
        if word_lower in text_lower:
            return True
    return False


def capture_screenshot(screenshot_dir):
    try:
        screenshot_path = os.path.join(screenshot_dir, f"screenshot_{time.strftime('%Y-%m-%d_%H-%M-%S')}.png")
        pyscreenshot.grab().save(screenshot_path)
        return screenshot_path
    except Exception as e:
        print("Error capturing screenshot:", e)
        return None


def close_application(window_title):
    try:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd != 0:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            return True
        return False
    except Exception as e:
        print("Error closing application:", e)
        return False


if __name__ == "__main__":
    # Specify the directory where logs, webcam images, and screenshots should be stored
    log_dir = "logs"
    image_dir = "images"
    screenshot_dir = "screenshots"

    # Create directories if they don't exist
    os.makedirs(log_dir, exist_ok=True)
    os.makedirs(image_dir, exist_ok=True)
    os.makedirs(screenshot_dir, exist_ok=True)

    log_file = os.path.join(log_dir, "activity_log.txt")  # Name of the log file
    active_apps = {}  # Dictionary to store opening times of applications

    screenshot_interval = 600  # Interval for capturing screenshots (in seconds)
    last_screenshot_time = time.time()

    # Infinite loop to continuously monitor activity
    try:    
        while True:
            try:
                active_window = get_active_window()
                if active_window:
                    try:
                        app_name, sub_title = active_window.split(' - ', 1)
                        print("Active Application:", app_name)
                        print("Sub Title:", sub_title)
                        if contains_sensitive_words(app_name):
                            screenshot_file = capture_screenshot(screenshot_dir)
                            if screenshot_file:
                                write_to_log(log_file, f"Acessed Restricted site Screenshot captured at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} - {screenshot_file}")
                            print("Detected sensitive words in the application name:", app_name)
                            if close_application(active_window):
                                print("Closed the application window:", active_window)
                            else:
                                print("Failed to close the application window:", active_window)
                        elif app_name not in active_apps:
                            active_apps[app_name] = time.time()  # Record opening time
                            # Capture images when you open an application
                            webcam_image_file = capture_image(image_dir)
                            if webcam_image_file:
                                write_to_log(log_file, f"{sub_title} Application Opened {app_name} opened at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} Webcam Image captured: {webcam_image_file}")
                        else:
                            active_apps[app_name] = time.time()
                    except ValueError:
                        # If splitting active_window throws an error, log only the application name
                        app_name = active_window
                        print("Active Application:", app_name)
                        if app_name not in active_apps:
                            active_apps[app_name] = time.time()  # Record opening time
                            # Capture images when you open an application
                            webcam_image_file = capture_image(image_dir)
                            if webcam_image_file:
                                write_to_log(log_file, f"Application Opened {app_name} opened at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} Webcam Image captured: {webcam_image_file}")
                        else:
                            active_apps[app_name] = time.time()  # Update active time

                else:
                    for app, start_time in list(active_apps.items()):
                        write_to_log(log_file, f"Application Closed {app} closed at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())}")
                        del active_apps[app]  # Remove the closed application from active list
    
            # Capture screenshot every 10 minutes
                current_time = time.time()
                if current_time - last_screenshot_time >= screenshot_interval:
                    screenshot_file = capture_screenshot(screenshot_dir)
                    if screenshot_file:
                        write_to_log(log_file, f"Screenshot captured at {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())} - {screenshot_file}")
                    last_screenshot_time = current_time

                time.sleep(1)  # Adjust the interval as needed
            except KeyboardInterrupt:
                print("Logging stopped by user.")
                break
    except:
        print("Failed to capture screenshot")