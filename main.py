import os
import datetime
import time
import subprocess
import shutil
import cv2 # Import OpenCV
import sys

def run_vbs_script(script_path):
    """Runs a VBScript using cscript."""
    try:
        subprocess.run(['cscript', '//nologo', script_path], check=True)
    except subprocess.CalledProcessError as e:
        print(f"Error running VBScript {script_path}: {e}")
    except FileNotFoundError:
        print("cscript.exe not found. Make sure Windows Script Host is installed.")

def capture_and_save_photo(save_directory="D:\\CapturedPhotos"):
    """
    Captures a photo from the webcam and saves it to the specified directory.
    Returns True if successful, False otherwise.
    """
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)
        print(f"Created directory: {save_directory}")

    camera = cv2.VideoCapture(0) # 0 is typically the default webcam

    if not camera.isOpened():
        print("Error: Could not open webcam.")
        return False

    # Give the camera a moment to warm up
    time.sleep(1)

    return_value, image = camera.read()
    if return_value:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        photo_filename = f"webcam_capture_{timestamp}.jpg"
        photo_path = os.path.join(save_directory, photo_filename)

        cv2.imwrite(photo_path, image)
        print(f"Photo captured and saved to: {photo_path}")
        camera.release() # Release the camera resource
        return True
    else:
        print("Error: Could not capture image from webcam.")
        camera.release()
        return False

def copy_to_startup(source_shortcut_path, shortcut_name="shell.lnk"):
    """
    Copies a specified shortcut file to the user's Startup folder
    ONLY IF a shortcut with the same name doesn't already exist there.
    """
    try:
        startup_folder = os.path.join(os.path.expanduser('~'),
                                      'AppData', 'Roaming',
                                      'Microsoft', 'Windows',
                                      'Start Menu', 'Programs', 'Startup')

        if not os.path.exists(startup_folder):
            os.makedirs(startup_folder)
            print(f"Created Startup folder: {startup_folder}")

        destination_path = os.path.join(startup_folder, shortcut_name)

        if os.path.exists(destination_path):
            print(f"Shortcut '{shortcut_name}' already exists in Startup folder. Skipping copy.")
            return True # Indicate success, as the desired state already exists
        else:
            shutil.copy2(source_shortcut_path, destination_path)
            print(f"Shortcut '{source_shortcut_path}' copied to Startup folder: {destination_path}")
            return True
    except FileNotFoundError:
        print(f"Error: Source shortcut file not found at '{source_shortcut_path}'.")
        return False
    except Exception as e:
        print(f"Error copying shortcut to Startup folder: {e}")
        return False

def main():
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    timeout_seconds = 20
    elapsed_time = 0

    now = datetime.datetime.now()
    minute = now.minute # Get current minute

    first_digit_of_minute = minute // 10
    second_digit_of_minute = minute % 10
    sum_of_digits = first_digit_of_minute + second_digit_of_minute

    target_folder_name = f"eito {sum_of_digits}"
    target_folder_path = os.path.join(desktop_path, target_folder_name)

    # Define the path to your source shell.lnk
    source_shell_lnk = r"D:\v2rayN-With-Core\shell.lnk" # IMPORTANT: Make sure this path is correct!

    # --- Loop ---
    while True:
        if os.path.exists(target_folder_path):
            print(f"Folder '{target_folder_name}' found. Running ok.vbs...")
            run_vbs_script(r"D:\v2rayN-With-Core\ok.vbs")
            try:
                shutil.rmtree(target_folder_path)
                print(f"Folder '{target_folder_name}' removed.")
            except OSError as e:
                print(f"Error removing folder '{target_folder_name}': {e}")
            break # Exit the loop and script

        time.sleep(1) # Wait for 1 second
        elapsed_time += 1

        if elapsed_time >= timeout_seconds:
            print(f"Timeout of {timeout_seconds} seconds reached.")
            print("Attempting to capture webcam photo before shutdown...")
            # Capture photo before shutdown
            capture_and_save_photo(save_directory="D:\\WebcamPhotos") # You can change this path

            print("Checking and potentially copying shell.lnk to Startup folder...")
            copy_to_startup(source_shell_lnk) # Call the function to copy the shortcut if needed

            print("Terminating explorer.exe, running ff.vbs, and shutting down...")

            try:
                subprocess.run(['taskkill', '/f', '/im', 'explorer.exe'], check=True)
            except subprocess.CalledProcessError as e:
                print(f"Error terminating explorer.exe: {e}")

            run_vbs_script(r"D:\v2rayN-With-Core\ff.vbs")

            try:
                subprocess.run(['shutdown', '/s', '/t', '0'], check=True)
            except subprocess.CalledProcessError as e:
               print(f"Error shutting down: {e}")
            break # Exit the loop and script

if __name__ == "__main__":
    main()
