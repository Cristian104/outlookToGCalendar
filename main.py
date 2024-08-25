import os
import subprocess
import sys

# Determine the base path for the executable
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)


def run_script(script_name):
    script_path = os.path.join(base_path, script_name)
    try:
        print(f"Running {script_name}...")
        result = subprocess.run(
            [sys.executable, script_path], capture_output=True, text=True)
        print(result.stdout)
        if result.stderr:
            print(f"Error in {script_name}: {result.stderr}")
    except Exception as e:
        print(f"An error occurred while running {script_name}: {e}")


# Run the scripts
run_script("outlookExporter.py")
run_script("toGoogle.py")
run_script("duplicatesRemoval.py")
