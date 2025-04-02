import os
import webbrowser
import subprocess
import time

# List of Streamlit apps to run
apps = [
    "finale2.py",
    "size_wise_rej.py",
    "new_proj.py"  # Add more if needed
]

# Open a new browser tab for each app
for index, app in enumerate(apps):
    port = 8501 + index  # Assign different ports to each app
    command = f"start /b streamlit run {app} --server.port {port}"  # Run in background
    subprocess.Popen(command, shell=True)
    time.sleep(2)  # Small delay to allow app startup
    webbrowser.open(f"http://localhost:{port}")  # Open in browser
