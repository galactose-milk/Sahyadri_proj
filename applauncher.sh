#!/bin/bash

# List of Streamlit apps
APPS=("finale_2.py" "new_proj.py" "size_wise_rej.py")  # Add more apps if needed
PORT=8501  # Starting port

# Open each app in a new process
for APP in "${APPS[@]}"; do
    echo "Starting $APP on port $PORT..."
    streamlit run "$APP" --server.port $PORT >/dev/null 2>&1 &
    sleep 2  # Give it some time to start
    xdg-open "http://localhost:$PORT"  # Open in browser
    ((PORT++))
done

echo "All apps started!"

#UNDER PROCESSSS