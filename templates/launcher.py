import webbrowser
import threading
import sys
from app import app
import os

def open_browser():
    webbrowser.open('http://127.0.0.1:5000')

if __name__ == '__main__':
    # Open browser after 1 second
    threading.Timer(1, open_browser).start()
    
    # Run the app
    app.run(host='127.0.0.1', port=5000, debug=False)