#!/usr/bin/env python3
import sys
print("Python version:", sys.version)
print("Starting test server...")

from flask import Flask
app = Flask(__name__)

@app.route('/health')
def health():
    return {"status": "ok"}

if __name__ == '__main__':
    print("Running Flask app...")
    app.run(host='0.0.0.0', port=5000, debug=True)
