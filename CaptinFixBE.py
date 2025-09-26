from flask import Flask, render_template, request, jsonify
import subprocess
import json
import os
import sys
from datetime import datetime

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('CaptainFixFE.html')

@app.route('/start_test', methods=['POST'])
def start_test():
    try:
        data = request.get_json()
        if not data:
            return "No JSON payload received", 400

        url = data.get('url', '').strip()
        if not url:
            return jsonify({"error": "url is required"}), 400

        ts = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
        config_filename = f"test_config_{ts}.json"
        # Put configs in a subfolder for cleanliness
        configs_dir = os.path.join(os.getcwd(), "test_configs")
        os.makedirs(configs_dir, exist_ok=True)
        config_path = os.path.join(configs_dir, config_filename)

        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        # Resolve absolute path to CaptainFixQA.py
        script_path = os.path.join(os.getcwd(), "CaptainFixQA.py")
        if not os.path.exists(script_path):
            return jsonify({"error": f"Script not found: {script_path}"}), 500

        # Use same Python interpreter running Flask
        python_exe = sys.executable  # ensures same interpreter
        logs_dir = os.path.join(os.getcwd(), "logs")
        os.makedirs(logs_dir, exist_ok=True)
        stdout_log = os.path.join(logs_dir, f"captain_{ts}.log")

        # Launch process and redirect stdout/stderr to log file so we can debug
        log_file = open(stdout_log, "a", encoding="utf-8")
        # Pass current env to child (so OPENAI_API_KEY etc. are available)
        env = os.environ.copy()

        proc = subprocess.Popen(
            [python_exe, script_path, config_path],
            stdout=log_file,
            stderr=log_file,
            cwd=os.getcwd(),
            env=env,
            close_fds=True
        )

        return jsonify({
            "message": "Test started",
            "pid": proc.pid,
            "config": config_filename,
            "log": os.path.relpath(stdout_log, os.getcwd())
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)
