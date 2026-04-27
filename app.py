
from flask import Flask, request, jsonify, send_from_directory
import subprocess
import os

app = Flask(__name__, static_folder='.', static_url_path='')

@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

@app.route('/run-script', methods=['POST'])
def run_script():
    # Recebe os dados do frontend (se necessário)
    data = request.json or {}
    # Executa o script Python existente
    script_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'gerador_script_nokia.py'))
    result = subprocess.run(['python', script_path], capture_output=True, text=True)
    return jsonify({
        'stdout': result.stdout,
        'stderr': result.stderr,
        'returncode': result.returncode
    })

if __name__ == '__main__':
    app.run(debug=True)
