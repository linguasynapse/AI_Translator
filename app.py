"""
app.py  -  XLIFF / DOCX / XLSX  AI Translator
Demarrer : python app.py
"""
import os
import sys
import threading
import tempfile
import webbrowser
import time
from pathlib import Path
from flask import Flask, request, jsonify, send_file, abort, after_this_request
from werkzeug.utils import secure_filename
from translator_engine import translate_file, SUPPORTED

# Gestion du chemin de base pour les ressources (dev vs PyInstaller)
def get_base_path():
    """Retourne le chemin de base (dev ou exécutable PyInstaller)"""
    if getattr(sys, 'frozen', False):
        # Mode exécutable : les fichiers sont dans _MEIPASS
        return sys._MEIPASS
    else:
        # Mode développement : dossier du script
        return Path(__file__).parent
    
# Configuration Flask pour servir les fichiers statiques et templates depuis le bon répertoire
base_path = get_base_path()
app = Flask(__name__, 
           template_folder=str(base_path),
           static_folder=str(base_path))

# Variable globale pour suivre l'état de la tâche de traduction
task = {
    'running': False, 'current': 0, 'total': 0,
    'message': '', 'done': False, 'error': None, 
    'output_paths': [], 'stop_requested': False
}
task_lock = threading.Lock()

# Route pour la page d'accueil
@app.route('/')
def index():
    html_path = Path(get_base_path()) / 'index.html'
    if not html_path.exists():
        return f"Erreur: index.html non trouvé dans {html_path}", 404
    return html_path.read_text(encoding='utf-8')


# Route pour le favicon (évite les erreurs 404 dans la console du navigateur)
@app.route('/favicon.ico')
def favicon():
    return '', 204

# Route pour gérer l'upload de fichiers
@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f:
        return jsonify({'ok': False, 'error': 'Pas de fichier'})
    suffix = Path(f.filename).suffix.lower()
    if suffix not in SUPPORTED:
        return jsonify({'ok': False, 'error': 'Format non supporte : ' + suffix})
    tmp_dir = Path(tempfile.mkdtemp()).resolve()
    input_path = str(tmp_dir / secure_filename(f.filename))
    f.save(input_path)
    return jsonify({'ok': True, 'filename': input_path})

# Route pour lancer la traduction
@app.route('/translate', methods=['POST'])
def translate():
    global task
    data = request.json
    with task_lock:
        if task['running']:
            return jsonify({'ok': False, 'error': 'Traduction deja en cours'})
        task = {
            'running': True, 'current': 0, 'total': 0,
            'message': 'Demarrage...', 'done': False, 'error': None, 
            'output_paths': [], 'stop_requested': False
        }

    def run():
        global task
        try:
            inp = data['filename']
            base_out = str(Path(inp).parent / (Path(inp).stem + '_translated'))
            formats = data.get('output_format', ['original'])
            if isinstance(formats, str):
                formats = [formats]

            def cb(current, total, message, done):
                with task_lock:
                    task['current'] = current
                    task['total'] = total
                    task['message'] = message
                    if done:
                        task['done'] = True
                        task['running'] = False
                        task['stop_requested'] = False

            def check_stop():
                with task_lock:
                    return task.get('stop_requested', False)

            generated = translate_file(
                input_path=inp,
                output_path=base_out,
                api_key=data['api_key'],
                source_lang=data['source_lang'],
                target_lang=data['target_lang'],
                model=data.get('model', 'google/gemini-2.5-flash-exp'),
                delay=data.get('delay', 1.0),
                skip_translated=data.get('skip_translated', True),
                style_instructions=data.get('style_instructions', ''),
                output_formats=formats,
                progress_cb=cb,
                stop_check_cb=check_stop,
                use_segmentation=data.get('use_segmentation', False),
                context_size=data.get('context_size', 3)
            )
            with task_lock:
                task['output_paths'] = generated
        except Exception as e:
            with task_lock:
                task['error'] = str(e)
                task['running'] = False
                task['done'] = True
                task['message'] = str(e)

    threading.Thread(target=run, daemon=True).start()
    return jsonify({'ok': True})

# Route pour demander l'arrêt de la traduction en cours
@app.route('/stop', methods=['POST'])
def stop_translation():
    global task
    with task_lock:
        if not task['running']:
            return jsonify({'ok': False, 'error': 'Aucune traduction en cours'})
        task['stop_requested'] = True
        task['message'] = 'Arret demande...'
    return jsonify({'ok': True})

# Route pour obtenir l'état de la traduction
@app.route('/progress')
def progress():
    with task_lock:
        return jsonify(task.copy())

# Route pour télécharger un fichier généré
@app.route('/download')
def download():
    path = request.args.get('f', '').strip()
    if not path:
        abort(400)
    if not os.path.exists(path):
        abort(404)

    full_path = Path(path)

    @after_this_request
    def cleanup(response):
        threading.Timer(2.0, lambda: full_path.unlink(missing_ok=True)).start()
        return response

    response = send_file(
        str(full_path),
        as_attachment=True,
        download_name=full_path.name,
        mimetype='application/octet-stream',
    )
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '0'
    return response


if __name__ == '__main__':
    def open_browser():
        time.sleep(1.2)
        webbrowser.open('http://localhost:5000')

    threading.Thread(target=open_browser, daemon=True).start()
    app.run(debug=False, port=5000, threaded=True, use_reloader=False)