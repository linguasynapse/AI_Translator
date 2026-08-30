"""
app.py  -  XLIFF / DOCX / XLSX  AI Translator
Demarrer : python app.py
"""
import logging
from logging.handlers import RotatingFileHandler
import os
import sys
import shutil
import atexit
import socket
import secrets
import random
import json
import threading
import tempfile
import webbrowser
import time
from pathlib import Path
from flask import Flask, request, jsonify, send_file, abort, after_this_request, Response
from werkzeug.utils import secure_filename
from werkzeug.exceptions import HTTPException
from i18n_routes import register_i18n
from translator_engine import translate_file, SUPPORTED

# ============================================================
# LOGGING
# ============================================================
def get_log_dir() -> Path:
    """
    Retourne un dossier de logs accessible en écriture.
    Important pour PyInstaller : ne pas logger dans sys._MEIPASS.
    """
    if getattr(sys, "frozen", False): 
        # Mode PyInstaller
        if sys.platform == "win32":
            base = Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()))
        elif sys.platform == "darwin":
            base = Path.home() / "Library" / "Logs"
        else:
            base = Path(os.environ.get("XDG_STATE_HOME", Path.home() / ".local" / "state"))

        return base / "AITranslator"

    # Mode développement
    return Path(__file__).resolve().parent / "logs"

def setup_logging() -> logging.Logger:
    log_dir = get_log_dir()
    log_dir.mkdir(parents=True, exist_ok=True)

    log_file = log_dir / "app.log"

    logger = logging.getLogger("aitranslator")
    logger.setLevel(
        logging.DEBUG if os.environ.get("AI_TRANSLATOR_DEBUG") else logging.INFO
    )

    # Évite d'ajouter plusieurs handlers si setup_logging() est appelé plusieurs fois
    if not logger.handlers:
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
        )

        file_handler = RotatingFileHandler(
            log_file,
            maxBytes=2 * 1024 * 1024,  # 2 MB
            backupCount=3,
            encoding="utf-8",
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

        # Console seulement en développement
        if not getattr(sys, "frozen", False) and sys.stderr is not None:
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(formatter)
            logger.addHandler(console_handler)

    return logger

logger = setup_logging()

# Pour réduire le bruit de Werkzeug
logging.getLogger("werkzeug").setLevel(logging.WARNING) 
logging.getLogger("werkzeug").setLevel(logging.ERROR)

# ============================================
# CHECK OLLAMA
# ============================================
def check_ollama():
    """Vérifie si Ollama local est disponible"""
    try:
        import requests
        response = requests.get("http://localhost:11434/api/tags", timeout=2)
        if response.status_code == 200:
            models = response.json().get('models', [])
            # logger.info(f"Local Ollama available with {len(models)} models")
            return True
    except:
        logger.warning("Local Ollama is not available")
    return False

OLLAMA_AVAILABLE = check_ollama()

# ============================================
# FONCTIONS DE BASE
# ============================================
# Fichier de sortie
INPUT_DIR = Path(__file__).resolve().parent / "uploads"
OUTPUT_DIR = Path(__file__).resolve().parent / "outputs"
INPUT_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
def get_output_path(inp):
    """Génère le chemin de sortie pour un fichier source"""
    inp_path = Path(inp)
    # Utiliser le nom du fichier sans extension
    base_name = inp_path.stem
    return OUTPUT_DIR / f"{base_name}_translated"

def clear_input_folder():
    """Supprime tous les fichiers restants dans le dossier uploads."""
    if INPUT_DIR.exists():
        for file_path in INPUT_DIR.iterdir():
            if file_path.is_file():
                try:
                    file_path.unlink()
                    logger.info(f"Cleanup: Deleted all remaining files in the uploads folder.")
                except Exception as e:
                    logger.warning(f"Could not delete remaining files in the uploads folder.: {e}")

# Nettoyer les reliquats d'une ancienne session au lancement de l'application
clear_input_folder() # Comment when debugging to keep inputs

# Enregistrer la fonction pour qu'elle s'exécute automatiquement à la fermeture de Python
atexit.register(clear_input_folder)

# Fonction pour obtenir un port libre
def get_free_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        return s.getsockname()[1]
    
# Gestion du chemin de base pour les ressources (dev vs PyInstaller)
def get_base_path():
    """Retourne le chemin de base (dev ou exécutable PyInstaller)"""
    if getattr(sys, 'frozen', False):
        # Mode exécutable : les fichiers sont dans _MEIPASS
        return sys._MEIPASS
    else:
        # Mode développement : dossier du script
        return Path(__file__).parent

def resource_path(filename):
    """Retourne le chemin absolu d'un fichier, en dev ou en mode PyInstaller."""
    base = get_base_path()
    return Path(base) / filename

# Gestion du port et du token de session
def init_session():
    """Génère un token de session unique pour sécuriser les requêtes"""
    # Créer une nouvelle session
    PORT = get_free_port()
    # Génération d'un token de session unique pour sécuriser les requêtes 
    SESSION_TOKEN = secrets.token_hex(32)
    return PORT, SESSION_TOKEN

# ============================================
# INITIALISATION
# ============================================

# Initialisation du port et du token de session
PORT, SESSION_TOKEN = init_session()

# Configuration Flask pour servir les fichiers statiques et templates depuis le bon répertoire
base_path = get_base_path()
#app = Flask(__name__)
app = Flask(__name__, 
           template_folder=str(base_path),
           static_folder=str(base_path))

# Limite de taille de fichier
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# Enregistrement des routes i18n pour la traduction de l'interface
register_i18n(app, resource_path) # Ici if faut passer la fonction resource_path pour PyInstaller

# Variable globale pour suivre l'état de la tâche de traduction
task = {
    'running': False, 'current': 0, 'total': 0,
    'message': '', 'done': False, 'error': None, 
    'output_paths': [], 'stop_requested': False
}
task_lock = threading.Lock()

# ============================================
# ROUTES FLASK
# ============================================

# Sécurité : vérification du token de session pour toutes les requêtes sauf la page principale
@app.before_request
def check_token():
    # Routes exemptees du token (ressources statiques publiques)
    exempt = ('/', '/favicon.ico', '/i18n.js')
    if request.path in exempt:
        return
    # Exempter aussi les fichiers de locales
    if request.path.startswith('/locales/'):
        return
    # Vérifier le token à l'en-tête HTTP et au paramètre URL. 
    # Requêtes fetch() : envoient le token via l'en-tête X-Session-Token; 
    # téléchargement direct (clic sur lien) : envoie le token via le paramètre ?token=
    token = request.headers.get('X-Session-Token') or request.args.get('token')
    if token != SESSION_TOKEN:
        logger.exception("Session token mismatch. Session aborted")
        abort(403)

# Sécurité : ajout des en-têtes HTTP pour renforcer la sécurité
@app.after_request
def security_headers(response):
    # Interdit aux autres origines d'appeler l'API
    response.headers['X-Content-Type-Options'] = 'nosniff'
    # Bloque le chargement dans une iframe
    response.headers['X-Frame-Options'] = 'DENY'
    # Pas de fuite du referer
    response.headers['Referrer-Policy'] = 'no-referrer'
    # CSP strict — aucun script externe autorisé
    response.headers['Content-Security-Policy'] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-inline'; "
        "style-src 'self' 'unsafe-inline'; "
        "font-src 'self' data:; "
        "connect-src 'self'; "
        "img-src 'self' data:;"
    )
    return response

# Route pour la page d'accueil
@app.route('/')
def index():
    html_path = Path(base_path) / 'index.html'
    if not html_path.exists():
        logger.exception("index.html path not found")
        return f"Error: index.html not found", 404
    html = html_path.read_text(encoding='utf-8')
    html = html.replace('__SESSION_TOKEN__', SESSION_TOKEN) # Injecter le token dans le JS
    html = html.replace('__PORT__', str(PORT)) # Injecter le port dans le HTML
    return html

# Route pour le favicon (évite les erreurs 404 dans la console du navigateur)
@app.route('/favicon.ico')
def favicon():
    return '', 204

@app.errorhandler(Exception)
def handle_exception(e):
    if isinstance(e, HTTPException):
        return e
    logger.exception("Unhandled Flask exception")
    return jsonify({"ok": False, "error": "Internal server error"}), 500

@app.route('/ollama/models')
def get_ollama_models():
    """Récupère la liste des modèles Ollama installés"""
    try:
        import requests
        response = requests.get("http://localhost:11434/api/tags", timeout=2)
        if response.status_code == 200:
            data = response.json()
            models = [model['name'] for model in data.get('models', [])]
            return jsonify({'ok': True, 'models': models})
        else:
            return jsonify({'ok': False, 'error': 'Ollama API error'}), 500
    except requests.exceptions.ConnectionError:
        return jsonify({'ok': False, 'error': 'Ollama not running'}), 503
    except Exception as e:
        return jsonify({'ok': False, 'error': str(e)}), 500
    
# Route pour gérer upload de fichiers
@app.route('/upload', methods=['POST'])
def upload():
    f = request.files.get('file')
    if not f:
        return jsonify({'ok': False, 'error': 'No file to upload'})
    suffix = Path(f.filename).suffix.lower()
    if suffix not in SUPPORTED:
        return jsonify({'ok': False, 'error': 'Unsupported format: ' + suffix})
    # Utiliser le dossier INPUT_DIR au lieu de tempfile
    input_path = str(INPUT_DIR / secure_filename(f.filename))
    f.save(input_path)
    logger.info(f"File uploaded")
    return jsonify({'ok': True, 'filename': input_path})

# Route pour lancer la traduction
@app.route('/translate', methods=['POST'])
def translate():
    global task
    data = request.get_json(silent=True)
    if not data:
        logger.warning("Translate request rejected: invalid or missing JSON")
        return jsonify({'ok': False, 'error': 'Invalid JSON payload'}), 400
    # Vérification du modèle et de la clé API
    model = data.get('model', '')
    is_ollama = model.startswith("ollama_local/")
    requires_api_key = not is_ollama # clé optionnelle pour ollama
    if requires_api_key and not data.get('api_key'):
        # logger.warning("Translate request rejected: missing API key for model %s", model)
        return jsonify({'ok': False, 'error': 'Missing API key for model ' + model}), 400

    required = ['filename', 'source_lang', 'target_lang']
    if requires_api_key:
        required.append('api_key')
        
    missing = [k for k in required if not data.get(k)]
    if missing:
        #logger.warning("Translate request rejected, missing: %s", missing)
        #logger.warning("Translate request rejected: missing file name, API key, source language or target language.")
        return jsonify({'ok': False, 'error': 'Missing fields: ' + ', '.join(missing)}), 400
    try:
        delay = max(0.0, float(data.get('delay', 1.0) or 0))
    except (TypeError, ValueError):
        delay = 1.0
    try:
        context_size = int(data.get('context_size', 3))
    except (TypeError, ValueError):
        context_size = 3
    if context_size not in (0, 3, 5):
        context_size = 3

    formats = data.get('output_format', ['original'])
    if isinstance(formats, str):
        formats = [formats]
    
    for field in required:
        if field not in data or not data[field]:
            # logger.warning("Missing field: %s")
            return jsonify({'ok': False, 'error': f'Missing field: {field}'}), 400
    with task_lock:
        if task['running']:
            # logger.warning("Translation request rejected: already running")
            return jsonify({'ok': False, 'error': 'Translation already in progress'})
        task = {
            'running': True, 'current': 0, 'total': 0,
            'message': 'Starting translation...', 'done': False, 'error': None, 
            'output_paths': [], 'stop_requested': False
        }

    def run():
        global task
        try:
            inp = data['filename']
            # Utiliser INPUT_DIR pour les fichiers d'entrée
            base_out = str(get_output_path(inp))
            logger.info("Translation started")            
            
            formats = data.get('output_format', ['original'])
            if isinstance(formats, str):
                formats = [formats]

            #logger.debug("Output formats: %s", formats)
            def cb(current, total, message, done):
                with task_lock:
                    task['current'] = current
                    task['total'] = total
                    task['message'] = message
                    if done:
                        task['running'] = False
                        task['stop_requested'] = False
                if done:
                    logger.info("Translation finished: %s", message)
                else:
                    logger.debug("Progress %s/%s: %s", current, total, message)    

            def check_stop():
                with task_lock:
                    return task.get('stop_requested', False)

            try:
                delay = max(0.0, float(data.get('delay', 1.0) or 0))
            except (TypeError, ValueError):
                delay = 1.0

            if delay < 0:
                delay = 0

            try:
                context_size = int(data.get('context_size', 3))
            except (TypeError, ValueError):
                context_size = 3

            if context_size not in (0, 3, 5):
                context_size = 3

            generated = translate_file(
                input_path=inp,
                output_path=base_out,
                api_key=data['api_key'] if not is_ollama else None,
                source_lang=data['source_lang'],
                target_lang=data['target_lang'],
                model=data.get('model', 'google/gemini-2.5-flash-exp'),
                delay=delay,
                skip_translated=data.get('skip_translated', True),
                style_instructions=data.get('style_instructions', ''),
                output_formats=formats,
                progress_cb=cb,
                stop_check_cb=check_stop,
                use_segmentation=data.get('use_segmentation', False),
                context_size=context_size
            )
            with task_lock:
                task['output_paths'] = generated
                task['done'] = True
                task['running'] = False
            #logger.info("Translation generated %s file(s)", len(generated))

        except Exception as e:
            logger.exception("Translation failed")
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
            logger.error("No translation in progress.")
            return jsonify({'ok': False, 'error': 'No translation in progress'})
        task['stop_requested'] = True
        task['message'] = 'Stop requested...'
    return jsonify({'ok': True})

# Route pour obtenir l'état de la traduction
@app.route('/progress')
def progress():
    with task_lock:
        return jsonify(task.copy())

# Route pour télécharger un fichier généré
ALLOWED_DOWNLOAD_ROOTS = [
    OUTPUT_DIR,
    Path(tempfile.gettempdir()).resolve()
]

@app.route('/download')
def download():
    path = request.args.get('f', '').strip()
    if not path:
        logger.warning("Download request with empty path")
        abort(400)

    full_path = Path(path).resolve()
    logger.info("Download requested")

    # Vérifier que le chemin est autorisé
    allowed = any(
        str(full_path).startswith(str(root))
        for root in ALLOWED_DOWNLOAD_ROOTS
    )

    if not allowed:
        logger.warning("Download denied: path not in allowed roots")
        abort(403)

    if not full_path.exists() or not full_path.is_file():
        logger.warning("Download failed: path not found")
        abort(404)
    # Le token est déjà vérifié par @app.before_request
    # Il accepte soit X-Session-Token (header) soit ?token= (paramètre URL)
    try:
        # Lire le fichier en mémoire
        with open(full_path, 'rb') as f:
            file_data = f.read()
        
        # Créer la réponse avec les données en mémoire
        response = Response(
            file_data,
            mimetype='application/octet-stream',
            headers={
                'Content-Disposition': f'attachment; filename="{full_path.name}"',
                'Cache-Control': 'no-store, no-cache, must-revalidate, max-age=0',
                'Pragma': 'no-cache',
                'Expires': '0',
                'Content-Length': str(len(file_data))
            }
        )
        return response
        
    except Exception as e:
        logger.error(f"Download error: {e}")
        abort(500)

# ============================================
# POINT D'ENTRÉE
# ============================================

if __name__ == '__main__':    
    # Ouvrir le navigateur
    def open_browser():
        time.sleep(1.2)
        webbrowser.open(f'http://127.0.0.1:{PORT}')

    threading.Thread(target=open_browser, daemon=True).start()
    # Démarrer le serveur (BLOQUANT - rien après ne s'exécute)
    app.run(host='127.0.0.1', debug=False, port=PORT, threaded=True, use_reloader=False)