"""
i18n_routes.py — Routes Flask pour le moteur i18n
"""

from flask import Response, abort
import logging

logger = logging.getLogger("aitranslator.i18nroutes")


def register_i18n(app, resource_path_fn):
    """
    Enregistre les routes i18n sur l'app Flask.

    Args:
        app             : instance Flask
        resource_path_fn: fonction resource_path(filename) -> Path
                          deja presente dans app.py pour PyInstaller
    """
    # Récupérer les langues supportées dans le dossier locales pour rédure le risque de désynchronisation entre i18n.js et Flask.
    locales_dir = resource_path_fn("locales")

    if locales_dir.exists() and locales_dir.is_dir():
        SUPPORTED_LANGS = {p.stem for p in locales_dir.glob("*.json")}
    else:
        SUPPORTED_LANGS = set()

    @app.route("/i18n.js")
    def serve_i18n_js():
        """Utiliser le moteur i18n.js."""
        path = resource_path_fn("i18n.js")
        if not path.exists():
            abort(404)
        return Response(
            path.read_text(encoding="utf-8"),
            mimetype="application/javascript; charset=utf-8",
        )

    @app.route("/locales/<lang>.json")
    def serve_locale(lang):
        """Utiliser le fichier de traduction pour la langue demandee."""
        if lang not in SUPPORTED_LANGS:
            abort(404)
        locale_path = resource_path_fn(f"locales/{lang}.json")
        if not locale_path.exists():
            abort(404)
        return Response(
            locale_path.read_text(encoding="utf-8"),
            mimetype="application/json; charset=utf-8",
            headers={
                "Cache-Control": "no-store",  # Pas de cache. Toujours la version fraiche
            },
        )