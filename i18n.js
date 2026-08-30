/**
 * i18n.js — Moteur i18n universel pour applis Flask locales
 */

const I18N = (() => {
    const SUPPORTED = ['fr', 'en', 'de', 'es', 'it', 'pt', 'pt-BR', 'nl', 'pl', 'ru', 'tr', 'zhs', 'zht', 'ja', 'ko'];
    const STORAGE_KEY = 'i18n_lang';
    const FALLBACK = 'en';

    let _locale = {};
    let _lang = FALLBACK;
    let _onChangeCb = null;

    // ── Chargement du fichier JSON ────────────────────────────────────────
    async function load(lang) {
        if (!SUPPORTED.includes(lang)) {
            console.warn(`[i18n] "${lang}" not supported, fallback -> ${FALLBACK}`);
            lang = FALLBACK;
        }
        try {
            const res = await fetch(`/locales/${lang}.json?v=${Date.now()}`);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            _locale = await res.json();
            _lang = lang;
            localStorage.setItem(STORAGE_KEY, lang);
            document.documentElement.lang = lang;
            _apply();
            _updateSwitcher();
            if (_onChangeCb) _onChangeCb(lang);
        } catch (err) {
            console.error(`[i18n] Unable to load ${lang}.json :`, err);
        }
    }

    // ── Traduction avec interpolation ─────────────────────────────────────
    function t(key, vars) {
        let str = _locale[key];
        if (str === undefined) {
            console.warn(`[i18n] Missing key : "${key}"`);
            return key;  // Affiche la cle brute comme signal visible
        }
        if (vars) {
            Object.entries(vars).forEach(([k, v]) => {
                str = str.replace(new RegExp(`\\{${k}\\}`, 'g'), v);
            });
        }
        return str;
    }

    // ── Application au DOM ────────────────────────────────────────────────
    function _apply() {
        document.querySelectorAll('[data-i18n]').forEach(el => {
            el.textContent = t(el.getAttribute('data-i18n'));
        });
        document.querySelectorAll('[data-i18n-placeholder]').forEach(el => {
            el.placeholder = t(el.getAttribute('data-i18n-placeholder'));
        });
        document.querySelectorAll('[data-i18n-title]').forEach(el => {
            el.title = t(el.getAttribute('data-i18n-title'));
        });
        document.querySelectorAll('[data-i18n-html]').forEach(el => {
            el.innerHTML = t(el.getAttribute('data-i18n-html'));
        });
    }

    // ── Mise a jour visuelle du switcher ──────────────────────────────────
    function _updateSwitcher() {

        // --- Boutons data-i18n-lang ---
        document.querySelectorAll('[data-i18n-lang]').forEach(btn => {
            btn.classList.toggle('lang-active', btn.getAttribute('data-i18n-lang') === _lang);
        });

        // --- Selects data-i18n-select-lang ---
        document.querySelectorAll('[data-i18n-select-lang]').forEach(select => {
            // Met à jour la valeur sélectionnée
            select.value = _lang;

            // Optionnel : classe active sur les <option>
            select.querySelectorAll('option').forEach(opt => {
                opt.classList.toggle('lang-active', opt.value === _lang);
            });
        });
    }

    // ── Init au chargement ────────────────────────────────────────────────
    function init(defaultLang) {
        const saved = localStorage.getItem(STORAGE_KEY);
        const detected = saved || defaultLang || FALLBACK;
        load(detected);

        // Branchement automatique des boutons data-i18n-lang
        document.addEventListener('click', e => {
            const btn = e.target.closest('[data-i18n-lang]');
            if (btn) load(btn.getAttribute('data-i18n-lang'));
        });
        document.addEventListener('change', e => {
            const select = e.target.closest('[data-i18n-select-lang]');
            if (select) load(select.value);
        });
    }

    return {
        init,
        load,
        t,
        getLang: () => _lang,
        onLangChange: (cb) => { _onChangeCb = cb; },
        supported: SUPPORTED,
    };
})();

// Raccourcis globaux utilisables directement dans index.html
const t = (key, vars) => I18N.t(key, vars);
const i18nInit = (lang) => I18N.init(lang);
const i18nLoad = (lang) => I18N.load(lang);

I18N.onLangChange(function (lang) {
    // Rafraîchir les éléments dynamiques
    checkOllamaStatus();
    // Re-appliquer les traductions sur les éléments statiques
    if (typeof I18N !== 'undefined' && I18N._apply) {
        I18N._apply();
    }
});
