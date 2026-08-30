# 🌐 AI Translator - Outil de Traduction Assistée par IA

**Version :** 1.1.0  
**Licence :** MIT  
**Public cible :** Traducteurs professionnels, chefs de projet traduction, linguistes

---

## 📖 Présentation

AI Translator est une application de bureau qui utilise l'intelligence artificielle pour traduire vos documents tout en préservant leur mise en forme. Cette application :

- ✅ **UI en français, en anglais et en chinois**
- ✅ **Préserve la structure complète** de vos fichiers DOCX, XLSX et XLIFF
- ✅ **Maintient les balises XLIFF** (g, x, bpt, ept, ph, mrk, sub) intactes
- ✅ **Respecte votre terminologie** via des instructions de style personnalisées
- ✅ **Ignore les segments déjà traduits** dans vos fichiers XLIFF
- ✅ **Exporte en formats professionnels** : TMX (mémoires de traduction), XLSX bilingue, TSV, XLIFF 1.2

---

## 🚀 Fonctionnalités

### Formats supportés

| Format | Traduction | Préservation balises | Export bilingue |
| -------- | ----------- | --------------------- | ----------------- |
| **DOCX** | ✅ Texte, tableaux | ✅ Styles, polices | ✅ |
| **XLSX** | ✅ Toutes cellules | ✅ Formules, formats | ✅ |
| **XLIFF/XLF** | ✅ Segments | ✅ Balises internes | ✅ |

### Moteurs IA disponibles

- **Google Gemini** 3.5 Flash / 3 Flash Preview / 2.5 Flash
- **OpenAI** GPT-5-mini / GPT-5
- **Anthropic Claude** Sonnet 4.6 / Haiku 4.5
- **DeepSeek** V4 Pro / V4 Flash
- **Mistral** Dernière version
- **Ollama** Cloud et local

### Options professionnelles

- 🎨 **Styles de traduction** : Formel, Juridique, Marketing, Technique, Littéraire
- 📝 **Instructions personnalisées** : Règles jusqu'à 300 mots pour guider l'IA
- ⏱️ **Délai configurable** entre les appels API (respect des limites de rate)
- 🔄 **Skip traduit** : Ignore les segments déjà traduits (state="translated", "signed-off", "final")
- 🛑 **Arrêt d'urgence** : Interrompre la traduction à tout moment
- 🖥️ **Mode plein écran** : Interface épurée façon application native

### Formats d'export

- **Original** : Fichier traduit dans son format d'origine
- **Bilingue XLSX** : Tableau source/cible pour révision
- **TMX 1.4** : Mémoire de traduction réutilisable
- **TSV** : Format tabulé universel
- **XLIFF 1.2** : Standard d'échange pour outils de TAO

---

## 📦 Installation

### Pour les utilisateurs (uniquement pour Windows v1.0.0)

1. Téléchargez l'exécutable `AI_Translator.exe`
2. Double-cliquez pour lancer
3. L'application s'ouvre automatiquement dans votre navigateur

### Pour les développeurs

```bash
# Cloner le dépôt
git clone https://github.com/linguasynapse/AI_Translator.git
cd AI_Translator

# Optionnel: Créer un environnement virtuel (première installation) et activer l'environnement
python -m venv .aitranslator
.aitranslator\Scripts\activate

# Installer les dépendances
pip install -r requirements.txt

# Lancer l'application
python app.py
```
