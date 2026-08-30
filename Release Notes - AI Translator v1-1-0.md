# Release Notes - AI Translator v1.1.0

**Release Date:** August 2026

---

## Overview

We are excited to announce the release of AI Translator v1.1.0, a significant update that brings enhanced security, expanded provider support, improved XLIFF handling, and a fully localized user interface. This release focuses on making the application more robust, flexible, and accessible to a wider audience.

---

## New Features

### 🌐 Multilingual User Interface
The application now supports **English**, **French**, and **Simplified Chinese** for the user interface. Language can be switched seamlessly via the language selector in the header, with persistent preference storage.

### 🏠 Local Model Support with Ollama
We have added full support for **Ollama**, enabling users to run translation models **locally** without requiring an API key. This provides:
- Complete privacy (no data leaves your machine)
- Zero recurring costs
- Support for popular models: Llama 3, Mistral, Gemma, DeepSeek-R1, and more

**Cloud vs Local distinction** has been clearly separated in the interface for intuitive model selection.

### 🤖 New AI Providers
Expanded provider support to offer more choice and flexibility:
- **Mistral AI** (Mistral Large, Medium, Small)
- **Ollama** (local models)
- Updated model lists for all providers with the latest available versions

### 🔒 Enhanced Security
A **session-based token system** now protects all API endpoints:
- Each application session generates a unique 256-bit token
- All API requests (except static assets) require token validation
- Tokens are automatically injected into the frontend
- Download links include token parameters for secure file retrieval

### 📁 New Output Formats
Expanded export options to support various localization workflows:
- **TMX** (Translation Memory eXchange) - Industry standard for translation memory
- **TSV** (Tab-Separated Values) - Simple tabular format for spreadsheet import

---

## Improvements

### XLIFF Reconstruction
The XLIFF parsing and rebuilding engine has been significantly strengthened:
- **Fixed placeholder extraction** with better handling of nested tags
- **Robust tag reconstruction** that preserves inline formatting
- **Validation** of translated content to prevent tag corruption
- Proper handling of `g`, `x`, `bpt`, `ept`, `ph`, `it`, `mrk`, and `sub` elements

### DOCX Processing
Enhanced document parsing with support for:
- **Content Controls** (`<w:sdt>` structured document tags) - frequently found in forms, templates, and highlighted text blocks
- **Text boxes** and shapes
- **Tables** with proper cell handling
- **Segmented paragraph translation** for improved accuracy on long texts

### Translation Engine
- **Context-aware translation** with configurable context window (0, 3, or 5 sentences)
- **Sentence segmentation** using `pysbd` for improved accuracy on long paragraphs
- **Retry logic** for LLM failures with validation checks
- **Bilingual source cleaning** to handle documents with embedded translations

### Logging & Monitoring
Improved log management with:
- **Rotating file handler** (configurable size and retention)
- **Automatic log cleanup** for files older than 30 days
- **Compression** of logs older than 7 days (saves disk space)
- Separate log levels for different application components

### User Experience
- **Fullscreen mode** for distraction-free operation
- **Progress polling** with real-time status updates
- **Cancel translation** capability (stop running jobs)
- **Download links** with timestamp to prevent caching
- **Better error handling** with descriptive user messages

### Security Headers
All responses now include enhanced security headers:
- `X-Content-Type-Options: nosniff`
- `X-Frame-Options: DENY`
- `Referrer-Policy: no-referrer`
- `Content-Security-Policy` with strict default settings

---

## Bug Fixes

### XLIFF
- Fixed `NameError: name 'tree' is not defined` in DOCX translation
- Fixed `NameError: name 'idx' is not defined` in `_rebuild` function
- Resolved placeholder extraction for nested tags

### API Key Handling
- API key now correctly optional for local models (Ollama)
- Proper toggle between "required" and "not required" states in UI
- Fixed `api_key` validation for Ollama local models

### Download
- Files now correctly served from `outputs/` directory
- Fixed `403 Forbidden` errors on download (token now included in URL)
- Files are read into memory before deletion for reliable transfer
- Extended cleanup timers to prevent premature file deletion

### Session Management
- Fixed `Session aborted` errors by ensuring token validation across all routes
- Token now properly injected into all API requests
- Download links include token parameter

### DOCX Processing
- Fixed extraction of highlighted text in inside `<w:sdt>` tags
- Proper handling of text in tables and shapes

### Translation Engine
- Fixed `Unsupported format` error for `.xlf` files (now correctly handled)
- Improved `_clean_translation` function to handle source-target concatenation
- Better validation of translation quality before applying changes

---

## Breaking Changes

### API Key Requirement
**Ollama local models no longer require an API key.** If you were using a dummy key for local models, it can now be left empty.

### Model Naming
For consistency, model names have been standardized:
- Ollama local: `ollama_local/{model_name}`
- Ollama cloud: `ollama_cloud/{model_name}`
- Other providers: `{provider}/{model_name}`

**Action required:** Update your selection if you had manually entered model names.

---

## Known Limitations

- **DOCX**: Text in embedded objects (charts, SmartArt) may not be extracted
- **XLIFF**: Some complex nested structures may still require manual review
- **Ollama**: Must be installed and running separately (not bundled with the application)
- **Large files**: Processing time scales with document size and API rate limits

---

## Upgrade Instructions

1. **Backup your files** (especially `outputs/` and `logs/` directories)
2. **Update dependencies**:
   ```bash
   pip install -r requirements.txt --upgrade
   ```
3. **For Ollama users**: Ensure Ollama is installed and running:
   ```bash
   ollama serve
   ollama pull <your-preferred-model>
   ```
4. **Restart the application**

---

## Contributing

We welcome contributions! Please visit our GitHub repository at [github.com/linguasynapse](https://github.com/linguasynapse) to report issues or submit pull requests.

---

## Acknowledgments

Special thanks to all users who provided feedback and reported issues. Your input has been invaluable in shaping this release.

---

**© 2026 Lingua Synapse. All rights reserved.**