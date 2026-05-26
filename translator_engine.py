"""
translator_engine.py
Moteur LLM pour DOCX, XLSX, XLIFF
Support multi-format, reconstruction de balises, exports propres
"""
import time
import re
import pysbd
import csv
import xml.etree.ElementTree as ET
from pathlib import Path

SUPPORTED = {'.docx', '.xlsx', '.xliff', '.xlf'}
NS_X = 'urn:oasis:names:tc:xliff:document:1.2'
ET.register_namespace('', NS_X)

#── Segmentation multilingue ───────────────────────────────────────────────────────

def segment_text(text, source_lang="", use_segmentation=True):
    if not use_segmentation or not text or not text.strip():
        return [text] if text else []

    lang = source_lang[:2].lower() if source_lang else "en"
    
    try:
        seg = pysbd.Segmenter(language=lang, clean=False)
        sents = seg.segment(text)
        sents = [s.strip() for s in sents if s.strip()]
        if sents:
            return sents
    except Exception:
        pass  # Repli automatique en cas d'erreur pysbd

    # Fallback léger : regex occidentale standard
    pattern = r"(?<=[.!?])\s+(?=[A-ZÀ-ÖØ-öø-ÿ])"
    sents = re.split(pattern, text)
    sents = [s.strip() for s in sents if s.strip()]
    return sents if sents else [text]

def segment_text_regex(text, source_lang="", use_segmentation=True):
    if not use_segmentation or not text or not text.strip():
        return [text] if text else []

    text = text.strip()
    # Texte très court → pas de coupe risquée
    if len(text) < 100 and text.count(".") <= 1 and text.count("。") <= 1:
        return [text]

    lang = source_lang[:2].lower() if source_lang else ""

    # CJK
    if lang in ("zh", "ja", "ko"):
        parts = re.split(r"([。！？])", text)
        out = []
        for i in range(0, len(parts), 2):
            s = parts[i].strip()
            if not s:
                continue
            if i + 1 < len(parts):
                s += parts[i + 1]
            out.append(s)
        return out if len(out) > 1 else [text]

    # Occidental
    # 1. Protéger temporairement les cas pièges
    temp = text
    temp = re.sub(r"(\d)\.(\d)", r"\1DOT\2", temp)          # 3.14 → 3DOT14
    temp = re.sub(r"\b([A-ZÀ-ÖØ-öø-ÿ]{1,3})\.", r"\1ABR", temp) # Dr. → DrABR
    temp = re.sub(r"\.{3,}", "ELLIPSIS", temp)               # ... → ELLIPSIS

    # 2. Segmentation sur ponctuation + majuscule
    pattern = r"(?<=[.!?])\s+(?=[A-ZÀ-ÖØ-öø-ÿ])"
    sents = re.split(pattern, temp)

    # 3. Nettoyage & restauration
    result = []
    for s in sents:
        s = s.strip()
        if not s:
            continue
        s = s.replace("DOT", ".").replace("ABR", ".").replace("ELLIPSIS", "...")
        result.append(s)

    return result if len(result) > 1 else [text]

# ── Exceptions ───────────────────────────────────────────────────────────────
class TranslationStoppedException(Exception):
    """Exception levée quand l'utilisateur demande l'arrêt"""
    pass

# ── Routeur LLM ────────────────────────────────────────────────────────────
def translate_with_llm(text, source_lang, target_lang, context, 
                       style_instructions='', model_str="", api_key=''):
    if not text or not text.strip():
        return text
    if isinstance(text, bytes):
        text = text.decode('utf-8', errors='ignore')
    text = text.replace('\x00', '')

    if '/' not in model_str:
        provider, model = model_str, 'default'
    else:
        provider, model = model_str.split("/", 1)
    
    provider = provider.strip()
    allowed = ("google", "openai", "qwen", "deepseek", "anthropic")
    if provider not in allowed:
        raise ValueError(f"Fournisseur non supporté : {provider}")

    style_block = f"\nSTYLE / TONE:\n{style_instructions.strip()}\n" if style_instructions and style_instructions.strip() else ""
    ctx_block = ""
    if context:
        ctx_block = "Contexte récent:\n" + "\n".join(f"- {c}" for c in context[-3:]) + "\n"

    prompt = (
        f"Professional translator. Translate from {source_lang} to {target_lang}.\n"
        f"Returns ONLY the translation. Preserves the placeholders {{0}}, {{1}} exactly.\n"
        f"{style_block}{ctx_block}Text to be translated:\n{text}"
    )
    if provider == "google":
        from google import genai
        client = genai.Client(api_key=api_key)
        r = client.models.generate_content(model=model, contents=prompt)
        return r.text.strip()
    elif provider in ("openai", "qwen", "deepseek"):
        from openai import OpenAI
        urls = {
            "openai": "https://api.openai.com/v1",
            "qwen": "https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
            "deepseek": "https://api.deepseek.com/v1"
        }
        client = OpenAI(api_key=api_key, base_url=urls[provider])
        r = client.chat.completions.create(
            model=model, messages=[{"role": "user", "content": prompt}],
            temperature=0.2, max_tokens=4000
        )
        return r.choices[0].message.content.strip()
    elif provider == "anthropic":
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)
        r = client.messages.create(
            model=model, max_tokens=4000,
            messages=[{"role": "user", "content": prompt}]
        )
        return r.content[0].text.strip()
    raise ValueError(f"Fournisseur inconnu : {provider}")

# ── Utilitaires XLIFF ──────────────────────────────────────────────────────
def _extract(el):
    tags, parts = [], []
    def walk(e):
        if e.text: parts.append(e.text)
        for ch in e:
            loc = ch.tag.split('}')[-1] if '}' in ch.tag else ch.tag
            if loc in ('g','x','bpt','ept','ph','it','mrk','sub'):
                parts.append(f'{{{len(tags)}}}')
                tags.append(ch)
            else:
                walk(ch)
            if ch.tail: parts.append(ch.tail)
    walk(el)
    return ''.join(parts), tags

def _rebuild(el, tr, tags):
    for child in list(el): el.remove(child)
    el.text = None
    for part in re.split(r'(\{\d+\})', tr):
        m = re.match(r'^\{(\d+)\}$', part)
        if m:
            idx = int(m.group(1))
            if idx < len(tags):
                el.append(tags[idx])
            else:
                if len(el) > 0:
                    list(el)[-1].tail = (list(el)[-1].tail or '') + part
                else:
                    el.text = (el.text or '') + part
        else:
            if len(el) > 0:
                list(el)[-1].tail = (list(el)[-1].tail or '') + part
            else:
                el.text = (el.text or '') + part

# ── Formats de sortie ──────────────────────────────────────────────────────
def generate_bilingual_xlsx(pairs, sl, tl, out_path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bilingue"
    ws.append([f"Source ({sl})", f"Target ({tl})"])
    for s, t in pairs:
        ws.append([s, t])
    ws.column_dimensions['A'].width = 50
    ws.column_dimensions['B'].width = 50
    out = str(Path(out_path).with_suffix('.xlsx'))
    wb.save(out)
    return out

def generate_tmx(pairs, source_lang, target_lang, out_path):
    import xml.dom.minidom as minidom
    from datetime import datetime
    doc = minidom.Document()
    tmx = doc.createElement('tmx')
    tmx.setAttribute('version', '1.4')
    doc.appendChild(tmx)

    header = doc.createElement('header')
    header.setAttribute('creationtool', 'AI Translator')
    header.setAttribute('adminlang', 'EN')
    header.setAttribute('srclang', source_lang[:2].lower())
    header.setAttribute('creationdate', datetime.now().strftime('%Y%m%dT%H%M%S'))
    tmx.appendChild(header)

    body = doc.createElement('body')
    tmx.appendChild(body)
    for src, tgt in pairs:
        tu = doc.createElement('tu')
        for lang, txt in [(source_lang, src), (target_lang, tgt)]:
            tuv = doc.createElement('tuv')
            tuv.setAttribute('xml:lang', lang[:2].lower())
            seg = doc.createElement('seg')
            seg.appendChild(doc.createTextNode(txt))
            tuv.appendChild(seg)
            tu.appendChild(tuv)
        body.appendChild(tu)

    out = str(Path(out_path).with_suffix('.tmx'))
    with open(out, 'w', encoding='utf-8') as f:
        f.write(doc.toprettyxml(indent='  ', encoding='utf-8').decode('utf-8'))
    return out

def generate_tsv(pairs, out_path):
    out = str(Path(out_path).with_suffix('.tsv'))
    with open(out, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter='\t')
        writer.writerow(['Source', 'Target'])
        for s, t in pairs:
            writer.writerow([s, t])
    return out

def generate_xliff(pairs, source_lang, target_lang, out_path, original_filename=""):
    """Génère un fichier XLIFF 1.2 standard à partir de paires source/cible"""
    import xml.etree.ElementTree as ET

    xliff = ET.Element('xliff', version="1.2", xmlns=NS_X)
    file_elem = ET.SubElement(xliff, 'file', {
        'original': Path(original_filename).name if original_filename else "source",
        'source-language': source_lang,
        'target-language': target_lang,
        'datatype': 'plaintext'
    })
    body = ET.SubElement(file_elem, 'body')

    for i, (src, tgt) in enumerate(pairs, 1):
        tu = ET.SubElement(body, 'trans-unit', id=str(i))
        ET.SubElement(tu, 'source').text = src
        tgt_el = ET.SubElement(tu, 'target', {'state': 'translated'})
        tgt_el.text = tgt

    out = str(Path(out_path).with_suffix('.xliff'))
    ET.ElementTree(xliff).write(out, encoding='utf-8', xml_declaration=True)
    return out

# ── DOCX ───────────────────────────────────────────────────────────────────
def translate_docx(inp, out, api_key, sl, tl, model, delay, style_instructions, 
                   output_formats, cb=None, stop_check=None, 
                   use_segmentation=False, context_size=3):
    """
    Traduit un fichier DOCX.
    
    Args:
        use_segmentation: Si True, segmente les paragraphes en phrases
        context_size: Nombre de phrases de contexte (0, 3, ou 5)
    """
    from docx import Document
    doc = Document(inp)
    
    segments = []
    
    # Parser les paragraphes
    for p in doc.paragraphs:
        if p.text.strip():
            text = p.text.strip()
            
            if use_segmentation:
                # Segmenter le paragraphe en phrases
                sentences = segment_text(text, sl, use_segmentation=True)
                for sentence in sentences:
                    if sentence.strip():
                        segments.append(('p', p, sentence, True))  # True = phrase segmentée
            else:
                segments.append(('p', p, text, False))  # False = paragraphe entier
    
    # Parser les tableaux
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if p.text.strip():
                        text = p.text.strip()
                        
                        if use_segmentation:
                            sentences = segment_text(text, sl, use_segmentation=True)
                            for sentence in sentences:
                                if sentence.strip():
                                    segments.append(('cell', p, sentence, True))
                        else:
                            segments.append(('cell', p, text, False))
    
    total = len(segments)
    ctx = []
    pairs = []
    translated_buffer = {}  # Buffer pour reconstruire les paragraphes segmentés
    
    for i, (seg_type, element, original, is_sentence) in enumerate(segments):
        # Vérifier si l'arrêt est demandé
        if stop_check and stop_check():
            if cb: cb(total, total, "Traduction arrêtée par l'utilisateur.", True)
            raise TranslationStoppedException("Arrêt demandé par l'utilisateur")
        
        if cb:
            msg = f"Phrase {i+1}/{total}" if is_sentence else f"Segment {i+1}/{total}"
            cb(i, total, msg, False)
        
        try:
            # Construire le contexte selon le paramètre context_size
            context_for_llm = []
            if context_size > 0:
                context_for_llm = ctx[-context_size:] if len(ctx) >= context_size else ctx[:]
            
            t = translate_with_llm(original, sl, tl, context_for_llm, 
                                  style_instructions, model, api_key)
            
            pairs.append((original, t))
            ctx.append(t)
            
            # Limiter le contexte selon la taille choisie
            if len(ctx) > max(context_size, 3):
                ctx.pop(0)
            
            # Si c'est une phrase segmentée, on stocke la traduction
            if is_sentence:
                if element not in translated_buffer:
                    translated_buffer[element] = []
                translated_buffer[element].append(t)
            else:
                # Paragraphe entier, on applique directement
                if element.runs:
                    element.runs[0].text = t
                    for r in element.runs[1:]:
                        r.text = ""
                else:
                    element.text = t
                    
        except Exception as e:
            if cb: cb(i, total, f"Erreur {i+1}: {e}", False)
            # En cas d'erreur, garder l'original
            if is_sentence:
                if element not in translated_buffer:
                    translated_buffer[element] = []
                translated_buffer[element].append(original)
            pairs.append((original, original))
        
        if i < total - 1:
            time.sleep(delay)
    
    # Reconstruire les paragraphes segmentés
    for element, translated_sentences in translated_buffer.items():
        full_translation = ' '.join(translated_sentences)
        if element.runs:
            element.runs[0].text = full_translation
            for r in element.runs[1:]:
                r.text = ""
        else:
            element.text = full_translation
    
    # Générer les fichiers de sortie
    paths = []
    for fmt in output_formats:
        if fmt == 'original':
            out_path = str(Path(out).with_suffix('.docx'))
            doc.save(out_path)
            paths.append(out_path)
        elif fmt == 'bilingual_xlsx':
            paths.append(generate_bilingual_xlsx(pairs, sl, tl, out + '_bilingual'))
        elif fmt == 'tmx':
            paths.append(generate_tmx(pairs, sl, tl, out + '_bilingual'))
        elif fmt == 'tsv':
            paths.append(generate_tsv(pairs, out + '_bilingual'))
        elif fmt == 'xliff':
            paths.append(generate_xliff(pairs, sl, tl, out + '_bilingual', inp))
    
    if cb: cb(total, total, f"DOCX terminé — {len(paths)} fichier(s) généré(s).", True)
    return paths

# ── XLSX ───────────────────────────────────────────────────────────────────
def translate_xlsx(inp, out, api_key, sl, tl, model, delay, style_instructions, 
                   output_formats, cb=None, stop_check=None,
                   use_segmentation=False, context_size=3):
    """
    Traduit un fichier XLSX.
    
    Args:
        use_segmentation: Si True, segmente les cellules en phrases
        context_size: Nombre de phrases de contexte (0, 3, ou 5)
    """
    import openpyxl
    wb = openpyxl.load_workbook(inp)
    
    cells = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.strip():
                    text = cell.value.strip()
                    
                    if use_segmentation and len(text) > 100:
                        sentences = segment_text(text, sl, use_segmentation=True)
                        for sentence in sentences:
                            if sentence.strip():
                                cells.append((ws, cell, sentence, True))
                    else:
                        cells.append((ws, cell, text, False))
    
    total = len(cells)
    ctx = []
    pairs = []
    cell_buffer = {}  # Buffer pour reconstruire les cellules segmentées
    
    for i, (ws, cell, original, is_sentence) in enumerate(cells):
        # Vérifier si l'arrêt est demandé
        if stop_check and stop_check():
            if cb: cb(total, total, "Traduction arrêtée par l'utilisateur.", True)
            raise TranslationStoppedException("Arrêt demandé par l'utilisateur")
        
        if cb:
            msg = f"Phrase {i+1}/{total}" if is_sentence else f"Cellule {i+1}/{total}"
            cb(i, total, msg, False)
        
        try:
            # Construire le contexte
            context_for_llm = []
            if context_size > 0:
                context_for_llm = ctx[-context_size:] if len(ctx) >= context_size else ctx[:]
            
            t = translate_with_llm(original, sl, tl, context_for_llm, 
                                  style_instructions, model, api_key)
            
            pairs.append((original, t))
            ctx.append(t)
            
            if len(ctx) > max(context_size, 3):
                ctx.pop(0)
            
            if is_sentence:
                if cell not in cell_buffer:
                    cell_buffer[cell] = []
                cell_buffer[cell].append(t)
            else:
                cell.value = t
                
        except Exception as e:
            if cb: cb(i, total, f"Erreur {i+1}: {e}", False)
            if is_sentence:
                if cell not in cell_buffer:
                    cell_buffer[cell] = []
                cell_buffer[cell].append(original)
            pairs.append((original, original))
        
        if i < total - 1:
            time.sleep(delay)
    
    # Reconstruire les cellules segmentées
    for cell, translated_sentences in cell_buffer.items():
        cell.value = ' '.join(translated_sentences)
    
    # Générer les fichiers de sortie
    paths = []
    for fmt in output_formats:
        if fmt == 'original':
            out_path = str(Path(out).with_suffix('.xlsx'))
            wb.save(out_path)
            paths.append(out_path)
        elif fmt == 'bilingual_xlsx':
            paths.append(generate_bilingual_xlsx(pairs, sl, tl, out + '_bilingual'))
        elif fmt == 'tmx':
            paths.append(generate_tmx(pairs, sl, tl, out + '_bilingual'))
        elif fmt == 'tsv':
            paths.append(generate_tsv(pairs, out + '_bilingual'))
        elif fmt == 'xliff':
            paths.append(generate_xliff(pairs, sl, tl, out + '_bilingual', inp))
    
    if cb: cb(total, total, f"XLSX terminé — {len(paths)} fichier(s) généré(s).", True)
    return paths

# ── XLIFF ──────────────────────────────────────────────────────────────────
def translate_xliff_file(inp, out, api_key, sl, tl, model, delay, skip_translated, style_instructions, output_formats, cb=None, stop_check=None):
    tree = ET.parse(inp)
    root = tree.getroot()
    units = root.findall(f'.//{{{NS_X}}}trans-unit')
    
    todo = []
    for u in units:
        src = u.find(f'{{{NS_X}}}source')
        if src is None: continue
        txt, tags = _extract(src)
        if not txt.strip(): continue
        if skip_translated:
            tgt = u.find(f'{{{NS_X}}}target')
            if tgt is not None and tgt.get('state', '') in ('translated', 'signed-off', 'final'):
                continue
        todo.append((u, src, txt, tags))

    total = len(todo)
    if total == 0:
        if cb: cb(0, 0, "Aucun segment à traduire.", True)
        return []

    ctx = []
    pairs = []
    for i, (u, src, txt, itags) in enumerate(todo):
        if stop_check and stop_check():
            if cb: cb(total, total, "Traduction arrêtée par l'utilisateur.", True)
            raise TranslationStoppedException("Arrêt demandé par l'utilisateur")
        if cb: cb(i, total, f"Segment {i+1}/{total}", False)
        try:
            tr = translate_with_llm(txt, sl, tl, ctx, style_instructions, model, api_key)
            if not tr.strip(): tr = txt

            tgt = u.find(f'{{{NS_X}}}target')
            if tgt is None: tgt = ET.SubElement(u, f'{{{NS_X}}}target')
            tgt.set('state', 'translated')
            _rebuild(tgt, tr, itags)
            pairs.append((txt, tr))
            ctx.append(tr)
            if len(ctx) > 3: ctx.pop(0)
        except Exception as e:
            if cb: cb(i, total, f"Erreur segment {i+1}: {e}", False)
        if i < total - 1: time.sleep(delay)

    paths = []
    for fmt in output_formats:
        if fmt in ('original', 'xliff'):
            out_path = str(Path(out).with_suffix('.xliff' if Path(inp).suffix != '.xlf' else '.xlf'))
            tree.write(out_path, encoding='utf-8', xml_declaration=True)
            paths.append(out_path)
        elif fmt == 'bilingual_xlsx': paths.append(generate_bilingual_xlsx(pairs, sl, tl, out + '_bilingual'))
        elif fmt == 'tmx': paths.append(generate_tmx(pairs, sl, tl, out + '_bilingual'))
        elif fmt == 'tsv': paths.append(generate_tsv(pairs, out + '_bilingual'))

    if cb: cb(total, total, f"XLIFF terminé — {len(paths)} fichier(s) généré(s).", True)
    return paths

# ── Dispatcher ─────────────────────────────────────────────────────────────
def translate_file(input_path, output_path, api_key, source_lang, target_lang,
                   model, delay, skip_translated, style_instructions,
                   output_formats=None, progress_cb=None, stop_check_cb=None,
                   use_segmentation=False, context_size=3):
    """
    Dispatcher principal.
    
    Nouveaux paramètres:
        use_segmentation: Active la segmentation par phrase
        context_size: Taille du contexte (0, 3, ou 5 phrases)
    """
    if not output_formats:
        output_formats = ['original']
    
    ext = Path(input_path).suffix.lower()
    if ext not in SUPPORTED:
        raise ValueError(f"Format non supporté : {ext}")
    
    try:
        if ext == '.docx':
            return translate_docx(
                input_path, output_path, api_key, source_lang, target_lang,
                model, delay, style_instructions, output_formats, 
                progress_cb, stop_check_cb, use_segmentation, context_size
            )
        elif ext == '.xlsx':
            return translate_xlsx(
                input_path, output_path, api_key, source_lang, target_lang,
                model, delay, style_instructions, output_formats,
                progress_cb, stop_check_cb, use_segmentation, context_size
            )
        else:
            return translate_xliff_file(
                input_path, output_path, api_key, source_lang, target_lang,
                model, delay, skip_translated, style_instructions, output_formats,
                progress_cb, stop_check_cb
            )
    except TranslationStoppedException:
        return []