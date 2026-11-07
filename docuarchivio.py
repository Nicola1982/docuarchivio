#!/usr/bin/env python3
"""
DocuArchivio - passo 1 (Versione Ottimizzata per PC Datati / Celeron N2840)
Archivio documenti locale (stile Papra): ricerca nel testo, tag, ambiti Lavoro/Casa,
cartella osservata che importa da sola i nuovi file.

Installazione (una volta sola):   pip install pypdf python-docx
Avvio:                            python docuarchivio.py
Si apre il browser su http://127.0.0.1:8765  (chiudi la finestra nera per fermarlo)

Dati salvati nella cartella DocuArchivio_dati accanto a questo file.
Importazione da Papra: nella barra a sinistra incolli la chiave API (Papra > Impostazioni utente,
permessi organizations:read, documents:read, tags:read). La chiave non viene salvata su disco.

Domande al documento (IA locale, tutto sul tuo PC, nessuna installazione):
  apri Impostazioni (⚙) e scegli un modello: Ultra Leggero (390 MB), Leggero (1 GB), Bilanciato (2,5 GB) o Potente (4,7 GB).
  Il programma scarica da solo il motore llama.cpp e il modello. Puoi anche mettere un tuo .gguf
  nella cartella DocuArchivio_dati/modello. Windows 64 bit.

OCR (scansioni e immagini): installa Tesseract una sola volta con
  winget install UB-Mannheim.TesseractOCR
poi in Impostazioni > OCR premi "Installa componenti PDF" e "Scarica lingua italiana".
"""
import os, re, json, time, shutil, hashlib, sqlite3, threading, webbrowser, mimetypes
import math, sys, subprocess, importlib.util, atexit, zipfile, html, socket
import urllib.request, urllib.error
from collections import Counter
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote
from pathlib import Path

BASE = Path(__file__).resolve().parent / "DocuArchivio_dati"
FILES = BASE / "files"
DBPATH = BASE / "archivio.db"
PORT = 8765
LOCK = threading.Lock()
SCAN_LOCK = threading.Lock()
TESTO = {".txt", ".md", ".csv", ".log", ".json"}
MODELLO = BASE / "modello"
HF = "https://huggingface.co/bartowski/"
MODELLI = [
    {"id": "ultraleggero", "nome": "Ultra Leggero (0.5B)", "mb": 390, "file": "Qwen2.5-0.5B-Instruct-Q4_K_M.gguf",
     "url": HF + "Qwen2.5-0.5B-Instruct-GGUF/resolve/main/Qwen2.5-0.5B-Instruct-Q4_K_M.gguf",
     "desc": "Qwen2.5 0.5B. Pesa pochissimo (390 MB), velocissimo su CPU vecchie (es. Celeron/Atom senza AVX2)."},
    {"id": "leggero", "nome": "Leggero (1.5B)", "mb": 1000, "file": "Qwen2.5-1.5B-Instruct-Q4_K_M.gguf",
     "url": HF + "Qwen2.5-1.5B-Instruct-GGUF/resolve/main/Qwen2.5-1.5B-Instruct-Q4_K_M.gguf",
     "desc": "Qwen2.5 1,5B. Buon bilanciamento per PC standard. Risposte coerenti e veloci."},
    {"id": "bilanciato", "nome": "Bilanciato (4B)", "mb": 2500, "file": "Qwen_Qwen3-4B-Q4_K_M.gguf",
     "url": HF + "Qwen_Qwen3-4B-GGUF/resolve/main/Qwen_Qwen3-4B-Q4_K_M.gguf",
     "desc": "Qwen3 4B. Risposte molto migliori; servono almeno 6 GB di RAM libera e su CPU vecchie è lento."},
    {"id": "potente", "nome": "Potente (7B)", "mb": 4700, "file": "Qwen2.5-7B-Instruct-Q4_K_M.gguf",
     "url": HF + "Qwen2.5-7B-Instruct-GGUF/resolve/main/Qwen2.5-7B-Instruct-Q4_K_M.gguf",
     "desc": "Qwen2.5 7B. La qualità migliore; servono almeno 10 GB di RAM libera e molta pazienza senza GPU."},
]
OCR_DIR = BASE / "ocr"
TESSDATA = OCR_DIR / "tessdata"
LINGUA_URL = "https://github.com/tesseract-ocr/tessdata_fast/raw/main/ita.traineddata"
IMG_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
OCR_DPI = 200
CHARS = 4000  # Ridotto a 4000 caratteri per velocizzare l'elaborazione su CPU lente
SISTEMA = ("Sei un assistente che risponde in italiano a domande su un documento. Usa SOLO le "
           "informazioni dell'estratto del documento fornito. Se la risposta non e' nell'estratto, "
           "dillo chiaramente senza inventare. Rispondi in modo chiaro e conciso.")
STOP = set("che come dove quando quale quali cosa della delle dello degli nel nella per con sono del dei "
           "una uno questo questa documento alla alle dal dalla nei sul sulla chi mi ti si".split())


def db():
    c = sqlite3.connect(DBPATH, timeout=30)
    c.row_factory = sqlite3.Row
    return c


def pulisci_testo(t):
    """Normalizza il testo per ridurre la memoria e migliorare FTS."""
    if not t:
        return ""
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t]+", " ", t)).strip()


def init():
    FILES.mkdir(parents=True, exist_ok=True)
    with db() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS docs(id INTEGER PRIMARY KEY, name TEXT, ext TEXT, size INTEGER,
            hash TEXT UNIQUE, added TEXT, stored TEXT, ambito TEXT DEFAULT '', tags TEXT DEFAULT '', ocr TEXT DEFAULT '');
        CREATE TABLE IF NOT EXISTS seen(path TEXT PRIMARY KEY, mtime REAL, size INTEGER);
        CREATE TABLE IF NOT EXISTS settings(k TEXT PRIMARY KEY, v TEXT);
        CREATE VIRTUAL TABLE IF NOT EXISTS fts USING fts5(name, body, tags,
            tokenize='unicode61 remove_diacritics 2');
        """)
        try:
            c.execute("ALTER TABLE docs ADD COLUMN ocr TEXT DEFAULT ''")
        except sqlite3.OperationalError:
            pass


def getset(k):
    with db() as c:
        r = c.execute("SELECT v FROM settings WHERE k=?", (k,)).fetchone()
    return r["v"] if r else ""


def setset(k, v):
    with LOCK, db() as c:
        c.execute("INSERT OR REPLACE INTO settings VALUES(?,?)", (k, v))


def estrai(path, ext):
    try:
        if ext == ".pdf":
            from pypdf import PdfReader
            return pulisci_testo("\n".join((p.extract_text() or "") for p in PdfReader(str(path)).pages))
        if ext == ".docx":
            import docx
            d = docx.Document(str(path))
            parti = [p.text for p in d.paragraphs]
            for t in d.tables:
                for r in t.rows:
                    parti.append(" | ".join(c.text for c in r.cells))
            return pulisci_testo("\n".join(parti))
        if ext in TESTO:
            return pulisci_testo(Path(path).read_text(errors="ignore"))
    except ImportError as e:
        print("Libreria mancante (pip install pypdf python-docx):", e)
    except Exception as e:
        print("Errore lettura testo", path, e)
    return ""


def aggiungi(src, nome, ambito="", extra=""):
    """Copia il file nell'archivio. Restituisce l'id, o None se esiste gia'."""
    h = hashlib.sha256()
    with open(src, "rb") as f:
        for blocco in iter(lambda: f.read(1 << 20), b""):
            h.update(blocco)
    h = h.hexdigest()
    with db() as c:
        if c.execute("SELECT id FROM docs WHERE hash=?", (h,)).fetchone():
            return None
    ext = Path(nome).suffix.lower()
    stored = f"{h[:2]}/{h}{ext}"
    dst = FILES / stored
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    testo = estrai(dst, ext) or pulisci_testo(extra)
    ocr = "coda" if (getset("ocr_auto") != "0" and (ext in IMG_EXT or (ext == ".pdf" and len(testo.strip()) < 60))) else ""
    with LOCK, db() as c:
        try:
            cur = c.execute(
                "INSERT INTO docs(name,ext,size,hash,added,stored,ambito,ocr) VALUES(?,?,?,?,?,?,?,?)",
                (nome, ext, dst.stat().st_size, h, time.strftime("%Y-%m-%d %H:%M"), stored, ambito, ocr))
        except sqlite3.IntegrityError:
            return None
        c.execute("INSERT INTO fts(rowid,name,body,tags) VALUES(?,?,?,?)",
                  (cur.lastrowid, nome, testo, ""))
        return cur.lastrowid


def scansiona():
    cartella = getset("folder")
    if not cartella or not Path(cartella).is_dir():
        return 0
    n = 0
    with SCAN_LOCK:
        for root, _, nomi in os.walk(cartella):
            for fn in nomi:
                if fn.startswith((".", "~$")):
                    continue
                p = Path(root) / fn
                try:
                    st = p.stat()
                except OSError:
                    continue
                if time.time() - st.st_mtime < 3:  # ancora in copia
                    continue
                with db() as c:
                    s = c.execute("SELECT mtime,size FROM seen WHERE path=?", (str(p),)).fetchone()
                if s and s["mtime"] == st.st_mtime and s["size"] == st.st_size:
                    continue
                parti = p.relative_to(cartella).parts
                amb = parti[0].lower() if len(parti) > 1 and parti[0].lower() in ("lavoro", "casa") else ""
                try:
                    if aggiungi(p, fn, amb):
                        n += 1
                except Exception as e:
                    print("Errore import", p, e)
                    continue
                with LOCK, db() as c:
                    c.execute("INSERT OR REPLACE INTO seen VALUES(?,?,?)", (str(p), st.st_mtime, st.st_size))
    return n


def sorvegliante():
    while True:
        try:
            scansiona()
        except Exception as e:
            print("Errore scansione:", e)
        time.sleep(10)


def query_fts(q):
    return " ".join('"%s"*' % t for t in re.findall(r"\w+", q))


def elenco(q, amb, tag):
    with db() as c:
        if query_fts(q):
            righe = c.execute(
                "SELECT d.*, snippet(fts,1,char(1),char(2),' … ',18) AS snip FROM fts "
                "JOIN docs d ON d.id=fts.rowid WHERE fts MATCH ? ORDER BY rank LIMIT 300",
                (query_fts(q),)).fetchall()
        else:
            righe = c.execute("SELECT *, '' AS snip FROM docs ORDER BY id DESC LIMIT 500").fetchall()
    out = []
    for r in righe:
        tags = [t for t in r["tags"].split(",") if t]
        if amb and r["ambito"] != amb:
            continue
        if tag and tag not in tags:
            continue
        out.append(dict(id=r["id"], name=r["name"], ext=r["ext"], size=r["size"], added=r["added"],
                        ambito=r["ambito"], tags=tags, snip=r["snip"], ocr=r["ocr"] or ""))
    return out


def riepilogo():
    with db() as c:
        righe = c.execute("SELECT ambito, tags FROM docs").fetchall()
    amb, tags = {"": 0, "lavoro": 0, "casa": 0}, {}
    for r in righe:
        amb[r["ambito"]] = amb.get(r["ambito"], 0) + 1
        for t in r["tags"].split(","):
            if t:
                tags[t] = tags.get(t, 0) + 1
    return {"totale": len(righe), "ambiti": amb, "tags": dict(sorted(tags.items()))}


def salva_meta(i, tags, amb):
    tags = sorted({t.replace(",", " ").strip().lower() for t in tags if t.strip()})
    with LOCK, db() as c:
        c.execute("UPDATE docs SET tags=?, ambito=? WHERE id=?", (",".join(tags), amb, i))
        c.execute("UPDATE fts SET tags=? WHERE rowid=?", (" ".join(tags), i))


def elimina(i):
    with LOCK, db() as c:
        r = c.execute("SELECT stored FROM docs WHERE id=?", (i,)).fetchone()
        if not r:
            return
        c.execute("DELETE FROM docs WHERE id=?", (i,))
        c.execute("DELETE FROM fts WHERE rowid=?", (i,))
    try:
        (FILES / r["stored"]).unlink()
    except OSError:
        pass


PAPRA = {"run": False, "msg": "", "n": 0, "tot": 0, "nuovi": 0, "dup": 0, "err": 0}


def papra_get(base, key, path, params=None):
    url = base.rstrip("/") + path + ("?" + urllib.parse.urlencode(params) if params else "")
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + key,
                                               "User-Agent": "DocuArchivio/1.0"})
    return urllib.request.urlopen(req, timeout=60)


def papra_json(base, key, path, params=None):
    with papra_get(base, key, path, params) as resp:
        return json.loads(resp.read().decode("utf-8"))


def importa_papra(base, key):
    P = PAPRA
    P.update(run=True, msg="Connessione a Papra…", n=0, tot=0, nuovi=0, dup=0, err=0)
    try:
        orgs = papra_json(base, key, "/api/organizations").get("organizations", [])
        if not orgs:
            P["msg"] = "Nessuna organizzazione trovata in Papra."
            return
        tmpdir = BASE / "tmp"
        tmpdir.mkdir(exist_ok=True)
        for org in orgs:
            oid, onome = org["id"], (org.get("name") or "").strip()
            amb = onome.lower() if onome.lower() in ("lavoro", "casa") else ""
            tag_org = "" if amb else onome.lower()
            docs, pagina = [], 0
            while True:
                r = papra_json(base, key, f"/api/organizations/{oid}/documents",
                               {"pageIndex": pagina, "pageSize": 100})
                blocco = r.get("documents", [])
                docs += blocco
                P["tot"] += len(blocco)
                if not blocco or len(docs) >= r.get("documentsCount", 0):
                    break
                pagina += 1
            for d in docs:
                P["msg"] = f"Importo da «{onome}»: {d.get('name', '')}"
                try:
                    nome = os.path.basename(d.get("name") or d["id"])
                    if not Path(nome).suffix:
                        nome += mimetypes.guess_extension(d.get("mimeType") or "") or ""
                    tmp = tmpdir / f"pp_{time.time_ns()}"
                    try:
                        with papra_get(base, key, f"/api/organizations/{oid}/documents/{d['id']}/file") as resp, \
                                open(tmp, "wb") as f:
                            shutil.copyfileobj(resp, f)
                        i = aggiungi(tmp, nome, amb, d.get("content") or "")
                    finally:
                        tmp.unlink(missing_ok=True)
                    if i:
                        tags = [(t.get("name") if isinstance(t, dict) else str(t)) or ""
                                for t in (d.get("tags") or [])]
                        if tag_org:
                            tags.append(tag_org)
                        if tags:
                            salva_meta(i, tags, amb)
                        P["nuovi"] += 1
                    else:
                        P["dup"] += 1
                except Exception as e:
                    P["err"] += 1
                    print("Errore import Papra", d.get("name"), e)
                P["n"] += 1
        P["msg"] = "Importazione completata."
    except urllib.error.HTTPError as e:
        P["msg"] = ("Papra ha rifiutato la richiesta (%d). Controlla la chiave API: servono i permessi "
                    "organizations:read, documents:read e tags:read." % e.code)
    except Exception as e:
        P["msg"] = "Errore: %s" % e
    finally:
        P["run"] = False


MOTORE = BASE / "motore"
MOTORE_FISSO = "b11669"
LLM = {"proc": None, "port": 0, "modello": ""}
LLM_LOCK = threading.Lock()
AVVIO_LOCK = threading.Lock()
DL = {"run": False, "pct": 0, "msg": "", "fase": "", "id": ""}
OCR = {"doc": 0, "nome": "", "pagina": 0, "tot": 0, "attesa": "", "errore": "", "pip": ""}


def arresta_tutto():
    ferma_motore()

atexit.register(arresta_tutto)


def trova_motore():
    t = list(MOTORE.rglob("llama-server.exe")) if MOTORE.exists() else []
    return t[0] if t else None


def elenco_modelli():
    out = [dict(m, presente=(MODELLO / m["file"]).exists()) for m in MODELLI]
    noti = {m["file"] for m in MODELLI}
    for f in (sorted(MODELLO.glob("*.gguf")) if MODELLO.exists() else []):
        if f.name not in noti:
            out.append({"id": "file:" + f.name, "nome": f.name, "mb": f.stat().st_size // 1048576, "file": f.name,
                        "url": "", "desc": "File messo a mano nella cartella modello.", "presente": True})
    return out


def modello_scelto():
    el, mid = elenco_modelli(), getset("modello_id") or "ultraleggero"
    return next((m for m in el if m["id"] == mid), el[0])


def trova_modello():
    p = MODELLO / modello_scelto()["file"]
    return p if p.exists() else None


def url_motore(tag):
    return ("https://github.com/ggml-org/llama.cpp/releases/download/%s/llama-%s-bin-win-cpu-x64.zip"
            % (tag, tag))


def ultimo_tag():
    intest = {"User-Agent": "DocuArchivio/1.0"}
    try:
        with urllib.request.urlopen(urllib.request.Request(
                "https://github.com/ggml-org/llama.cpp/releases/latest", headers=intest), timeout=30) as r:
            tag = r.geturl().rsplit("/", 1)[1]
        if not re.fullmatch(r"b\d+", tag):
            with urllib.request.urlopen(urllib.request.Request(
                    "https://github.com/ggml-org/llama.cpp/releases/download/%s/nightly-tag.txt" % tag,
                    headers=intest), timeout=30) as r:
                tag = r.read().decode().strip()
        if re.fullmatch(r"b\d+", tag):
            return tag
    except Exception:
        pass
    return MOTORE_FISSO


def scarica(url, dest, fase):
    DL.update(fase=fase, pct=0)
    part = Path(str(dest) + ".part")
    req = urllib.request.Request(url, headers={"User-Agent": "DocuArchivio/1.0"})
    with urllib.request.urlopen(req, timeout=60) as r, open(part, "wb") as f:
        tot, fatto = int(r.headers.get("Content-Length", 0)), 0
        while True:
            b = r.read(1 << 20)
            if not b:
                break
            f.write(b)
            fatto += len(b)
            if tot:
                DL["pct"] = int(fatto * 100 / tot)
    if tot and fatto != tot:
        raise IOError("file incompleto")
    part.replace(dest)


def scarica_modello(mid=""):
    DL.update(run=True, pct=0, msg="", fase="")
    try:
        m = next((x for x in MODELLI if x["id"] == (mid or getset("modello_id") or "ultraleggero")), MODELLI[0])
        DL["id"] = m["id"]
        MOTORE.mkdir(parents=True, exist_ok=True)
        MODELLO.mkdir(parents=True, exist_ok=True)
        if not trova_motore():
            if os.name != "nt":
                raise RuntimeError("il download automatico del motore c'e' solo per Windows")
            zp, errore = MOTORE / "motore.zip", None
            for tag in dict.fromkeys([ultimo_tag(), MOTORE_FISSO]):
                try:
                    scarica(url_motore(tag), zp, "il motore")
                    errore = None
                    break
                except Exception as e:
                    errore = e
            if errore:
                raise errore
            DL.update(fase="il motore (estrazione)", pct=100)
            with zipfile.ZipFile(zp) as z:
                z.extractall(MOTORE)
            zp.unlink(missing_ok=True)
        if not (MODELLO / m["file"]).exists():
            scarica(m["url"], MODELLO / m["file"], "il modello " + m["nome"])
        setset("modello_id", m["id"])
        threading.Thread(target=avvia_motore, daemon=True).start()
    except Exception as e:
        DL["msg"] = ("Download non riuscito (%s). Puoi copiare a mano llama-server.exe (con le sue DLL) in "
                     "DocuArchivio_dati/motore e un file .gguf in DocuArchivio_dati/modello." % e)
    finally:
        DL["run"] = False


def motore_pronto():
    try:
        with urllib.request.urlopen("http://127.0.0.1:%d/health" % LLM["port"], timeout=3) as r:
            return r.status == 200
    except Exception:
        return False


def avvia_motore():
    with AVVIO_LOCK:
        return _avvia_motore()


def libera_porta():
    with socket.socket() as sk:
        sk.bind(("127.0.0.1", 0))
        return sk.getsockname()[1]


def ferma_motore():
    p = LLM["proc"]
    if p and p.poll() is None:
        p.terminate()
        try:
            p.wait(5)
        except Exception:
            p.kill()
    LLM.update(proc=None, port=0, modello="")


def _avvia_motore():
    exe, m = trova_motore(), trova_modello()
    if not exe or not m:
        return "Mancano il motore o il modello IA: apri Impostazioni (⚙) e scarica un modello."
    p = LLM["proc"]
    if p and p.poll() is None and LLM["modello"] == m.name:
        if motore_pronto():
            return ""
    else:
        ferma_motore()
        porta = libera_porta()
        log = open(BASE / "motore.log", "wb")
        # Calcola max 2 thread per evitare saturazione della CPU su PC dual-core (es. N2840)
        num_threads = min(2, os.cpu_count() or 2)
        p = subprocess.Popen([str(exe), "-m", str(m), "-c", "2048", "--host", "127.0.0.1",
                              "--port", str(porta), "-t", str(num_threads)],
                             stdout=log, stderr=log, cwd=str(exe.parent))
        LLM.update(proc=p, port=porta, modello=m.name)
    fine = time.time() + 900
    while time.time() < fine:
        if p.poll() is not None:
            LLM.update(proc=None, port=0, modello="")
            return ("Il motore IA si è chiuso (codice %s). Dettagli nel file DocuArchivio_dati/motore.log"
                    % p.returncode)
        if motore_pronto():
            return ""
        time.sleep(1)
    ferma_motore()
    return "Il motore IA non è partito entro 15 minuti."


def senza_think(gen):
    dentro, buf = False, ""
    for t in gen:
        buf += t
        while True:
            if dentro:
                k = buf.find("</think>")
                if k < 0:
                    buf = buf[-8:]
                    break
                buf, dentro = buf[k + 8:].lstrip("\n"), False
            else:
                k = buf.find("<think>")
                if k < 0:
                    sicuro = len(buf) - 6
                    if sicuro > 0:
                        yield buf[:sicuro]
                        buf = buf[sicuro:]
                    break
                if k:
                    yield buf[:k]
                buf, dentro = buf[k + 7:], True
    if not dentro and buf:
        yield buf


def chat_stream(msgs, max_tokens=600):
    return senza_think(_flusso(msgs, max_tokens))


def _flusso(msgs, max_tokens=600):
    dati = json.dumps({"messages": msgs, "stream": True, "temperature": 0.2, "max_tokens": max_tokens,
                        "chat_template_kwargs": {"enable_thinking": False}}).encode()
    req = urllib.request.Request("http://127.0.0.1:%d/v1/chat/completions" % LLM["port"], data=dati,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=900) as r:
        for riga in r:
            riga = riga.decode("utf-8", "ignore").strip()
            if not riga.startswith("data:"):
                continue
            corpo = riga[5:].strip()
            if corpo == "[DONE]":
                break
            try:
                t = json.loads(corpo)["choices"][0]["delta"].get("content")
            except (ValueError, KeyError, IndexError):
                continue
            if t:
                yield t


def scegli_estratto(testo, q, limite=CHARS):
    if len(testo) <= limite:
        return testo
    pezzi = [testo[k:k + 800] for k in range(0, len(testo), 700)]
    n = max(1, limite // 800)
    if re.search(r"riassum|sintesi|sintetizz|di cosa parla|di che cosa|argomento", q.lower()):
        passo = len(pezzi) / n
        scelti = sorted({int(k * passo) for k in range(n)})
    else:
        rad = {w[:5] for w in re.findall(r"\w{3,}", q.lower()) if w not in STOP}
        tok = [re.findall(r"\w{3,}", p.lower()) for p in pezzi]
        idf = {rd: math.log(1 + len(pezzi) / (1 + sum(1 for t in tok if any(w.startswith(rd) for w in t))))
               for rd in rad}
        punti = []
        for k, t in enumerate(tok):
            sc = 0.0
            for rd in rad:
                tf = sum(1 for w in t if w.startswith(rd))
                if tf:
                    sc += idf[rd] * (1 + math.log(tf))
            punti.append((sc, k))
        punti.sort(reverse=True)
        scelti = sorted(k for sc, k in punti[:n] if sc > 0) or list(range(n))
    return "\n…\n".join(pezzi[k] for k in scelti)


def chat_once(msgs, max_tokens=60):
    dati = json.dumps({"messages": msgs, "stream": False, "temperature": 0, "max_tokens": max_tokens,
                        "chat_template_kwargs": {"enable_thinking": False}}).encode()
    req = urllib.request.Request("http://127.0.0.1:%d/v1/chat/completions" % LLM["port"], data=dati,
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=900) as r:
        return re.sub(r"<think>.*?</think>", "", json.loads(r.read().decode("utf-8"))["choices"][0]["message"]["content"],
                      flags=re.S).strip()


DEF_WEB = "Dai un'opinione sui prezzi e sulle voci di questo documento, confrontandoli con ciò che si trova online."
SISTEMA_WEB = SISTEMA + (" Ti vengono dati anche risultati di una ricerca web. Distingui chiaramente cio' che viene "
                         "dal documento da cio' che viene dal web, cita il sito per ogni dato preso dal web, non "
                         "inventare prezzi e, se i risultati non bastano per un parere, dillo. Dai opinioni "
                         "prudenti e indicative.")
SISTEMA_QUERY = ("Devi preparare ricerche web per valutare prezzi o informazioni citati in un documento. Scrivi al "
                 "massimo 2 query di ricerca brevi in italiano, una per riga, sui prodotti, servizi o voci "
                 "principali. NON includere nomi di persone o di aziende private, indirizzi, numeri di fattura, "
                 "codici fiscali o altri dati personali. Rispondi solo con le query.")


def pulisci_html(t):
    return html.unescape(re.sub(r"<[^>]+>", "", t)).strip()


def cerca_web(q, n=4):
    dati = urllib.parse.urlencode({"q": q, "kl": "it-it"}).encode()
    req = urllib.request.Request("https://html.duckduckgo.com/html/", data=dati, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0 Safari/537.36",
        "Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=20) as r:
        h = r.read().decode("utf-8", "ignore")
    ancore = list(re.finditer(r'<a\b([^>]*class="[^"]*result__a[^"]*"[^>]*)>(.*?)</a>', h, re.S))
    snip = list(re.finditer(r'<a\b([^>]*class="[^"]*result__snippet[^"]*"[^>]*>(.*?)</a>', h, re.S))
    out = []
    for k, am in enumerate(ancore):
        attr, tit = am.group(1), am.group(2)
        m = re.search(r'href="([^"]+)"', attr)
        if not m:
            continue
        url = html.unescape(m.group(1))
        if url.startswith("//"):
            url = "https:" + url
        uddg = parse_qs(urlparse(url).query).get("uddg")
        if uddg:
            url = uddg[0]
        if "duckduckgo.com/y.js" in url:
            continue
        fine = ancore[k + 1].start() if k + 1 < len(ancore) else len(h)
        sm = next((x for x in snip if am.end() <= x.start() < fine), None)
        out.append({"titolo": pulisci_html(tit), "url": url, "testo": pulisci_html(sm.group(1)) if sm else ""})
        if len(out) >= n:
            break
    return out


def prepara_query_web(estratto, q):
    try:
        t = chat_once([{"role": "system", "content": SISTEMA_QUERY},
                       {"role": "user", "content": "Documento (estratto):\n%s\n\nRichiesta: %s" % (estratto[:1500], q)}], 60)
    except Exception:
        return []
    out = []
    for riga in t.splitlines():
        riga = re.sub(r"^\s*(?:(?:[-*•]|\d+[.)])\s*)+", "", riga).strip(" \"«»'")
        if len(riga) >= 4:
            out.append(riga[:100])
    return out[:2]


def rispondi_ia(i, q, storia, web=False):
    if not LLM_LOCK.acquire(blocking=False):
        yield "Il modello sta già elaborando un'altra domanda. Riprova tra poco."
        return
    try:
        errore = avvia_motore()
        if errore:
            yield errore
            return
        with db() as c:
            r = c.execute("SELECT d.name, f.body FROM docs d JOIN fts f ON f.rowid=d.id WHERE d.id=?",
                          (i,)).fetchone()
        if not r:
            yield "Documento non trovato."
            return
        testo = (r["body"] or "").strip()
        if not testo:
            yield "Questo documento non ha testo leggibile (forse è una scansione: servirebbe l'OCR)."
            return
        q = q.strip() or (DEF_WEB if web else "")
        if not q:
            yield "Scrivi una domanda."
            return
        ex = scegli_estratto(testo, q, 3000 if web else CHARS)
        blocco_web, fonti = "", []
        if web:
            query = prepara_query_web(ex, q)
            if not query:
                yield "Non sono riuscito a preparare una ricerca online: rispondo solo con il documento.\n\n"
            else:
                yield "🔎 Ricerca online: " + " · ".join("«%s»" % x for x in query) + "\n\n"
                for x in query:
                    try:
                        res = cerca_web(x)
                    except Exception:
                        res = []
                    if res:
                        blocco_web += "Ricerca «%s»:\n" % x + "\n".join(
                            "- %s (%s): %s" % (v["titolo"], urlparse(v["url"]).netloc, v["testo"][:300])
                            for v in res) + "\n"
                        fonti += res
                if not fonti:
                    yield ("Nessun risultato online (manca la connessione o la ricerca è bloccata): "
                           "rispondo solo con il documento.\n\n")
        msgs = [{"role": "system", "content": SISTEMA_WEB if fonti else SISTEMA}]
        for t in storia[-2:]:
            msgs += [{"role": "user", "content": str(t.get("q", ""))[:300]},
                     {"role": "assistant", "content": str(t.get("a", ""))[:400]}]
        contenuto = "Estratto del documento «%s»:\n---\n%s\n---\n" % (r["name"], ex)
        if blocco_web:
            contenuto += "Risultati della ricerca web:\n%s---\n" % blocco_web
        msgs.append({"role": "user", "content": contenuto + "Domanda: " + q})
        yield from chat_stream(msgs)
        if fonti:
            uniche = list({v["url"]: v for v in fonti}.values())[:6]
            yield "\n\nFonti web:\n" + "\n".join("- %s — %s" % (v["titolo"], v["url"]) for v in uniche)
    except Exception as e:
        yield "\n[Errore del modello: %s]" % e
    finally:
        LLM_LOCK.release()


SISTEMA_CERCA = ("Sei l'assistente di un archivio di documenti. L'utente sta cercando un documento. Ti do l'elenco "
                 "dei documenti candidati, numerati, con un estratto. Indica quale (o quali) corrisponde meglio "
                 "alla richiesta, citando il numero tra parentesi quadre e spiegando in una frase perche'. Se "
                 "nessuno sembra giusto, dillo. Rispondi in italiano, in modo breve.")


def cerca_ia(q):
    q = q.strip()
    if not q:
        return
    if not LLM_LOCK.acquire(blocking=False):
        yield "Il modello sta già elaborando un'altra richiesta. Riprova tra poco."
        return
    try:
        errore = avvia_motore()
        if errore:
            yield errore
            return
        try:
            kw = chat_once([{"role": "system", "content":
                             "Estrai da una richiesta di ricerca di documenti le parole chiave utili, includendo "
                             "sinonimi e termini correlati che potrebbero comparire nel documento. Rispondi SOLO "
                             "con le parole separate da virgola, massimo 10, senza spiegazioni."},
                            {"role": "user", "content": q}], 40)
        except Exception:
            kw = ""
        termini = [w for w in dict.fromkeys(re.findall(r"\w{3,}", (q + " " + kw).lower())) if w not in STOP][:14]
        if not termini:
            yield "Scrivi qualche parola in più per la ricerca."
            return
        with db() as c:
            righe = c.execute(
                "SELECT d.id, d.name, d.ext, d.ambito, d.tags, d.added, snippet(fts,-1,'','',' … ',24) AS sn "
                "FROM fts JOIN docs d ON d.id=fts.rowid WHERE fts MATCH ? ORDER BY rank LIMIT 8",
                (" OR ".join('"%s"*' % t for t in termini),)).fetchall()
        yield "\x01" + json.dumps([{"id": r["id"], "name": r["name"], "ext": r["ext"]} for r in righe]) + "\x02"
        if not righe:
            yield "Non ho trovato documenti collegati a questa richiesta."
            return
        elenco_txt = "\n".join(
            "[%d] %s | %s | tag: %s | aggiunto %s | …%s" % (n, r["name"], r["ambito"] or "-", r["tags"] or "-",
                                                          r["added"], (r["sn"] or "")[:250].replace("\n", " "))
            for n, r in enumerate(righe, 1))
        yield from chat_stream([{"role": "system", "content": SISTEMA_CERCA},
                                {"role": "user", "content": "Richiesta: %s\n\nCandidati:\n%s" % (q, elenco_txt)}], 400)
    except Exception as e:
        yield "\n[Errore del modello: %s]" % e
    finally:
        LLM_LOCK.release()


def trova_tesseract():
    t = shutil.which("tesseract")
    if t:
        return t
    basi = [os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")]
    if os.environ.get("LOCALAPPDATA"):
        basi.append(os.path.join(os.environ["LOCALAPPDATA"], "Programs"))
    for b in basi:
        if b and (Path(b) / "Tesseract-OCR" / "tesseract.exe").exists():
            return str(Path(b) / "Tesseract-OCR" / "tesseract.exe")
    portatile = OCR_DIR / "tesseract" / "tesseract.exe"
    return str(portatile) if portatile.exists() else None


def ocr_componenti():
    return {"pdfium": importlib.util.find_spec("pypdfium2") is not None,
            "pillow": importlib.util.find_spec("PIL") is not None}


def ocr_lingua_ok():
    return (TESSDATA / "ita.traineddata").exists()


def ocr_manca(ext):
    if not trova_tesseract():
        return "Tesseract non è installato"
    if not ocr_lingua_ok():
        return "manca la lingua italiana"
    c = ocr_componenti()
    if ext == ".pdf" and not (c["pdfium"] and c["pillow"]):
        return "mancano i componenti per i PDF (pypdfium2 e pillow)"
    return ""


def scarica_lingua():
    OCR["errore"] = ""
    try:
        TESSDATA.mkdir(parents=True, exist_ok=True)
        req = urllib.request.Request(LINGUA_URL, headers={"User-Agent": "DocuArchivio/1.0"})
        with urllib.request.urlopen(req, timeout=60) as r:
            dati = r.read()
        if len(dati) < 100000:
            raise IOError("file non valido")
        (TESSDATA / "ita.traineddata").write_bytes(dati)
    except Exception as e:
        OCR["errore"] = "Download della lingua non riuscito: %s" % e


def installa_pip(pacchetti):
    OCR["pip"] = "Installo %s…" % " ".join(pacchetti)
    try:
        r = subprocess.run([sys.executable, "-m", "pip", "install", *pacchetti],
                           capture_output=True, text=True, timeout=900)
        importlib.invalidate_caches()
        OCR["pip"] = "Componenti installati." if r.returncode == 0 else \
            "Installazione non riuscita: " + (r.stderr or r.stdout)[-200:]
    except Exception as e:
        OCR["pip"] = "Installazione non riuscita: %s" % e


def tess_testo(img):
    r = subprocess.run([trova_tesseract(), str(img), "stdout", "-l", "ita", "--tessdata-dir", str(TESSDATA),
                        "--psm", "3"], capture_output=True, timeout=900,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    if r.returncode != 0:
        raise RuntimeError((r.stderr or b"").decode("utf-8", "ignore")[-200:] or "errore di Tesseract")
    return r.stdout.decode("utf-8", "ignore")


def prepara_img(im):
    from PIL import ImageOps
    im = ImageOps.exif_transpose(im).convert("L")
    if max(im.size) < 1600:
        im = im.resize((im.width * 2, im.height * 2))
    return im


def ocr_file(path, ext, aggiorna):
    tmp = BASE / "tmp"
    tmp.mkdir(exist_ok=True)
    parti = []

    def una(im):
        f = tmp / ("ocr_%d.png" % time.time_ns())
        try:
            prepara_img(im).save(f)
            parti.append(tess_testo(f))
        finally:
            f.unlink(missing_ok=True)

    if ext == ".pdf":
        import pypdfium2 as pdfium
        pdf = pdfium.PdfDocument(str(path))
        try:
            n = len(pdf)
            for k in range(n):
                aggiorna(k + 1, n)
                w, h = pdf[k].get_size()
                una(pdf[k].render(scale=min(OCR_DPI / 72, 4000 / max(w, h, 1))).to_pil())
        finally:
            pdf.close()
    else:
        try:
            from PIL import Image, ImageSequence
        except ImportError:
            aggiorna(1, 1)
            return tess_testo(path)
        with Image.open(path) as im0:
            frames = [fr.copy() for fr in ImageSequence.Iterator(im0)][:50]
        for k, fr in enumerate(frames):
            aggiorna(k + 1, len(frames))
            una(fr)
    return pulisci_testo("\n\n".join(parti))


def accoda_ocr(i=None):
    estensioni = sorted(IMG_EXT)
    marche = ",".join("?" * len(estensioni))
    with LOCK, db() as c:
        if i:
            cur = c.execute("UPDATE docs SET ocr='coda' WHERE id=? AND (ext='.pdf' OR ext IN (%s))" % marche,
                            (i, *estensioni))
        else:
            cur = c.execute(
                "UPDATE docs SET ocr='coda' WHERE ocr IN ('','errore') AND (ext='.pdf' OR ext IN (%s)) "
                "AND length(trim(coalesce((SELECT body FROM fts WHERE fts.rowid=docs.id),''))) < 60" % marche,
                estensioni)
        return cur.rowcount


def ocr_lavoratore():
    while True:
        try:
            with db() as c:
                r = c.execute("SELECT id,name,ext,stored FROM docs WHERE ocr='coda' ORDER BY id LIMIT 1").fetchone()
            if not r:
                OCR["attesa"] = ""
                time.sleep(3)
                continue
            manca = ocr_manca(r["ext"])
            if manca:
                OCR["attesa"] = "OCR in attesa: " + manca + "."
                time.sleep(5)
                continue
            OCR.update(doc=r["id"], nome=r["name"], pagina=0, tot=0, attesa="")
            try:
                testo = ocr_file(FILES / r["stored"], r["ext"], lambda k, n: OCR.update(pagina=k, tot=n))
                with LOCK, db() as c:
                    c.execute("UPDATE docs SET ocr='ok' WHERE id=?", (r["id"],))
                    c.execute("UPDATE fts SET body=? WHERE rowid=?", (testo, r["id"]))
            except Exception as e:
                print("Errore OCR", r["name"], e)
                OCR["errore"] = "Errore OCR su %s: %s" % (r["name"], e)
                with LOCK, db() as c:
                    c.execute("UPDATE docs SET ocr='errore' WHERE id=?", (r["id"],))
            OCR.update(doc=0, nome="", pagina=0, tot=0)
        except Exception as e:
            print("Errore nel lavoratore OCR", e)
            time.sleep(5)


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def rispondi(self, code, corpo, tipo="application/json", extra=None):
        if isinstance(corpo, (dict, list)):
            corpo = json.dumps(corpo).encode()
        elif isinstance(corpo, str):
            corpo = corpo.encode()
        self.send_response(code)
        self.send_header("Content-Type", tipo)
        self.send_header("Content-Length", str(len(corpo)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(corpo)

    def stream_ia(self, gen):
        self.send_response(200)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Cache-Control", "no-cache")
        self.end_headers()
        try:
            for pezzo in gen:
                try:
                    self.wfile.write(pezzo.encode("utf-8"))
                    self.wfile.flush()
                except OSError:
                    break
        finally:
            gen.close()

    def corpo(self):
        return self.rfile.read(int(self.headers.get("Content-Length", 0)))

    def do_GET(self):
        u = urlparse(self.path)
        qs = parse_qs(u.query)
        g = lambda k: qs.get(k, [""])[0]
        if u.path == "/":
            return self.rispondi(200, PAGINA, "text/html; charset=utf-8")
        if u.path == "/api/docs":
            return self.rispondi(200, elenco(g("q"), g("ambito"), g("tag")))
        if u.path == "/api/riepilogo":
            return self.rispondi(200, riepilogo())
        if u.path == "/api/settings":
            return self.rispondi(200, {"folder": getset("folder"), "papra_url": getset("papra_url")})
        if u.path == "/api/papra/status":
            return self.rispondi(200, PAPRA)
        if u.path == "/api/ia/stato":
            m = trova_modello()
            return self.rispondi(200, {"motore": trova_motore() is not None, "attivo": motore_pronto(),
                                       "modello": m.name if m else "", "scarica": DL,
                                       "modelli": elenco_modelli(), "scelto": modello_scelto()["id"]})
        if u.path == "/api/ocr/stato":
            with db() as c:
                coda = c.execute("SELECT count(*) FROM docs WHERE ocr='coda'").fetchone()[0]
            return self.rispondi(200, {"tesseract": trova_tesseract() is not None, "lingua": ocr_lingua_ok(),
                                       "componenti": ocr_componenti(), "auto": getset("ocr_auto") != "0",
                                       "coda": coda, "corrente": {k: OCR[k] for k in ("doc", "nome", "pagina", "tot")},
                                       "attesa": OCR["attesa"], "errore": OCR["errore"], "pip": OCR["pip"]})
        m = re.match(r"/api/testo/(\d+)$", u.path)
        if m:
            with db() as c:
                r = c.execute("SELECT body FROM fts WHERE rowid=?", (int(m[1]),)).fetchone()
            return self.rispondi(200, (r["body"] if r else "") or "", "text/plain; charset=utf-8")
        m = re.match(r"/file/(\d+)$", u.path)
        if m:
            with db() as c:
                r = c.execute("SELECT name,stored FROM docs WHERE id=?", (int(m[1]),)).fetchone()
            if r and (FILES / r["stored"]).exists():
                tipo = mimetypes.guess_type(r["name"])[0] or "application/octet-stream"
                return self.rispondi(200, (FILES / r["stored"]).read_bytes(), tipo, {
                    "Content-Disposition": ("attachment" if g("scarica") else "inline")
                    + "; filename*=UTF-8''" + quote(r["name"])})
        self.rispondi(404, {})

    def do_POST(self):
        if self.headers.get("X-DA") != "1":
            return self.rispondi(403, {})
        u = urlparse(self.path)
        qs = parse_qs(u.query)
        if u.path == "/api/upload":
            nome = os.path.basename(qs.get("name", ["file"])[0]) or "file"
            tmp = BASE / "tmp"
            tmp.mkdir(exist_ok=True)
            p = tmp / f"up_{time.time_ns()}"
            p.write_bytes(self.corpo())
            try:
                i = aggiungi(p, nome, qs.get("ambito", [""])[0])
            finally:
                p.unlink(missing_ok=True)
            return self.rispondi(200, {"id": i, "duplicato": i is None})
        if u.path == "/api/settings":
            d = json.loads(self.corpo() or b"{}")
            setset("folder", d.get("folder", "").strip().strip('"'))
            return self.rispondi(200, {"ok": True})
        if u.path == "/api/chiedi":
            d = json.loads(self.corpo() or b"{}")
            return self.stream_ia(rispondi_ia(int(d.get("id", 0)), str(d.get("q", "")), d.get("storia") or [],
                                              bool(d.get("web"))))
        if u.path == "/api/cerca-ia":
            d = json.loads(self.corpo() or b"{}")
            return self.stream_ia(cerca_ia(str(d.get("q", ""))))
        if u.path == "/api/ia/scarica":
            d = json.loads(self.corpo() or b"{}")
            if not DL["run"]:
                DL.update(run=True, pct=0, msg="", fase="")
                threading.Thread(target=scarica_modello, args=(str(d.get("id", "")),), daemon=True).start()
            return self.rispondi(200, {"ok": True})
        if u.path == "/api/ia/modello":
            d = json.loads(self.corpo() or b"{}")
            if LLM_LOCK.locked():
                return self.rispondi(200, {"ok": False, "msg": "Il modello sta rispondendo: riprova tra poco."})
            if d.get("id") not in {m["id"] for m in elenco_modelli()}:
                return self.rispondi(200, {"ok": False, "msg": "Modello sconosciuto."})
            setset("modello_id", d["id"])
            threading.Thread(target=avvia_motore, daemon=True).start()
            return self.rispondi(200, {"ok": True, "msg": "Cambio modello: il caricamento richiede un po'."})
        if u.path == "/api/ia/elimina":
            d = json.loads(self.corpo() or b"{}")
            m = next((x for x in elenco_modelli() if x["id"] == d.get("id")), None)
            if not m or m["id"] == modello_scelto()["id"]:
                return self.rispondi(200, {"ok": False, "msg": "Non posso eliminare il modello in uso."})
            try:
                (MODELLO / m["file"]).unlink(missing_ok=True)
            except OSError:
                return self.rispondi(200, {"ok": False, "msg": "Il file è ancora in uso, riprova tra poco."})
            return self.rispondi(200, {"ok": True, "msg": "Modello eliminato."})
        if u.path == "/api/ocr/impostazioni":
            d = json.loads(self.corpo() or b"{}")
            setset("ocr_auto", "1" if d.get("auto") else "0")
            return self.rispondi(200, {"ok": True})
        if u.path == "/api/ocr/lingua":
            threading.Thread(target=scarica_lingua, daemon=True).start()
            return self.rispondi(200, {"ok": True})
        if u.path == "/api/ocr/componenti":
            if not OCR["pip"].startswith("Installo"):
                threading.Thread(target=installa_pip, args=(["pypdfium2", "pillow"],), daemon=True).start()
            return self.rispondi(200, {"ok": True})
        if u.path == "/api/ocr/tutti":
            return self.rispondi(200, {"n": accoda_ocr()})
        mo = re.match(r"/api/ocr/(\d+)$", u.path)
        if mo:
            return self.rispondi(200, {"n": accoda_ocr(int(mo[1]))})
        if u.path == "/api/papra/start":
            if PAPRA["run"]:
                return self.rispondi(200, {"ok": False, "msg": "Importazione già in corso."})
            d = json.loads(self.corpo() or b"{}")
            base = (d.get("url") or "https://api.papra.app").strip()
            key = (d.get("key") or "").strip()
            if not key:
                return self.rispondi(200, {"ok": False, "msg": "Inserisci la chiave API di Papra."})
            setset("papra_url", base)
            PAPRA["run"] = True
            threading.Thread(target=importa_papra, args=(base, key), daemon=True).start()
            return self.rispondi(200, {"ok": True})
        if u.path == "/api/scan":
            return self.rispondi(200, {"n": scansiona()})
        m = re.match(r"/api/meta/(\d+)$", u.path)
        if m:
            d = json.loads(self.corpo())
            salva_meta(int(m[1]), d.get("tags", []), d.get("ambito", ""))
            return self.rispondi(200, {"ok": True})
        self.rispondi(404, {})

    def do_DELETE(self):
        if self.headers.get("X-DA") != "1":
            return self.rispondi(403, {})
        m = re.match(r"/api/docs/(\d+)$", urlparse(self.path).path)
        if m:
            elimina(int(m[1]))
            return self.rispondi(200, {"ok": True})
        self.rispondi(404, {})


PAGINA = r"""<!doctype html>
<html lang="it"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>DocuArchivio</title>
<style>
:root{--bg:#f6f7f8;--panel:#fff;--ink:#1c2328;--mute:#69747c;--line:#e1e5e8;--acc:#0f766e;--accbg:#e3f3f1;--mark:#fff1b8}
@media (prefers-color-scheme:dark){:root{--bg:#14181b;--panel:#1b2125;--ink:#e6eaec;--mute:#94a0a8;--line:#2b343a;--acc:#4fd1c5;--accbg:#173432;--mark:#5a4a12}}
*{box-sizing:border-box}
body{margin:0;font:15px/1.45 system-ui,"Segoe UI",sans-serif;background:var(--bg);color:var(--ink);height:100vh;display:grid;grid-template-columns:230px 1fr 320px}
aside,section,article{overflow:auto;padding:16px}
aside{background:var(--panel);border-right:1px solid var(--line)}
article{background:var(--panel);border-left:1px solid var(--line)}
h1{font-size:18px;margin:0 0 16px}
nav a{display:flex;justify-content:space-between;padding:7px 10px;border-radius:6px;color:var(--ink);text-decoration:none;cursor:pointer}
nav a:hover{background:var(--bg)} nav a.on{background:var(--accbg);color:var(--acc);font-weight:600}
nav small{color:var(--mute)} h3{font-size:13px;color:var(--mute);margin:18px 0 6px;font-weight:600}
input,select,button{font:inherit;color:inherit} 
input[type=text],input[type=password],select{width:100%;padding:8px 10px;border:1px solid var(--line);border-radius:6px;background:var(--bg)}
button{padding:7px 12px;border:1px solid var(--line);border-radius:6px;background:var(--panel);cursor:pointer}
button.p{background:var(--acc);border-color:var(--acc);color:var(--panel)} button.d{color:#c0392b}
button:focus-visible,input:focus-visible,select:focus-visible,a:focus-visible{outline:2px solid var(--acc);outline-offset:1px}
.bar{display:flex;gap:8px;margin-bottom:14px}.bar button{white-space:nowrap}
.doc{background:var(--panel);border:1px solid var(--line);border-radius:8px;padding:11px 14px;margin-bottom:8px;cursor:pointer}
.doc.on{border-color:var(--acc)} .doc b{word-break:break-word}
.meta{color:var(--mute);font-size:13px;margin-top:2px}
.tag{display:inline-block;background:var(--accbg);color:var(--acc);border-radius:10px;padding:0 8px;font-size:12px;margin-right:4px}
.snip{font-size:13px;color:var(--mute);margin-top:5px} mark{background:var(--mark);color:inherit}
.vuoto{color:var(--mute);text-align:center;padding:50px 10px;border:2px dashed var(--line);border-radius:10px}
body.drag section{outline:3px dashed var(--acc);outline-offset:-10px}
#msg{z-index:20;position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:var(--ink);color:var(--bg);padding:8px 16px;border-radius:8px;display:none}
.row{display:flex;gap:8px;margin-top:10px}
#vw{position:fixed;inset:0;z-index:5;background:var(--bg);display:flex;flex-direction:column}
#vw[hidden]{display:none}
#vw header{display:flex;gap:8px;align-items:center;padding:10px 16px;background:var(--panel);border-bottom:1px solid var(--line)}
#vw header b{flex:1;word-break:break-word}
.vwb{flex:1;display:grid;grid-template-columns:1fr 380px;min-height:0}
#vwDoc{overflow:auto;background:var(--panel);min-height:0}
#vwDoc iframe{width:100%;height:100%;border:0}
#vwDoc img{max-width:100%;display:block;margin:0 auto}
#vwDoc pre{margin:0;padding:18px;white-space:pre-wrap;word-break:break-word;font:14px/1.6 system-ui,"Segoe UI",sans-serif}
#vwChat{display:flex;flex-direction:column;border-left:1px solid var(--line);background:var(--panel);min-height:0}
#vwM{flex:1;overflow:auto;padding:14px}
.mq{background:var(--accbg);border-radius:8px;padding:8px 12px;margin:10px 0 6px 8%}
.ma{white-space:pre-wrap;padding:4px 2px}
#top,#scrim,.solomob,.vtabs{display:none}
.card{background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:16px;margin-bottom:14px}
.card h2{font-size:16px;margin:0 0 10px}
#set{position:fixed;inset:0;z-index:7;background:var(--bg);display:flex;flex-direction:column}
#set[hidden]{display:none}
#set header{display:flex;gap:8px;align-items:center;padding:10px 16px;background:var(--panel);border-bottom:1px solid var(--line)}
#set header b{flex:1}
.setb{flex:1;overflow:auto;padding:16px;width:100%;max-width:760px;margin:0 auto}
@media(max-width:900px){
 body{display:block;height:auto;min-height:100vh}
 #top{display:flex;align-items:center;gap:8px;position:sticky;top:0;z-index:4;background:var(--panel);border-bottom:1px solid var(--line);padding:8px 12px}
 #top b{flex:1}
 aside{position:fixed;top:0;bottom:0;left:0;width:82%;max-width:320px;z-index:8;transform:translateX(-105%);transition:transform .2s;box-shadow:2px 0 12px rgba(0,0,0,.3)}
 body.menu aside{transform:none}
 body.menu #scrim,body.det #scrim{display:block;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:6}
 article{position:fixed;left:0;right:0;bottom:0;max-height:70vh;z-index:7;transform:translateY(105%);transition:transform .2s;border-left:0;border-top:1px solid var(--line);box-shadow:0 -2px 12px rgba(0,0,0,.3)}
 body.det article{transform:none}
 .solomob{display:inline-block}
 section{padding:12px}
 .bar{flex-wrap:wrap}.bar input{flex:1 1 100%}
 button,input[type=text],input[type=password],select{min-height:42px}
 .vtabs{display:flex;gap:6px}
 #vw header{flex-wrap:wrap}
 .vwb{grid-template-columns:1fr}
 #vw.tabchat #vwDoc{display:none}
 #vw:not(.tabchat) #vwChat{display:none}
 #vwChat{border-left:0}
}
</style></head><body>
<header id="top"><button id="menuB" aria-label="Menu">☰</button><b>DocuArchivio</b><button id="setB2" aria-label="Impostazioni">⚙</button></header>
<div id="scrim"></div>
<aside>
 <h1>DocuArchivio</h1>
 <nav id="amb"></nav>
 <h3>Tag</h3><nav id="tags"></nav>
 <h3>Assistente IA</h3>
 <div class="meta iaSt" role="status"></div>
 <div class="row"><button class="iaDl" hidden>Scarica motore e modello IA</button></div>
 <div class="row" style="margin-top:16px"><button id="setB">⚙ Impostazioni</button></div>
</aside>
<section>
 <div class="bar"><input type="text" id="q" placeholder="Cerca, o descrivi cosa ti serve"><button id="qIA">Chiedi all'IA</button><button class="p" id="up">Carica file</button></div>
 <div class="doc" id="iaBox" hidden style="cursor:default"><button id="iaX" style="float:right" aria-label="Chiudi">×</button><b>Risposta dell'IA</b><div class="ma" id="iaTxt" role="status"></div><div id="iaDocs" style="margin-top:8px"></div></div>
 <input type="file" id="file" multiple hidden>
 <div id="lista"></div>
</section>
<article id="det"><div class="meta">Seleziona un documento per vedere i dettagli.</div></article>
<div id="vw" hidden>
 <header><b id="vwT"></b><div class="vtabs"><button id="tD">Documento</button><button id="tC">Domande</button></div><button id="vwD">Scarica</button><button id="vwX">Chiudi</button></header>
 <div class="vwb">
  <div id="vwDoc"></div>
  <div id="vwChat">
   <div id="vwM"></div>
   <div class="meta iaSt" style="padding:0 14px" role="status"></div>
   <div class="row" style="padding:10px 14px"><input type="text" id="vwQ" placeholder="Fai una domanda su questo documento"><button class="p" id="vwG">Chiedi</button></div>
   <div class="row" style="padding:0 14px"><button id="vwW">Chiedi + cerca online</button></div>
   <div class="meta" style="padding:6px 14px 0">Con «cerca online» il documento non viene inviato in rete: il modello prepara brevi ricerche generiche, che vedi scritte nella risposta.</div>
   <div class="row" style="padding:0 14px 12px"><button class="iaDl" hidden>Scarica motore e modello IA</button></div>
  </div>
 </div>
</div>
<div id="set" hidden>
 <header><b>Impostazioni</b><button id="setX">Chiudi</button></header>
 <div class="setb">
  <div class="card"><h2>Modello IA</h2>
   <div class="meta iaSt" role="status"></div>
   <div id="mods" style="margin-top:8px"></div>
  </div>
  <div class="card"><h2>Cartella osservata</h2>
 <input type="text" id="folder" placeholder="C:\Documenti\Da archiviare">
 <div class="row"><button id="salva">Salva</button><button id="scan">Importa ora</button></div>
  </div>
  <div class="card"><h2>Importa da Papra</h2>
 <input type="text" id="pUrl" placeholder="https://api.papra.app">
 <input type="password" id="pKey" placeholder="Chiave API di Papra" style="margin-top:6px">
 <div class="row"><button id="pGo">Importa tutto</button></div>
 <div class="meta" id="pSt" role="status"></div>
  </div>
  <div class="card"><h2>Riconoscimento testo (OCR)</h2>
   <div class="meta" id="ocrSt" style="white-space:pre-line" role="status"></div>
   <label class="row" style="align-items:center"><input type="checkbox" id="ocrAuto"> OCR automatico per scansioni e immagini</label>
   <div class="row"><button id="ocrComp" hidden>Installa componenti PDF</button><button id="ocrLing" hidden>Scarica lingua italiana</button></div>
   <div class="row"><button id="ocrTutti">OCR su tutti i documenti senza testo</button></div>
   <div class="meta" style="margin-top:8px">Tesseract va installato una volta sola: da terminale <code>winget install UB-Mannheim.TesseractOCR</code>, oppure dal sito github.com/UB-Mannheim/tesseract/wiki. Il resto lo scarica il programma.</div>
  </div>
 </div>
</div>
<div id="msg" role="status"></div>
<script>
const $=s=>document.querySelector(s);
const api=(u,o={})=>fetch(u,{...o,headers:{'X-DA':'1'}}).then(r=>r.json());
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const st={q:'',amb:'',tag:'',sel:null,docs:[],vw:null,storia:[],busy:false,ia:false};
const IMG=['.png','.jpg','.jpeg','.tif','.tiff','.bmp','.webp'];
const ocrBadge=d=>d.ocr==='coda'?'<span class="tag">OCR in coda</span>':d.ocr==='ok'?'<span class="tag">OCR</span>':d.ocr==='errore'?'<span class="tag">OCR non riuscito</span>':'';
const kb=n=>n>1048576?(n/1048576).toFixed(1)+' MB':Math.max(1,Math.round(n/1024))+' KB';
function msg(t){const m=$('#msg');m.textContent=t;m.style.display='block';clearTimeout(msg.t);msg.t=setTimeout(()=>m.style.display='none',3500)}
async function carica(){
 st.docs=await api('/api/docs?'+new URLSearchParams({q:st.q,ambito:st.amb,tag:st.tag}));
 const r=await api('/api/riepilogo');
 const A=[['','Tutti',r.totale],['lavoro','Lavoro',r.ambiti.lavoro||0],['casa','Casa',r.ambiti.casa||0]];
 $('#amb').innerHTML=A.map(a=>`<a tabindex=0 data-a="${a[0]}" class="${st.amb===a[0]?'on':''}">${a[1]}<small>${a[2]}</small></a>`).join('');
 const T=Object.entries(r.tags);
 $('#tags').innerHTML=T.length?T.map(t=>`<a tabindex=0 data-t="${esc(t[0])}" class="${st.tag===t[0]?'on':''}">${esc(t[0])}<small>${t[1]}</small></a>`).join(''):'<div class="meta">Nessun tag ancora</div>';
 lista();dettaglio();
}
function lista(){
 $('#lista').innerHTML=st.docs.length?st.docs.map(d=>`<div class="doc ${st.sel===d.id?'on':''}" data-id="${d.id}">
  <b>${esc(d.name)}</b>
  <div class="meta">${d.ambito?esc(d.ambito)+' · ':''}${esc(d.added)} · ${kb(d.size)} ${d.tags.map(t=>`<span class="tag">${esc(t)}</span>`).join('')}${ocrBadge(d)}</div>
  ${d.snip?`<div class="snip">${esc(d.snip).replace(/\x01/g,'<mark>').replace(/\x02/g,'</mark>')}</div>`:''}</div>`).join('')
 :`<div class="vuoto">${st.q||st.amb||st.tag?'Nessun documento trovato.':'L\'archivio è vuoto. Trascina qui i file o premi Carica file.'}</div>`;
}
function dettaglio(){
 const d=st.docs.find(x=>x.id===st.sel);
 if(!d){$('#det').innerHTML='<div class="meta">Seleziona un documento per vedere i dettagli.</div>';return}
 $('#det').innerHTML=`<button class="solomob" id="dX" style="float:right">Chiudi</button><h3 style="margin-top:0">Documento</h3><b style="word-break:break-word">${esc(d.name)}</b>
 <div class="meta">${esc(d.added)} · ${kb(d.size)}</div>
 <h3>Ambito</h3><select id="dA"><option value="">Nessuno</option><option value="lavoro">Lavoro</option><option value="casa">Casa</option></select>
 <h3>Tag (separati da virgola)</h3><input type="text" id="dT" value="${esc(d.tags.join(', '))}">
 <div class="row"><button class="p" id="dS">Salva modifiche</button><button id="dO">Apri</button></div>
 ${(d.ext==='.pdf'||IMG.includes(d.ext))?'<div class="row"><button id="dR">Riconosci testo (OCR)</button></div>':''}
 <div class="row"><button class="d" id="dD">Elimina dall'archivio</button></div>`;
 $('#dA').value=d.ambito;
 $('#dX').onclick=()=>document.body.classList.remove('det');
 if($('#dR'))$('#dR').onclick=async()=>{const r=await api('/api/ocr/'+d.id,{method:'POST'});msg(r.n?"Documento messo in coda per l'OCR":'OCR non applicabile a questo file');carica()};
 $('#dS').onclick=async()=>{await api('/api/meta/'+d.id,{method:'POST',body:JSON.stringify({tags:$('#dT').value.split(','),ambito:$('#dA').value})});msg('Modifiche salvate');carica()};
 $('#dO').onclick=()=>apri(d);
 $('#dD').onclick=async()=>{if(confirm('Eliminare "'+d.name+'" dall\'archivio?')){await api('/api/docs/'+d.id,{method:'DELETE'});st.sel=null;document.body.classList.remove('det');msg('Documento eliminato');carica()}};
}
async function invia(files){
 let n=0,dup=0;
 for(const f of files){const r=await api('/api/upload?'+new URLSearchParams({name:f.name,ambito:st.amb}),{method:'POST',body:f});r.duplicato?dup++:n++}
 msg(n+' caricati'+(dup?', '+dup+' già presenti':''));carica();
}
$('#amb').onclick=e=>{const a=e.target.closest('a');if(a){st.amb=a.dataset.a;document.body.classList.remove('menu');carica()}};
$('#tags').onclick=e=>{const a=e.target.closest('a');if(a){st.tag=st.tag===a.dataset.t?'':a.dataset.t;document.body.classList.remove('menu');carica()}};
$('#lista').ondblclick=e=>{const d=e.target.closest('.doc');if(d){const x=st.docs.find(y=>y.id===+d.dataset.id);if(x)apri(x)}};
$('#lista').onclick=e=>{const d=e.target.closest('.doc');if(d){st.sel=+d.dataset.id;document.body.classList.add('det');lista();dettaglio()}};
let tm;$('#q').oninput=e=>{clearTimeout(tm);tm=setTimeout(()=>{st.q=e.target.value;carica()},250)};
$('#up').onclick=()=>$('#file').click();
$('#file').onchange=e=>{invia([...e.target.files]);e.target.value=''};
$('#salva').onclick=async()=>{await api('/api/settings',{method:'POST',body:JSON.stringify({folder:$('#folder').value})});msg('Cartella salvata')};
$('#scan').onclick=async()=>{const r=await api('/api/scan',{method:'POST'});msg(r.n+' nuovi documenti importati');carica()};
addEventListener('dragover',e=>{e.preventDefault();document.body.classList.add('drag')});
addEventListener('dragleave',e=>{if(!e.relatedTarget)document.body.classList.remove('drag')});
addEventListener('drop',e=>{e.preventDefault();document.body.classList.remove('drag');if(e.dataTransfer.files.length)invia([...e.dataTransfer.files])});
async function apri(d){
 document.body.classList.remove('det','menu');$('#vw').classList.remove('tabchat');st.vw=d;st.storia=[];$('#vwM').innerHTML='';$('#vw').hidden=false;$('#vwT').textContent=d.name;
 const box=$('#vwDoc');
 if(d.ext==='.pdf')box.innerHTML=`<iframe src="/file/${d.id}" title="Anteprima"></iframe>`;
 else if(['.png','.jpg','.jpeg','.gif','.webp','.bmp'].includes(d.ext))box.innerHTML=`<img src="/file/${d.id}" alt="">`;
 else{const t=await fetch('/api/testo/'+d.id,{headers:{'X-DA':'1'}}).then(r=>r.text());
  if(t.trim()){box.innerHTML='<pre></pre>';box.firstChild.textContent=t}
  else box.innerHTML='<div class="vuoto" style="margin:20px">Anteprima non disponibile per questo tipo di file. Usa Scarica.</div>'}
 iaStato();
}
async function iaStato(){
 const s=await api('/api/ia/stato');
 const t=s.scarica.run?'Scarico '+(s.scarica.fase||'')+'… '+s.scarica.pct+'%':(!s.motore||!s.modello)?(s.scarica.msg||'Mancano il motore e il modello IA (si scaricano una volta sola).'):(s.attivo?'IA pronta, in memoria: '+s.modello:'IA pronta: si avvia alla prima domanda ('+s.modello+')');
 document.querySelectorAll('.iaSt').forEach(e=>e.textContent=t);
 document.querySelectorAll('.iaDl').forEach(b=>b.hidden=s.scarica.run||(s.motore&&s.modello));
 renderModelli(s);
 st.ia=!!(s.motore&&s.modello&&!s.scarica.run);
 if(s.scarica.run||(st.ia&&!s.attivo&&(iaStato.n||0)<30))setTimeout(()=>{iaStato.n=(iaStato.n||0)+1;iaStato()},2000);
}
function renderModelli(s){
 const el=$('#mods');if(!el)return;
 const gb=mb=>mb>=1000?(mb/1000).toFixed(1)+' GB':mb+' MB';
 el.innerHTML=s.modelli.map(m=>{
  const uso=m.id===s.scelto&&m.presente,scar=s.scarica.run&&s.scarica.id===m.id;
  let az='';
  if(scar)az='<span class="meta">Scarico '+esc(s.scarica.fase||'')+'… '+s.scarica.pct+'%</span>';
  else if(m.presente&&!uso)az=`<button data-u="${esc(m.id)}">Usa questo</button><button class="d" data-e="${esc(m.id)}">Elimina</button>`;
  else if(!m.presente&&m.url)az=`<button class="p" data-d="${esc(m.id)}"${s.scarica.run?' disabled':''}>Scarica (circa ${gb(m.mb)})</button>`;
  return `<div class="doc ${uso?'on':''}" style="cursor:default"><b>${esc(m.nome)}</b> ${uso?'<span class="tag">In uso</span>':m.presente?'<span class="tag">Scaricato</span>':''}<div class="meta">${esc(m.desc)}</div><div class="row">${az}</div></div>`;
 }).join('')+(s.scarica.msg?`<div class="meta">${esc(s.scarica.msg)}</div>`:'');
}
async function ocrStato(){
 const s=await api('/api/ocr/stato'),c=s.componenti,ok=c.pdfium&&c.pillow;
 let t=(s.tesseract?'✓ Tesseract installato':'✗ Tesseract non trovato')+'\n'+(s.lingua?'✓ lingua italiana presente':'✗ lingua italiana da scaricare')+'\n'+(ok?'✓ componenti per i PDF installati':'✗ componenti per i PDF da installare (pypdfium2, pillow)');
 if(s.corrente.doc)t+='\nOCR in corso: '+s.corrente.nome+(s.corrente.tot?' (pagina '+s.corrente.pagina+' di '+s.corrente.tot+')':'');
 if(s.coda)t+='\nDocumenti in coda: '+s.coda;
 for(const x of [s.attesa,s.pip,s.errore])if(x)t+='\n'+x;
 $('#ocrSt').textContent=t;$('#ocrAuto').checked=s.auto;$('#ocrLing').hidden=s.lingua;$('#ocrComp').hidden=ok;
 clearTimeout(ocrStato.t);
 if(!$('#set').hidden&&(s.corrente.doc||s.coda||s.pip.startsWith('Installo')))ocrStato.t=setTimeout(ocrStato,2000);
}
function apriImp(){document.body.classList.remove('menu');$('#set').hidden=false;iaStato();ocrStato()}
$('#setB').onclick=apriImp;$('#setB2').onclick=apriImp;
$('#setX').onclick=()=>{$('#set').hidden=true};
$('#menuB').onclick=()=>document.body.classList.toggle('menu');
$('#scrim').onclick=()=>document.body.classList.remove('menu','det');
$('#tD').onclick=()=>$('#vw').classList.remove('tabchat');
$('#tC').onclick=()=>$('#vw').classList.add('tabchat');
$('#mods').onclick=async e=>{
 const b=e.target.closest('button');if(!b)return;let r;
 if(b.dataset.u)r=await api('/api/ia/modello',{method:'POST',body:JSON.stringify({id:b.dataset.u})});
 else if(b.dataset.e){if(!confirm('Eliminare questo modello dal disco?'))return;r=await api('/api/ia/elimina',{method:'POST',body:JSON.stringify({id:b.dataset.e})})}
 else if(b.dataset.d)r=await api('/api/ia/scarica',{method:'POST',body:JSON.stringify({id:b.dataset.d})});
 if(r&&r.msg)msg(r.msg);iaStato();
};
$('#ocrAuto').onchange=e=>api('/api/ocr/impostazioni',{method:'POST',body:JSON.stringify({auto:e.target.checked})});
$('#ocrLing').onclick=async()=>{await api('/api/ocr/lingua',{method:'POST'});msg('Scarico la lingua italiana…');setTimeout(ocrStato,2500)};
$('#ocrComp').onclick=async()=>{await api('/api/ocr/componenti',{method:'POST'});setTimeout(ocrStato,500)};
$('#ocrTutti').onclick=async()=>{const r=await api('/api/ocr/tutti',{method:'POST'});msg(r.n+' documenti messi in coda');ocrStato()};
const DEFWEB="Dai un'opinione sui prezzi e sulle voci di questo documento, confrontandoli con ciò che si trova online.";
async function leggi(url,corpo,out,fine){
 const r=await fetch(url,{method:'POST',headers:{'X-DA':'1'},body:JSON.stringify(corpo)});
 const rd=r.body.getReader(),dec=new TextDecoder();let t='';
 for(;;){const x=await rd.read();if(x.done)break;t+=dec.decode(x.value,{stream:true});out(t)}
 return t;
}
async function chiedi(web){
 const q=$('#vwQ').value.trim()||(web?DEFWEB:'');if(!q||st.busy)return;
 if(!st.ia){msg('Il modello IA non è ancora pronto');return}
 $('#vwQ').value='';st.busy=true;
 const M=$('#vwM'),qd=document.createElement('div'),a=document.createElement('div');
 qd.className='mq';qd.textContent=q+(web?'  (+ online)':'');a.className='ma';a.textContent='Sto leggendo il documento… (la prima volta il modello deve avviarsi: può servire un minuto)';
 M.append(qd,a);M.scrollTop=M.scrollHeight;
 try{
  const t=await leggi('/api/chiedi',{id:st.vw.id,q:web&&q===DEFWEB?'':q,storia:st.storia.slice(-2),web:!!web},t=>{a.textContent=t;M.scrollTop=M.scrollHeight});
  st.storia.push({q,a:t});
 }catch(e){a.textContent='Errore: '+e}
 st.busy=false;
}
function mostraDocs(docs){
 st.iaDocs=docs;
 $('#iaDocs').innerHTML=docs.map(d=>`<div class="doc" data-ia="${d.id}"><b>${esc(d.name)}</b></div>`).join('');
}
async function cercaIA(){
 const q=$('#q').value.trim();if(!q||st.busy)return;
 if(!st.ia){msg('Il modello IA non è ancora pronto: usa «Scarica motore e modello IA» nella barra a sinistra');return}
 st.busy=true;$('#iaBox').hidden=false;$('#iaDocs').innerHTML='';
 $('#iaTxt').textContent="Cerco nell'archivio con l'aiuto del modello… (la prima volta il modello deve avviarsi)";
 try{
  let docs=null;
  await leggi('/api/cerca-ia',{q},t=>{
   if(!docs&&t.includes('\x02')){const i=t.indexOf('\x02');docs=JSON.parse(t.slice(1,i));mostraDocs(docs)}
   if(docs)t=t.slice(t.indexOf('\x02')+1);
   if(docs||!t.startsWith('\x01'))$('#iaTxt').textContent=t;
  });
 }catch(e){$('#iaTxt').textContent='Errore: '+e}
 st.busy=false;
}
$('#qIA').onclick=cercaIA;
$('#iaX').onclick=()=>{$('#iaBox').hidden=true};
$('#iaDocs').onclick=e=>{const el=e.target.closest('[data-ia]');if(el){const d=(st.iaDocs||[]).find(x=>x.id===+el.dataset.ia);if(d)apri(d)}};
$('#vwX').onclick=()=>{$('#vw').hidden=true;$('#vwDoc').innerHTML=''};
$('#vwD').onclick=()=>{location.href='/file/'+st.vw.id+'?scarica=1'};
$('#vwG').onclick=()=>chiedi(false);$('#vwW').onclick=()=>chiedi(true);
$('#vwQ').onkeydown=e=>{if(e.key==='Enter')chiedi(false)};
document.querySelectorAll('.iaDl').forEach(b=>b.onclick=async()=>{await api('/api/ia/scarica',{method:'POST'});iaStato()});
addEventListener('keydown',e=>{if(e.key!=='Escape')return;if(!$('#set').hidden)$('#setX').onclick();else if(!$('#vw').hidden)$('#vwX').onclick()});
let pt;
async function pStato(){const s=await api('/api/papra/status');
 $('#pSt').textContent=s.run?`${s.msg} (${s.n}/${s.tot})`:(s.msg?`${s.msg} Nuovi: ${s.nuovi}, già presenti: ${s.dup}, errori: ${s.err}`:'');
 if(s.run)pt=setTimeout(pStato,1500);else if(s.msg)carica()}
$('#pGo').onclick=async()=>{const r=await api('/api/papra/start',{method:'POST',body:JSON.stringify({url:$('#pUrl').value,key:$('#pKey').value})});
 if(!r.ok){msg(r.msg||'Importazione non avviata');return}$('#pKey').value='';pStato()};
api('/api/settings').then(s=>{$('#folder').value=s.folder;$('#pUrl').value=s.papra_url||''});
carica();pStato();iaStato();setInterval(()=>{if(document.activeElement.id!=='q'&&document.activeElement.id!=='dT')carica()},15000);
</script></body></html>
"""


def main():
    init()
    threading.Thread(target=sorvegliante, daemon=True).start()
    threading.Thread(target=ocr_lavoratore, daemon=True).start()
    if trova_motore() and trova_modello():
        threading.Thread(target=avvia_motore, daemon=True).start()
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), H)
    url = f"http://127.0.0.1:{PORT}"
    print("DocuArchivio attivo su", url, "- chiudi questa finestra per fermarlo")
    webbrowser.open(url)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
