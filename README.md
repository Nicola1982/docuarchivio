# 🗂️ DocuArchivio

**DocuArchivio** è un archivio di documenti locale, leggero e ottimizzato per PC datati (es. processori Celeron o Atom senza istruzioni AVX2). Gira interamente sul tuo computer in locale, garantendo la massima privacy.

## ✨ Caratteristiche Principali
* **Ricerca testuale avanzata:** Indicizzazione e ricerca rapida nel testo dei documenti (grazie a SQLite FTS5).
* **Gestione Tag e Ambiti:** Organizzazione dei file suddividendoli per ambiti (es. *Lavoro* / *Casa*) e tag personalizzati.
* **Cartella Osservata:** Una cartella monitorata in background che importa automaticamente i nuovi file aggiunti.
* **Intelligenza Artificiale Locale:** Chat integrata con i documenti basata su `llama.cpp` (completamente offline, senza inviare dati in cloud). Puoi scegliere il modello in base alle risorse del tuo PC (da 0.5B a 7B).
* **Riconoscimento OCR:** Integrazione con Tesseract per estrarre testo da immagini e PDF scansionati.
* **Importazione da Papra:** Sincronizzazione opzionale tramite chiave API con la piattaforma Papra.

---

## 🚀 Requisiti e Installazione

### 1. Requisiti di sistema
* **Python 3.10 o superiore** installato sul sistema.
* **Windows 64-bit** (consigliato per il download automatico del server LLM).

### 2. Installazione delle librerie Python
Apri il terminale (o il prompt dei comandi) nella cartella del programma e installa le dipendenze necessarie con un solo comando:
```bash
pip install pypdf python-docx

🖥️ Avvio
Per avviare l'applicazione, posizionati nella cartella del progetto ed esegui:

Bash
python docuarchivio.py
Il programma aprirà automaticamente il browser predefinito all'indirizzo http://127.0.0.1:8765.
Per chiudere l'applicazione e fermare il server, basta chiudere la finestra del terminale (la "finestra nera").
