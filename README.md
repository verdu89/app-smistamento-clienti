# 📦 App Smistamento Clienti

Questa applicazione è uno script Google Apps Script + strumenti di supporto che permette di **gestire, smistare e monitorare i clienti** in maniera automatizzata tramite Google Sheets, Gmail e notifiche email.  
È pensata per team commerciali che devono gestire rapidamente lead, follow-up e attività di vendita.

---

## 🚀 Funzionalità principali

- 📋 Raccolta automatica dati da **Google Form** o input manuale su Google Sheets.
- 🔄 Smistamento clienti ai venditori con regole di priorità.
- 🧹 Deduplica dei record per evitare duplicati nel database.
- ✉️ Notifiche email ai venditori e reminder automatici.
- 📊 Dashboard di monitoraggio delle attività di vendita.
- 🤖 Integrazione con OpenAI per analisi dei dati e supporto smart.

---

## 🛠️ Requisiti

- **Google Workspace** con accesso a Google Sheets, Gmail e Apps Script.
- **Account OpenAI** per usare le funzionalità AI (opzionale).
- Node.js (se si lavora su componenti locali).

---

## 📂 Struttura del progetto

```
app-smistamento-clienti/
├── README.md
├── .gitignore
├── Code.gs              # Script principale per Apps Script
├── apenai-api-key.txt   # ⚠️ File locale con chiave API (ignorato da Git)
└── package.json         # Configurazione Node.js (se usata)
```

---

## 🔑 Configurazione chiavi API

⚠️ **Importante:** Non salvare mai le chiavi API direttamente nel codice.  
Puoi gestirle in sicurezza in due modi:

### 🔐 Su Google Apps Script

- Vai su **Editor Apps Script → Project Settings → Script properties**
- Aggiungi una proprietà chiamata `OPENAI_API_KEY` con la tua chiave.
- Nel codice recuperala con:
  ```javascript
  const apiKey =
    PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  ```

### 🔐 Su sviluppo locale (Node.js)

- Crea un file `.env`:
  ```env
  OPENAI_API_KEY=sk-xxxxxxxx
  ```
- Caricalo nel codice con:
  ```js
  require("dotenv").config();
  const apiKey = process.env.OPENAI_API_KEY;
  ```

---

## ▶️ Come avviare

1. Clona il repository:

   ```bash
   git clone https://github.com/verdu89/app-smistamento-clienti.git
   cd app-smistamento-clienti
   ```

2. Configura il progetto Apps Script:

   - Apri `Code.gs` nell’editor di Google Apps Script.
   - Collega il foglio Google desiderato.
   - Imposta le proprietà dello script (chiavi API, email admin, ecc.).

3. Testa l’applicazione su Google Sheets:
   - Aggiungi un nuovo cliente.
   - Verifica che venga smistato correttamente.
   - Controlla la ricezione delle email automatiche.

---

## 🛡️ Sicurezza

- `.gitignore` è configurato per evitare di caricare file con chiavi (`apenai-api-key.txt`, `.env`, ecc.).
- Se accidentalmente una chiave finisce nella history, **revocarla subito** e usare [git-filter-repo](https://github.com/newren/git-filter-repo).
- Le chiavi devono sempre essere gestite tramite variabili d’ambiente o `PropertiesService`.

---
