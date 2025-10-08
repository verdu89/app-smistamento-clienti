# 📦 Smistamento Clienti

_(Customer Routing & Automation System)_

Questa applicazione è realizzata interamente in **Google Apps Script** e consente di **gestire automaticamente lead, assegnarli ai venditori, inviare notifiche e richiedere recensioni** tramite Google Sheets + Gmail + WhatsApp.

---

## 🚀 Funzionalità principali / Main Features

| Funzionalità IT                         | Feature EN                            |
| --------------------------------------- | ------------------------------------- |
| 📋 Raccolta lead da Google Sheet / Form | Collects leads from Sheets / Forms    |
| 🔄 Smistamento automatico ai venditori  | Auto-routing to assigned vendors      |
| 🧹 Rimozione duplicati                  | Deduplication to avoid double entries |
| ✉️ Invio email e WhatsApp automatici    | Automatic Email & WhatsApp messages   |
| 📊 Dashboard di monitoraggio            | Sales / Lead tracking dashboard       |
| ⭐ Richiesta recensioni a fine lavoro   | End-of-job review requests            |

---

## 🛠️ Requisiti / Requirements

- ✅ **Google Workspace** con accesso a Sheets + Gmail + Apps Script
- ✅ (Opzionale) **Chiave OpenAI** salvata in _Script Properties_

---

## 📂 Struttura del Progetto / Project Structure

SmistamentoClienti (Apps Script Project)
├── 00_config.gs # Getter chiavi / API (es. getOpenAIKey)
├── 01_utils.gs # Funzioni generiche riutilizzabili
├── 02_logging.gs # Sistema log centralizzato
├── 10_server.gs # Endpoint Webhook (doGet / doPost)
├── 20_whatsapp.gs # Invio messaggi WhatsApp
├── 30_email_queue.gs # Coda email e invio asincrono
├── 31_gmail_reconcile.gs# Allineamento Gmail ↔ Main
├── 40_meta_leads.gs # Import da Meta/Facebook
├── 50_vendors.gs # Logica assegnazione Vendor
├── 60_core.gs # Logica principale di smistamento
├── 70_dashboard.gs # Aggiornamento Dashboard / Report
├── 98_triggers_setup.gs # Creazione trigger automatici
└── 99_trigger_handlers.gs # Funzioni chiamate dai trigger

yaml
Copia codice

---

## 🔑 Configurazione chiavi API (OpenAI)

1. Vai su **Editor Apps Script → Project Settings → Script Properties**
2. Aggiungi:

| Key            | Value      |
| -------------- | ---------- |
| OPENAI_API_KEY | sk-xxxxxxx |

3. Recuperabile nel codice tramite:

```js
function getOpenAIKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
}
🔄 Flusso Principale Lead / Main Lead Flow
🇮🇹

Il lead entra nel Foglio Main

avviaProgramma() lo smista al venditore corretto

Parte la comunicazione automatica (Email / WhatsApp)

Il contatto viene tracciato in Dashboard

🇬🇧

Lead is inserted into the Main sheet

avviaProgramma() assigns it to the correct Vendor

Automatic messaging starts (Email / WhatsApp)

Lead status is tracked in Dashboard

⭐ Flusso Recensioni / Review Request Flow
Evento IT	Azione IT
Compilazione di una riga nel foglio "Recensioni Extra"	Invio automatico richiesta recensione

Event EN	Action EN
A row is filled in "Recensioni Extra" sheet	Automatic review request is sent

⚙️ Trigger Automatici / Active Triggers
Funzione	Scopo IT	Purpose EN
onEditInstalled	Reagisce alle modifiche nel foglio	Reacts to sheet edits
processEmailQueue	Invia le email in attesa	Sends queued emails
setupDailyReminderTrigger	Reminder giornalieri	Daily reminders
setupDashboardFridayTrigger	Report settimanale	Weekly report
createOnEditTrigger()	Crea trigger sulle modifiche	Creates onEdit trigger

▶️ Test Manuali / Manual Tests
Funzione	Cosa fa IT	What it does EN
avviaProgramma()	Avvia lo smistamento lead	Starts full routing
processEmailQueue()	Forza invio email	Forces email sending
updateDashboardFromMain()	Aggiorna dashboard	Refresh dashboard

🛡️ Sicurezza / Security
✔️ Nessuna chiave API è salvata nel codice
✔️ Le chiavi sono gestite tramite Script Properties
✔️ Nessun file viene esposto pubblicamente
```
