# Rilievi in campo – prototipo 0.1

## Contenuto della cartella

| File | A cosa serve |
|---|---|
| `index.html`, `app.js` | l'app |
| `config.js` | configurazione: va inserito il `clientId` di Microsoft (passo 3) |
| `sw.js`, `manifest.webmanifest`, `icons/` | installazione sulla Home e funzionamento offline |
| `specie.json` | elenco delle 240 specie del Protocollo ARETÉ (autocompletamento) |
| `vendor/` | librerie: MSAL (login Microsoft) e piexifjs (metadati EXIF delle foto) |

## 1. Metterla online (GitHub Pages, gratis)

L'app deve stare su un indirizzo **https**: è obbligatorio per usare GPS, bussola e fotocamera.

1. Crea un account su github.com (se non lo hai).
2. Crea un nuovo repository pubblico, per esempio `rilievi-campo`.
3. *Add file → Upload files*: trascina **tutto il contenuto** di questa cartella (non la cartella stessa) e conferma con *Commit*.
4. *Settings → Pages*: in *Branch* scegli `main` e cartella `/ (root)`, poi *Save*.
5. Dopo un minuto l'app è su `https://<tuo-utente>.github.io/rilievi-campo/`.

## 2. Installarla sul telefono

- **iPhone (Safari)**: apri l'indirizzo, poi *Condividi → Aggiungi alla schermata Home*. Usala sempre dall'icona: così Safari non cancella i rilievi salvati sul telefono.
- **Android (Chrome)**: apri l'indirizzo, poi menu ⋮ → *Installa app*.

Al primo uso tocca **Attiva bussola** e accetta i permessi per posizione, movimento e fotocamera.

## 3. Collegare OneDrive (registrazione dell'app su Microsoft)

Serve una sola volta e vale per tutti i colleghi.

Da agosto 2024 Microsoft non permette più di registrare app con un semplice account personale: serve una *directory* (tenant) Microsoft Entra. Il modo più semplice è creare un account Azure gratuito (chiede una carta solo per la verifica; la registrazione di un'app è gratuita).

1. Vai su <https://entra.microsoft.com> (o portal.azure.com) ed entra con il tuo account; se te lo chiede, crea la directory / l'account Azure gratuito.
2. *Identità → Applicazioni → Registrazioni app → Nuova registrazione*.
3. Nome: `Rilievi in campo`.
4. Tipi di account supportati: **Account in qualsiasi directory organizzativa e account Microsoft personali**.
5. URI di reindirizzamento: piattaforma **Applicazione a pagina singola (SPA)**, indirizzo esatto dell'app, per esempio `https://<tuo-utente>.github.io/rilievi-campo/index.html`.
6. *Registra*. Copia l'**ID applicazione (client)**.
7. Apri `config.js`, incolla l'ID tra le virgolette di `clientId` e ricarica il file su GitHub.

Poi nell'app: *Impostazioni → Collega OneDrive*. Ognuno entra con il proprio account OneDrive personale e accetta i permessi "Leggi il profilo" e "Accedi ai tuoi file".

**Finché OneDrive non è collegato** l'app funziona lo stesso: i rilievi restano sul telefono e dal dettaglio di ogni rilievo il pulsante *Condividi file* li manda all'app OneDrive (o a Mail, WhatsApp, File…).

## Cosa salva l'app su OneDrive

```
Rilievi/                         ← nome cartella modificabile in Impostazioni
  <Lavoro>/
    registro.csv                 ← una riga per albero, si apre con Excel
    gs4521_VA-20261118-1032/
      dati.json                  ← tutti i dati della scheda
      foto_01.jpg                ← foto pulita (con GPS e azimut nell'EXIF)
      foto_01_N.jpg              ← stessa foto con la freccia N (se attiva)
      allegato_….                ← file dello strumento (valutazione strumentale)
```

## Cosa provare nella Fase 0

1. **Bussola**: confrontala con una bussola vera o con l'app Bussola del telefono, sia con il telefono in piano sia in verticale (fotocamera puntata). Prova su iPhone e su almeno un Android.
2. **GPS sotto chioma**: guarda la precisione (±m) e usa *Fissa posizione* quando è buona.
3. **Fotocamera**: la fotocamera interna registra la direzione nell'istante dello scatto. *Foto di sistema* dà la qualità piena del telefono, ma la direzione viene letta al tocco del pulsante.
4. **Offline**: modalità aereo → salva 2 rilievi → riattiva la rete → controlla che arrivino su OneDrive.
5. Annota tutto quello che è scomodo in campo: il prossimo giro di modifiche parte da lì.

## Limiti noti del prototipo

- La declinazione magnetica è un valore fisso (+3,5°, modificabile in Impostazioni), non ancora calcolata dalla posizione.
- Con la fotocamera interna la risoluzione dipende dal browser (spesso 1920×1080 o 4032×3024).
- Su iPhone la sincronizzazione parte solo con l'app aperta (limite di Safari).
- Il registro è in CSV; la scheda fac-simile ARETÉ compilata in automatico è prevista nella Fase 3.
