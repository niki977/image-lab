# The Ultraspeaker Image Lab

Add-in per PowerPoint che regola le foto delle slide, come «Regola immagine» di Keynote:
istogramma con livelli (nero 0%, mezzitoni 50%, bianco 100%), esposizione, contrasto, saturazione,
luci, ombre, nitidezza, riduzione rumore, temperatura e tinta, stili rapidi e punto neutro.
Interfaccia in italiano, inglese, spagnolo, francese e tedesco.

## Come funziona nella slide

1. Selezioni una foto: il pannello la legge e mostra anteprima e istogramma.
2. Muovi i cursori: l’anteprima nel pannello si aggiorna subito.
3. **Applica alla slide**: la foto regolata viene inserita nella stessa posizione, con la stessa
   dimensione e rotazione. L’originale resta nella slide, **nascosto** (lo vedi nel Riquadro di selezione).
4. Se riselezioni la foto modificata, il pannello riparte dall’originale con le regolazioni salvate:
   puoi cambiarle quante volte vuoi senza perdere qualità, oppure premere **Originale** per tornare indietro.

## Requisiti

PowerPoint di Microsoft 365 con le API PowerPoint 1.10:
Windows versione 2601 o successiva, Mac 16.105 o successiva, oppure PowerPoint sul web.
Su versioni precedenti il pannello mostra un messaggio che chiede di aggiornare.

## Pubblicazione su GitHub Pages

1. Crea su GitHub un repository pubblico chiamato **image-lab** (account `niki977`).
2. Carica tutti i file di questa cartella (anche le cartelle `assets` e `fonts`).
3. In *Settings → Pages* scegli *Deploy from a branch*, ramo `main`, cartella `/ (root)`.
4. Dopo un minuto il pannello è raggiungibile su `https://niki977.github.io/image-lab/`
   (aprendolo nel browser parte in modalità prova con una foto di esempio).

Se usi un altro nome di repository, sostituisci `https://niki977.github.io/image-lab/` in `image-lab-manifest.xml`.

## Installazione per le prove (sideload)

- **PowerPoint sul web**: Home → Componenti aggiuntivi → Altri componenti aggiuntivi → I miei componenti
  aggiuntivi → Carica il mio componente aggiuntivo → scegli `image-lab-manifest.xml`.
- **Mac**: copia `image-lab-manifest.xml` nella cartella
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef` (creala se non esiste), poi riapri PowerPoint:
  lo trovi in Home → Componenti aggiuntivi → I miei componenti aggiuntivi.
- **Windows**: Home → Componenti aggiuntivi → Altri componenti aggiuntivi → I miei componenti aggiuntivi →
  Carica il mio componente aggiuntivo (nelle versioni recenti); in alternativa usa una cartella condivisa
  come «Catalogo di componenti aggiuntivi attendibili» nelle impostazioni del Centro protezione.

Il pulsante **Image Lab** compare nella scheda Home, nel gruppo «The Ultraspeaker».

## File

| File | Cosa contiene |
|---|---|
| `index.html` | Il pannello (struttura e stile) |
| `app.js` | Collegamento con PowerPoint, anteprima, istogramma, applicazione |
| `imaging.js` | Elaborazione dei pixel (livelli, tono, colore, nitidezza, rumore) |
| `i18n.js` | Testi nelle 5 lingue |
| `image-lab-manifest.xml` | Manifest dell’add-in per PowerPoint |
| `assets/` | Loghi SVG (Image Lab, The Ultraspeaker, Arena, cursori) e icone PNG (16–300 px) |
| `fonts/` | Quicksand in woff2 (400, 500, 600, 700) e licenza OFL |
| `GUIDA-INSTALLAZIONE.html` | Guida passo passo |
| `support.html`, `privacy.html`, `terms.html`, `legal.css` | Pagine di supporto, privacy e condizioni d’uso (IT/EN) richieste dallo store |
| `store/` | Testi della scheda nelle 5 lingue, screenshot 1366×768, presentazione di prova, istruzioni per Partner Center |

## Note per il Marketplace Microsoft

Per la pubblicazione su AppSource servono anche: pagina di privacy e condizioni d’uso (URL),
icona 300×300 (`assets/image-lab-icon-300.png`), screenshot del pannello e una descrizione estesa.
L’add-in non invia le immagini a nessun server: tutta l’elaborazione avviene nel pannello.
